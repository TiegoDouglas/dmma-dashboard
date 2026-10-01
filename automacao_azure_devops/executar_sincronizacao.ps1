[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$repo = "TiegoDouglas/dmma-dashboard"
$branch = "main"
$workflow = "sync-azure-work-item.yml"
$repositoryPath = "automacao_azure_devops/vagas.xlsx"
$localWorkbook = Join-Path $PSScriptRoot "vagas.xlsx"

function Invoke-Gh {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,
        [switch]$AllowFailure
    )

    $output = & gh @Arguments 2>&1
    $exitCode = $LASTEXITCODE
    if (-not $AllowFailure -and $exitCode -ne 0) {
        $message = ($output | Out-String).Trim()
        throw "GitHub CLI falhou (codigo $exitCode): $message"
    }
    return [pscustomobject]@{
        ExitCode = $exitCode
        Output = ($output | Out-String).Trim()
    }
}

function Get-RemoteWorkbook {
    $result = Invoke-Gh -Arguments @(
        "api",
        "--method", "GET",
        "repos/$repo/contents/$repositoryPath",
        "-f", "ref=$branch"
    )
    try {
        return $result.Output | ConvertFrom-Json
    }
    catch {
        throw "Resposta invalida ao consultar a planilha na branch ${branch}: $($_.Exception.Message)"
    }
}

function Get-Sha256Hex {
    param([Parameter(Mandatory = $true)][byte[]]$Bytes)

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($sha256.ComputeHash($Bytes))).Replace("-", "")
    }
    finally {
        $sha256.Dispose()
    }
}

function Publish-Workbook {
    if (-not (Test-Path -LiteralPath $localWorkbook -PathType Leaf)) {
        throw "Planilha nao encontrada ao lado do BAT: $localWorkbook"
    }

    Write-Host "Consultando a planilha atual na branch $branch..."
    $remote = Get-RemoteWorkbook
    if (-not $remote.sha -or -not $remote.content) {
        throw "A API do GitHub nao retornou SHA e conteudo da planilha remota."
    }

    $localBytes = [System.IO.File]::ReadAllBytes($localWorkbook)
    try {
        $remoteBytes = [Convert]::FromBase64String(($remote.content -replace "\s", ""))
    }
    catch {
        throw "O conteudo remoto da planilha nao esta em Base64 valido."
    }

    if ((Get-Sha256Hex $localBytes) -eq (Get-Sha256Hex $remoteBytes)) {
        Write-Host "A planilha local ja e identica a versao da branch main; nenhum commit foi criado."
        return [string]$remote.sha
    }

    Write-Host "Publicando a planilha local na branch $branch..."
    $payload = @{
        message = "Atualiza planilha de vagas pela sincronizacao Windows"
        content = [Convert]::ToBase64String($localBytes)
        branch = $branch
        sha = [string]$remote.sha
    } | ConvertTo-Json -Compress

    $tempFile = [System.IO.Path]::GetTempFileName()
    try {
        $utf8WithoutBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($tempFile, $payload, $utf8WithoutBom)
        $result = Invoke-Gh -Arguments @(
            "api",
            "--method", "PUT",
            "repos/$repo/contents/$repositoryPath",
            "--input", $tempFile
        )
    }
    finally {
        Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
    }

    try {
        $published = $result.Output | ConvertFrom-Json
    }
    catch {
        throw "Resposta invalida ao publicar a planilha: $($_.Exception.Message)"
    }
    $publishedSha = [string]$published.content.sha
    if (-not $publishedSha) {
        throw "A API do GitHub nao confirmou o SHA da planilha publicada."
    }
    if ($published.commit.html_url) {
        Write-Host "Commit criado: $($published.commit.html_url)"
    }

    Write-Host "Aguardando a planilha publicada ficar disponivel na branch main..."
    for ($attempt = 1; $attempt -le 30; $attempt++) {
        $current = Get-RemoteWorkbook
        if ([string]$current.sha -eq $publishedSha) {
            Write-Host "Publicacao confirmada."
            return $publishedSha
        }
        Start-Sleep -Seconds 2
    }
    throw "A planilha publicada nao ficou disponivel na branch main apos 60 segundos."
}

function Start-And-WatchWorkflow {
    $invocationId = [Guid]::NewGuid().ToString("N")
    $expectedTitle = "sync-azure-$invocationId"
    $dispatchTime = [DateTime]::UtcNow

    Write-Host "Disparando o workflow na branch $branch..."
    Invoke-Gh -Arguments @(
        "workflow", "run", $workflow,
        "--repo", $repo,
        "--ref", $branch,
        "-f", "invocation_id=$invocationId"
    ) | Out-Null

    Write-Host "Localizando exclusivamente a execucao criada por este processo..."
    $run = $null
    for ($attempt = 1; $attempt -le 30; $attempt++) {
        $result = Invoke-Gh -Arguments @(
            "run", "list",
            "--repo", $repo,
            "--workflow", $workflow,
            "--branch", $branch,
            "--event", "workflow_dispatch",
            "--limit", "50",
            "--json", "databaseId,createdAt,displayTitle,url"
        )
        try {
            $runs = @($result.Output | ConvertFrom-Json)
        }
        catch {
            throw "Resposta invalida ao localizar a execucao do workflow."
        }
        $run = $runs |
            Where-Object {
                $_.displayTitle -eq $expectedTitle -and
                ([DateTime]$_.createdAt).ToUniversalTime() -ge $dispatchTime.AddSeconds(-5)
            } |
            Sort-Object { [DateTime]$_.createdAt } -Descending |
            Select-Object -First 1
        if ($run) {
            break
        }
        Start-Sleep -Seconds 2
    }
    if (-not $run) {
        throw "A execucao identificada por $invocationId nao foi localizada apos 60 segundos."
    }

    Write-Host "Execucao encontrada: $($run.databaseId)"
    Write-Host "URL: $($run.url)"
    if ($run.url -and -not $env:SYNC_AZURE_NO_BROWSER) {
        Start-Process ([string]$run.url)
    }

    Write-Host "Aguardando a conclusao..."
    $watch = Invoke-Gh -Arguments @(
        "run", "watch", [string]$run.databaseId,
        "--repo", $repo,
        "--exit-status"
    ) -AllowFailure
    if ($watch.Output) {
        Write-Host $watch.Output
    }
    return $watch.ExitCode
}

$exitCode = 1
try {
    Write-Host ""
    Write-Host "Sincronizacao de vaga com Azure DevOps"
    Write-Host "Pasta local: $PSScriptRoot"
    Write-Host ""

    if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
        throw "GitHub CLI (gh) nao foi encontrado. Instale em https://cli.github.com/."
    }

    Write-Host "Verificando autenticacao no GitHub..."
    $auth = Invoke-Gh -Arguments @("auth", "status") -AllowFailure
    if ($auth.ExitCode -ne 0) {
        throw "GitHub CLI nao esta autenticado. Execute 'gh auth login' e tente novamente."
    }

    Publish-Workbook | Out-Null
    $exitCode = Start-And-WatchWorkflow
    Write-Host ""
    if ($exitCode -eq 0) {
        Write-Host "SUCESSO: planilha publicada e sincronizacao concluida."
    }
    else {
        Write-Host "ERRO: o workflow terminou com falha (codigo $exitCode)."
    }
}
catch {
    Write-Host ""
    Write-Host "ERRO: $($_.Exception.Message)"
    $exitCode = 1
}

exit $exitCode
