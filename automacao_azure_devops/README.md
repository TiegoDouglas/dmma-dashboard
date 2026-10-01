# Sincronização de vaga com Azure DevOps

Esta automação lê `automacao_azure_devops/vagas.xlsx` e atualiza o work item
`3133` da organização `AccountMSFT`, projeto `Esteira de Vagas`.

A aba `Planilha1` deve ter exatamente os cabeçalhos `Description`, `state` e
`Skills`, além de exatamente uma linha preenchida. O script rejeita linhas com
campos vazios, fórmulas, colunas adicionais preenchidas ou mais de uma linha de
dados.

Antes da atualização, o script consulta o tipo do work item e os estados
permitidos para esse tipo. O valor de `state` aceita correspondência exata sem
diferenciar maiúsculas e minúsculas. Também são reconhecidos aliases por
categoria, incluindo `Em andamento` para `InProgress`, `Novo`/`New` para
`Proposed` e `Concluído`/`Closed` para `Completed`. Se não houver uma resolução
única, a execução falha e lista os nomes válidos retornados pelo Azure DevOps.

## Configuração no GitHub

Crie o secret de Actions `ADO_PAT` com um Personal Access Token do Azure DevOps
que tenha somente a permissão necessária **Work Items: Read & write**. O PAT é
lido exclusivamente pela variável de ambiente no workflow, não deve ser
incluído na planilha, no código, nos logs ou em variáveis comuns do repositório.
Se o secret estiver ausente, a execução falhará explicitamente.

Por padrão, o script consulta a API de campos do projeto para descobrir o
`referenceName` do campo exibido como `Skills`. Se houver campos ambíguos ou a
consulta não retornar esse nome, configure a variável de repositório ou ambiente
`AZURE_SKILLS_FIELD_REFERENCE_NAME` com o `referenceName` correto, por exemplo
`Custom.Skills`.

## Execução

O workflow **Sincronizar vaga com Azure DevOps**:

- executa automaticamente no início de cada hora;
- pode ser iniciado em **Actions > Sincronizar vaga com Azure DevOps > Run
  workflow**;
- permite marcar `dry_run` na execução manual para validar a planilha sem
  acessar ou alterar o Azure DevOps.

### Execução por duplo clique no Windows

A pasta `automacao_azure_devops` é um pacote autocontido para Windows. O BAT
publica o `vagas.xlsx` que estiver na mesma pasta diretamente no caminho
`automacao_azure_devops/vagas.xlsx` da branch `main` e, depois de confirmar a
publicação, dispara e acompanha o workflow. Se o arquivo local já for idêntico
ao remoto, nenhum commit desnecessário é criado.

Passo a passo:

1. baixe o ZIP do repositório pelo GitHub (ou clone o repositório);
2. copie a pasta inteira `automacao_azure_devops` para o Desktop, sem separar os
   arquivos `.bat`, `.ps1` e `vagas.xlsx`;
3. instale o [GitHub CLI](https://cli.github.com/);
4. abra um terminal uma única vez e execute `gh auth login` com uma conta que
   possa gravar conteúdo e executar Actions em `TiegoDouglas/dmma-dashboard`;
5. edite e salve `vagas.xlsx` dentro da pasta copiada no Desktop;
6. dê duplo clique em `automacao_azure_devops\executar_sincronizacao.bat`.

O iniciador aceita caminhos com espaços e caracteres acentuados, abre a página
da execução no navegador, acompanha o resultado até o fim e mantém a janela
aberta com uma mensagem de sucesso ou erro. O BAT da raiz continua disponível
como atalho quando o repositório completo é usado. Nenhum PAT é salvo
localmente: o Azure DevOps continua usando somente o secret `ADO_PAT` do GitHub
Actions, enquanto a publicação da planilha usa a autenticação existente do
GitHub CLI.

Para executar localmente:

```powershell
python -m pip install -r automacao_azure_devops\requirements.txt
python -m unittest discover -s automacao_azure_devops -p "test_*.py" -v
python automacao_azure_devops\sync_vagas.py --dry-run
```

O dry-run não exige `ADO_PAT`. A execução real foi projetada para o workflow e
recebe `ADO_PAT` no GitHub Actions somente por meio do secret descrito acima;
nunca grave o token em arquivo ou variável comum.
