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

O arquivo `executar_sincronizacao.bat`, na raiz do repositório, dispara o
workflow na branch `main`, localiza a execução criada e acompanha o resultado
até o fim. O token do Azure DevOps não é armazenado no arquivo: a execução
continua usando o secret `ADO_PAT` configurado no GitHub Actions.

Antes de usar:

1. publique a versão atualizada de `automacao_azure_devops/vagas.xlsx` na branch
   `main` (por commit/pull request); o workflow usa o arquivo do GitHub, não uma
   alteração que exista apenas no computador;
2. instale o [GitHub CLI](https://cli.github.com/);
3. autentique-se uma vez com `gh auth login` usando uma conta com permissão para
   executar Actions neste repositório;
4. dê duplo clique em `executar_sincronizacao.bat`.

A janela permanece aberta ao final e informa sucesso ou erro. Em caso de falha,
o arquivo preserva o código de saída retornado pelo acompanhamento do workflow.

Para executar localmente:

```powershell
python -m pip install -r automacao_azure_devops\requirements.txt
python -m unittest discover -s automacao_azure_devops -p "test_*.py" -v
python automacao_azure_devops\sync_vagas.py --dry-run
```

O dry-run não exige `ADO_PAT`. A execução real foi projetada para o workflow e
recebe `ADO_PAT` no GitHub Actions somente por meio do secret descrito acima;
nunca grave o token em arquivo ou variável comum.
