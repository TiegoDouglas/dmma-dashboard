# Sincronização de vaga com Azure DevOps

Esta automação lê `automacao_azure_devops/vagas.xlsx` e atualiza o work item
`3133` da organização `AccountMSFT`, projeto `Esteira de Vagas`.

A aba `Planilha1` deve ter exatamente os cabeçalhos `Description`, `state` e
`Skills`, além de exatamente uma linha preenchida. O script rejeita linhas com
campos vazios, fórmulas, colunas adicionais preenchidas ou mais de uma linha de
dados.

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

Para executar localmente:

```powershell
python -m pip install -r automacao_azure_devops\requirements.txt
python automacao_azure_devops\sync_vagas.py --dry-run
```

O dry-run não exige `ADO_PAT`. A execução real foi projetada para o workflow e
recebe `ADO_PAT` no GitHub Actions somente por meio do secret descrito acima;
nunca grave o token em arquivo ou variável comum.
