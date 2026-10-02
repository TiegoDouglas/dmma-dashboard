# Sincronização de vagas com Azure DevOps

Esta automação lê `automacao_azure_devops/vagas.xlsx` e sincroniza itens
relacionados no Azure DevOps da organização `AccountMSFT`, projeto
`Esteira de Vagas`.

## Layout da planilha

A aba única `Planilha1` usa duas linhas de cabeçalho:

- linha 1: tipo de work item de cada coluna (`Project` ou `Position`);
- linha 2: nome de exibição do campo no Azure DevOps;
- linha 3 em diante: valores. A coluna A contém apenas os rótulos
  `Tipo de Item`, `Campos` e `valor` e não é enviada ao Azure DevOps.

O arquivo atual possui 30 campos mapeados:

| Tipo | Chave natural | Campos |
|---|---|---|
| `Project` | `OppID` | `OppID`, `Account.`, `PROJETO`, `Win Probability`, `Estimated Close Date`, `Status EALA`, `Sales Stage`, `Must Win Deal`, `Forecast category`, `Estimated Project Start Date`, `Estimated Project End Date`, `Channel` |
| `Position` | `Id MyScheduling` | `Id MyScheduling`, `Status Myscheduling`, `Cliente`, `RoleTitle`, `Skillls`, `Idioma`, `Staffing`, `Role Description`, `Modalidade de Trabalho`, `RolePrimaryContact`, `level_from`, `level_to`, `Prática`, `Sub Prática`, `Role Start Date`, `Role End Date`, `Role Is Overdue`, `DataAberturaScheduling` |

Datas, números, booleanos e textos são preservados de acordo com o tipo da
célula. Fórmulas não são aceitas. Linhas vazias são ignoradas. Cada linha deve
informar as duas chaves naturais.

Mais de uma `Position` pode compartilhar o mesmo `Project`. Quando uma chave se
repete, os demais valores do item também precisam ser idênticos; divergências
falham antes de qualquer chamada de escrita.

## Identidade, criação e relações

A sincronização não usa mais o work item fixo `3133`.

Antes de escrever, o script:

1. consulta os campos válidos de cada tipo pela API;
2. resolve cada nome da linha 2 para um único `referenceName`;
3. rejeita campos ausentes, ambíguos ou somente leitura;
4. localiza `Project` por `OppID` e `Position` por `Id MyScheduling`;
5. falha se houver mais de um item com a mesma chave.

A resolução de campos usa comparação canônica exata: ignora caixa, espaços,
hífens, underscores, pontuação e acentos. Assim, por exemplo, `OppID`, `Opp ID`
e `opp-id` são equivalentes. Também é considerado o último segmento de um
`referenceName`, permitindo relacionar uma coluna a `Custom.OppID`. Não há
fuzzy matching. Se mais de um campo produzir a mesma forma canônica, a execução
falha por ambiguidade. As mensagens de erro listam os candidatos próximos e os
campos graváveis disponíveis, sem exibir valores da planilha ou credenciais.
Todos os cabeçalhos inválidos de um item são agregados na mesma mensagem para
evitar ciclos de diagnóstico campo a campo.

Equivalências de negócio comprovadas pelo processo atual são mantidas em um
mapa explícito no código, incluindo `Project.PROJETO -> System.Title`,
`Project.Account. -> Custom.Account`,
`Project.Win Probability -> Custom.WinProb`, as datas estimadas do projeto para
os campos de agendamento `Start Date`/`Finish Date` e
`Position.RoleTitle -> System.Title`. Configurações em
`AZURE_FIELD_REFERENCE_OVERRIDES` sempre têm precedência sobre esse mapa.

Um item existente é atualizado somente quando algum valor mudou. Se não
existir, ele é criado. `System.Title` é preenchido por `PROJETO` no `Project` e
por `RoleTitle` na `Position` quando nenhuma coluna já resolver para esse campo.

Cada `Position` recebe o `Project` da mesma linha como pai pela relação
`System.LinkTypes.Hierarchy-Reverse`. A relação é criada somente quando ainda
não existe. A automação recusa trocar silenciosamente uma `Position` que já
tenha outro pai.

Se um nome de exibição não for único no processo do Azure DevOps, configure a
variável de repositório ou ambiente `AZURE_FIELD_REFERENCE_OVERRIDES` com um
objeto JSON. As chaves seguem o formato `Tipo.Nome da coluna`:

```json
{
  "Position.Skillls": "Custom.Skills",
  "Project.Account.": "Custom.Account"
}
```

## Configuração no GitHub

Crie o secret de Actions `ADO_PAT` com um Personal Access Token que tenha apenas
`Work Items: Read & write`. O token é lido exclusivamente pelo workflow e nunca
deve ser incluído na planilha, no código, nos logs ou em variáveis comuns.

O workflow **Sincronizar vaga com Azure DevOps**:

- executa automaticamente no início de cada hora;
- pode ser iniciado em **Actions > Sincronizar vaga com Azure DevOps > Run
  workflow**;
- oferece `dry_run`, que valida e mostra os itens/relações planejados sem exigir
  `ADO_PAT` e sem fazer qualquer chamada ao Azure DevOps.

Na execução real, todas as descobertas e buscas são concluídas antes da primeira
criação ou atualização. Assim, um erro de campo ou uma chave duplicada não deixa
uma sincronização parcialmente iniciada.

## Execução por duplo clique no Windows

A pasta `automacao_azure_devops` continua sendo um pacote portátil. O BAT
publica o `vagas.xlsx` que estiver na mesma pasta em
`automacao_azure_devops/vagas.xlsx` da branch `main` e dispara o workflow. Se o
arquivo local já for idêntico ao remoto, nenhum commit desnecessário é criado.

1. Baixe o ZIP do repositório ou clone o projeto.
2. Copie a pasta inteira `automacao_azure_devops` para o Desktop.
3. Instale o [GitHub CLI](https://cli.github.com/).
4. Execute uma vez `gh auth login` com uma conta que possa gravar conteúdo e
   executar Actions em `TiegoDouglas/dmma-dashboard`.
5. Edite e salve `vagas.xlsx` ao lado dos arquivos `.bat` e `.ps1`.
6. Dê duplo clique em `executar_sincronizacao.bat`.

O iniciador aceita caminhos com espaços e caracteres acentuados, abre a
execução no navegador, acompanha o resultado e mantém a janela aberta com a
mensagem final. O BAT da raiz continua disponível como atalho. Nenhum PAT do
Azure DevOps é salvo localmente.

## Validação local

```powershell
python -m pip install -r automacao_azure_devops\requirements.txt
python -m unittest discover -s automacao_azure_devops -p "test_*.py" -v
python automacao_azure_devops\sync_vagas.py --dry-run
```

O dry-run é completamente offline. A execução real requer `ADO_PAT` e foi
projetada para ocorrer no GitHub Actions.
