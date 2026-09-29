# Automação Universidade Corporativa Senior (UCS)

Automação em Python para criação/verificação de usuários na plataforma UCS, associação de cursos, envio de e-mail informativo e atualização do status em uma planilha.

O projeto pode usar duas origens de dados:

- Planilha Excel no SharePoint/OneDrive, acessada pelo link completo via Microsoft Graph.
- Arquivo Excel local, usado somente quando `SHAREPOINT_PATH` não estiver definido.

## Pré-requisitos

- Windows.
- Python 3.10 ou superior.
- Google Chrome instalado.
- Outlook desktop configurado, para envio dos e-mails via COM.
- Acesso de supervisor na plataforma UCS.
- Para SharePoint: App Registration no Azure com permissões Microsoft Graph para ler/escrever arquivos.

## Instalação

Crie e ative um ambiente virtual:

```powershell
py -3 -m venv venv
.\venv\Scripts\Activate.ps1
```

Instale as dependências:

```powershell
pip install -r requirements.txt
```

## Configuração

Copie `.env.example` para `.env` e preencha os valores reais:

```powershell
Copy-Item .env.example .env
```

Variáveis disponíveis:

```env
AZURE_TENANT_ID=
AZURE_CLIENT_ID=
AZURE_CLIENT_SECRET=

SHAREPOINT_PATH=
SHAREPOINT_SHEET=Sheet1

LOCAL_FILE_PATH=

UCS_USER=
UCS_PASSWORD=
STD_PASS=
```

### SharePoint

Para usar a planilha do SharePoint, defina `SHAREPOINT_PATH` com o link completo da planilha:

```env
SHAREPOINT_PATH=https://...
SHAREPOINT_SHEET=NomeDaAba
```

Quando `SHAREPOINT_PATH` estiver preenchido, o arquivo local é ignorado.

O App Registration do Azure deve ter permissões de aplicação compatíveis com leitura e escrita de arquivos, por exemplo `Files.ReadWrite.All` ou `Sites.ReadWrite.All`, com consentimento administrativo concedido.

### Arquivo local

Para usar arquivo local, deixe `SHAREPOINT_PATH` vazio e configure:

```env
LOCAL_FILE_PATH=C:\caminho\para\planilha.xlsx
```

Também é possível passar o arquivo local por argumento:

```powershell
py -3 main.py C:\caminho\para\planilha.xlsx
```

## Estrutura esperada da planilha

A planilha precisa conter uma coluna `Status`. Linhas com `Status` igual a `Ok` ou `N Ok` são ignoradas. As demais são processadas.

Colunas usadas pela automação:

- `Nome Completo`
- `CPF`
- `E-mail`
- `Status`
- `Outros cursos`
- `Cursos Disponíveis dentro da Gestão Empresarial | ERP SENIOR`

Os cursos devem estar separados por `;` quando houver mais de um curso na mesma célula.

## Execução

Com o `.env` configurado:

```powershell
py -3 main.py
```

Fluxo de execução:

1. Carrega a planilha do SharePoint, se `SHAREPOINT_PATH` estiver definido.
2. Caso contrário, carrega o arquivo local.
3. Filtra usuários pendentes pelo `Status`.
4. Abre o Chrome e faz login na UCS.
5. Verifica se o usuário já existe.
6. Se existir, coleta todos os cursos já atribuídos, percorrendo as páginas da lista de cursos.
7. Associa apenas os cursos faltantes.
8. Se não existir, cria o usuário e associa os cursos solicitados.
9. Envia e-mail informativo pelo Outlook.
10. Atualiza o `Status` da linha para `Ok` ou `N Ok`.
11. Salva a planilha na mesma origem usada na leitura.

## Arquivos principais

- `main.py`: ponto de entrada e escolha da origem da planilha.
- `dados.py`: leitura, filtro de pendentes e salvamento da planilha.
- `graph_api.py`: integração com Microsoft Graph para SharePoint/OneDrive.
- `navegador.py`: abertura do Chrome e login na UCS.
- `ucs.py`: orquestração da execução.
- `usuarios.py`: criação/verificação de usuários, status e regras de cursos.
- `cursos.py`: associação dos cursos na UCS.
- `informativo.py`: envio de e-mail pelo Outlook.
- `utils.py`: logger, leitura do `.env`, credenciais e funções auxiliares.

## Segurança

O arquivo `.env` contém senhas e chaves reais e não deve ser versionado. Ele já está listado no `.gitignore`.

Versione somente `.env.example`, sem valores sensíveis.

Antes de subir para o Git, confira:

```powershell
git status --short
```

Se `.env` aparecer como arquivo rastreado, remova-o do controle de versão antes de enviar o projeto e rotacione os segredos caso eles já tenham sido publicados.

## Logs

Os logs são gravados na pasta `logs/`, com nome no formato:

```text
AT_UCS_DD-MM-YYYY_HH-MM-SS.log
```

A pasta `logs/` também está ignorada pelo Git.

## Observações Técnicas

- A leitura da planilha SharePoint usa `usedRange()` do Microsoft Graph.
- O salvamento é explícito: URL salva via Graph; caminho local salva via `pandas.to_excel`.
- A automação usa seletores Selenium da tela atual da UCS. Mudanças na interface da plataforma podem exigir ajuste de XPaths/IDs.
- O envio de e-mail depende do Outlook desktop estar instalado e autenticado no Windows.
