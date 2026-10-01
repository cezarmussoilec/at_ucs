# Automação Universidade Corporativa Senior (UCS)

Automação em Python para processar solicitações de acesso à Universidade Corporativa Senior (UCS). O projeto lê uma planilha de usuários pendentes, cria ou localiza usuários na UCS, associa cursos, envia e-mails informativos pelo Outlook e atualiza o status de processamento na própria planilha.

## Funcionalidades

- Leitura de planilha Excel via SharePoint/OneDrive usando Microsoft Graph.
- Leitura alternativa de arquivo Excel local.
- Processamento apenas de linhas pendentes, ignorando registros com status `Ok` ou `N Ok`.
- Login automatizado na UCS com Selenium e Google Chrome.
- Criação de novos usuários com senha padrão.
- Verificação de usuários já existentes pelo CPF.
- Associação apenas dos cursos ainda não atribuídos ao usuário.
- Envio de e-mail de primeiro acesso para novos usuários.
- Envio opcional de e-mail para usuários existentes quando novos cursos forem liberados.
- Atualização da planilha de origem com `Ok` ou `N Ok`.
- Geração de logs por execução.
- Envio opcional de alerta/resumo da execução por e-mail.

## Tecnologias

- Python 3.10+
- Selenium
- pandas
- openpyxl
- python-dotenv
- requests
- Microsoft Graph API
- Outlook Desktop via COM (`pywin32`)

## Pré-requisitos

- Windows.
- Python 3.10 ou superior.
- Google Chrome instalado.
- Outlook Desktop instalado, autenticado e configurado no perfil do Windows que executará a automação.
- Acesso de supervisor na plataforma UCS.
- Planilha Excel com as colunas esperadas.
- Para uso com SharePoint/OneDrive: App Registration no Azure com permissões Microsoft Graph para leitura e escrita do arquivo.

## Instalação

Clone o repositório e acesse a pasta do projeto:

```powershell
git clone <url-do-repositorio>
cd at_ucs
```

Crie e ative um ambiente virtual:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Instale as dependências:

```powershell
pip install -r requirements.txt
```

## Configuração

Copie o arquivo de exemplo e preencha as variáveis com os dados reais do ambiente:

```powershell
Copy-Item .env.example .env
```

Variáveis disponíveis:

```env
# Credenciais Azure AD / Microsoft Graph
AZURE_TENANT_ID=
AZURE_CLIENT_ID=
AZURE_CLIENT_SECRET=

# Planilha no SharePoint/OneDrive
SHAREPOINT_PATH=
SHAREPOINT_SHEET=Sheet1

# Planilha local
LOCAL_FILE_PATH=/path/to/local/file.xlsx

# Execução do navegador
HEADLESS_BROWSER=false

# Credenciais UCS
UCS_USER=
UCS_PASSWORD=
STD_PASS=

# E-mails informativos
OUTLOOK_FROM=
INFORMATIVE_EMAIL_BCC=
SEND_EXISTING_USER_EMAIL=true

# Alertas da execução
ALERT_EMAIL_TO=
ALERT_EMAIL_FROM=
ALERT_ON_SUCCESS=true
ALERT_ON_ERROR=true
```

### Variáveis principais

| Variável | Obrigatória | Descrição |
| --- | --- | --- |
| `UCS_USER` | Sim | Usuário usado para login na UCS. |
| `UCS_PASSWORD` | Sim | Senha do usuário usado para login na UCS. |
| `STD_PASS` | Sim | Senha padrão definida para novos usuários criados. |
| `SHAREPOINT_PATH` | Não | URL completa da planilha no SharePoint/OneDrive. Quando preenchida, tem prioridade sobre arquivo local. |
| `SHAREPOINT_SHEET` | Não | Nome da aba da planilha usada no SharePoint/OneDrive. Valor padrão: `Sheet1`. |
| `LOCAL_FILE_PATH` | Não | Caminho da planilha local, usado quando `SHAREPOINT_PATH` estiver vazio. |
| `HEADLESS_BROWSER` | Não | Define se o Chrome será executado sem janela visível. Aceita `true`, `false`, `1`, `0`, `sim`, `não`, `yes` ou `no`. |
| `OUTLOOK_FROM` | Não | Remetente usado nos e-mails informativos e, por padrão, nos alertas. |
| `INFORMATIVE_EMAIL_BCC` | Não | Cópia oculta dos e-mails informativos. Separe múltiplos destinatários por `;`. |
| `SEND_EXISTING_USER_EMAIL` | Não | Envia e-mail para usuário já existente quando novos cursos forem atribuídos. |
| `ALERT_EMAIL_TO` | Não | Destinatários do resumo da execução. Separe múltiplos destinatários por `;` ou `,`. |
| `ALERT_EMAIL_FROM` | Não | Remetente dos alertas. Se vazio, usa `OUTLOOK_FROM`. |
| `ALERT_ON_SUCCESS` | Não | Envia resumo ao final da execução. |
| `ALERT_ON_ERROR` | Não | Envia alerta quando ocorre erro fatal. |

## Origem dos dados

### SharePoint/OneDrive

Para processar uma planilha hospedada no SharePoint ou OneDrive, configure:

```env
SHAREPOINT_PATH=https://...
SHAREPOINT_SHEET=NomeDaAba
```

Quando `SHAREPOINT_PATH` estiver preenchida, o sistema ignora `LOCAL_FILE_PATH` e qualquer caminho local informado por argumento.

O App Registration do Azure deve ter permissões de aplicação compatíveis com leitura e escrita de arquivos, como `Files.ReadWrite.All` ou `Sites.ReadWrite.All`, com consentimento administrativo concedido.

### Arquivo local

Para processar uma planilha local, deixe `SHAREPOINT_PATH` vazia e configure:

```env
LOCAL_FILE_PATH=C:\caminho\para\planilha.xlsx
```

Também é possível informar o arquivo local diretamente na execução:

```powershell
py -3 main.py C:\caminho\para\planilha.xlsx
```

## Estrutura da planilha

A planilha precisa conter a coluna `Status`. Linhas com `Status` igual a `Ok` ou `N Ok` são ignoradas. As demais linhas são consideradas pendentes.

Colunas utilizadas pela automação:

| Coluna | Descrição |
| --- | --- |
| `Nome Completo` | Nome do usuário que será criado ou consultado na UCS. |
| `CPF` | CPF usado como login e chave de busca na UCS. |
| `E-mail` | Destinatário dos e-mails informativos. |
| `Status` | Controle de processamento. A automação grava `Ok` ou `N Ok`. |
| `Outros cursos` | Cursos adicionais solicitados. |
| `Cursos Disponíveis dentro da Gestão Empresarial ERP SENIOR` | Cursos específicos da trilha ERP Senior. |

Quando houver mais de um curso na mesma célula, separe os nomes com ponto e vírgula (`;`).

## Cursos suportados

A automação possui regras de associação para os cursos abaixo:

- `Gestão de Pessoas | HCM`
- `Otimização Logística`
- `Gestão de Armazenagem | WMS Senior`
- `Gestão de Mão de Obra no Armazém`
- `Gestão Empresarial | ERP XT - Trilha`
- `ERP XT Cadastros Iniciais`
- `ERP XT Mercado`
- `ERP XT Manufatura`
- `ERP XT Qualidade`
- `ERP XT Finanças`
- `ERP XT Custos`
- `ERP XT Integrações`
- `ERP XT Suprimentos`
- `ERP XT Serviços`
- `ERP XT Controladoria`

Se a planilha contiver um curso sem regra implementada, o registro será marcado como `N Ok` e o erro será registrado no log.

## Execução

Com o ambiente virtual ativado e o arquivo `.env` configurado, execute:

```powershell
py -3 main.py
```

Ou usando diretamente o Python do ambiente virtual:

```powershell
.\.venv\Scripts\python.exe main.py
```

Fluxo executado:

1. Carrega a planilha do SharePoint/OneDrive quando `SHAREPOINT_PATH` está configurada.
2. Caso contrário, carrega a planilha local.
3. Filtra usuários pendentes pela coluna `Status`.
4. Abre o Chrome e acessa a UCS.
5. Realiza login e troca para o perfil de supervisor.
6. Consulta o usuário pelo CPF.
7. Para usuário existente, coleta os cursos já atribuídos e associa apenas os cursos faltantes.
8. Para usuário novo, cria o cadastro, associa os cursos solicitados e envia o e-mail de primeiro acesso.
9. Envia e-mail para usuário existente quando novos cursos forem atribuídos e `SEND_EXISTING_USER_EMAIL=true`.
10. Atualiza a linha da planilha com `Ok` ou `N Ok`.
11. Salva a planilha na mesma origem usada na leitura.
12. Envia o alerta de execução, quando configurado.

## Execução em segundo plano

Para executar o Chrome sem janela visível:

```env
HEADLESS_BROWSER=true
```

Para acompanhar a automação na tela:

```env
HEADLESS_BROWSER=false
```

Mesmo em modo headless, o envio de e-mails depende do Outlook Desktop instalado, autenticado e acessível na sessão do Windows. Para agendamentos, recomenda-se executar a automação com o usuário do Windows conectado.

## E-mails

O envio é feito pelo Outlook Desktop via COM.

`OUTLOOK_FROM` define o remetente dos e-mails. Se o endereço estiver configurado como conta no Outlook, o envio usa essa conta. Caso contrário, a automação tenta usar envio delegado com `SentOnBehalfOfName`.

Para o destinatário visualizar exatamente o outro endereço como remetente, o usuário do Outlook precisa ter permissão `Send As` no Exchange/Microsoft 365. Com permissão `Send on behalf`, o destinatário pode ver a indicação "em nome de".

`INFORMATIVE_EMAIL_BCC` permite enviar cópia oculta dos e-mails informativos:

```env
INFORMATIVE_EMAIL_BCC=email1@empresa.com;email2@empresa.com
```

E-mails informativos enviados:

- Usuário novo: informa link da UCS, login, senha inicial e orientação de primeiro acesso.
- Usuário existente: informa que o cadastro já estava ativo e lista os novos cursos liberados.

## Alertas de execução

Configure `ALERT_EMAIL_TO` para receber o resumo da execução:

```env
ALERT_EMAIL_TO=responsavel@empresa.com
ALERT_EMAIL_FROM=
ALERT_ON_SUCCESS=true
ALERT_ON_ERROR=true
```

O alerta de sucesso contém:

- total de usuários pendentes;
- total processado com `Ok`;
- total processado com `N Ok`;
- caminho do arquivo de log;
- usuários processados;
- cursos solicitados;
- cursos já atribuídos anteriormente;
- cursos atribuídos na execução;
- erro por usuário, quando houver.

Use `ALERT_ON_SUCCESS=false` para desativar o resumo de execuções concluídas. Use `ALERT_ON_ERROR=false` para desativar alertas de erro fatal.

## Estrutura do projeto

```text
.
├── alertas.py          # Envio de alertas e resumo da execução
├── cursos.py           # Regras de acesso ao catálogo e associação de cursos
├── dados.py            # Leitura, filtro e salvamento da planilha
├── graph_api.py        # Integração com Microsoft Graph
├── informativo.py      # Envio de e-mails informativos pelo Outlook
├── main.py             # Ponto de entrada da aplicação
├── navegador.py        # Configuração do Chrome, login e perfil supervisor
├── ucs.py              # Orquestração da automação na UCS
├── usuarios.py         # Cadastro, consulta e processamento dos usuários
├── utils.py            # Logger, variáveis de ambiente e funções auxiliares
├── requirements.txt    # Dependências Python
└── .env.example        # Exemplo de configuração
```

## Logs

Os logs são gravados na pasta `logs/` com o padrão:

```text
AT_UCS_DD-MM-YYYY_HH-MM-SS.log
```

A pasta `logs/` é ignorada pelo Git.

## Segurança

O arquivo `.env` contém senhas, chaves e dados sensíveis. Ele não deve ser versionado.

Boas práticas:

- Versione apenas `.env.example`, sem valores reais.
- Não publique credenciais da UCS, Azure ou Outlook.
- Revogue e rotacione segredos caso algum valor sensível seja exposto.
- Antes de enviar alterações ao repositório, confira:

```powershell
git status --short
```

## Observações técnicas

- A leitura da planilha no SharePoint/OneDrive usa `usedRange()` do Microsoft Graph.
- O salvamento no SharePoint/OneDrive atualiza o intervalo usado da planilha via Microsoft Graph.
- O salvamento local usa `pandas.to_excel`.
- A automação depende dos seletores atuais da interface da UCS. Mudanças na plataforma podem exigir manutenção de XPaths, IDs e tempos de espera.
- O envio de e-mail depende do Outlook Desktop e da sessão do Windows.
- O CPF é normalizado para conter apenas números antes da busca e criação do login.
