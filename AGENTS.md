\# AGENTS.md



\## Project name



Outlook Mail Merge Assistant



\## Project goal



Build a Windows desktop application that replaces a difficult Outlook/Word mail merge workflow.



The app reads an Excel spreadsheet with accounting clients, validates the data, generates personalized cobrança emails, attaches the correct file for each company, and creates Outlook draft emails.



The app must create drafts only. It must not send emails automatically in V1.



\## Target user



The primary user is a non-technical accounting professional who already uses Microsoft Outlook desktop and Excel on Windows.



The UX must be simple, safe, and clear.



The user should not need to understand Word mail merge, Outlook mail merge, SMTP, OAuth, Microsoft Graph, or programming.



\## Target environment



\- Windows 10 or Windows 11

\- Microsoft Outlook desktop installed and already configured

\- Microsoft Excel files in `.xlsx` format

\- Python 3.11+



Do not implement macOS or Linux support in V1.



\## Tech stack



Use:



\- Python 3.11+

\- customtkinter for the desktop GUI

\- openpyxl for Excel reading/writing

\- pywin32 for Outlook COM automation

\- pytest for tests

\- ruff for linting

\- black for formatting

\- mypy for type checking

\- pyinstaller for building the Windows executable



Do not use:



\- Microsoft Graph

\- SMTP

\- Azure app registration

\- web apps

\- cloud services

\- external databases

\- credentials storage

\- automatic email sending



\## Product principles



1\. Never send emails automatically in V1.

2\. Always create Outlook drafts.

3\. Validate all rows before creating any drafts.

4\. Invalid rows must not generate drafts.

5\. Every run must generate a report.

6\. Error messages must be understandable by a non-technical user.

7\. Avoid destructive or irreversible actions.

8\. Do not store passwords.

9\. Do not ask for Outlook credentials.

10\. Keep Outlook-specific code isolated from GUI code.

11\. Keep business logic outside the UI layer.

12\. Prefer safety over automation.



\## Real files



The user will provide these sample files inside the `samples/` folder:



\- `samples/Cadastro de Empresas Contabilizuum - maio26.xlsx`

\- `samples/Mala Direta - Email Cobrança CZ2.docx`



The Excel file is the source of client data.



The Word file is the base cobrança email model. The app does not need to parse the Word file dynamically in V1. Instead, implement the default cobrança template described below.



\## Real Excel worksheet



Use the worksheet named:



`Clientes`



This worksheet must be selected by default.



If the workbook does not contain `Clientes`, show a friendly error:



`A planilha não possui uma aba chamada "Clientes". Verifique o arquivo selecionado.`



Ignore other worksheets in V1.



\## Real Excel columns



The actual spreadsheet has columns similar to:



\- ID

\- CNPJ

\- RAZÃO SOCIAL

\- Proprietário/Dirigente

\- ME, EPP

\- DATA ABERTURA

\- CPF do Sócio

\- Vencimento Certificado Digital

\- E-Mail 1

\- E-Mail 2

\- E-Mail 3

\- Valor Jan26

\- Valor fev26

\- Dif

\- Extenso

\- Dia de Pagamento

\- PIX

\- ArquivoAnexo



Column names must be normalized by trimming leading/trailing whitespace.



Keep accents and original names supported.



\## Required columns for V1



The following columns are required:



\- RAZÃO SOCIAL

\- Proprietário/Dirigente

\- E-Mail 1

\- Valor fev26

\- Dia de Pagamento

\- PIX

\- ArquivoAnexo



If any required column is missing, stop validation and show a friendly error listing the missing columns.



\## Optional columns



These columns are optional:



\- E-Mail 2

\- E-Mail 3

\- CNPJ

\- ME, EPP

\- Valor Jan26

\- Dif

\- Extenso

\- Vencimento Certificado Digital



\## Recipient behavior



For each valid row:



\- `E-Mail 1` is required.

\- `E-Mail 2` is optional.

\- `E-Mail 3` is optional.



Combine non-empty email fields into the `To` list.



Example:



\- E-Mail 1: cliente@email.com

\- E-Mail 2: financeiro@email.com

\- E-Mail 3: empty



Result:



`To = cliente@email.com; financeiro@email.com`



Do not put these emails in CC by default. They are alternate or additional client emails.



\## Email subject



Use this default subject:



`Honorários contábeis - {RAZÃO SOCIAL}`



Example:



`Honorários contábeis - A F SERVICOS DE TRANSPORTES LTDA`



\## Email body template



Use this default cobrança template:



Olá {Proprietário/Dirigente},



Já enviamos para o seu e-mail nossa nota fiscal de honorários contábeis, os fechamentos na Prefeitura e Receita Federal no Lucro Presumido da sua empresa {RAZÃO SOCIAL}. Caso você tenha tributos a pagar, sugiro já programar esse pagamento até dia {Dia de Pagamento}.



Portanto, por favor, acuse o recebimento. Então, por meio deste, solicito o referido pagamento.



A chave PIX para esse pagamento ({Valor fev26}) é {PIX}, pagamento até dia {Dia de Pagamento} desse mês.



Favorecido: MM GONCALVES ASSESSORIA E CONSULTORIA CONTABIL E FINANCEIRA LTDA, CNPJ 63.648.448/0001-36.



Mas, caso você já tenha feito esse pagamento, por favor, desconsidere essa mensagem e desde já agradecemos seu pagamento, pois só assim poderemos continuar oferecendo nossos serviços contábeis.



Atenciosamente.



Márcio Marques Gonçalves

GONÇALVES CONSULTORIA

(92) 98159-1780 / 3029-6945



\## Template variables



The template must support these variables:



\- `{Proprietário/Dirigente}`

\- `{RAZÃO SOCIAL}`

\- `{Valor fev26}`

\- `{Dia de Pagamento}`

\- `{PIX}`



The app may support additional variables later, but V1 only needs these.



\## Currency formatting



The field `Valor fev26` must be formatted as Brazilian currency.



Examples:



\- `300` becomes `R$ 300,00`

\- `400` becomes `R$ 400,00`

\- `390.5` becomes `R$ 390,50`

\- `1,500.75` should become `R$ 1.500,75` if parsed successfully



If the value cannot be parsed as a number, the row is invalid.



Do not insert raw numeric values like `300` into the email body.



\## Date/payment day behavior



The field `Dia de Pagamento` can be a number, text, or date-like value.



For V1, convert it to a clean string suitable for the email.



Examples:



\- `10` becomes `10`

\- `dia 10` remains `dia 10`

\- empty value is invalid



If the cell contains an Excel error, the row is invalid.



\## Excel error handling



Never insert raw Excel errors into emails.



Treat these values as invalid when they appear in any field used by the email or attachment logic:



\- `#REF!`

\- `#NAME?`

\- `#VALUE!`

\- `#N/A`

\- `#DIV/0!`

\- `#NULL!`

\- `#NUM!`



If a row contains one of these errors in a required field, do not create a draft for that row.



Show a friendly message such as:



`Linha 8: o campo "Dia de Pagamento" está com erro na planilha. Corrija no Excel antes de gerar o e-mail.`



\## Attachment behavior



Use the `ArquivoAnexo` column as the attachment path.



A row is invalid if:



\- `ArquivoAnexo` is empty

\- the file does not exist

\- the path points to a directory instead of a file

\- any path contains an Excel error



Support multiple attachments separated by semicolon.



Example:



`C:\\Clientes\\arquivo1.pdf; C:\\Clientes\\arquivo2.pdf`



If one attachment is invalid, the entire row is invalid.



\## Validation rules



A row is invalid if:



\- `RAZÃO SOCIAL` is empty

\- `Proprietário/Dirigente` is empty

\- `E-Mail 1` is empty

\- any email address is malformed

\- `Valor fev26` is empty

\- `Valor fev26` cannot be parsed as currency

\- `Dia de Pagamento` is empty

\- `PIX` is empty

\- `ArquivoAnexo` is empty

\- any attachment file does not exist

\- any attachment path is a directory

\- any required field contains an Excel error

\- the rendered subject is empty

\- the rendered body is empty



Validation must happen before Outlook draft creation.



\## Email validation



Accept multiple emails separated by semicolon.



Valid examples:



\- `cliente@email.com`

\- `cliente@email.com; financeiro@email.com`



Invalid examples:



\- `cliente@email`

\- `cliente.com`

\- `@email.com`

\- empty string



\## Outlook behavior



Use pywin32 and Outlook COM automation.



For each valid row, create one Outlook MailItem.



Set:



\- `To`

\- `Subject`

\- `Body`

\- `Attachments`



Then call:



`mail.Save()`



Do not call:



`mail.Send()`



in V1.



\## Test mode



Implement a test mode.



When test mode is enabled:



\- All drafts must be addressed to a single test email provided by the user.

\- Original recipients must not be placed in `To`, `CC`, or `BCC`.

\- Prefix the subject with `\[TESTE]`.

\- Prepend the body with the original recipient information.



Example subject:



`\[TESTE] Honorários contábeis - A F SERVICOS DE TRANSPORTES LTDA`



Example body prefix:



MODO DE TESTE

Destinatário original: cliente@email.com; financeiro@email.com



\---



Then include the normal body.



Attachments should still be added normally.



\## Main UI



Build a simple desktop window with these controls:



1\. Button: `Selecionar planilha`

2\. Label showing selected file path

3\. Button: `Validar clientes`

4\. Summary:

&#x20;  - total rows found

&#x20;  - valid rows

&#x20;  - invalid rows

5\. Table with:

&#x20;  - Status

&#x20;  - Linha

&#x20;  - Razão Social

&#x20;  - Proprietário/Dirigente

&#x20;  - Emails

&#x20;  - Valor

&#x20;  - Dia de Pagamento

&#x20;  - Anexo

&#x20;  - Mensagem

6\. Checkbox: `Modo de teste`

7\. Text field: `E-mail de teste`

8\. Button: `Criar rascunhos no Outlook`

9\. Button: `Abrir pasta de relatórios`



Before creating drafts, show a confirmation dialog:



`Serão criados X rascunhos no Outlook. Deseja continuar?`



\## Reports



Every validation/draft run must generate a report in the `reports/` folder.



Filename format:



`relatorio\_envio\_YYYYMMDD\_HHMMSS.xlsx`



Report columns:



\- Data\_Hora

\- Linha

\- Razão Social

\- Proprietário/Dirigente

\- Emails

\- Assunto

\- Valor

\- Dia de Pagamento

\- Anexo

\- Status

\- Mensagem



Possible statuses:



\- `OK`

\- `ERRO`

\- `RASCUNHO\_CRIADO`

\- `IGNORADO`



The report must include both successful and failed rows.



\## Architecture



Use this project structure:



outlook-mail-merge-assistant/

├── AGENTS.md

├── PRD.md

├── README.md

├── pyproject.toml

├── requirements.txt

├── requirements-dev.txt

├── .gitignore

├── samples/

│   ├── Cadastro de Empresas Contabilizuum - maio26.xlsx

│   └── Mala Direta - Email Cobrança CZ2.docx

├── reports/

│   └── .gitkeep

├── scripts/

│   ├── build\_exe.ps1

│   └── create\_sample\_excel.py

├── src/

│   └── mailmerge\_assistant/

│       ├── \_\_init\_\_.py

│       ├── app.py

│       ├── config.py

│       ├── models.py

│       ├── excel\_reader.py

│       ├── clientes\_mapper.py

│       ├── validators.py

│       ├── template\_engine.py

│       ├── outlook\_client.py

│       ├── report\_writer.py

│       ├── controller.py

│       └── ui/

│           ├── \_\_init\_\_.py

│           └── main\_window.py

└── tests/

&#x20;   ├── test\_validators.py

&#x20;   ├── test\_template\_engine.py

&#x20;   ├── test\_clientes\_mapper.py

&#x20;   ├── test\_excel\_reader.py

&#x20;   └── test\_report\_writer.py



\## Module responsibilities



\### `models.py`



Define dataclasses for:



\- spreadsheet rows

\- validation results

\- email drafts

\- processing results

\- report rows



\### `excel\_reader.py`



Responsible for:



\- opening `.xlsx` files

\- selecting the `Clientes` worksheet

\- reading headers

\- normalizing headers

\- preserving original Excel row numbers

\- ignoring fully empty rows

\- returning structured row objects



\### `clientes\_mapper.py`



Responsible for:



\- mapping a row from the `Clientes` worksheet into an email draft

\- combining `E-Mail 1`, `E-Mail 2`, and `E-Mail 3`

\- formatting `Valor fev26` as BRL

\- creating the subject

\- rendering the body

\- detecting Excel errors

\- validating business-specific required fields



\### `validators.py`



Responsible for:



\- validating email lists

\- parsing semicolon-separated values

\- validating attachment paths

\- detecting Excel error strings

\- validating required values



\### `template\_engine.py`



Responsible for:



\- replacing `{variavel}` placeholders

\- detecting missing placeholders

\- not using `eval`

\- not executing code



\### `outlook\_client.py`



Responsible for:



\- all pywin32 Outlook integration

\- creating draft emails

\- converting technical Outlook errors to friendly runtime errors



\### `report\_writer.py`



Responsible for:



\- writing `.xlsx` report files

\- creating the `reports/` folder if missing

\- returning the report path



\### `controller.py`



Responsible for:



\- orchestrating reading, validation, mapping, draft creation, and reporting

\- keeping the UI thin



\### `ui/main\_window.py`



Responsible only for:



\- desktop interface

\- user actions

\- showing validation summaries

\- showing friendly errors

\- calling controller methods



Do not put business logic in the UI.



\## Tests



Write tests for:



\- required column detection

\- row parsing from `Clientes`

\- combining `E-Mail 1`, `E-Mail 2`, `E-Mail 3`

\- invalid email blocking

\- Brazilian currency formatting

\- Excel error detection

\- invalid attachment detection

\- subject generation

\- body template rendering

\- test mode behavior

\- report generation



Outlook must be mocked in tests.



Unit tests must not require real Outlook.



\## Commands



Install environment:



```powershell

py -3.11 -m venv .venv

.\\.venv\\Scripts\\Activate.ps1

python -m pip install --upgrade pip setuptools wheel

pip install -r requirements.txt

pip install -r requirements-dev.txt

