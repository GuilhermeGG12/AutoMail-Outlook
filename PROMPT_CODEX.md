\# Prompt para Codex



Implemente este projeto do zero seguindo exatamente o `AGENTS.md` e o `PRD.md`.



O objetivo é criar um aplicativo desktop Windows em Python que substitui a mala direta do Outlook/Word.



\## Contexto real



O usuário tem uma planilha Excel com clientes contábeis e um modelo Word de cobrança.



Os arquivos reais estão em:



\- `samples/Cadastro de Empresas Contabilizuum - maio26.xlsx`

\- `samples/Mala Direta - Email Cobrança CZ2.docx`



A V1 não precisa ler dinamicamente o Word. Use o template definido no `AGENTS.md` e no `PRD.md`.



\## O que implementar



Crie a estrutura completa:



outlook-mail-merge-assistant/

├── AGENTS.md

├── PRD.md

├── README.md

├── pyproject.toml

├── requirements.txt

├── requirements-dev.txt

├── .gitignore

├── samples/

├── reports/

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



\## Funcionalidades obrigatórias



1\. Ler arquivo `.xlsx`.

2\. Usar a aba `Clientes`.

3\. Validar colunas obrigatórias:

&#x20;  - `RAZÃO SOCIAL`

&#x20;  - `Proprietário/Dirigente`

&#x20;  - `E-Mail 1`

&#x20;  - `Valor fev26`

&#x20;  - `Dia de Pagamento`

&#x20;  - `PIX`

&#x20;  - `ArquivoAnexo`

4\. Combinar destinatários:

&#x20;  - `E-Mail 1`

&#x20;  - `E-Mail 2`

&#x20;  - `E-Mail 3`

5\. Gerar assunto:

&#x20;  - `Honorários contábeis - {RAZÃO SOCIAL}`

6\. Gerar corpo usando o template do PRD.

7\. Formatar `Valor fev26` como moeda brasileira.

8\. Validar anexos da coluna `ArquivoAnexo`.

9\. Bloquear erros de Excel:

&#x20;  - `#REF!`

&#x20;  - `#NAME?`

&#x20;  - `#VALUE!`

&#x20;  - `#N/A`

&#x20;  - `#DIV/0!`

&#x20;  - `#NULL!`

&#x20;  - `#NUM!`

10\. Criar rascunhos no Outlook usando pywin32.

11\. Usar `mail.Save()`.

12\. Nunca usar `mail.Send()` na V1.

13\. Implementar modo de teste.

14\. Gerar relatório `.xlsx` em `reports/`.

15\. Implementar interface gráfica simples com customtkinter.

16\. Implementar testes automatizados.

17\. Implementar script de build com PyInstaller.



\## Interface esperada



A janela deve ter:



\- botão `Selecionar planilha`;

\- label com caminho do arquivo;

\- botão `Validar clientes`;

\- resumo de válidos/inválidos;

\- tabela de prévia;

\- checkbox `Modo de teste`;

\- campo `E-mail de teste`;

\- botão `Criar rascunhos no Outlook`;

\- botão `Abrir pasta de relatórios`.



Antes de criar rascunhos, mostrar confirmação:



`Serão criados X rascunhos no Outlook. Deseja continuar?`



\## Regras críticas



\- Não enviar e-mails automaticamente.

\- Não implementar SMTP.

\- Não implementar Microsoft Graph.

\- Não pedir login.

\- Não armazenar senha.

\- Não colocar lógica de negócio dentro da UI.

\- Mockar Outlook nos testes.

\- Testes não podem depender de Outlook instalado.



\## Arquivos de configuração



Crie `pyproject.toml` com black, ruff, mypy e pytest configurados.



Crie `requirements.txt` com:



\- customtkinter

\- openpyxl

\- pywin32

\- pydantic

\- python-dotenv



Crie `requirements-dev.txt` com:



\- pytest

\- pytest-cov

\- ruff

\- black

\- mypy

\- pyinstaller



\## Comandos que devem funcionar



Instalação:



```powershell

py -3.11 -m venv .venv

.\\.venv\\Scripts\\Activate.ps1

python -m pip install --upgrade pip setuptools wheel

pip install -r requirements.txt

pip install -r requirements-dev.txt

