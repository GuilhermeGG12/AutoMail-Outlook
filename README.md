# Outlook Mail Merge Assistant

Aplicativo desktop Windows para validar uma planilha Excel de clientes contábeis e criar rascunhos personalizados no Microsoft Outlook.

O app usa Outlook desktop via COM. Ele não usa SMTP, não usa Microsoft Graph e não solicita credenciais.

## Instalação

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip setuptools wheel
pip install -r requirements.txt
pip install -r requirements-dev.txt
```

## Executar

```powershell
python -m mailmerge_assistant.app
```

## Testes

```powershell
pytest
ruff check .
black --check .
mypy src
```

## Build do executável

```powershell
.\scripts\build_exe.ps1
```

Se a execução de scripts estiver bloqueada:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_exe.ps1
```

## Fluxo

1. Selecione um arquivo `.xlsx`.
2. Valide a aba `Clientes`.
3. Revise linhas válidas e inválidas.
4. Opcionalmente ative o modo de teste.
5. Use `Visualizar e-mail` para conferir formatação, negritos e logo.
6. Crie rascunhos no Outlook ou envie e-mails pelo Outlook.
7. Para envio automático, confirme digitando `ENVIAR`.
8. Consulte o relatório gerado em `reports/`.
