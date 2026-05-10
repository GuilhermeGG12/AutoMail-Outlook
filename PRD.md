\# PRD — Outlook Mail Merge Assistant



\## 1. Produto



Outlook Mail Merge Assistant.



Aplicativo desktop para Windows que cria rascunhos no Outlook a partir de uma planilha Excel de clientes contábeis.



O objetivo é substituir o fluxo tradicional de mala direta do Word/Outlook por um fluxo mais simples, seguro e revisável.



\## 2. Problema



O usuário final tem dificuldade para usar mala direta pelo Outlook/Word.



O fluxo atual é frágil porque:



\- depende de Word, Outlook e Excel funcionando juntos;

\- é difícil revisar cada envio antes;

\- anexos personalizados por cliente são difíceis de controlar;

\- erros na planilha podem ir parar no e-mail;

\- caminhos de arquivo quebrados podem passar despercebidos;

\- há risco de enviar cobrança errada para cliente errado;

\- a interface de mala direta é confusa para usuário leigo.



\## 3. Objetivo da V1



Criar um app desktop que:



1\. lê a planilha Excel de clientes;

2\. usa a aba `Clientes`;

3\. valida os dados necessários;

4\. monta e-mails de cobrança personalizados;

5\. anexa o arquivo indicado em `ArquivoAnexo`;

6\. cria rascunhos no Outlook;

7\. gera relatório de sucesso e erro.



A V1 não envia e-mails automaticamente.



\## 4. Usuário-alvo



Usuário de escritório contábil que usa:



\- Windows;

\- Excel;

\- Outlook desktop;

\- planilha de clientes;

\- documentos anexos personalizados.



O usuário não é técnico.



\## 5. Fluxo principal



O fluxo esperado é:



1\. Usuário abre o programa.

2\. Clica em `Selecionar planilha`.

3\. Escolhe o arquivo `.xlsx`.

4\. O programa lê a aba `Clientes`.

5\. O programa valida os dados.

6\. O programa mostra uma tabela com linhas válidas e inválidas.

7\. Usuário corrige a planilha, se necessário.

8\. Usuário clica em `Criar rascunhos no Outlook`.

9\. O programa pede confirmação.

10\. O programa cria rascunhos apenas para linhas válidas.

11\. O programa gera relatório em Excel.

12\. Usuário revisa e envia os rascunhos pelo Outlook.



\## 6. Arquivos reais de referência



Os arquivos reais devem ficar em `samples/`:



\- `samples/Cadastro de Empresas Contabilizuum - maio26.xlsx`

\- `samples/Mala Direta - Email Cobrança CZ2.docx`



O Excel contém os dados dos clientes.



O Word contém o modelo textual usado como base para a cobrança.



Na V1, o programa não precisa ler dinamicamente o Word. O template pode ser implementado diretamente no código, com possibilidade de edição futura.



\## 7. Escopo da V1



\### Dentro do escopo



\- Aplicativo desktop Windows.

\- Seleção de planilha Excel.

\- Leitura da aba `Clientes`.

\- Validação de dados.

\- Tabela de pré-visualização.

\- Criação de rascunhos no Outlook.

\- Modo de teste.

\- Relatório Excel.

\- Build `.exe`.



\### Fora do escopo



\- Envio automático.

\- SMTP.

\- Microsoft Graph.

\- Login.

\- Banco de dados.

\- App web.

\- Multiusuário.

\- Agendamento.

\- Rastreamento de abertura.

\- Integração com nuvem.

\- Leitura dinâmica avançada do Word.



\## 8. Requisitos funcionais



\## RF01 — Selecionar planilha



O app deve permitir selecionar um arquivo `.xlsx`.



Critérios:



\- aceitar somente `.xlsx`;

\- mostrar o caminho selecionado;

\- exibir erro amigável se o arquivo não puder ser aberto.



\## RF02 — Ler aba `Clientes`



O app deve usar a aba `Clientes`.



Critérios:



\- se a aba não existir, mostrar erro;

\- ignorar outras abas;

\- preservar número original da linha no Excel;

\- ignorar linhas totalmente vazias.



\## RF03 — Validar colunas obrigatórias



Colunas obrigatórias:



\- `RAZÃO SOCIAL`

\- `Proprietário/Dirigente`

\- `E-Mail 1`

\- `Valor fev26`

\- `Dia de Pagamento`

\- `PIX`

\- `ArquivoAnexo`



Critérios:



\- se faltar alguma coluna, interromper a validação;

\- listar todas as colunas ausentes;

\- permitir colunas extras.



\## RF04 — Validar e-mails



O app deve validar:



\- `E-Mail 1`

\- `E-Mail 2`

\- `E-Mail 3`



Critérios:



\- `E-Mail 1` é obrigatório;

\- `E-Mail 2` e `E-Mail 3` são opcionais;

\- e-mails vazios opcionais devem ser ignorados;

\- e-mails inválidos tornam a linha inválida;

\- múltiplos e-mails separados por `;` devem ser aceitos.



\## RF05 — Montar destinatários



O app deve combinar os e-mails da linha.



Exemplo:



\- `E-Mail 1 = cliente@email.com`

\- `E-Mail 2 = financeiro@email.com`

\- `E-Mail 3 = vazio`



Resultado:



`To = cliente@email.com; financeiro@email.com`



\## RF06 — Gerar assunto



O assunto deve ser:



`Honorários contábeis - {RAZÃO SOCIAL}`



Exemplo:



`Honorários contábeis - A F SERVICOS DE TRANSPORTES LTDA`



\## RF07 — Gerar corpo do e-mail



O corpo deve usar o template:



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



\## RF08 — Formatar valor



`Valor fev26` deve virar moeda brasileira.



Exemplos:



\- `300` → `R$ 300,00`

\- `400` → `R$ 400,00`

\- `390.5` → `R$ 390,50`



Se o valor for inválido, a linha deve ser inválida.



\## RF09 — Validar erros do Excel



O app deve bloquear linhas com erros do Excel em campos usados no e-mail.



Erros bloqueados:



\- `#REF!`

\- `#NAME?`

\- `#VALUE!`

\- `#N/A`

\- `#DIV/0!`

\- `#NULL!`

\- `#NUM!`



\## RF10 — Validar anexo



O app deve usar `ArquivoAnexo`.



Critérios:



\- obrigatório;

\- deve existir;

\- deve ser arquivo;

\- pode conter múltiplos caminhos separados por `;`;

\- se um anexo estiver inválido, a linha inteira é inválida.



\## RF11 — Pré-visualização



A interface deve mostrar tabela com:



\- Status

\- Linha

\- Razão Social

\- Proprietário/Dirigente

\- Emails

\- Valor

\- Dia de Pagamento

\- Anexo

\- Mensagem



\## RF12 — Criar rascunhos no Outlook



Para cada linha válida:



\- criar um MailItem no Outlook;

\- preencher destinatários;

\- preencher assunto;

\- preencher corpo;

\- adicionar anexos;

\- salvar como rascunho.



Deve usar:



`mail.Save()`



Não usar:



`mail.Send()`



\## RF13 — Modo de teste



Quando ativo:



\- todos os rascunhos vão para o e-mail de teste;

\- destinatários reais não são usados;

\- assunto recebe prefixo `\[TESTE]`;

\- corpo informa destinatários originais.



\## RF14 — Relatório



Gerar relatório `.xlsx` em `reports/`.



Nome:



`relatorio\_envio\_YYYYMMDD\_HHMMSS.xlsx`



Colunas:



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



\## 9. Requisitos não funcionais



\## RNF01 — Segurança



\- Não enviar e-mails automaticamente.

\- Criar somente rascunhos.

\- Pedir confirmação antes da criação em massa.

\- Não armazenar senha.

\- Não pedir credenciais.



\## RNF02 — Usabilidade



\- Interface simples.

\- Mensagens em português.

\- Nada de traceback para o usuário.

\- Botões claros.

\- Fluxo guiado.



\## RNF03 — Confiabilidade



\- Não processar linha inválida.

\- Continuar processamento se uma linha falhar.

\- Registrar tudo no relatório.

\- Não criar rascunhos duplicados dentro da mesma execução.



\## RNF04 — Performance



A V1 deve lidar bem com:



\- até 500 linhas;

\- até 5 anexos por e-mail.



\## 10. Stack técnica



\- Python 3.11+

\- customtkinter

\- openpyxl

\- pywin32

\- pytest

\- ruff

\- black

\- mypy

\- pyinstaller



\## 11. Estrutura esperada



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

├── src/

│   └── mailmerge\_assistant/

└── tests/



\## 12. Critérios de aceite finais



A V1 está pronta quando:



1\. O app abre no Windows.

2\. O usuário seleciona o Excel.

3\. O app lê a aba `Clientes`.

4\. O app valida colunas e linhas.

5\. O app mostra tabela de prévia.

6\. O app bloqueia linhas inválidas.

7\. O app cria rascunhos no Outlook para linhas válidas.

8\. O app não envia nenhum e-mail automaticamente.

9\. O app gera relatório Excel.

10\. O modo de teste funciona.

11\. Os testes automatizados passam.

12\. O build `.exe` é gerado.

