from __future__ import annotations

from pathlib import Path

APP_NAME = "Outlook Mail Merge Assistant"
WORKSHEET_NAME = "Clientes"
REPORTS_DIR = Path("reports")

REQUIRED_COLUMNS = [
    "RAZÃO SOCIAL",
    "Proprietário/Dirigente",
    "E-Mail 1",
    "Valor fev26",
    "Dia de Pagamento",
    "PIX",
]

OPTIONAL_COLUMNS = [
    "E-Mail 2",
    "E-Mail 3",
    "CNPJ",
    "ME, EPP",
    "Valor Jan26",
    "Dif",
    "Extenso",
    "Vencimento Certificado Digital",
    "ArquivoAnexo",
]

EMAIL_SUBJECT_TEMPLATE = "Honorários contábeis - {RAZÃO SOCIAL}"

EMAIL_BODY_TEMPLATE = "\n\n".join(
    [
        "Olá {Proprietário/Dirigente},",
        (
            "Já enviamos para o seu e-mail nossa nota fiscal de honorários contábeis, "
            "os fechamentos na Prefeitura e Receita Federal no Lucro Presumido da sua "
            "empresa {RAZÃO SOCIAL}. Caso você tenha tributos a pagar, sugiro já "
            "programar esse pagamento até dia {Dia de Pagamento}."
        ),
        (
            "Portanto, por favor, acuse o recebimento. Então, por meio deste, solicito "
            "o referido pagamento."
        ),
        (
            "A chave PIX para esse pagamento ({Valor fev26}) é {PIX}, pagamento até dia "
            "{Dia de Pagamento} desse mês."
        ),
        (
            "Favorecido: MM GONCALVES ASSESSORIA E CONSULTORIA CONTABIL E FINANCEIRA "
            "LTDA, CNPJ 63.648.448/0001-36."
        ),
        (
            "Mas, caso você já tenha feito esse pagamento, por favor, desconsidere essa "
            "mensagem e desde já agradecemos seu pagamento, pois só assim poderemos "
            "continuar oferecendo nossos serviços contábeis."
        ),
        "Atenciosamente.",
        "Márcio Marques Gonçalves\nGONÇALVES CONSULTORIA\n(92) 98159-1780 / 3029-6945",
    ]
)

EXCEL_ERROR_VALUES = {
    "#REF!",
    "#NAME?",
    "#VALUE!",
    "#N/A",
    "#DIV/0!",
    "#NULL!",
    "#NUM!",
}
