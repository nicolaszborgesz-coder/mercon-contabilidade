from pathlib import Path
from decimal import Decimal, InvalidOperation
from datetime import datetime
import xml.etree.ElementTree as ET
import csv
import unicodedata
import re

from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill,
    Font,
    Alignment,
    Border,
    Side
)
from openpyxl.utils import get_column_letter
MES = "2026-10"
from pathlib import 

BASE = Path(__file__).parent.parent

ARQUIVO_FINAL = BASE / f"MERCON_CONCILIACAO_V5_1_{MES}.xlsx"

CLIENTES_POR_CNPJ = {
    "12345678000101": BASE / "01-BARBOSA-MERCADO - 12345678000101",
    "12345678000102": BASE / "02-OFICINA-SAO-JOSE - 12345678000102",
    "12345678000103": BASE / "03-FARMACIA-CENTRAL - 12345678000103",
    "12345678000104": BASE / "04-MERCADINHO-GOIAS - 12345678000104",
    "12345678000105": BASE / "05-LOJA-ROUPAS-MARIA - 12345678000105",
    "12345678000106": BASE / "06-PADARIA-PAO-QUENTE - 12345678000106",
    "12345678000107": BASE / "07-AGROPECUARIA-RURAL - 12345678000107",
    "12345678000108": BASE / "08-CLINICA-VETERINARIA - 12345678000108",
    "12345678000109": BASE / "09-AUTO-PECAS-GOIANESIA - 12345678000109",
    "12345678000110": BASE / "10-SALAO-BELEZA-ELITE - 12345678000110",
    "12345678000111": BASE / "11-RESTAURANTE-SABOR - 12345678000111",
    "12345678000112": BASE / "12-MATERIAL-CONSTRUCAO - 12345678000112",
    "12345678000113": BASE / "13-DISTRIBUIDORA-BEBIDAS - 12345678000113",
    "12345678000114": BASE / "14-LAN-HOUSE-GAMES - 12345678000114",
    "12345678000115": BASE / "15-ESCRITORIO-ADVOCACIA - 12345678000115",
    "12345678000116": BASE / "16-CONTABILIDADE-PARCEIRA - 12345678000116",
}

AZUL_ESCURO = "17365D"
AZUL_MEDIO = "4472C4"
AZUL_CLARO = "D9EAF7"
BRANCO = "FFFFFF"
CINZA = "F2F4F7"
CINZA_TEXTO = "666666"
VERDE = "C6EFCE"
VERDE_TEXTO = "006100"
AMARELO = "FFEB9C"
AMARELO_TEXTO = "9C6500"
VERMELHO = "FFC7CE"
VERMELHO_TEXTO = "9C0006"
LARANJA = "FCE4D6"

FILL_AZUL = PatternFill("solid", fgColor=AZUL_ESCURO)
FILL_AZUL_MEDIO = PatternFill("solid", fgColor=AZUL_MEDIO)
FILL_AZUL_CLARO = PatternFill("solid", fgColor=AZUL_CLARO)
FILL_CINZA = PatternFill("solid", fgColor=CINZA)
FILL_VERDE = PatternFill("solid", fgColor=VERDE)
FILL_AMARELO = PatternFill("solid", fgColor=AMARELO)
FILL_VERMELHO = PatternFill("solid", fgColor=VERMELHO)
FILL_LARANJA = PatternFill("solid", fgColor=LARANJA)

FONTE_BRANCA = Font(
    name="Arial",
    size=10,
    bold=True,
    color=BRANCO
)

FONTE_NORMAL = Font(
    name="Arial",
    size=10,
    color="333333"
)

BORDA_FINA = Border(
    bottom=Side(style="thin", color="D9E1F2")
)

FORMATO_MOEDA = '"R$" #,##0.00;[Red]-"R$" #,##0.00'
def normalizar_texto(texto):
    texto = str(texto).strip().lower()

    texto = unicodedata.normalize(
        "NFKD", texto
    ).encode("ASCII", "ignore").decode("ASCII")

    return texto


def decimal_seguro(valor):
    """
    Converte valores monetários para Decimal.

    Aceita:
    1.234,56
    1234,56
    1234.56
    R$ 1.234,56
    """

    if valor is None:
        return Decimal("0")

    texto = str(valor).strip()

    if not texto:
        return Decimal("0")

    texto = texto.replace("R$", "")
    texto = texto.replace(" ", "")

    negativo = False

    if texto.startswith("(") and texto.endswith(")"):
        negativo = True
        texto = texto[1:-1]

    if texto.startswith("-"):
        negativo = True
        texto = texto[1:]

    texto = re.sub(r"[^0-9,.]", "", texto)

    if not texto:
        return Decimal("0")

    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "")
            texto = texto.replace(",", ".")
        else:
            texto = texto.replace(",", "")

    elif "," in texto:
        texto = texto.replace(".", "")
        texto = texto.replace(",", ".")

    try:
        resultado = Decimal(texto)

        if negativo:
            resultado = -resultado

        return resultado

    except InvalidOperation:
        return Decimal("0")


def nome_tag(elemento):
    return elemento.tag.split("}")[-1]


def configurar_aba(ws):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90

    ws.freeze_panes = "A6"

    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    ws.sheet_properties.pageSetUpPr.autoPageBreaks = False

    ws.oddHeader.center.text = "&BM E R C O N"
    ws.oddFooter.center.text = "Página &P de &N"


def criar_titulo(ws, titulo, subtitulo, ultima_coluna):
    ws.merge_cells(
        start_row=1,
        start_column=1,
        end_row=1,
        end_column=ultima_coluna
    )

    celula = ws.cell(1, 1, titulo)
    celula.fill = FILL_AZUL
    celula.font = Font(
        name="Arial",
        size=18,
        bold=True,
        color=BRANCO
    )
    celula.alignment = Alignment(
        horizontal="left",
        vertical="center"
    )

    ws.row_dimensions[1].height = 38

    ws.merge_cells(
        start_row=2,
        start_column=1,
        end_row=2,
        end_column=ultima_coluna
    )

    celula = ws.cell(2, 1, subtitulo)
    celula.font = Font(
        name="Arial",
        size=10,
        italic=True,
        color=CINZA_TEXTO
    )
    celula.alignment = Alignment(
        vertical="center"
    )

    ws.row_dimensions[2].height = 25


def criar_cabecalho(ws, linha, quantidade_colunas):
    for coluna in range(1, quantidade_colunas + 1):
        celula = ws.cell(linha, coluna)

        celula.fill = FILL_AZUL_MEDIO
        celula.font = FONTE_BRANCA
        celula.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        celula.border = BORDA_FINA

    ws.row_dimensions[linha].height = 32


def estilizar_dados(ws, linha_inicio, linha_fim, coluna_inicio, coluna_fim):
    for linha in range(linha_inicio, linha_fim + 1):
        for coluna in range(coluna_inicio, coluna_fim + 1):

            celula = ws.cell(linha, coluna)

            celula.font = FONTE_NORMAL
            celula.border = BORDA_FINA
            celula.alignment = Alignment(
                vertical="center"
            )

            if linha % 2 == 0:
                celula.fill = FILL_CINZA


def colorir_status(celula):
    status = str(celula.value or "").upper()

    if status == "OK - CONCILIADO":
        celula.fill = FILL_VERDE
        celula.font = Font(
            bold=True,
            color=VERDE_TEXTO
        )

    elif status in (
        "PENDENTE PARCIAL",
        "PENDENTE TOTAL",
        "EXTRATO AUSENTE",
        "COMPETENCIA AUSENTE"
    ):
        celula.fill = FILL_AMARELO
        celula.font = Font(
            bold=True,
            color=AMARELO_TEXTO
        )

    elif status in (
        "DIVERGENCIA",
        "SOBRA BANCARIA"
    ):
        celula.fill = FILL_VERMELHO
        celula.font = Font(
            bold=True,
            color=VERMELHO_TEXTO
        )

    elif status == "SEM MOVIMENTO":
        celula.fill = FILL_LARANJA
        celula.font = Font(
            bold=True,
            color="C65911"
        )


def ajustar_larguras(ws, larguras):
    for coluna, largura in larguras.items():
        ws.column_dimensions[coluna].width = largura


def criar_cartao(ws, intervalo, titulo, valor, formato=None):
    ws.merge_cells(intervalo)

    celula = ws[intervalo.split(":")[0]]
    celula.value = f"{titulo}\n{valor}"

    celula.fill = FILL_AZUL_CLARO

    celula.font = Font(
        name="Arial",
        size=13,
        bold=True,
        color=AZUL_ESCURO
    )

    celula.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )
  def ler_xml(caminho):
    """
    Extrai valor total da NF-e, quantidade de itens e CFOPs.
    """

    arvore = ET.parse(caminho)
    raiz = arvore.getroot()

    valor_total = Decimal("0")
    quantidade_itens = 0
    cfops = []

    for elemento in raiz.iter():

        tag = nome_tag(elemento)

        if tag == "vNF":
            valor_total = decimal_seguro(elemento.text)

        elif tag == "det":
            quantidade_itens += 1

        elif tag == "CFOP":
            cfop = (elemento.text or "").strip()

            if cfop:
                cfops.append(cfop)

    if valor_total == 0:
        raise ValueError(
            "Valor total da NF-e não encontrado ou igual a zero."
        )

    return {
        "valor": valor_total,
        "itens": quantidade_itens,
        "cfops": sorted(set(cfops))
    }


def processar_xmls(pasta):
    total = Decimal("0")
    quantidade = 0
    itens = 0

    lista_cfops = set()
    erros = []

    arquivos = list(pasta.glob("*.xml"))

    for arquivo in arquivos:

        try:
            dados = ler_xml(arquivo)

            total += dados["valor"]
            quantidade += 1
            itens += dados["itens"]

            lista_cfops.update(dados["cfops"])

        except Exception as erro:

            erros.append({
                "arquivo": arquivo.name,
                "erro": str(erro)
            })

    return {
        "total": total,
        "quantidade": quantidade,
        "itens": itens,
        "cfops": sorted(lista_cfops),
        "erros": erros
    }
  def ler_extrato(caminho):
    """
    Lê extratos CSV com separadores ; , ou tabulação.

    Prioriza colunas de crédito e débito quando existem.
    Se existir somente uma coluna Valor, considera valores
    positivos como créditos e negativos como débitos.
    """

    encodings = [
        "utf-8-sig",
        "utf-8",
        "latin-1"
    ]

    delimitadores = [
        ";",
        ",",
        "\t"
    ]

    linhas = None

    for encoding in encodings:

        for delimitador in delimitadores:

            try:
                with open(
                    caminho,
                    "r",
                    encoding=encoding,
                    newline=""
                ) as arquivo:

                    leitor = csv.DictReader(
                        arquivo,
                        delimiter=delimitador
                    )

                    dados = list(leitor)

                    if leitor.fieldnames and dados:
                        linhas = dados
                        break

            except (UnicodeDecodeError, csv.Error):
                continue

        if linhas is not None:
            break

    if linhas is None:
        raise ValueError(
            "Não foi possível interpretar o arquivo CSV."
        )

    creditos = Decimal("0")
    debitos = Decimal("0")
    lancamentos = 0

    for linha in linhas:

        valores = {
            normalizar_texto(chave): valor
            for chave, valor in linha.items()
            if chave is not None
        }

        valor_credito = None
        valor_debito = None

        for chave, valor in valores.items():

            if chave in (
                "credito",
                "creditos",
                "entrada",
                "entradas"
            ):
                valor_credito = decimal_seguro(valor)

            elif chave in (
                "debito",
                "debitos",
                "saida",
                "saidas"
            ):
                valor_debito = decimal_seguro(valor)

        if valor_credito is not None or valor_debito is not None:

            creditos += valor_credito or Decimal("0")
            debitos += valor_debito or Decimal("0")

            lancamentos += 1

        else:

            valor = None

            for chave, conteudo in valores.items():

                if chave in (
                    "valor",
                    "valor lancamento",
                    "valor movimento"
                ):
                    valor = decimal_seguro(conteudo)
                    break

            if valor is not None:

                if valor >= 0:
                    creditos += valor
                else:
                    debitos += abs(valor)

                lancamentos += 1

    return {
        "creditos": creditos,
        "debitos": debitos,
        "lancamentos": lancamentos
    }
def processar_clientes():

    resultados = []
    erros_xml = []
    pastas_nao_encontradas = []

    for cnpj, pasta_cliente in CLIENTES_POR_CNPJ.items():

        nome_cliente = pasta_cliente.name

        pasta_competencia = pasta_cliente / MES

        if not pasta_cliente.exists():

            pastas_nao_encontradas.append({
                "cnpj": cnpj,
                "cliente": nome_cliente,
                "pasta": str(pasta_cliente),
                "motivo": "PASTA DO CLIENTE NÃO ENCONTRADA"
            })

            continue

        if not pasta_competencia.exists():

            resultados.append({
                "cnpj": cnpj,
                "cliente": nome_cliente,
                "quantidade": 0,
                "valor_nfe": Decimal("0"),
                "creditos": Decimal("0"),
                "debitos": Decimal("0"),
                "diferenca": Decimal("0"),
                "status": "COMPETENCIA AUSENTE",
                "itens": 0,
                "erros_cfop": 0,
                "xml_erros": 0,
                "lancamentos": 0,
                "cfops": "",
                "extrato": "NÃO"
            })

            pastas_nao_encontradas.append({
                "cnpj": cnpj,
                "cliente": nome_cliente,
                "pasta": str(pasta_competencia),
                "motivo": "PASTA DA COMPETÊNCIA NÃO ENCONTRADA"
            })

            continue

        dados_xml = processar_xmls(pasta_competencia)

        caminho_extrato = pasta_competencia / "extrato.csv"

        extrato_encontrado = caminho_extrato.exists()

        if extrato_encontrado:

            try:
                dados_extrato = ler_extrato(caminho_extrato)

            except Exception as erro:

                dados_extrato = {
                    "creditos": Decimal("0"),
                    "debitos": Decimal("0"),
                    "lancamentos": 0
                }

                erros_xml.append({
                    "cnpj": cnpj,
                    "cliente": nome_cliente,
                    "arquivo": "extrato.csv",
                    "erro": f"ERRO NA LEITURA DO EXTRATO: {erro}"
                })

                extrato_encontrado = False

        else:

            dados_extrato = {
                "creditos": Decimal("0"),
                "debitos": Decimal("0"),
                "lancamentos": 0
            }

        for erro in dados_xml["erros"]:

            erros_xml.append({
                "cnpj": cnpj,
                "cliente": nome_cliente,
                "arquivo": erro["arquivo"],
                "erro": erro["erro"]
            })

        valor_nfe = dados_xml["total"]
        creditos = dados_extrato["creditos"]
        debitos = dados_extrato["debitos"]

        diferenca = creditos - valor_nfe

        status = classificar_status(
            valor_nfe,
            creditos,
            extrato_encontrado
        )

        resultados.append({
            "cnpj": cnpj,
            "cliente": nome_cliente,
            "quantidade": dados_xml["quantidade"],
            "valor_nfe": valor_nfe,
            "creditos": creditos,
            "debitos": debitos,
            "diferenca": diferenca,
            "status": status,
            "itens": dados_xml["itens"],
            "erros_cfop": 0,
            "xml_erros": len(dados_xml["erros"]),
            "lancamentos": dados_extrato["lancamentos"],
            "cfops": ", ".join(dados_xml["cfops"]),
            "extrato": "SIM" if extrato_encontrado else "NÃO"
        })

    return resultados, erros_xml, pastas_nao_encontradas
  def gerar_excel(resultados, erros_xml, pastas_nao_encontradas):

    wb = Workbook()
      ws = wb.active
    ws.title = "CONCILIACAO"

    configurar_aba(ws)

    colunas = [
        "CNPJ",
        "CLIENTE",
        "QTD. NF-e",
        "VALOR NF-e",
        "CRÉDITOS",
        "DÉBITOS",
        "DIFERENÇA",
        "STATUS",
        "QTD. ITENS",
        "ERROS CFOP",
        "XML COM ERRO",
        "LANÇAMENTOS",
        "CFOPs",
        "EXTRATO"
    ]

    criar_titulo(
        ws,
        "MERCON | CONCILIAÇÃO FISCAL",
        f"Relatório de conferência - Competência: {MES}",
        len(colunas)
    )

    for coluna, nome in enumerate(colunas, start=1):
        ws.cell(5, coluna, nome)

    criar_cabecalho(ws, 5, len(colunas))

    for resultado in resultados:

        valores = [
            resultado["cnpj"],
            resultado["cliente"],
            resultado["quantidade"],
            float(resultado["valor_nfe"]),
            float(resultado["creditos"]),
            float(resultado["debitos"]),
            float(resultado["diferenca"]),
            resultado["status"],
            resultado["itens"],
            resultado["erros_cfop"],
            resultado["xml_erros"],
            resultado["lancamentos"],
            resultado["cfops"],
            resultado["extrato"]
        ]

        ws.append(valores)

    ultima_linha = ws.max_row

    if ultima_linha >= 6:

        estilizar_dados(
            ws,
            6,
            ultima_linha,
            1,
            len(colunas)
        )

        for linha in range(6, ultima_linha + 1):

            for coluna in (4, 5, 6, 7):
                ws.cell(linha, coluna).number_format = FORMATO_MOEDA

            colorir_status(ws.cell(linha, 8))

            ws.cell(linha, 1).number_format = "@"

    ws.auto_filter.ref = f"A5:N{ultima_linha}"

    ajustar_larguras(
        ws,
        {
            "A": 19,
            "B": 39,
            "C": 12,
            "D": 17,
            "E": 17,
            "F": 17,
            "G": 17,
            "H": 23,
            "I": 12,
            "J": 13,
            "K": 14,
            "L": 14,
            "M": 25,
            "N": 12
        }
    )
resumo = wb.create_sheet("RESUMO")

    configurar_aba(resumo)

    criar_titulo(
        resumo,
        "MERCON | PAINEL GERENCIAL",
        f"Resumo da conciliação - Competência: {MES}",
        8
    )

    total_clientes = len(resultados)

    total_nfe = sum(
        (r["valor_nfe"] for r in resultados),
        Decimal("0")
    )

    total_creditos = sum(
        (r["creditos"] for r in resultados),
        Decimal("0")
    )

    total_debitos = sum(
        (r["debitos"] for r in resultados),
        Decimal("0")
    )

    conciliados = sum(
        1 for r in resultados
        if r["status"] == "OK - CONCILIADO"
    )

    pendentes = sum(
        1 for r in resultados
        if r["status"] in (
            "PENDENTE PARCIAL",
            "PENDENTE TOTAL",
            "COMPETENCIA AUSENTE"
        )
    )

    divergencias = sum(
        1 for r in resultados
        if r["status"] in (
            "DIVERGENCIA",
            "SOBRA BANCARIA"
        )
    )

    criar_cartao(
        resumo,
        "A4:B6",
        "CLIENTES PROCESSADOS",
        total_clientes
    )

    criar_cartao(
        resumo,
        "C4:D6",
        "CONCILIADOS",
        conciliados
    )

    criar_cartao(
        resumo,
        "E4:F6",
        "PENDENTES",
        pendentes
    )

    criar_cartao(
        resumo,
        "G4:H6",
        "DIVERGÊNCIAS",
        divergencias
    )

    resumo.merge_cells("A8:B8")
    resumo["A8"] = "INDICADORES FINANCEIROS"

    for celula in ("A8", "B8"):
        resumo[celula].fill = FILL_AZUL_MEDIO
        resumo[celula].font = FONTE_BRANCA

    resumo["A9"] = "Total NF-e"
    resumo["B9"] = float(total_nfe)

    resumo["A10"] = "Total de créditos"
    resumo["B10"] = float(total_creditos)

    resumo["A11"] = "Total de débitos"
    resumo["B11"] = float(total_debitos)

    resumo["A12"] = "Diferença geral"
    resumo["B12"] = float(total_creditos - total_nfe)

    for linha in range(9, 13):
        resumo.cell(linha, 1).font = FONTE_NORMAL
        resumo.cell(linha, 2).font = Font(
            bold=True,
            color=AZUL_ESCURO
        )

        resumo.cell(linha, 2).number_format = FORMATO_MOEDA

        resumo.cell(linha, 1).border = BORDA_FINA
        resumo.cell(linha, 2).border = BORDA_FINA

    # Tabela de status

    resumo["D8"] = "STATUS"
    resumo["E8"] = "QUANTIDADE"

    for celula in ("D8", "E8"):
        resumo[celula].fill = FILL_AZUL_MEDIO
        resumo[celula].font = FONTE_BRANCA
        resumo[celula].alignment = Alignment(
            horizontal="center"
        )

    status_lista = [
        "OK - CONCILIADO",
        "PENDENTE PARCIAL",
        "PENDENTE TOTAL",
        "DIVERGENCIA",
        "SOBRA BANCARIA",
        "EXTRATO AUSENTE",
        "COMPETENCIA AUSENTE",
        "SEM MOVIMENTO"
    ]

    linha_status = 9

    for status in status_lista:

        quantidade = sum(
            1 for r in resultados
            if r["status"] == status
        )

        resumo.cell(linha_status, 4, status)
        resumo.cell(linha_status, 5, quantidade)

        colorir_status(resumo.cell(linha_status, 4))

        resumo.cell(linha_status, 5).alignment = Alignment(
            horizontal="center"
        )

        resumo.cell(linha_status, 4).border = BORDA_FINA
        resumo.cell(linha_status, 5).border = BORDA_FINA

        linha_status += 1

    resumo["A15"] = "INFORMAÇÕES DO PROCESSAMENTO"
    resumo["A15"].fill = FILL_AZUL_MEDIO
    resumo["A15"].font = FONTE_BRANCA

    resumo["A16"] = "Competência"
    resumo["B16"] = MES

    resumo["A17"] = "Data de geração"
    resumo["B17"] = datetime.now().strftime("%d/%m/%Y %H:%M")

    resumo["A18"] = "XMLs com erro"
    resumo["B18"] = len(erros_xml)

    resumo["A19"] = "Pastas não encontradas"
    resumo["B19"] = len(pastas_nao_encontradas)

    for linha in range(16, 20):
        resumo.cell(linha, 1).font = FONTE_NORMAL
        resumo.cell(linha, 2).font = FONTE_NORMAL

    ajustar_larguras(
        resumo,
        {
            "A": 28,
            "B": 22,
            "C": 5,
            "D": 27,
            "E": 16,
            "F": 5,
            "G": 18,
            "H": 18
        }
    )

    resumo.freeze_panes = "A8"
ws_erros = wb.create_sheet("ERROS_XML")

    configurar_aba(ws_erros)

    criar_titulo(
        ws_erros,
        "MERCON | REGISTRO DE ERROS",
        "Arquivos XML ou extratos que apresentaram problemas",
        4
    )

    cabecalho_erros = [
        "CNPJ",
        "CLIENTE",
        "ARQUIVO",
        "DESCRIÇÃO DO ERRO"
    ]

    for coluna, nome in enumerate(cabecalho_erros, start=1):
        ws_erros.cell(5, coluna, nome)

    criar_cabecalho(ws_erros, 5, 4)

    for erro in erros_xml:

        ws_erros.append([
            erro["cnpj"],
            erro["cliente"],
            erro["arquivo"],
            erro["erro"]
        ])

    if ws_erros.max_row >= 6:
        estilizar_dados(
            ws_erros,
            6,
            ws_erros.max_row,
            1,
            4
        )

    ws_erros.auto_filter.ref = f"A5:D{ws_erros.max_row}"

    ajustar_larguras(
        ws_erros,
        {
            "A": 20,
            "B": 40,
            "C": 35,
            "D": 75
        }
    )
ws_pastas = wb.create_sheet("PASTAS_NAO_ENCONTRADAS")

    configurar_aba(ws_pastas)

    criar_titulo(
        ws_pastas,
        "MERCON | PASTAS NÃO LOCALIZADAS",
        "Clientes ou competências que não foram encontrados",
        4
    )

    cabecalho_pastas = [
        "CNPJ",
        "CLIENTE",
        "CAMINHO",
        "MOTIVO"
    ]

    for coluna, nome in enumerate(cabecalho_pastas, start=1):
        ws_pastas.cell(5, coluna, nome)

    criar_cabecalho(ws_pastas, 5, 4)

    for pasta in pastas_nao_encontradas:

        ws_pastas.append([
            pasta["cnpj"],
            pasta["cliente"],
            pasta["pasta"],
            pasta["motivo"]
        ])

    if ws_pastas.max_row >= 6:
        estilizar_dados(
            ws_pastas,
            6,
            ws_pastas.max_row,
            1,
            4
        )

    ws_pastas.auto_filter.ref = f"A5:D{ws_pastas.max_row}"

    ajustar_larguras(
        ws_pastas,
        {
            "A": 20,
            "B": 40,
            "C": 95,
            "D": 45
        }
    )
for aba in wb.worksheets:

        aba.sheet_properties.pageSetUpPr.fitToPage = True

        aba.sheet_view.showGridLines = False

        aba.sheet_properties.tabColor = AZUL_MEDIO

    wb.properties.creator = "Mercon - Sistema de Conciliação"
    wb.properties.title = "Relatório de Conciliação Fiscal"
    wb.properties.subject = f"Conciliação da competência {MES}"

    BASE.mkdir(
        parents=True,
        exist_ok=True
    )

    wb.save(ARQUIVO_FINAL)

    print("\n======================================")
    print("       MERCON - CONCILIADOR")
    print("======================================")

    print(f"Competência: {MES}")
    print(f"Clientes processados: {total_clientes}")
    print(f"Clientes conciliados: {conciliados}")
    print(f"Pendências: {pendentes}")
    print(f"Divergências: {divergencias}")
    print(f"Erros encontrados: {len(erros_xml)}")

    print("\nRelatório gerado com sucesso!")
    print(f"Arquivo: {ARQUIVO_FINAL}")
if __name__ == "__main__":

    print("Iniciando conciliação Mercon...")
    print(f"Competência: {MES}")

    resultados, erros_xml, pastas_nao_encontradas = (
        processar_clientes()
    )

    gerar_excel(
        resultados,
        erros_xml,
        pastas_nao_encontradas
    )
