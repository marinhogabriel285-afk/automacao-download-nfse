import os
import ssl
import gzip
import base64
import time
import urllib3
import requests
import xml.etree.ElementTree as ET
from requests_pkcs12 import Pkcs12Adapter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ==============================================================================
# CONFIGURAÇÕES FIXAS
# ==============================================================================
CERTIFICADO_PFX = "COLOQUE_AQUI_O_NOME_DO_CERTIFICADO.pfx"
SENHA_CERTIFICADO = "COLOQUE_AQUI_A_SENHA_DO_CERTIFICADO"
MEU_CNPJ = "COLOQUE_AQUI_O_CNPJ_SO_NUMEROS"
PASTA_SAIDA = "notas_baixadas"
# ==============================================================================

urllib3.disable_warnings()


class GovernoAdapter(Pkcs12Adapter):
    def init_poolmanager(self, *args, **kwargs):
        ctx = ssl.create_default_context()
        ctx.minimum_version = ssl.TLSVersion.TLSv1_2
        ctx.maximum_version = ssl.TLSVersion.TLSv1_2
        ctx.set_ciphers("DEFAULT:@SECLEVEL=1")
        kwargs["ssl_context"] = ctx
        return super().init_poolmanager(*args, **kwargs)


def criar_sessao():
    sessao = requests.Session()
    sessao.mount(
        "https://",
        GovernoAdapter(pkcs12_filename=CERTIFICADO_PFX, pkcs12_password=SENHA_CERTIFICADO)
    )
    return sessao


def solicitar_parametros():
    print("=" * 60)
    print("   DOWNLOAD DE NFS-e RECEBIDAS + PLANILHA EXCEL")
    print("=" * 60)

    while True:
        try:
            mes = int(input("Digite o MÊS desejado (1 a 12): ").strip())
            if 1 <= mes <= 12:
                break
            print("Por favor, digite um número entre 1 e 12.")
        except ValueError:
            print("Entrada inválida. Digite apenas números (ex: 10).")

    while True:
        try:
            ano = int(input("Digite o ANO desejado (ex: 2026): ").strip())
            if 2020 <= ano <= 2100:
                break
            print("Por favor, digite um ano válido com 4 dígitos.")
        except ValueError:
            print("Entrada inválida. Digite o ano com 4 dígitos (ex: 2026).")

    resp_pdf = input("Deseja gerar também os PDFs das notas? [S/N]: ").strip().upper()
    print("-" * 60)
    return mes, ano, (resp_pdf == "S")


def buscar_tag(elemento, nome_tag):
    if elemento is None:
        return ""
    for elem in elemento.iter():
        if elem.tag.endswith(nome_tag) and elem.text:
            return elem.text.strip()
    return ""


def para_float(valor_str):
    if not valor_str:
        return 0.0
    try:
        return float(valor_str.replace(",", "."))
    except ValueError:
        return 0.0


def analisar_xml(xml_bytes, meu_cnpj):
    dados = {
        "ano": None,
        "mes": None,
        "eh_recebida": False,
        "numero": "",
        "data_emissao": "",
        "municipio": "",
        "prest_cnpj": "",
        "prest_nome": "",
        "toma_cnpj": "",
        "toma_nome": "",
        "descricao": "",
        "valor_bruto": 0.0,
        "previ_retida": 0.0,
        "contrib_sociais_retidas": 0.0,
        "valor_liquido_calc": 0.0,
        "valor_iss": 0.0
    }

    try:
        raiz = ET.fromstring(xml_bytes)
        dados["numero"] = buscar_tag(raiz, "nNFSe") or buscar_tag(raiz, "nDPS")
        dados["municipio"] = buscar_tag(raiz, "xLocEmi")
        data_raw = buscar_tag(raiz, "dhEmi") or buscar_tag(raiz, "dCompet")

        if len(data_raw) >= 10:
            dados["ano"] = int(data_raw[0:4])
            dados["mes"] = int(data_raw[5:7])
            dados["data_emissao"] = f"{data_raw[8:10]}/{data_raw[5:7]}/{data_raw[0:4]}"

        for elem in raiz.iter():
            if elem.tag.endswith("emit") or elem.tag.endswith("prest"):
                cnpj_encontrado = "".join(filter(str.isdigit, buscar_tag(elem, "CNPJ")))
                if cnpj_encontrado:
                    dados["prest_cnpj"] = cnpj_encontrado
                    if not dados["prest_nome"]:
                        dados["prest_nome"] = buscar_tag(elem, "xNome")

        for elem in raiz.iter():
            if elem.tag.endswith("toma"):
                dados["toma_cnpj"] = "".join(filter(str.isdigit, buscar_tag(elem, "CNPJ")))
                dados["toma_nome"] = buscar_tag(elem, "xNome")
                break

        dados["descricao"] = buscar_tag(raiz, "xDescServ")

        # Extração dos valores numéricos para cálculo exato
        v_serv = para_float(buscar_tag(raiz, "vServ") or buscar_tag(raiz, "vBC"))
        v_ret_cp = para_float(buscar_tag(raiz, "vRetCP"))  # Contribuição Previdenciária Retida

        # Contribuições Sociais Retidas (PIS, COFINS, CSLL)
        v_ret_pis = para_float(buscar_tag(raiz, "vRetPIS"))
        v_ret_cofins = para_float(buscar_tag(raiz, "vRetCOFINS"))
        v_ret_csll = para_float(buscar_tag(raiz, "vRetCSLL"))
        v_contrib_sociais = round(v_ret_pis + v_ret_cofins + v_ret_csll, 2)

        # Fallback caso o XML traga apenas o total retido federal sem detalhar PIS/COFINS/CSLL
        if v_contrib_sociais == 0.0:
            v_ret_fed = para_float(buscar_tag(raiz, "vRetFed"))
            if v_ret_fed > v_ret_cp:
                v_contrib_sociais = round(v_ret_fed - v_ret_cp, 2)

        v_liquido_calc = round(v_serv - v_ret_cp - v_contrib_sociais, 2)

        dados["valor_bruto"] = v_serv
        dados["previ_retida"] = v_ret_cp
        dados["contrib_sociais_retidas"] = v_contrib_sociais
        dados["valor_liquido_calc"] = v_liquido_calc
        dados["valor_iss"] = para_float(buscar_tag(raiz, "vISSQN"))

    except Exception:
        pass

    meu_cnpj_limpo = "".join(filter(str.isdigit, meu_cnpj))
    # É recebida se o prestador NÃO for o próprio condomínio
    dados["eh_recebida"] = (dados["prest_cnpj"] != "" and dados["prest_cnpj"] != meu_cnpj_limpo)
    return dados


def formatar_cnpj(cnpj):
    c = "".join(filter(str.isdigit, str(cnpj)))
    if len(c) == 14:
        return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"
    return c


def gerar_pdf_do_xml(dados, chave_acesso, pasta_destino_pdf):
    nome_arquivo = f"NFSe_{dados['numero']}_{chave_acesso[:10]}.pdf"
    caminho_pdf = os.path.join(pasta_destino_pdf, nome_arquivo)
    c = canvas.Canvas(caminho_pdf, pagesize=A4)
    largura, altura = A4

    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, altura - 50, "DANFSe - DOCUMENTO AUXILIAR DA NFS-e NACIONAL")
    c.setFont("Helvetica", 9)
    c.drawString(40, altura - 68, f"Chave de Acesso: {chave_acesso}")
    c.line(40, altura - 78, largura - 40, altura - 78)

    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, altura - 100, f"Número da NFS-e: {dados['numero']}")
    c.drawString(220, altura - 100, f"Data de Emissão: {dados['data_emissao']}")
    c.drawString(400, altura - 100, f"Município: {dados['municipio']}")
    c.line(40, altura - 112, largura - 40, altura - 112)

    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, altura - 132, "PRESTADOR DE SERVIÇOS (EMITENTE)")
    c.setFont("Helvetica", 10)
    c.drawString(40, altura - 148, f"Razão Social: {dados['prest_nome']}")
    c.drawString(40, altura - 164, f"CNPJ: {formatar_cnpj(dados['prest_cnpj'])}")
    c.line(40, altura - 176, largura - 40, altura - 176)

    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, altura - 196, "TOMADOR DE SERVIÇOS (CLIENTE)")
    c.setFont("Helvetica", 10)
    c.drawString(40, altura - 212, f"Razão Social: {dados['toma_nome']}")
    c.drawString(40, altura - 228, f"CNPJ: {formatar_cnpj(dados['toma_cnpj'])}")
    c.line(40, altura - 240, largura - 40, altura - 240)

    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, altura - 260, "VALORES E RETENÇÕES FEDERAIS")
    c.setFont("Helvetica", 10)
    c.drawString(40, altura - 278, f"Valor Bruto: R$ {dados['valor_bruto']:.2f}")
    c.drawString(210, altura - 278, f"Previdência Ret.: R$ {dados['previ_retida']:.2f}")
    c.drawString(390, altura - 278, f"Contrib. Sociais: R$ {dados['contrib_sociais_retidas']:.2f}")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, altura - 296, f"Valor Líquido (Bruto - Retenções): R$ {dados['valor_liquido_calc']:.2f}")
    c.line(40, altura - 308, largura - 40, altura - 308)

    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, altura - 328, "DISCRIMINAÇÃO DOS SERVIÇOS")
    c.setFont("Helvetica", 9)

    y = altura - 346
    linha_atual = ""
    for palavra in dados["descricao"].split():
        if len(linha_atual + " " + palavra) <= 95:
            linha_atual = (linha_atual + " " + palavra).strip()
        else:
            c.drawString(40, y, linha_atual)
            y -= 14
            linha_atual = palavra
            if y < 60:
                break
    if linha_atual and y >= 60:
        c.drawString(40, y, linha_atual)

    c.line(40, 50, largura - 40, 50)
    c.setFont("Helvetica-Oblique", 8)
    c.drawString(40, 38, "Documento gerado localmente a partir do XML oficial assinado digitalmente.")
    c.save()
    print(f"   -> PDF gerado: {nome_arquivo}")


def gerar_planilha_excel(lista_notas, mes, ano, pasta_destino):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Recebidas_{mes:02d}_{ano}"

    colunas = [
        "Número da Nota",
        "Data de Emissão",
        "Nome do Prestador",
        "CNPJ do Prestador",
        "Valor Bruto (Serviço)",
        "Contrib. Previdenciária Retida",
        "Contrib. Sociais Retidas",
        "Valor Líquido (Bruto - Retenções)"
    ]
    ws.append(colunas)

    # Estilização do cabeçalho
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_header = Font(color="FFFFFF", bold=True)

    for col_idx in range(1, len(colunas) + 1):
        celula = ws.cell(row=1, column=col_idx)
        celula.fill = fill_header
        celula.font = font_header
        celula.alignment = Alignment(horizontal="center", vertical="center")

    formato_moeda = 'R$ #,##0.00'

    # Inserção dos dados das notas recebidas
    for idx, nota in enumerate(lista_notas, start=2):
        ws.append([
            nota["numero"],
            nota["data_emissao"],
            nota["prest_nome"],
            formatar_cnpj(nota["prest_cnpj"]),
            nota["valor_bruto"],
            nota["previ_retida"],
            nota["contrib_sociais_retidas"],
            nota["valor_liquido_calc"]
        ])

        ws.cell(row=idx, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=idx, column=2).alignment = Alignment(horizontal="center")
        ws.cell(row=idx, column=4).alignment = Alignment(horizontal="center")

        for col_valor in (5, 6, 7, 8):
            cel = ws.cell(row=idx, column=col_valor)
            cel.number_format = formato_moeda

    # Linha de Totais no final (usando fórmulas do Excel)
    if lista_notas:
        linha_total = len(lista_notas) + 2
        ws.cell(row=linha_total, column=1, value="TOTAL").font = Font(bold=True)
        for col_idx, letra_col in [(5, "E"), (6, "F"), (7, "G"), (8, "H")]:
            cel_total = ws.cell(row=linha_total, column=col_idx)
            cel_total.value = f"=SUM({letra_col}2:{letra_col}{linha_total - 1})"
            cel_total.font = Font(bold=True)
            cel_total.number_format = formato_moeda

    # Ajuste automático da largura das colunas
    larguras = {
        "A": 18,
        "B": 18,
        "C": 42,
        "D": 22,
        "E": 24,
        "F": 30,
        "G": 26,
        "H": 32
    }
    for col_letra, largura in larguras.items():
        ws.column_dimensions[col_letra].width = largura

    nome_excel = f"Relatorio_Recebidas_{ano}-{mes:02d}.xlsx"
    caminho_excel = os.path.join(pasta_destino, nome_excel)
    wb.save(caminho_excel)
    return caminho_excel


def iniciar_automacao():
    mes_desejado, ano_desejado, gerar_pdf = solicitar_parametros()

    pasta_base_periodo = os.path.join(PASTA_SAIDA, f"{ano_desejado}-{mes_desejado:02d}")
    pasta_recebidas_xml = os.path.join(pasta_base_periodo, "recebidas_xml")
    os.makedirs(pasta_recebidas_xml, exist_ok=True)

    if gerar_pdf:
        pasta_recebidas_pdf = os.path.join(pasta_base_periodo, "recebidas_pdf")
        os.makedirs(pasta_recebidas_pdf, exist_ok=True)

    servidor = "adn.nfse.gov.br"
    url_base = f"https://{servidor}/contribuintes"

    sessao = criar_sessao()
    headers = {"Accept": "application/json", "User-Agent": "Mozilla/5.0"}

    nsu_atual = 0
    notas_recebidas_planilha = []

    print(f"\nBuscando notas RECEBIDAS de {mes_desejado:02d}/{ano_desejado} para o CNPJ {MEU_CNPJ}...\n")

    while True:
        nsu_formatado = f"{nsu_atual:015d}"
        url_dfe = f"{url_base}/DFe/{nsu_formatado}?cnpjConsulta={MEU_CNPJ}&lote=true"

        try:
            resposta = sessao.get(url_dfe, headers=headers, timeout=45)
        except Exception as e:
            print(f"Erro de conexão com o portal: {e}")
            break

        if resposta.status_code == 200:
            dados_lote = resposta.json()
            lote = dados_lote.get("LoteDFe", [])

            if not lote:
                break

            for doc in lote:
                nsu_doc = int(doc["NSU"])
                chave = doc["ChaveAcesso"]
                xml_base64 = doc.get("ArquivoXml")

                if xml_base64:
                    xml_bytes = gzip.decompress(base64.b64decode(xml_base64))
                    info_nota = analisar_xml(xml_bytes, MEU_CNPJ)

                    if info_nota["ano"] == ano_desejado and info_nota["mes"] == mes_desejado:
                        if info_nota["eh_recebida"]:
                            nome_xml = f"NFSe_{info_nota['numero']}_{chave[:10]}.xml"
                            caminho_xml = os.path.join(pasta_recebidas_xml, nome_xml)

                            with open(caminho_xml, "wb") as f:
                                f.write(xml_bytes)

                            notas_recebidas_planilha.append(info_nota)
                            print(f"[RECEBIDA] Nota {info_nota['numero']} ({info_nota['prest_nome'][:25]}) salva!")

                            if gerar_pdf:
                                gerar_pdf_do_xml(info_nota, chave, pasta_recebidas_pdf)
                        else:
                            print(f"[IGNORADA - EMITIDA] Nota {info_nota['numero']}")
                    else:
                        data_str = f"{info_nota['mes']:02d}/{info_nota['ano']}" if info_nota["ano"] else "Desconhecida"
                        print(f"[PULADA] NSU {nsu_doc} ({data_str})")

                if nsu_doc > nsu_atual:
                    nsu_atual = nsu_doc

            time.sleep(0.5)

        elif resposta.status_code in (204, 404):
            break
        else:
            print(f"\nErro retornado pela API (HTTP {resposta.status_code}):")
            print(resposta.text)
            break

    # Gera a planilha Excel consolidada
    caminho_planilha = gerar_planilha_excel(notas_recebidas_planilha, mes_desejado, ano_desejado, pasta_base_periodo)

    print("\n" + "=" * 60)
    print(f"Busca finalizada para {mes_desejado:02d}/{ano_desejado}!")
    print(f"Notas Recebidas salvas: {len(notas_recebidas_planilha)}")
    print(f"Planilha gerada em:     {caminho_planilha}")
    print(f"Pasta dos arquivos:     {pasta_base_periodo}")
    print("=" * 60)
    input("\nPressione ENTER para fechar...")


if __name__ == "__main__":
    iniciar_automacao()