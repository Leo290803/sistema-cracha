from flask import Flask, request, send_file, jsonify, redirect
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
import io
import base64
import fitz
import unicodedata
import re
import html

app = Flask(__name__)


def h(valor):
    return html.escape(str(valor or ""))


def limpar_texto(txt):
    txt = str(txt).upper().strip()
    txt = unicodedata.normalize("NFD", txt)
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
    txt = re.sub(r"[^A-Z0-9 ]", " ", txt)
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip()


def valor_linha(row, coluna, padrao=""):
    if not coluna:
        return padrao
    valor = row.get(coluna, padrao)
    if pd.isna(valor):
        return padrao
    valor = str(valor).strip()
    if valor.lower() == "nan":
        return padrao
    return valor


def valor_valido(valor):
    texto = limpar_texto(valor)
    return bool(texto and texto not in ["---", "-", "NA", "NAN", "NONE", "NULL"])


def encontrar_coluna(df, nomes_possiveis):
    for col in df.columns:
        col_limpa = limpar_texto(col)
        for nome in nomes_possiveis:
            if limpar_texto(nome) in col_limpa:
                return col
    return None


def normalizar_lista_escolas(valores):
    escolas = []

    if not valores:
        return escolas

    if isinstance(valores, list):
        partes = valores
    else:
        partes = re.split(r"[\n,;]+", str(valores))

    for item in partes:
        escola = limpar_texto(item)
        if escola:
            escolas.append(escola)

    return escolas


def escola_permitida(escola, escolas_escolhidas):
    if not escolas_escolhidas:
        return True

    escola_limpa = limpar_texto(escola)

    for escolhida in escolas_escolhidas:
        if escolhida == escola_limpa:
            return True

    return False


def normalizar_sexo(valor):
    sexo = limpar_texto(valor)

    if sexo in ["M", "MASC", "MASCULINO"]:
        return "MASCULINO"

    if sexo in ["F", "FEM", "FEMININO"]:
        return "FEMININO"

    return None


def calcular_categoria(data_nascimento):
    nascimento = pd.to_datetime(data_nascimento, dayfirst=True, errors="coerce")

    if pd.isna(nascimento):
        return None

    ano = nascimento.year

    if 2009 <= ano <= 2011:
        return "15 A 17 ANOS"

    if 2012 <= ano <= 2014:
        return "12 A 14 ANOS"

    return None


def texto_pagina_fitz(pdf_bytes, pagina_index):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        texto = doc[pagina_index].get_text()
        doc.close()
        return texto or ""
    except Exception:
        return ""


def buscar_linha_por_nome(df, texto_pagina, col_nome):
    texto_pagina_limpo = limpar_texto(texto_pagina)

    melhor_idx = None
    melhor_row = None
    melhor_pontuacao = 0

    for idx, row in df.iterrows():
        nome_excel = limpar_texto(row[col_nome])

        if not nome_excel:
            continue

        partes = nome_excel.split()

        if len(partes) < 2:
            continue

        acertos = sum(1 for parte in partes if parte in texto_pagina_limpo)
        pontuacao = acertos / len(partes)

        primeiro_nome_ok = partes[0] in texto_pagina_limpo
        segundo_nome_ok = partes[1] in texto_pagina_limpo

        if primeiro_nome_ok and segundo_nome_ok and pontuacao > melhor_pontuacao:
            melhor_pontuacao = pontuacao
            melhor_idx = idx
            melhor_row = row

    if melhor_pontuacao >= 0.50:
        return melhor_idx, melhor_row

    return None, None


def extrair_numero_credencial(texto_pagina):
    candidatos = re.findall(r"\b\d{6,20}\b", str(texto_pagina))
    if candidatos:
        return candidatos[0]
    return ""


def identificar_tipo_pessoa(row, col_funcao=None, col_tipo_usuario=None, texto_pagina=""):
    funcao = valor_linha(row, col_funcao, "")
    tipo_usuario = valor_linha(row, col_tipo_usuario, "")
    texto_pdf = limpar_texto(texto_pagina)

    if valor_valido(funcao):
        return limpar_texto(funcao)

    if "CHEFE DE DELEGACAO" in texto_pdf:
        return "CHEFE DE DELEGAÇÃO"

    if "TECNICO" in texto_pdf:
        return "TÉCNICO"

    if "OFICIAL" in texto_pdf:
        return "OFICIAL"

    if "PRESTADOR" in limpar_texto(tipo_usuario):
        return "DIRIGENTE"

    return "ATLETA"


def montar_categoria(row, col_nome, col_sexo, col_data, linha_excel):
    nome = str(row[col_nome]).strip()
    sexo = normalizar_sexo(row[col_sexo])
    categoria = calcular_categoria(row[col_data])

    if not sexo:
        raise Exception(f"Linha {linha_excel}: sexo inválido para {nome}")

    if not categoria:
        raise Exception(f"Linha {linha_excel}: data de nascimento inválida ou fora da categoria para {nome}")

    return f"{categoria} {sexo}"


def montar_texto_cracha(
    row,
    col_nome,
    col_sexo,
    col_data,
    col_escola,
    col_funcao,
    col_tipo_usuario,
    linha_excel,
    texto_pagina="",
    texto_atleta="categoria"
):
    tipo_pessoa = identificar_tipo_pessoa(row, col_funcao, col_tipo_usuario, texto_pagina)
    escola = valor_linha(row, col_escola, "")

    if tipo_pessoa != "ATLETA":
        if not valor_valido(escola):
            nome = str(row[col_nome]).strip()
            raise Exception(f"Linha {linha_excel}: escola inválida para {nome}")
        return escola.upper(), tipo_pessoa, escola

    if texto_atleta == "escola":
        if not valor_valido(escola):
            nome = str(row[col_nome]).strip()
            raise Exception(f"Linha {linha_excel}: escola inválida para {nome}")
        return escola.upper(), tipo_pessoa, escola

    if not col_sexo:
        raise Exception("Não encontrei a coluna SEXO na planilha.")

    if not col_data:
        raise Exception("Não encontrei a coluna DATA NASCIMENTO na planilha.")

    categoria = montar_categoria(row, col_nome, col_sexo, col_data, linha_excel)
    return categoria, tipo_pessoa, escola


def carregar_planilha(excel_file):
    df = pd.read_excel(excel_file, dtype=str)
    df.columns = df.columns.str.strip().str.upper()

    col_nome = encontrar_coluna(df, ["NOME"])
    col_sexo = encontrar_coluna(df, ["SEXO"])
    col_data = encontrar_coluna(df, ["DATA NASCIMENTO", "DATA DE NASCIMENTO", "NASCIMENTO", "DATA"])
    col_escola = encontrar_coluna(df, ["ESCOLA", "UNIDADE ESCOLAR", "INSTITUICAO", "INSTITUIÇÃO"])
    col_funcao = encontrar_coluna(df, ["FUNCAO", "FUNÇÃO", "CARGO"])
    col_tipo_usuario = encontrar_coluna(df, ["TIPO USUARIO", "TIPO USUÁRIO", "TIPO"])
    col_cpf = encontrar_coluna(df, ["CPF"])

    if not col_nome:
        raise Exception("Não encontrei a coluna NOME na planilha.")

    if not col_escola:
        raise Exception("Não encontrei a coluna ESCOLA na planilha.")

    return {
        "df": df,
        "col_nome": col_nome,
        "col_sexo": col_sexo,
        "col_data": col_data,
        "col_escola": col_escola,
        "col_funcao": col_funcao,
        "col_tipo_usuario": col_tipo_usuario,
        "col_cpf": col_cpf
    }


def analisar_escolas_pdf(excel_file, pdf_file):
    dados = carregar_planilha(excel_file)
    df = dados["df"]
    col_nome = dados["col_nome"]
    col_escola = dados["col_escola"]

    pdf_bytes = pdf_file.read()
    reader = PdfReader(io.BytesIO(pdf_bytes))

    escolas = {}
    nao_encontrados = []

    for i, page in enumerate(reader.pages):
        texto_pypdf = page.extract_text() or ""
        texto_fitz = texto_pagina_fitz(pdf_bytes, i)
        texto_pagina = texto_pypdf + "\n" + texto_fitz

        idx_excel, row = buscar_linha_por_nome(df, texto_pagina, col_nome)

        if row is None:
            nao_encontrados.append(i + 1)
            continue

        escola_original = valor_linha(row, col_escola, "SEM ESCOLA")
        escola_limpa = limpar_texto(escola_original)

        if not escola_limpa:
            escola_limpa = "SEM ESCOLA"
            escola_original = "SEM ESCOLA"

        if escola_limpa not in escolas:
            escolas[escola_limpa] = {
                "nome": escola_original.upper(),
                "valor": escola_limpa,
                "quantidade": 0,
                "paginas": []
            }

        escolas[escola_limpa]["quantidade"] += 1
        escolas[escola_limpa]["paginas"].append(i + 1)

    lista = sorted(escolas.values(), key=lambda x: x["nome"])

    return {
        "escolas": lista,
        "total_escolas": len(lista),
        "total_crachas_identificados": sum(e["quantidade"] for e in lista),
        "nao_encontrados": nao_encontrados
    }


def criar_pdf(
    excel_file,
    pdf_file,
    pos_x,
    pos_y,
    rotacao,
    fonte,
    somente_primeira_pagina=False,
    texto_atleta="categoria",
    escolas_selecionadas=None
):
    dados = carregar_planilha(excel_file)

    df = dados["df"]
    col_nome = dados["col_nome"]
    col_sexo = dados["col_sexo"]
    col_data = dados["col_data"]
    col_escola = dados["col_escola"]
    col_funcao = dados["col_funcao"]
    col_tipo_usuario = dados["col_tipo_usuario"]

    escolas_escolhidas = normalizar_lista_escolas(escolas_selecionadas)

    pdf_bytes = pdf_file.read()
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()

    erros = []
    paginas_adicionadas = 0

    for i in range(len(reader.pages)):
        page = reader.pages[i]

        texto_pypdf = page.extract_text() or ""
        texto_fitz = texto_pagina_fitz(pdf_bytes, i)
        texto_pagina = texto_pypdf + "\n" + texto_fitz

        idx_excel, row = buscar_linha_por_nome(df, texto_pagina, col_nome)

        if row is None:
            erros.append(f"Página {i + 1}: não encontrei a pessoa do crachá na planilha.")
            continue

        escola = valor_linha(row, col_escola, "")

        if not escola_permitida(escola, escolas_escolhidas):
            continue

        linha_excel = idx_excel + 2

        try:
            texto_cracha, tipo_pessoa, escola = montar_texto_cracha(
                row=row,
                col_nome=col_nome,
                col_sexo=col_sexo,
                col_data=col_data,
                col_escola=col_escola,
                col_funcao=col_funcao,
                col_tipo_usuario=col_tipo_usuario,
                linha_excel=linha_excel,
                texto_pagina=texto_pagina,
                texto_atleta=texto_atleta
            )
        except Exception as e:
            erros.append(str(e))
            continue

        packet = io.BytesIO()
        largura = float(page.mediabox.width)
        altura = float(page.mediabox.height)

        can = canvas.Canvas(packet, pagesize=(largura, altura))
        can.setFont("Helvetica-Bold", fonte)
        can.setFillColorRGB(1, 1, 1)

        can.saveState()
        can.translate(pos_x, pos_y)
        can.rotate(rotacao)
        can.drawCentredString(0, 0, texto_cracha)
        can.restoreState()

        can.save()
        packet.seek(0)

        overlay = PdfReader(packet)

        if len(overlay.pages) > 0:
            page.merge_page(overlay.pages[0])

        writer.add_page(page)
        paginas_adicionadas += 1

        if somente_primeira_pagina:
            break

    if paginas_adicionadas == 0:
        if escolas_escolhidas:
            raise Exception("Nenhum crachá encontrado para as escolas selecionadas.")
        raise Exception("Nenhuma página foi gerada.")

    output = io.BytesIO()
    writer.write(output)
    output.seek(0)

    return output


def criar_pdf_manual(pdf_file, texto_manual, pos_x, pos_y, rotacao, fonte, somente_primeira_pagina=False):
    texto_manual = str(texto_manual or "").strip()

    if not texto_manual:
        raise Exception("Digite o nome da escola ou texto que deseja colocar no crachá.")

    pdf_bytes = pdf_file.read()
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()

    paginas_adicionadas = 0

    for i in range(len(reader.pages)):
        page = reader.pages[i]

        packet = io.BytesIO()
        largura = float(page.mediabox.width)
        altura = float(page.mediabox.height)

        can = canvas.Canvas(packet, pagesize=(largura, altura))
        can.setFont("Helvetica-Bold", fonte)
        can.setFillColorRGB(1, 1, 1)

        can.saveState()
        can.translate(pos_x, pos_y)
        can.rotate(rotacao)
        can.drawCentredString(0, 0, texto_manual.upper())
        can.restoreState()

        can.save()
        packet.seek(0)

        overlay = PdfReader(packet)

        if len(overlay.pages) > 0:
            page.merge_page(overlay.pages[0])

        writer.add_page(page)
        paginas_adicionadas += 1

        if somente_primeira_pagina:
            break

    if paginas_adicionadas == 0:
        raise Exception("Nenhuma página foi gerada.")

    output = io.BytesIO()
    writer.write(output)
    output.seek(0)

    return output



def criar_pdf_preview_rapido(
    excel_file,
    pdf_file,
    pos_x,
    pos_y,
    rotacao,
    fonte,
    texto_atleta="categoria",
    escolas_selecionadas=None
):
    """
    Prévia leve para Render Free.
    Não varre o PDF inteiro e não fica comparando todas as linhas da planilha.
    Usa somente a primeira página do PDF e uma linha da planilha.
    Se tiver escola marcada, pega a primeira linha daquela escola.
    """
    dados = carregar_planilha(excel_file)

    df = dados["df"]
    col_nome = dados["col_nome"]
    col_sexo = dados["col_sexo"]
    col_data = dados["col_data"]
    col_escola = dados["col_escola"]
    col_funcao = dados["col_funcao"]
    col_tipo_usuario = dados["col_tipo_usuario"]

    escolas_escolhidas = normalizar_lista_escolas(escolas_selecionadas)

    if df.empty:
        raise Exception("A planilha está vazia.")

    row = None
    idx_excel = 0

    if escolas_escolhidas:
        for idx in range(len(df)):
            r = df.iloc[idx]
            escola = valor_linha(r, col_escola, "")
            if escola_permitida(escola, escolas_escolhidas):
                row = r
                idx_excel = idx
                break

        if row is None:
            raise Exception("Não encontrei pessoa na planilha para a escola selecionada.")
    else:
        row = df.iloc[0]
        idx_excel = 0

    pdf_bytes = pdf_file.read()
    reader = PdfReader(io.BytesIO(pdf_bytes))

    if len(reader.pages) == 0:
        raise Exception("O PDF não possui páginas.")

    page = reader.pages[0]
    linha_excel = idx_excel + 2

    texto_cracha, tipo_pessoa, escola = montar_texto_cracha(
        row=row,
        col_nome=col_nome,
        col_sexo=col_sexo,
        col_data=col_data,
        col_escola=col_escola,
        col_funcao=col_funcao,
        col_tipo_usuario=col_tipo_usuario,
        linha_excel=linha_excel,
        texto_pagina="",
        texto_atleta=texto_atleta
    )

    packet = io.BytesIO()
    largura = float(page.mediabox.width)
    altura = float(page.mediabox.height)

    can = canvas.Canvas(packet, pagesize=(largura, altura))
    can.setFont("Helvetica-Bold", fonte)
    can.setFillColorRGB(1, 1, 1)

    can.saveState()
    can.translate(pos_x, pos_y)
    can.rotate(rotacao)
    can.drawCentredString(0, 0, texto_cracha)
    can.restoreState()

    can.save()
    packet.seek(0)

    overlay = PdfReader(packet)

    if len(overlay.pages) > 0:
        page.merge_page(overlay.pages[0])

    writer = PdfWriter()
    writer.add_page(page)

    output = io.BytesIO()
    writer.write(output)
    output.seek(0)

    return output


def layout_home():
    return """
    <!doctype html>
    <html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Gerador de Crachás</title>
        <style>
            * { box-sizing: border-box; }

            body {
                margin: 0;
                font-family: Arial, sans-serif;
                background: #f3f4f6;
                color: #111827;
            }

            .top {
                background: linear-gradient(135deg, #0f172a, #166534);
                color: white;
                padding: 24px;
                text-align: center;
            }

            .top h1 {
                margin: 0;
                font-size: 32px;
            }

            .top p {
                margin: 8px 0 0;
                opacity: .9;
            }

            .wrap {
                max-width: 1250px;
                margin: 30px auto;
                padding: 0 16px;
                display: grid;
                grid-template-columns: 500px 1fr;
                gap: 20px;
            }

            .card {
                background: white;
                border-radius: 18px;
                padding: 24px;
                box-shadow: 0 10px 30px rgba(0,0,0,.08);
                border: 1px solid #e5e7eb;
            }

            label {
                font-weight: bold;
                display: block;
                margin-top: 16px;
                margin-bottom: 6px;
            }

            input, select {
                width: 100%;
                padding: 11px;
                border: 1px solid #cbd5e1;
                border-radius: 10px;
                font-size: 15px;
            }

            .grid {
                display: grid;
                grid-template-columns: repeat(2, 1fr);
                gap: 10px;
            }

            button {
                width: 100%;
                padding: 13px;
                border: 0;
                border-radius: 12px;
                background: #16a34a;
                color: white;
                font-weight: bold;
                font-size: 15px;
                cursor: pointer;
                margin-top: 12px;
            }

            button.secondary { background: #2563eb; }
            button.dark { background: #0f172a; }
            button.orange { background: #f97316; }
            button.purple { background: #7c3aed; }

            .help {
                background: #ecfdf5;
                border: 1px solid #bbf7d0;
                border-radius: 14px;
                padding: 14px;
                color: #166534;
                font-size: 14px;
                line-height: 1.5;
            }

            .help.manual {
                background: #f5f3ff;
                border-color: #ddd6fe;
                color: #5b21b6;
            }

            .preview-box {
                min-height: 420px;
                display: flex;
                align-items: center;
                justify-content: center;
                background: #f8fafc;
                border: 2px dashed #cbd5e1;
                border-radius: 16px;
                overflow: auto;
                padding: 15px;
            }

            .preview-box img {
                max-width: 100%;
                border-radius: 12px;
                box-shadow: 0 8px 25px rgba(0,0,0,.15);
            }

            .erro {
                white-space: pre-wrap;
                background: #fee2e2;
                color: #991b1b;
                padding: 14px;
                border-radius: 12px;
                font-weight: bold;
            }

            .ok {
                white-space: pre-wrap;
                background: #dcfce7;
                color: #166534;
                padding: 14px;
                border-radius: 12px;
                font-weight: bold;
            }

            .escolas-box {
                margin-top: 18px;
                border: 1px solid #e5e7eb;
                border-radius: 14px;
                padding: 14px;
                background: #f8fafc;
                max-height: 430px;
                overflow: auto;
            }

            .escola-item {
                display: flex;
                align-items: flex-start;
                gap: 10px;
                background: white;
                border: 1px solid #e5e7eb;
                border-radius: 12px;
                padding: 10px;
                margin-bottom: 8px;
            }

            .escola-item input {
                width: auto;
                margin-top: 3px;
            }

            .escola-nome {
                font-weight: bold;
                font-size: 14px;
            }

            .escola-qtd {
                color: #64748b;
                font-size: 13px;
                margin-top: 3px;
            }

            .linha-botoes {
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 8px;
            }

            .mini {
                font-size: 13px;
                padding: 10px;
                margin-top: 8px;
            }

            hr {
                border: none;
                border-top: 1px solid #e5e7eb;
                margin: 28px 0;
            }

            @media (max-width: 900px) {
                .wrap { grid-template-columns: 1fr; }
            }
        </style>
    </head>
    <body>
        <div class="top">
            <h1>Gerador de Crachás</h1>
            <p>Filtro por escola + modo manual para técnico/dirigente</p>
        </div>

        <div class="wrap">
            <div class="card">
                <h2>Modo Normal: Excel + PDF</h2>

                <div class="help">
                    1. Envie a planilha Excel e o PDF.<br>
                    2. Clique em <b>Ler escolas do PDF</b>.<br>
                    3. Marque as escolas desejadas.<br>
                    4. Gere o PDF final somente com essas escolas.
                </div>

                <form id="formulario" action="/gerar" method="POST" enctype="multipart/form-data">
                    <label>Planilha Excel</label>
                    <input type="file" name="excel" accept=".xlsx,.xls">

                    <label>PDF dos crachás</label>
                    <input type="file" name="pdf" accept=".pdf">

                    <label>Texto para ATLETA</label>
                    <select name="texto_atleta">
                        <option value="categoria" selected>Categoria</option>
                        <option value="escola">Escola</option>
                    </select>

                    <div class="grid">
                        <div>
                            <label>Posição X</label>
                            <input type="number" name="pos_x" value="118">
                        </div>

                        <div>
                            <label>Posição Y</label>
                            <input type="number" name="pos_y" value="300">
                        </div>

                        <div>
                            <label>Rotação</label>
                            <input type="number" name="rotacao" value="90">
                        </div>

                        <div>
                            <label>Tamanho da fonte</label>
                            <input type="number" name="fonte" value="10">
                        </div>
                    </div>

                    <button type="button" class="dark" onclick="analisarEscolas()">Ler escolas do PDF</button>

                    <div id="areaEscolas" class="escolas-box" style="display:none;"></div>

                    <button type="button" class="secondary" onclick="preview()">Baixar PDF teste da primeira página</button>
                    <button type="submit">Gerar PDF Final</button>
                    <button type="button" class="orange" onclick="exportarCredenciais()">Exportar Excel CPF + Credencial</button>
                </form>

                <hr>

                <h2>Modo Técnico / Manual: só PDF</h2>

                <div class="help manual">
                    Use quando você tiver somente o PDF, sem Excel.
                    Digite o nome da escola ou qualquer texto, e o sistema coloca em todos os crachás.
                </div>

                <form id="formManual" action="/gerar-manual" method="POST" enctype="multipart/form-data">
                    <label>PDF dos crachás</label>
                    <input type="file" name="pdf_manual" accept=".pdf" required>

                    <label>Nome da escola ou texto manual</label>
                    <input type="text" name="texto_manual" placeholder="Ex: ESCOLA ESTADUAL TESTE" required>

                    <div class="grid">
                        <div>
                            <label>Posição X</label>
                            <input type="number" name="pos_x_manual" value="118">
                        </div>

                        <div>
                            <label>Posição Y</label>
                            <input type="number" name="pos_y_manual" value="300">
                        </div>

                        <div>
                            <label>Rotação</label>
                            <input type="number" name="rotacao_manual" value="90">
                        </div>

                        <div>
                            <label>Tamanho da fonte</label>
                            <input type="number" name="fonte_manual" value="10">
                        </div>
                    </div>

                    <button type="button" class="purple" onclick="previewManual()">Baixar PDF teste manual</button>
                    <button type="submit" class="orange">Gerar PDF Manual</button>
                </form>
            </div>

            <div class="card">
                <h2>Teste</h2>
                <div id="preview" class="preview-box">
                    Clique no botão de teste para baixar somente a primeira página em PDF.
                </div>
            </div>
        </div>

        <script>
            function getForm() {
                return document.getElementById("formulario");
            }

            function getPreview() {
                return document.getElementById("preview");
            }

            function marcarTodas(valor) {
                document.querySelectorAll(".check-escola").forEach(cb => cb.checked = valor);
            }

            async function analisarEscolas() {
                const form = getForm();
                const dados = new FormData(form);
                const area = document.getElementById("areaEscolas");

                area.style.display = "block";
                area.innerHTML = '<div class="ok">Lendo PDF e identificando escolas...</div>';

                try {
                    const resp = await fetch("/analisar-escolas", {
                        method: "POST",
                        body: dados
                    });

                    const data = await resp.json();

                    if (!data.ok) {
                        area.innerHTML = '<div class="erro">' + data.erro + '</div>';
                        return;
                    }

                    let html = '';
                    html += '<h3>Escolas encontradas</h3>';
                    html += '<div class="ok">';
                    html += 'Total de escolas: ' + data.total_escolas + '\\n';
                    html += 'Crachás identificados: ' + data.total_crachas_identificados;
                    if (data.nao_encontrados.length > 0) {
                        html += '\\nPáginas sem localizar na planilha: ' + data.nao_encontrados.join(", ");
                    }
                    html += '</div>';

                    html += '<div class="linha-botoes">';
                    html += '<button type="button" class="mini secondary" onclick="marcarTodas(true)">Selecionar todas</button>';
                    html += '<button type="button" class="mini orange" onclick="marcarTodas(false)">Desmarcar todas</button>';
                    html += '</div>';

                    data.escolas.forEach(function(escola) {
                        html += '<label class="escola-item">';
                        html += '<input class="check-escola" type="checkbox" name="escolas_selecionadas" value="' + escola.valor + '">';
                        html += '<div>';
                        html += '<div class="escola-nome">' + escola.nome + '</div>';
                        html += '<div class="escola-qtd">' + escola.quantidade + ' crachá(s)</div>';
                        html += '</div>';
                        html += '</label>';
                    });

                    area.innerHTML = html;

                } catch (e) {
                    area.innerHTML = '<div class="erro">Erro ao analisar escolas.</div>';
                }
            }

            function preview() {
                const form = getForm();
                const actionOriginal = form.action;
                const targetOriginal = form.target;

                getPreview().innerHTML = '<div class="ok">Baixando PDF teste da primeira página...</div>';

                form.action = "/teste-primeira-pagina";
                form.target = "_blank";
                form.submit();

                form.action = actionOriginal || "/gerar";
                form.target = targetOriginal || "";
            }

            function previewManual() {
                const form = document.getElementById("formManual");
                const actionOriginal = form.action;
                const targetOriginal = form.target;

                getPreview().innerHTML = '<div class="ok">Baixando PDF teste manual da primeira página...</div>';

                form.action = "/teste-manual-primeira-pagina";
                form.target = "_blank";
                form.submit();

                form.action = actionOriginal || "/gerar-manual";
                form.target = targetOriginal || "";
            }

            function exportarCredenciais() {
                const form = getForm();
                form.action = "/gerar-excel-credenciais";
                form.submit();
                form.action = "/gerar";
            }
        </script>
    </body>
    </html>
    """


@app.route("/")
def index():
    return layout_home()


@app.route("/analisar-escolas", methods=["POST"])
def analisar_escolas():
    try:
        excel = request.files["excel"]
        pdf = request.files["pdf"]

        resultado = analisar_escolas_pdf(excel, pdf)

        return jsonify({
            "ok": True,
            "escolas": resultado["escolas"],
            "total_escolas": resultado["total_escolas"],
            "total_crachas_identificados": resultado["total_crachas_identificados"],
            "nao_encontrados": resultado["nao_encontrados"]
        })

    except Exception as e:
        return jsonify({
            "ok": False,
            "erro": str(e)
        })


@app.route("/teste-primeira-pagina", methods=["POST"])
def teste_primeira_pagina():
    try:
        excel = request.files["excel"]
        pdf = request.files["pdf"]

        if not excel or excel.filename == "":
            raise Exception("Selecione a planilha Excel.")

        if not pdf or pdf.filename == "":
            raise Exception("Selecione o PDF dos crachás.")

        pos_x = int(request.form.get("pos_x", 118))
        pos_y = int(request.form.get("pos_y", 300))
        rotacao = int(request.form.get("rotacao", 90))
        fonte = int(request.form.get("fonte", 10))
        texto_atleta = request.form.get("texto_atleta", "categoria")
        escolas_selecionadas = request.form.getlist("escolas_selecionadas")

        output = criar_pdf_preview_rapido(
            excel_file=excel,
            pdf_file=pdf,
            pos_x=pos_x,
            pos_y=pos_y,
            rotacao=rotacao,
            fonte=fonte,
            texto_atleta=texto_atleta,
            escolas_selecionadas=escolas_selecionadas
        )

        return send_file(
            output,
            as_attachment=True,
            download_name="TESTE_primeira_pagina.pdf",
            mimetype="application/pdf"
        )

    except Exception as e:
        return f"""
        <h1>Erro ao gerar PDF teste</h1>
        <pre>{h(e)}</pre>
        <a href="/">Voltar</a>
        """


@app.route("/teste-manual-primeira-pagina", methods=["POST"])
def teste_manual_primeira_pagina():
    try:
        pdf = request.files["pdf_manual"]

        if not pdf or pdf.filename == "":
            raise Exception("Selecione o PDF dos crachás.")

        texto_manual = request.form.get("texto_manual", "")
        pos_x = int(request.form.get("pos_x_manual", 118))
        pos_y = int(request.form.get("pos_y_manual", 300))
        rotacao = int(request.form.get("rotacao_manual", 90))
        fonte = int(request.form.get("fonte_manual", 10))

        output = criar_pdf_manual(
            pdf_file=pdf,
            texto_manual=texto_manual,
            pos_x=pos_x,
            pos_y=pos_y,
            rotacao=rotacao,
            fonte=fonte,
            somente_primeira_pagina=True
        )

        return send_file(
            output,
            as_attachment=True,
            download_name="TESTE_manual_primeira_pagina.pdf",
            mimetype="application/pdf"
        )

    except Exception as e:
        return f"""
        <h1>Erro ao gerar PDF teste manual</h1>
        <pre>{h(e)}</pre>
        <a href="/">Voltar</a>
        """


@app.route("/gerar", methods=["POST"])
def gerar():
    try:
        excel = request.files["excel"]
        pdf = request.files["pdf"]

        pos_x = int(request.form.get("pos_x", 118))
        pos_y = int(request.form.get("pos_y", 300))
        rotacao = int(request.form.get("rotacao", 90))
        fonte = int(request.form.get("fonte", 10))
        texto_atleta = request.form.get("texto_atleta", "categoria")
        escolas_selecionadas = request.form.getlist("escolas_selecionadas")

        output = criar_pdf(
            excel_file=excel,
            pdf_file=pdf,
            pos_x=pos_x,
            pos_y=pos_y,
            rotacao=rotacao,
            fonte=fonte,
            somente_primeira_pagina=False,
            texto_atleta=texto_atleta,
            escolas_selecionadas=escolas_selecionadas
        )

        return send_file(
            output,
            as_attachment=True,
            download_name="crachas_filtrados_por_escola.pdf",
            mimetype="application/pdf"
        )

    except Exception as e:
        return f"""
        <h1>Erro ao gerar PDF</h1>
        <pre>{h(e)}</pre>
        <a href="/">Voltar</a>
        """


@app.route("/gerar-manual", methods=["POST"])
def gerar_manual():
    try:
        pdf = request.files["pdf_manual"]

        texto_manual = request.form.get("texto_manual", "")
        pos_x = int(request.form.get("pos_x_manual", 118))
        pos_y = int(request.form.get("pos_y_manual", 300))
        rotacao = int(request.form.get("rotacao_manual", 90))
        fonte = int(request.form.get("fonte_manual", 10))

        output = criar_pdf_manual(
            pdf_file=pdf,
            texto_manual=texto_manual,
            pos_x=pos_x,
            pos_y=pos_y,
            rotacao=rotacao,
            fonte=fonte,
            somente_primeira_pagina=False
        )

        return send_file(
            output,
            as_attachment=True,
            download_name="crachas_texto_manual.pdf",
            mimetype="application/pdf"
        )

    except Exception as e:
        return f"""
        <h1>Erro ao gerar PDF manual</h1>
        <pre>{h(e)}</pre>
        <a href="/">Voltar</a>
        """


@app.route("/gerar-excel-credenciais", methods=["GET", "POST"])
def gerar_excel_credenciais():
    if request.method == "GET":
        return redirect("/")

    try:
        excel = request.files["excel"]
        pdf = request.files["pdf"]
        escolas_selecionadas = request.form.getlist("escolas_selecionadas")
        escolas_escolhidas = normalizar_lista_escolas(escolas_selecionadas)

        dados = carregar_planilha(excel)
        df = dados["df"]
        col_nome = dados["col_nome"]
        col_cpf = dados["col_cpf"]
        col_escola = dados["col_escola"]

        if not col_cpf:
            raise Exception("Não encontrei a coluna CPF na planilha.")

        pdf_bytes = pdf.read()
        reader = PdfReader(io.BytesIO(pdf_bytes))

        registros = []

        for i, page in enumerate(reader.pages):
            texto_pypdf = page.extract_text() or ""
            texto_fitz = texto_pagina_fitz(pdf_bytes, i)
            texto_pagina = texto_pypdf + "\n" + texto_fitz

            numero_credencial = extrair_numero_credencial(texto_pagina)

            cpf = ""
            escola = ""
            nome = ""

            idx_excel, row = buscar_linha_por_nome(df, texto_pagina, col_nome)

            if row is not None:
                cpf = valor_linha(row, col_cpf, "")
                escola = valor_linha(row, col_escola, "")
                nome = valor_linha(row, col_nome, "")

                if not escola_permitida(escola, escolas_escolhidas):
                    continue

            registros.append({
                "NOME": nome,
                "ESCOLA": escola,
                "CPF": cpf,
                "NUMERO_CREDENCIAL": numero_credencial
            })

        if not registros:
            raise Exception("Nenhuma credencial encontrada para as escolas selecionadas.")

        df_saida = pd.DataFrame(registros, columns=[
            "NOME",
            "ESCOLA",
            "CPF",
            "NUMERO_CREDENCIAL"
        ])

        output = io.BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_saida.to_excel(writer, index=False, sheet_name="Credenciais")

            ws = writer.sheets["Credenciais"]
            ws.column_dimensions["A"].width = 35
            ws.column_dimensions["B"].width = 45
            ws.column_dimensions["C"].width = 22
            ws.column_dimensions["D"].width = 25

        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="credenciais_filtradas_por_escola.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return f"""
        <h1>Erro ao gerar Excel de credenciais</h1>
        <pre>{h(e)}</pre>
        <a href="/">Voltar</a>
        """


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
