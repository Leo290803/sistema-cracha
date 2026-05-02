from flask import Flask, render_template, request, send_file, jsonify, redirect
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
import io
import base64
import fitz
import unicodedata
import re
import sqlite3
from datetime import datetime

app = Flask(__name__)

DB = "validacao_crachas.db"


# =========================
# BANCO DE DADOS
# =========================

def init_db():
    conn = sqlite3.connect(DB)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS atletas (
            codigo_qr TEXT PRIMARY KEY,
            nome TEXT,
            categoria TEXT,
            pagina INTEGER,
            escola TEXT,
            status TEXT DEFAULT 'PENDENTE',
            checkin_hora TEXT
        )
    """)

    conn.commit()
    conn.close()


def salvar_atleta(codigo_qr, nome, categoria, pagina, escola=""):
    if not codigo_qr:
        return

    conn = sqlite3.connect(DB)
    cur = conn.cursor()

    cur.execute("""
        INSERT OR REPLACE INTO atletas
        (codigo_qr, nome, categoria, pagina, escola, status, checkin_hora)
        VALUES (
            ?,
            ?,
            ?,
            ?,
            ?,
            COALESCE((SELECT status FROM atletas WHERE codigo_qr = ?), 'PENDENTE'),
            COALESCE((SELECT checkin_hora FROM atletas WHERE codigo_qr = ?), NULL)
        )
    """, (codigo_qr, nome, categoria, pagina, escola, codigo_qr, codigo_qr))

    conn.commit()
    conn.close()


def buscar_atleta_por_codigo(codigo_qr):
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("SELECT * FROM atletas WHERE codigo_qr = ?", (codigo_qr,))
    atleta = cur.fetchone()

    conn.close()
    return atleta


def registrar_checkin(codigo_qr):
    atleta = buscar_atleta_por_codigo(codigo_qr)

    if not atleta:
        return False, "ATLETA NÃO ENCONTRADO"

    if atleta["status"] == "ENTROU":
        return False, "JÁ ENTROU"

    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    conn = sqlite3.connect(DB)
    cur = conn.cursor()

    cur.execute("""
        UPDATE atletas
        SET status = 'ENTROU', checkin_hora = ?
        WHERE codigo_qr = ?
    """, (agora, codigo_qr))

    conn.commit()
    conn.close()

    return True, "CHECK-IN REALIZADO"


# =========================
# FUNÇÕES DE TEXTO
# =========================

def limpar_texto(txt):
    txt = str(txt).upper().strip()
    txt = unicodedata.normalize("NFD", txt)
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
    txt = re.sub(r"[^A-Z0-9 ]", " ", txt)
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip()


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


def encontrar_coluna(df, nomes_possiveis):
    for col in df.columns:
        col_limpa = limpar_texto(col)
        for nome in nomes_possiveis:
            if limpar_texto(nome) in col_limpa:
                return col
    return None


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

        acertos = 0

        for parte in partes:
            if parte in texto_pagina_limpo:
                acertos += 1

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


def extrair_codigo_qr_do_texto(texto_pagina):
    texto = limpar_texto(texto_pagina)
    partes = texto.split()

    candidatos = []

    for parte in partes:
        if parte.isdigit() and len(parte) >= 6:
            candidatos.append(parte)

    if candidatos:
        return candidatos[0]

    return None


def montar_categoria(row, col_nome, col_sexo, col_data, linha_excel):
    nome = str(row[col_nome]).strip()

    sexo = normalizar_sexo(row[col_sexo])
    categoria = calcular_categoria(row[col_data])

    if not sexo:
        raise Exception(f"Linha {linha_excel}: sexo inválido para {nome}")

    if not categoria:
        raise Exception(
            f"Linha {linha_excel}: data de nascimento inválida ou fora da categoria para {nome}"
        )

    return f"{categoria} {sexo}"


# =========================
# GERAR PDF E BASE
# =========================

def criar_pdf(excel_file, pdf_file, pos_x, pos_y, rotacao, fonte, somente_primeira_pagina=False):
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip().str.upper()

    col_nome = encontrar_coluna(df, ["NOME"])
    col_sexo = encontrar_coluna(df, ["SEXO"])
    col_data = encontrar_coluna(
        df,
        ["DATA NASCIMENTO", "DATA DE NASCIMENTO", "NASCIMENTO", "DATA"]
    )
    col_escola = encontrar_coluna(df, ["ESCOLA"])

    if not col_nome:
        raise Exception("Não encontrei a coluna NOME na planilha.")

    if not col_sexo:
        raise Exception("Não encontrei a coluna SEXO na planilha.")

    if not col_data:
        raise Exception("Não encontrei a coluna DATA NASCIMENTO na planilha.")

    reader = PdfReader(pdf_file)
    writer = PdfWriter()

    total_paginas = 1 if somente_primeira_pagina else len(reader.pages)

    erros = []

    for i in range(total_paginas):
        page = reader.pages[i]

        texto_pagina = page.extract_text() or ""
        codigo_qr = extrair_codigo_qr_do_texto(texto_pagina)

        idx_excel, row = buscar_linha_por_nome(df, texto_pagina, col_nome)

        if row is None:
            erros.append(f"Página {i + 1}: não encontrei o atleta do crachá na planilha.")
            continue

        linha_excel = idx_excel + 2

        try:
            categoria_texto = montar_categoria(row, col_nome, col_sexo, col_data, linha_excel)
        except Exception as e:
            erros.append(str(e))
            continue

        nome = str(row[col_nome]).strip()
        escola = str(row[col_escola]).strip() if col_escola else ""

        salvar_atleta(
            codigo_qr=codigo_qr,
            nome=nome,
            categoria=categoria_texto,
            pagina=i + 1,
            escola=escola
        )

        packet = io.BytesIO()
        largura = float(page.mediabox.width)
        altura = float(page.mediabox.height)

        can = canvas.Canvas(packet, pagesize=(largura, altura))
        can.setFont("Helvetica-Bold", fonte)

        can.saveState()
        can.translate(pos_x, pos_y)
        can.rotate(rotacao)
        can.drawCentredString(0, 0, categoria_texto)
        can.restoreState()

        can.save()
        packet.seek(0)

        overlay = PdfReader(packet)

        if len(overlay.pages) > 0:
            page.merge_page(overlay.pages[0])

        writer.add_page(page)

    if erros:
        raise Exception("\n".join(erros[:30]))

    output = io.BytesIO()
    writer.write(output)
    output.seek(0)

    return output


def montar_base_validacao(excel_file, pdf_file):
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip().str.upper()

    col_nome = encontrar_coluna(df, ["NOME"])
    col_sexo = encontrar_coluna(df, ["SEXO"])
    col_data = encontrar_coluna(df, ["DATA NASCIMENTO", "DATA DE NASCIMENTO", "NASCIMENTO", "DATA"])
    col_escola = encontrar_coluna(df, ["ESCOLA"])

    reader = PdfReader(pdf_file)

    base = []
    erros = []

    for i, page in enumerate(reader.pages):
        texto_pagina = page.extract_text() or ""

        codigo_qr = extrair_codigo_qr_do_texto(texto_pagina)
        idx_excel, row = buscar_linha_por_nome(df, texto_pagina, col_nome)

        if row is None:
            erros.append({
                "pagina": i + 1,
                "codigo_qr": codigo_qr,
                "erro": "Atleta não encontrado na planilha"
            })
            continue

        linha_excel = idx_excel + 2

        try:
            categoria = montar_categoria(row, col_nome, col_sexo, col_data, linha_excel)
        except Exception as e:
            erros.append({
                "pagina": i + 1,
                "codigo_qr": codigo_qr,
                "erro": str(e)
            })
            continue

        nome = str(row[col_nome]).strip()
        escola = str(row[col_escola]).strip() if col_escola else ""

        salvar_atleta(
            codigo_qr=codigo_qr,
            nome=nome,
            categoria=categoria,
            pagina=i + 1,
            escola=escola
        )

        base.append({
            "pagina": i + 1,
            "codigo_qr": codigo_qr,
            "nome": nome,
            "categoria": categoria,
            "escola": escola
        })

    return base, erros


# =========================
# ROTAS PRINCIPAIS
# =========================

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/preview", methods=["POST"])
def preview():
    try:
        excel = request.files["excel"]
        pdf = request.files["pdf"]

        pos_x = int(request.form.get("pos_x", 118))
        pos_y = int(request.form.get("pos_y", 300))
        rotacao = int(request.form.get("rotacao", 90))
        fonte = int(request.form.get("fonte", 10))

        pdf_preview = criar_pdf(
            excel,
            pdf,
            pos_x,
            pos_y,
            rotacao,
            fonte,
            somente_primeira_pagina=True
        )

        doc = fitz.open(stream=pdf_preview.getvalue(), filetype="pdf")
        page = doc[0]

        zoom = 1.2
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)

        img_bytes = pix.tobytes("png")
        img_base64 = base64.b64encode(img_bytes).decode("utf-8")

        doc.close()

        return jsonify({
            "ok": True,
            "imagem": f"data:image/png;base64,{img_base64}"
        })

    except Exception as e:
        return jsonify({
            "ok": False,
            "erro": str(e)
        })


@app.route("/gerar", methods=["POST"])
def gerar():
    try:
        excel = request.files["excel"]
        pdf = request.files["pdf"]

        pos_x = int(request.form.get("pos_x", 118))
        pos_y = int(request.form.get("pos_y", 300))
        rotacao = int(request.form.get("rotacao", 90))
        fonte = int(request.form.get("fonte", 10))

        output = criar_pdf(
            excel,
            pdf,
            pos_x,
            pos_y,
            rotacao,
            fonte,
            somente_primeira_pagina=False
        )

        return send_file(
            output,
            as_attachment=True,
            download_name="crachas_final.pdf",
            mimetype="application/pdf"
        )

    except Exception as e:
        return f"Erro ao gerar PDF:<br><pre>{e}</pre>"


# =========================
# VALIDAÇÃO DA BASE
# =========================

@app.route("/validar-base", methods=["POST"])
def validar_base():
    try:
        excel = request.files["excel"]
        pdf = request.files["pdf"]

        base, erros = montar_base_validacao(excel, pdf)

        html = """
        <html>
        <head>
            <title>Base de Validação</title>
            <style>
                body { font-family: Arial; padding: 20px; }
                table { border-collapse: collapse; width: 100%; margin-bottom: 25px; }
                th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
                th { background: #111827; color: white; }
                .erro { color: red; font-weight: bold; }
                .ok { color: green; font-weight: bold; }
            </style>
        </head>
        <body>
        """

        html += "<h1>Base de Validação dos Crachás</h1>"
        html += f"<p class='ok'>Atletas encontrados: {len(base)}</p>"

        html += "<table>"
        html += "<tr><th>Página</th><th>Código QR</th><th>Nome</th><th>Categoria</th><th>Escola</th><th>Link</th></tr>"

        for item in base:
            codigo = item["codigo_qr"]
            html += f"""
            <tr>
                <td>{item['pagina']}</td>
                <td>{codigo}</td>
                <td>{item['nome']}</td>
                <td>{item['categoria']}</td>
                <td>{item['escola']}</td>
                <td><a href="/atleta/{codigo}" target="_blank">Abrir</a></td>
            </tr>
            """

        html += "</table>"

        if erros:
            html += f"<h2 class='erro'>Erros encontrados: {len(erros)}</h2>"
            html += "<table>"
            html += "<tr><th>Página</th><th>Código QR</th><th>Erro</th></tr>"

            for erro in erros:
                html += f"""
                <tr>
                    <td>{erro['pagina']}</td>
                    <td>{erro['codigo_qr']}</td>
                    <td>{erro['erro']}</td>
                </tr>
                """

            html += "</table>"

        html += """
            <p><a href="/painel">Ir para o painel de check-in</a></p>
        </body>
        </html>
        """

        return html

    except Exception as e:
        return f"Erro ao validar base:<br><pre>{e}</pre>"


# =========================
# TELA DO ATLETA / CHECK-IN
# =========================

@app.route("/atleta/<codigo_qr>")
def atleta(codigo_qr):
    atleta = buscar_atleta_por_codigo(codigo_qr)

    if not atleta:
        return f"""
        <html>
        <head>
            <title>Atleta não encontrado</title>
            <style>
                body {{
                    font-family: Arial;
                    background: #111827;
                    color: white;
                    text-align: center;
                    padding: 40px;
                }}
                .card {{
                    background: #1f2937;
                    padding: 30px;
                    border-radius: 16px;
                    max-width: 500px;
                    margin: auto;
                }}
                .erro {{ color: #ef4444; font-size: 28px; font-weight: bold; }}
            </style>
        </head>
        <body>
            <div class="card">
                <div class="erro">ATLETA NÃO ENCONTRADO</div>
                <p>Código lido: {codigo_qr}</p>
            </div>
        </body>
        </html>
        """

    status = atleta["status"]

    cor = "#22c55e" if status == "PENDENTE" else "#f59e0b"
    texto_status = "LIBERADO PARA ENTRAR" if status == "PENDENTE" else "JÁ ENTROU"

    botao = ""

    if status == "PENDENTE":
        botao = f"""
        <form action="/checkin/{codigo_qr}" method="POST">
            <button type="submit">CONFIRMAR ENTRADA</button>
        </form>
        """
    else:
        botao = f"<p>Entrada registrada em: <strong>{atleta['checkin_hora']}</strong></p>"

    return f"""
    <html>
    <head>
        <title>{atleta['nome']}</title>
        <style>
            body {{
                font-family: Arial;
                background: #111827;
                color: white;
                text-align: center;
                padding: 30px;
            }}
            .card {{
                background: #1f2937;
                padding: 30px;
                border-radius: 18px;
                max-width: 520px;
                margin: auto;
                box-shadow: 0 10px 30px rgba(0,0,0,.35);
            }}
            h1 {{
                font-size: 28px;
                margin-bottom: 10px;
            }}
            .categoria {{
                font-size: 22px;
                font-weight: bold;
                margin: 20px 0;
            }}
            .status {{
                background: {cor};
                color: #111827;
                padding: 14px;
                border-radius: 12px;
                font-size: 22px;
                font-weight: bold;
                margin: 20px 0;
            }}
            button {{
                background: #22c55e;
                border: 0;
                padding: 16px 28px;
                border-radius: 12px;
                font-size: 20px;
                font-weight: bold;
                cursor: pointer;
            }}
            a {{
                color: #93c5fd;
            }}
        </style>
    </head>
    <body>
        <div class="card">
            <h1>{atleta['nome']}</h1>
            <p>Código QR: <strong>{atleta['codigo_qr']}</strong></p>
            <p>Escola: <strong>{atleta['escola']}</strong></p>
            <p>Página do crachá: <strong>{atleta['pagina']}</strong></p>

            <div class="categoria">{atleta['categoria']}</div>
            <div class="status">{texto_status}</div>

            {botao}

            <p style="margin-top: 25px;">
                <a href="/painel">Ver painel</a>
            </p>
        </div>
    </body>
    </html>
    """


@app.route("/checkin/<codigo_qr>", methods=["POST"])
def checkin(codigo_qr):
    registrar_checkin(codigo_qr)
    return redirect(f"/atleta/{codigo_qr}")


# =========================
# PAINEL
# =========================

@app.route("/painel")
def painel():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) as total FROM atletas")
    total = cur.fetchone()["total"]

    cur.execute("SELECT COUNT(*) as total FROM atletas WHERE status = 'ENTROU'")
    entrou = cur.fetchone()["total"]

    cur.execute("SELECT COUNT(*) as total FROM atletas WHERE status = 'PENDENTE'")
    pendente = cur.fetchone()["total"]

    cur.execute("SELECT * FROM atletas ORDER BY nome")
    atletas = cur.fetchall()

    conn.close()

    html = f"""
    <html>
    <head>
        <title>Painel de Check-in</title>
        <style>
            body {{ font-family: Arial; padding: 20px; background: #f3f4f6; }}
            .cards {{ display: flex; gap: 15px; margin-bottom: 20px; }}
            .card {{
                background: white;
                padding: 18px;
                border-radius: 14px;
                box-shadow: 0 4px 12px rgba(0,0,0,.08);
                flex: 1;
            }}
            .num {{ font-size: 32px; font-weight: bold; }}
            table {{ border-collapse: collapse; width: 100%; background: white; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; }}
            th {{ background: #111827; color: white; }}
            .entrou {{ color: #16a34a; font-weight: bold; }}
            .pendente {{ color: #dc2626; font-weight: bold; }}
            input {{
                padding: 12px;
                font-size: 18px;
                width: 280px;
                margin-bottom: 15px;
            }}
            button {{
                padding: 12px 18px;
                font-size: 18px;
                cursor: pointer;
            }}
        </style>
    </head>
    <body>
        <h1>Painel de Check-in</h1>

        <form onsubmit="event.preventDefault(); irAtleta();">
            <input id="codigo" placeholder="Digite ou escaneie o código QR" autofocus>
            <button type="submit">Buscar</button>
        </form>

        <script>
            function irAtleta() {{
                const codigo = document.getElementById("codigo").value.trim();
                if (codigo) {{
                    window.location.href = "/atleta/" + codigo;
                }}
            }}
        </script>

        <div class="cards">
            <div class="card">
                <div>Total</div>
                <div class="num">{total}</div>
            </div>
            <div class="card">
                <div>Entraram</div>
                <div class="num">{entrou}</div>
            </div>
            <div class="card">
                <div>Pendentes</div>
                <div class="num">{pendente}</div>
            </div>
        </div>

        <table>
            <tr>
                <th>Código QR</th>
                <th>Nome</th>
                <th>Categoria</th>
                <th>Escola</th>
                <th>Status</th>
                <th>Hora</th>
                <th>Abrir</th>
            </tr>
    """

    for a in atletas:
        classe = "entrou" if a["status"] == "ENTROU" else "pendente"

        html += f"""
        <tr>
            <td>{a['codigo_qr']}</td>
            <td>{a['nome']}</td>
            <td>{a['categoria']}</td>
            <td>{a['escola']}</td>
            <td class="{classe}">{a['status']}</td>
            <td>{a['checkin_hora'] or ''}</td>
            <td><a href="/atleta/{a['codigo_qr']}" target="_blank">Abrir</a></td>
        </tr>
        """

    html += """
        </table>
    </body>
    </html>
    """

    return html


if __name__ == "__main__":
    init_db()
    app.run(debug=True, host="0.0.0.0", port=5000)
