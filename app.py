from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
import io
import base64
import fitz

app = Flask(__name__)


def normalizar_sexo(valor):
    sexo = str(valor).strip().upper()

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
        col_limpa = str(col).strip().upper()
        for nome in nomes_possiveis:
            if nome in col_limpa:
                return col
    return None


def montar_categoria(row, col_nome, col_sexo, col_data, linha_excel):
    nome = str(row[col_nome]).strip() if col_nome else f"LINHA {linha_excel}"

    sexo = normalizar_sexo(row[col_sexo])
    categoria = calcular_categoria(row[col_data])

    if not sexo:
        raise Exception(f"Linha {linha_excel}: sexo inválido para {nome}")

    if not categoria:
        raise Exception(f"Linha {linha_excel}: data de nascimento inválida ou fora da categoria para {nome}")

    return f"{categoria} {sexo}"


def criar_pdf(excel_file, pdf_file, pos_x, pos_y, rotacao, fonte, somente_primeira_pagina=False):
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip().str.upper()

    col_nome = encontrar_coluna(df, ["NOME"])
    col_sexo = encontrar_coluna(df, ["SEXO"])
    col_data = encontrar_coluna(df, ["DATA NASCIMENTO", "NASCIMENTO", "DATA"])

    if not col_nome:
        raise Exception("Não encontrei a coluna NOME na planilha.")

    if not col_sexo:
        raise Exception("Não encontrei a coluna SEXO na planilha.")

    if not col_data:
        raise Exception("Não encontrei a coluna DATA NASCIMENTO na planilha.")

    reader = PdfReader(pdf_file)
    writer = PdfWriter()

    total_paginas = 1 if somente_primeira_pagina else min(len(reader.pages), len(df))

    erros = []

    for i in range(total_paginas):
        row = df.iloc[i]
        linha_excel = i + 2

        try:
            categoria_texto = montar_categoria(row, col_nome, col_sexo, col_data, linha_excel)
        except Exception as e:
            erros.append(str(e))
            continue

        page = reader.pages[i]

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
        raise Exception("\n".join(erros[:20]))

    output = io.BytesIO()
    writer.write(output)
    output.seek(0)

    return output


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


if __name__ == "__main__":
    app.run(debug=True)
