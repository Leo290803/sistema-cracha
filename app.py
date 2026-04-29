from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
import io
import base64
import fitz

app = Flask(__name__)

def get_categoria(ano, sexo):
    if 2009 <= ano <= 2011:
        cat = "15 a 17 anos"
    elif 2012 <= ano <= 2014:
        cat = "12 a 14 anos"
    else:
        cat = "FORA"

    return f"{cat} - {sexo}"

def criar_pdf(excel_file, pdf_file, pos_x, pos_y, rotacao, fonte, somente_primeira_pagina=False):
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip().str.upper()

    col_data = None
    col_sexo = None

    for col in df.columns:
        if "DATA" in col:
            col_data = col
        if "SEXO" in col:
            col_sexo = col

    if not col_data or not col_sexo:
        raise Exception("Não encontrei DATA ou SEXO na planilha!")

    reader = PdfReader(pdf_file)
    writer = PdfWriter()

    total_paginas = 1 if somente_primeira_pagina else len(reader.pages)

    for i in range(total_paginas):
        page = reader.pages[i]

        packet = io.BytesIO()

        largura = float(page.mediabox.width)
        altura = float(page.mediabox.height)

        can = canvas.Canvas(packet, pagesize=(largura, altura))

        # 1 atleta por página
        idx = i

        if idx < len(df):
            row = df.iloc[idx]

            try:
                ano = pd.to_datetime(row[col_data], dayfirst=True).year
                sexo = str(row[col_sexo]).strip().upper()
                categoria = get_categoria(ano, sexo)

                can.setFont("Helvetica-Bold", fonte)

                can.saveState()
                can.translate(pos_x, pos_y)
                can.rotate(rotacao)
                can.drawString(0, 0, categoria)
                can.restoreState()

            except:
                pass

        can.save()
        packet.seek(0)

        overlay = PdfReader(packet)

        if len(overlay.pages) > 0:
            page.merge_page(overlay.pages[0])

        writer.add_page(page)

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
        return f"Erro ao gerar PDF: {e}"

if __name__ == "__main__":
    app.run(debug=True)