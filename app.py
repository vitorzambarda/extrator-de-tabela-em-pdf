import PyPDF2
import requests
import re
import io
from docx import Document
from docx.shared import Inches
from flask import Flask, render_template, request, redirect, url_for, flash
import fitz
import msvcrt

app = Flask(__name__)
app.secret_key = "my_secret_key"

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/process_pdf", methods=["POST"])
def process_pdf():
    try:
        # Verificar se o arquivo enviado é um arquivo PDF
        if not request.files["pdf_file"].filename.endswith(".pdf"):
            flash("O arquivo enviado não é um arquivo PDF")
            return redirect(url_for("index"))

        # Ler o arquivo PDF
        pdf_file = request.files["pdf_file"].read()
        pdf_reader = PyPDF2.PdfFileReader(io.BytesIO(pdf_file))

        # Extrair tabelas do arquivo PDF
        tables = []
        for page_num in range(pdf_reader.getNumPages()):
            page = pdf_reader.getPage(page_num)
            page_text = page.extractText()
            page_tables = re.findall(r"[\d.,]+\s+[\w\s]+\s+[\d.,]+\s+[\d.,]+\s+[\d.,]+", page_text)
            for table_text in page_tables:
                table_rows = table_text.split("\n")
                table_rows = [row.strip() for row in table_rows if row.strip()]
                tables.append(table_rows)

        # Se não houver tabelas, exibir mensagem de erro
        if not tables:
            flash("Não foram encontradas tabelas no arquivo PDF")
            return redirect(url_for("index"))

        # Analisar tabelas com IA
        for table in tables:
            # encontrar os cabeçalhos
            item_header = "item" if "item" in table[0].lower() else ""
            descricao_header = "descrição" if "descrição" in table[0].lower() else ""
            marca_header = "marca" if "marca" in table[0].lower() else ""
            qtd_header = "quantidade" if "quantidade" in table[0].lower() else ""
            valor_unit_header = "valor unitário" if "valor unitário" in table[0].lower() else ""
            valor_total_header = "valor total" if "valor total" in table[0].lower() else ""

            # processar linhas da tabela
            for i in range(1, len(table)):
                item = ""
                descricao = ""
                marca = ""
                qtd = ""
                valor_unit = ""
                valor_total = ""
                row = table[i].split()

                # copiar valores das colunas correspondentes
                for j in range(len(row)):
                    if item_header and row[j].isdigit():
                        item = row[j]
                    elif descricao_header:
                        if descricao:
                            descricao += " "
                        descricao += row[j]
                    elif marca_header:
                        marca = ""
                    elif qtd_header:
                        if row[j].isdigit():
                            qtd = row[j]
                    elif valor_unit_header:
                        valor_unit = round(float(row[j].replace(",", ".")), 2)
                    elif valor_total_header:
                        valor_total = round(float(row[j].replace(",", ".")), 2)

                # calcular valor total
                if qtd and valor_unit:
                    valor_total = round(qtd * valor_unit, 2)
    finally:
                # adicionar linha na tabela do Word
                table.add_row().cells[0].text = item
                table.add_row().cells[1].text = descricao
                table.add_row().cells[2].text = marca
                table.add_row().cells[3].text = qtd
                table.add_row().cells[4].text = str(valor_unit)
                table.add_row()._cells[5].text = str(valor_total)