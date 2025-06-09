import os
import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape
import segno
import pdfkit
from playwright.sync_api import sync_playwright

# Coloque aqui o caminho exato do seu executável
path_wkhtmltopdf = r'C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'  # ajuste se necessário
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

# Caminhos
CSV_PATH = 'PLANILHA_ENSALAMENTO_DEFINITIVO - candidatos.csv'
TEMPLATE_DIR = 'templates/answer_sheet'
TEMPLATE_FILE = 'answer_sheet_static.html'
DOCS_DIR = 'docs'

# Garante que a pasta docs existe
os.makedirs(DOCS_DIR, exist_ok=True)

def _generate_qr_code(register):
    qr_content = (
        f"Código do Cargo: {register['id_cargo']}\n"
        f"Número de Inscrição: {register['id_inscricao']}\n"
        f"Candidato(a): {register['nome_candidato']}\n"
        f"CPF: {register['candidato_cpf']}"
    )
    return segno.make(qr_content).svg_inline(dark='#000000', scale=1.5, border=0)

def html_to_pdf(html_content, pdf_path):
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.set_content(html_content)
        page.pdf(path=pdf_path, format="A4")
        browser.close()

def main():
    # Carrega o template Jinja2
    env = Environment(
        loader=FileSystemLoader(TEMPLATE_DIR),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template(TEMPLATE_FILE)

    # Lê o CSV (apenas as colunas necessárias para o answer_sheet)
    df = pd.read_csv(CSV_PATH, dtype=str)
    # df = df.fillna('')

    for idx, row in df.iterrows():
        register = {
            'edital': row.get('CONCURSO', ''),
            'nome_escola': row.get('ESCO', ''),
            'bloco': row.get('BLOCO', ''),
            'andar': row.get('ANDAR', ''),
            'turno': row.get('TURNO', ''),
            'nome_sala': row.get('SALA', ''),
            'ordem': row.get('ORDEM', ''),
            'nome_cargo': row.get('DESC', ''),
            'id_cargo': row.get('DESC', '').split(" ")[1],
            'id_inscricao': row.get('INSC', ''),
            'nome_candidato': row.get('NOME', ''),
            'candidato_cpf': row.get('CPF', ''),
        }

        qr_code_svg = _generate_qr_code(register)

        # Monta o contexto para o template
        context = {
            'qr_code': qr_code_svg,
            'register': register
        }

        # Renderiza o HTML
        html_content = template.render(**context)

        # Converte para PDF
        pdf_path = os.path.join(DOCS_DIR, f"{register['nome_escola']+"_"+register['bloco']+"_"+register['andar']+"_"+register['nome_sala']+"_"+register['ordem']}.pdf")
        html_to_pdf(html_content, pdf_path)
        print(f'Gerado: {pdf_path}')

if __name__ == '__main__':
    main()
