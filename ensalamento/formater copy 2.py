import os
import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape
import segno
import pdfkit
from playwright.sync_api import sync_playwright
import re
from concurrent.futures import ThreadPoolExecutor
import threading
import base64

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

def sanitize_filename(name, max_length=60):
    # Remove caracteres inválidos para nomes de arquivos/pastas no Windows
    name = ''.join(c for c in name if c not in '\\/:*?"<>|')
    # Substitui espaços múltiplos por um único espaço
    name = re.sub(r'\s+', ' ', name)
    # Substitui espaços por sublinhado
    name = name.replace(' ', '_')
    # Limita o tamanho do nome
    return name[:max_length].strip('_')

def process_sala(sala_df, template, env, counter, total):
    nome_escola = sala_df.iloc[0]['ESCO']
    nome_sala = sala_df.iloc[0]['NOME SALA']
    turno = sala_df.iloc[0]['TURNO']
    
    # Carrega o template da capa
    cover_template = env.get_template('cover_sheet.html')
    
    # Prepara os dados para a capa
    cover_context = {
        'entity': sala_df.iloc[0]['CONCURSO'],
        'shift': f'01 - {turno}',
        'room__school__name': nome_escola,
        'room__block': sala_df.iloc[0]['BLOCO'],
        'room__floor': sala_df.iloc[0]['ANDAR'],
        'room__name': nome_sala,
        'room__id': sala_df.iloc[0]['ID_SALA'],
        'positions': {
            str(idx): {
                'id': row['DESC'].split(" ")[1],
                'name': row['DESC'],
                'count': len(sala_df[sala_df['DESC'] == row['DESC']])
            }
            for idx, row in sala_df.drop_duplicates('DESC').iterrows()
        },
        'prefeitura_logo': 'data:image/jpeg;base64,' + base64.b64encode(open('174.jpg', 'rb').read()).decode('utf-8'),
        'instituto_logo': 'data:image/png;base64,' + base64.b64encode(open('logo.png', 'rb').read()).decode('utf-8')
    }
    
    # Gera o HTML da capa
    cover_html = cover_template.render(**cover_context)
    
    # Lista para armazenar todo o conteúdo HTML da sala
    all_html_content = [cover_html]  # Adiciona a capa como primeiro elemento
    
    for idx, row in sala_df.iterrows():
        register = {
            'edital': row.get('CONCURSO', ''),
            'nome_escola': nome_escola,
            'bloco': row.get('BLOCO', ''),
            'andar': row.get('ANDAR', ''),
            'turno': row.get('TURNO', ''),
            'nome_sala': row.get('SALA', ''),
            'ordem': row.get('CART', ''),
            'nome_cargo': row.get('DESC', ''),
            'id_cargo': row.get('DESC', '').split(" ")[1],
            'id_inscricao': row.get('INSC', ''),
            'nome_candidato': row.get('NOME', ''),
            'candidato_cpf': row.get('CPF', ''),
            'ordem_geral': row.get('ORDEM', ''),
        }
        qr_code_svg = _generate_qr_code(register)
        context = {
            'qr_code': qr_code_svg,
            'register': register,
            'prefeitura_logo': 'data:image/jpeg;base64,' + base64.b64encode(open('174.jpg', 'rb').read()).decode('utf-8'),
            'instituto_logo': 'data:image/png;base64,' + base64.b64encode(open('logo.png', 'rb').read()).decode('utf-8')
        }
        html_content = template.render(**context)
        all_html_content.append(html_content)
        
        with counter['lock']:
            counter['count'] += 1
            print(f"Processado candidato {register['nome_candidato']} ({counter['count']}/{total})")
    
    # Combina todo o conteúdo HTML em um único documento
    combined_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            .page-break {{
                page-break-after: always;
            }}
        </style>
    </head>
    <body>
        {''.join(f'<div class="page-break">{html}</div>' for html in all_html_content)}
    </body>
    </html>
    """
    
    # Cria o diretório e gera o PDF único para a sala
    escola_dir = os.path.join(DOCS_DIR, sanitize_filename(nome_escola))
    sala_dir = os.path.join(escola_dir, f'{sanitize_filename(nome_sala)}_{turno}')
    os.makedirs(sala_dir, exist_ok=True)
    pdf_filename = f"answer_sheets_{sanitize_filename(nome_sala)}_{turno}.pdf"
    pdf_path = os.path.join(sala_dir, pdf_filename)
    
    # Gera o PDF com todas as folhas de resposta
    html_to_pdf(combined_html, pdf_path)
    print(f"Gerado PDF da sala: {pdf_path}")

def main():
    env = Environment(
        loader=FileSystemLoader(TEMPLATE_DIR),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template(TEMPLATE_FILE)
    df = pd.read_csv(CSV_PATH, dtype=str)
    grouped = df.groupby(['ESCO', 'NOME SALA', 'TURNO'])
    total = len(df)
    counter = {'count': 0, 'lock': threading.Lock()}
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = [executor.submit(process_sala, sala_df, template, env, counter, total) for _, sala_df in grouped]
        for future in futures:
            future.result()

if __name__ == '__main__':
    main()
