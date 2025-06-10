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
from datetime import datetime

# Coloque aqui o caminho exato do seu executável
path_wkhtmltopdf = r'C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'  # ajuste se necessário
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

# Caminhos
CSV_PATH = 'PLANILHA_ENSALAMENTO_DEFINITIVO - candidatos.csv'
TEMPLATE_DIR = 'templates/answer_sheet'
TEMPLATE_FILE = 'answer_sheet_static.html'
DOCS_DIR = 'docs'
LOG_FILE = 'erros_geracao.txt'

# Garante que a pasta docs existe
os.makedirs(DOCS_DIR, exist_ok=True)

def log_error(escola, error_message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(f"[{timestamp}] Escola: {escola} - Erro: {error_message}\n")

def _generate_qr_code(register):
    qr_content = (
        f"Código do Cargo: {register['id_cargo']}\n"
        f"Número de Inscrição: {register['id_inscricao']}\n"
        f"Candidato(a): {register['nome_candidato']}\n"
        f"CPF: {register['candidato_cpf']}"
    )
    return segno.make(qr_content).svg_inline(dark='#000000', scale=1.5, border=0)

def html_to_pdf(html_content, pdf_path):
    try:
        with sync_playwright() as p:
            # Configuração do browser com timeout aumentado para 2 minutos
            browser = p.chromium.launch()
            context = browser.new_context()
            page = context.new_page()
            
            # Configura timeout para 2 minutos
            page.set_default_timeout(120000)
            
            # Carrega o conteúdo e aguarda o carregamento completo
            page.set_content(html_content, wait_until='networkidle', timeout=120000)
            
            # Aguarda um pouco mais para garantir que tudo foi renderizado
            page.wait_for_load_state('networkidle', timeout=120000)
            
            # Gera o PDF com timeout aumentado
            page.pdf(path=pdf_path, format="A4")
            
            # Fecha o contexto e o browser
            context.close()
            browser.close()
            
    except Exception as e:
        print(f"Erro ao gerar PDF: {str(e)}")
        raise e

def sanitize_filename(name, max_length=60):
    # Remove caracteres inválidos para nomes de arquivos/pastas no Windows
    name = ''.join(c for c in name if c not in '\\/:*?"<>|')
    # Substitui espaços múltiplos por um único espaço
    name = re.sub(r'\s+', ' ', name)
    # Substitui espaços por sublinhado
    name = name.replace(' ', '_')
    # Limita o tamanho do nomecd ..
    return name[:max_length].strip('_')

def process_escola(escola_df, template, env, counter, total):
    try:
        nome_escola = escola_df.iloc[0]['ESCO']
        
        # Verifica se o PDF já existe
        escola_dir = os.path.join(DOCS_DIR, sanitize_filename(nome_escola))
        pdf_filename = f"answer_sheets_{sanitize_filename(nome_escola)}.pdf"
        pdf_path = os.path.join(escola_dir, pdf_filename)
        
        if os.path.exists(pdf_path):
            print(f"PDF da escola {nome_escola} já existe, pulando...")
            return
        
        print(f"Gerando PDF da escola {nome_escola}...")
        
        all_html_content = []
        total_candidatos_escola = len(escola_df)
        candidatos_processados = 0
        
        # Agrupa por sala
        for (nome_sala, turno), sala_df in escola_df.groupby(['NOME SALA', 'TURNO']):
            try:
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
                all_html_content.append(cover_html)
                
                for idx, row in sala_df.iterrows():
                    try:
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
                        
                        candidatos_processados += 1
                        porcentagem_escola = (candidatos_processados / total_candidatos_escola) * 100
                        porcentagem_geral = (counter['count'] / total) * 100
                        
                        with counter['lock']:
                            counter['count'] += 1
                            print(f"\rProgresso da escola {nome_escola}: {porcentagem_escola:.1f}% | Progresso geral: {porcentagem_geral:.1f}%", end='')
                            
                    except Exception as e:
                        print(f"\nErro ao processar candidato {row.get('NOME', '')} da escola {nome_escola}, sala {nome_sala}: {str(e)}")
                        continue
            except Exception as e:
                print(f"\nErro ao processar sala {nome_sala} da escola {nome_escola}: {str(e)}")
                continue
        
        print(f"\nGerando PDF final para a escola {nome_escola}...")
        
        if not all_html_content:
            print(f"Nenhum conteúdo gerado para a escola {nome_escola}")
            return
        
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
        
        # Cria o diretório e gera o PDF único para a escola
        os.makedirs(escola_dir, exist_ok=True)
        
        # Gera o PDF com todas as folhas de resposta
        html_to_pdf(combined_html, pdf_path)
        print(f"PDF gerado com sucesso: {pdf_path}")
    except Exception as e:
        error_message = str(e)
        print(f"\nErro ao processar escola {nome_escola}: {error_message}")
        log_error(nome_escola, error_message)

def main():
    # Limpa o arquivo de log se ele existir
    if os.path.exists(LOG_FILE):
        os.remove(LOG_FILE)
        
    env = Environment(
        loader=FileSystemLoader(TEMPLATE_DIR),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template(TEMPLATE_FILE)
    df = pd.read_csv(CSV_PATH, dtype=str)
    df = df.fillna('')
    grouped = df.groupby(['ESCO'])
    total = len(df)
    counter = {'count': 0, 'lock': threading.Lock()}
    
    # Lista para armazenar escolas que falharam
    failed_schools = []
    
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = [executor.submit(process_escola, escola_df, template, env, counter, total) for _, escola_df in grouped]
        for future in futures:
            try:
                future.result()
            except Exception as e:
                print(f"Erro ao processar escola: {str(e)}")
                failed_schools.append(future)
    
    # Tenta novamente as escolas que falharam
    if failed_schools:
        print("\nTentando novamente as escolas que falharam...")
        with ThreadPoolExecutor(max_workers=5) as executor:
            retry_futures = [executor.submit(process_escola, escola_df, template, env, counter, total) 
                           for _, escola_df in grouped if escola_df.iloc[0]['ESCO'] in failed_schools]
            for future in retry_futures:
                try:
                    future.result()
                except Exception as e:
                    print(f"Erro na segunda tentativa: {str(e)}")

if __name__ == '__main__':
    main()
