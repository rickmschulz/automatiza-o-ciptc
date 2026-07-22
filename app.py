import os
import time
import json
import io
import threading
import webbrowser
import logging
import re
from flask import Flask, jsonify, render_template, request, send_file
from docx import Document
from pptx import Presentation

# Desativa os logs padrão do Flask no terminal
log = logging.getLogger('werkzeug')
log.disabled = True

app = Flask(__name__)

last_heartbeat_time = time.time()
TIMEOUT_SECONDS = 10
PORT = 5000

# Expressão Regular para encontrar padrões como [TEXTO] ou {{TEXTO}}
TAG_PATTERN = r'\[.*?\]|\{\{.*?\}\}'

@app.route('/')
def home():
    global last_heartbeat_time
    last_heartbeat_time = time.time()
    return render_template('index.html')

@app.route('/heartbeat', methods=['POST'])
def heartbeat():
    global last_heartbeat_time
    last_heartbeat_time = time.time()
    return jsonify({"status": "alive"})

@app.route('/extrair_tags', methods=['POST'])
def extrair_tags():
    """Lê o arquivo temporariamente e devolve todas as tags encontradas."""
    if 'documento' not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado."}), 400
    
    arquivo = request.files['documento']
    if arquivo.filename == '':
        return jsonify({"error": "Arquivo vazio."}), 400

    extensao = arquivo.filename.lower()
    file_stream = io.BytesIO(arquivo.read())
    tags_encontradas = set() # Usamos 'set' para não repetir tags iguais

    try:
        if extensao.endswith('.docx'):
            doc = Document(file_stream)
            # Busca em parágrafos
            for p in doc.paragraphs:
                tags_encontradas.update(re.findall(TAG_PATTERN, p.text))
            # Busca em tabelas
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            tags_encontradas.update(re.findall(TAG_PATTERN, p.text))
                            
        elif extensao.endswith('.pptx'):
            prs = Presentation(file_stream)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for p in shape.text_frame.paragraphs:
                            tags_encontradas.update(re.findall(TAG_PATTERN, p.text))
                    if shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                for p in cell.text_frame.paragraphs:
                                    tags_encontradas.update(re.findall(TAG_PATTERN, p.text))
                                    
        # Converte o 'set' para uma lista normal do Python para enviar como JSON
        return jsonify({"tags": list(tags_encontradas)})
    except Exception as e:
        print(f"[ERRO na Extração] {str(e)}")
        return jsonify({"error": str(e)}), 500

def processar_word(file_stream, regras):
    doc = Document(file_stream)
    def substituir_texto(paragraphs):
        for p in paragraphs:
            for regra in regras:
                de, para = regra['de'], regra['para']
                if de in p.text:
                    for run in p.runs:
                        if de in run.text:
                            run.text = run.text.replace(de, para)
                    if de in p.text:
                        p.text = p.text.replace(de, para)
    substituir_texto(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                substituir_texto(cell.paragraphs)
    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    return output_stream

def processar_powerpoint(file_stream, regras):
    prs = Presentation(file_stream)
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        for regra in regras:
                            if regra['de'] in run.text:
                                run.text = run.text.replace(regra['de'], regra['para'])
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                for regra in regras:
                                    if regra['de'] in run.text:
                                        run.text = run.text.replace(regra['de'], regra['para'])
    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream

@app.route('/processar', methods=['POST'])
def processar_arquivo():
    if 'documento' not in request.files:
        return "Nenhum arquivo enviado.", 400
    arquivo = request.files['documento']
    if arquivo.filename == '':
        return "Nome de arquivo vazio.", 400
        
    try:
        regras_str = request.form.get('substituicoes', '[]')
        regras_json = json.loads(regras_str) 
        file_stream = io.BytesIO(arquivo.read())
        
        print(f"\n[INFO] Processando '{arquivo.filename}' na memória...")
        print(f"[INFO] Regras solicitadas: {regras_json}")
        
        extensao = arquivo.filename.lower()
        nome_arquivo_saida = f"Modificado_{arquivo.filename}"
        
        if extensao.endswith('.docx'):
            output_stream = processar_word(file_stream, regras_json)
            mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        elif extensao.endswith('.pptx'):
            output_stream = processar_powerpoint(file_stream, regras_json)
            mimetype = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        else:
            return "Formato não suportado.", 400
        
        # Injeta o nome sugerido no cabeçalho
        response = send_file(output_stream, mimetype=mimetype)
        response.headers['Content-Disposition'] = f'attachment; filename="{nome_arquivo_saida}"'
        return response
        
    except Exception as e:
        print(f"[ERRO] Falha ao processar: {str(e)}")
        return str(e), 500

def monitor_shutdown():
    while True:
        time.sleep(2)
        tempo_sem_comunicacao = time.time() - last_heartbeat_time
        if tempo_sem_comunicacao > TIMEOUT_SECONDS:
            print(f"\n[INFO] Aba fechada. Servidor inativo por {TIMEOUT_SECONDS}s. Desligando...")
            os._exit(0)

def open_browser():
    time.sleep(1)
    webbrowser.open(f"http://127.0.0.1:{PORT}")

if __name__ == '__main__':
    monitor_thread = threading.Thread(target=monitor_shutdown, daemon=True)
    monitor_thread.start()
    browser_thread = threading.Thread(target=open_browser, daemon=True)
    browser_thread.start()
    print(f"Iniciando automação web na porta {PORT}...")
    app.run(port=PORT, host="127.0.0.1")