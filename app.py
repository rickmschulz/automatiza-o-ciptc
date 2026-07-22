import os
import time
import json
import io
import threading
import webbrowser
import logging
from flask import Flask, jsonify, render_template, request, send_file
from docx import Document
from pptx import Presentation

# Desativa os logs padrão do Flask no terminal para manter a discrição na execução
log = logging.getLogger('werkzeug')
log.disabled = True

app = Flask(__name__)

# Variável global para armazenar a última vez que o navegador enviou um sinal
last_heartbeat_time = time.time()

# Constantes de configuração
TIMEOUT_SECONDS = 3  # Tempo limite sem resposta antes de desligar (em segundos)
PORT = 5000           # Porta local da aplicação

@app.route('/')
def home():
    """Rota principal que entrega a interface HTML e reseta o relógio de segurança."""
    global last_heartbeat_time
    # Reseta o tempo logo que a página carrega
    last_heartbeat_time = time.time()
    return render_template('index.html')

@app.route('/heartbeat', methods=['POST'])
def heartbeat():
    """Rota silenciosa que o JavaScript acessa de tempos em tempos."""
    global last_heartbeat_time
    # Atualiza o relógio sinalizando que o navegador ainda está aberto
    last_heartbeat_time = time.time()
    return jsonify({"status": "alive"})

def processar_word(file_stream, regras):
    """Abre o documento Word na memória e substitui os termos em parágrafos e tabelas."""
    doc = Document(file_stream)
    
    def substituir_texto(paragraphs):
        for p in paragraphs:
            for regra in regras:
                de, para = regra['de'], regra['para']
                if de in p.text:
                    # 1ª Tentativa: Substituir mantendo a formatação original (runs)
                    for run in p.runs:
                        if de in run.text:
                            run.text = run.text.replace(de, para)
                    
                    # 2ª Tentativa (Fallback): Se a palavra quebrou entre múltiplos runs
                    if de in p.text:
                        p.text = p.text.replace(de, para)
    
    # Substituir no corpo do texto
    substituir_texto(doc.paragraphs)
    
    # Substituir dentro de tabelas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                substituir_texto(cell.paragraphs)
                
    # Salva o resultado em um buffer de memória em vez de no disco
    output_stream = io.BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    return output_stream

def processar_powerpoint(file_stream, regras):
    """Abre o PowerPoint na memória e substitui termos nos slides e tabelas."""
    prs = Presentation(file_stream)
    
    for slide in prs.slides:
        for shape in slide.shapes:
            # Substituir em caixas de texto comuns
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        for regra in regras:
                            if regra['de'] in run.text:
                                run.text = run.text.replace(regra['de'], regra['para'])
            
            # Substituir dentro de tabelas nos slides
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                for regra in regras:
                                    if regra['de'] in run.text:
                                        run.text = run.text.replace(regra['de'], regra['para'])
                                        
    # Salva o resultado em um buffer de memória
    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    return output_stream

@app.route('/processar', methods=['POST'])
def processar_arquivo():
    """Rota que recebe o arquivo e realiza a automação 100% em memória."""
    if 'documento' not in request.files:
        return "Nenhum arquivo enviado.", 400
    
    arquivo = request.files['documento']
    
    if arquivo.filename == '':
        return "Nome de arquivo vazio.", 400
        
    try:
        # Captura as regras de substituição enviadas pelo JavaScript
        regras_str = request.form.get('substituicoes', '[]')
        regras_json = json.loads(regras_str) 
        
        # Lê o arquivo recebido diretamente para a memória
        file_stream = io.BytesIO(arquivo.read())
        
        print(f"\n[INFO] Processando '{arquivo.filename}' na memória...")
        print(f"[INFO] Regras de substituição solicitadas: {regras_json}")
        
        # Escolhe a função correta e processa em memória
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
        
        # Envia o arquivo resultante diretamente de volta para o cliente, sem tocar no disco rígido
        return send_file(
            output_stream,
            as_attachment=True,
            download_name=nome_arquivo_saida,
            mimetype=mimetype
        )
        
    except Exception as e:
        print(f"[ERRO] Falha ao processar: {str(e)}")
        return str(e), 500

def monitor_shutdown():
    """
    Roda num loop infinito em uma thread separada.
    Se o navegador não mandar um sinal (heartbeat) por mais tempo que o TIMEOUT_SECONDS,
    ele assume que a aba foi fechada e encerra todo o programa Python.
    """
    while True:
        time.sleep(2) # Verifica a cada 2 segundos
        tempo_sem_comunicacao = time.time() - last_heartbeat_time
        
        if tempo_sem_comunicacao > TIMEOUT_SECONDS:
            print(f"\n[INFO] Aba fechada. Servidor inativo por {TIMEOUT_SECONDS}s. Desligando...")
            # Força o encerramento limpo de todos os processos do sistema operacional
            os._exit(0)

def open_browser():
    """Espera 1 segundo para o Flask iniciar e abre o navegador padrão."""
    time.sleep(1)
    webbrowser.open(f"http://127.0.0.1:{PORT}")

if __name__ == '__main__':
    # 1. Inicia o monitor de desligamento automático em segundo plano
    monitor_thread = threading.Thread(target=monitor_shutdown, daemon=True)
    monitor_thread.start()
    
    # 2. Inicia uma thread para abrir o navegador
    browser_thread = threading.Thread(target=open_browser, daemon=True)
    browser_thread.start()
    
    # 3. Inicia o servidor Flask
    print(f"Iniciando automação web na porta {PORT}...")
    print("Mantenha a janela do terminal aberta por enquanto para ver os testes.")
    app.run(port=PORT, host="127.0.0.1")