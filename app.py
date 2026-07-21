import os
import time
import threading
import webbrowser
import logging
from flask import Flask, jsonify, render_template, request

# Desativa os logs padrão do Flask no terminal para manter a discrição na execução
log = logging.getLogger('werkzeug')
log.disabled = True

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

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

@app.route('/processar', methods=['POST'])
def processar_arquivo():
    """Rota que recebe o arquivo do frontend e realiza a automação."""
    if 'documento' not in request.files:
        return jsonify({"success": False, "error": "Nenhum arquivo enviado."}), 400
    
    arquivo = request.files['documento']
    
    if arquivo.filename == '':
        return jsonify({"success": False, "error": "Nome de arquivo vazio."}), 400
        
    try:
        # 1. Salva o arquivo na pasta temporária
        caminho_salvamento = os.path.join(UPLOAD_FOLDER, arquivo.filename)
        arquivo.save(caminho_salvamento)
        
        # 2. AQUI ENTRARÁ A LÓGICA DO PYTHON-DOCX / PYTHON-PPTX
        # Exemplo: 
        # novo_caminho = modificar_relatorio(caminho_salvamento)
        
        # Simula o tempo que o Python levaria lendo e editando o documento
        time.sleep(2) 
        print(f"[INFO] Arquivo {arquivo.filename} recebido e salvo em {caminho_salvamento}")
        
        return jsonify({
            "success": True, 
            "message": f"Arquivo {arquivo.filename} processado com sucesso!"
        })
        
    except Exception as e:
        print(f"[ERRO] Falha ao processar: {str(e)}")
        return jsonify({"success": False, "error": str(e)}), 500

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