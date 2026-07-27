import os
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from docx import Document
from pptx import Presentation

def ler_regras_excel(caminho_excel):
    """
    Lê as regras de substituição do arquivo Excel.
    Coluna A = de (Tag), Coluna B = para (Novo Texto).
    Trata espaços em branco e converte tipos de dados numéricos para string.
    """
    # data_only=True garante que lemos o valor final de células com fórmulas
    wb = openpyxl.load_workbook(caminho_excel, data_only=True)
    planilha = wb.active
    regras = []
    
    for linha in planilha.iter_rows(min_row=1, values_only=True):
        de_texto = linha[0]
        para_texto = linha[1]
        
        # Só processa se a coluna A não estiver vazia
        if de_texto is not None and str(de_texto).strip() != "":
            de_formatado = str(de_texto).strip()
            # Se a coluna B for vazia, substitui por string vazia, caso contrário converte para texto
            para_formatado = str(para_texto).strip() if para_texto is not None else ""
            regras.append({'de': de_formatado, 'para': para_formatado})
            
    return regras

def processar_word(caminho_entrada, caminho_saida, regras):
    """Aplica as substituições em documentos Word (.docx) iterando sobre parágrafos e tabelas."""
    doc = Document(caminho_entrada)
    
    def substituir_texto(paragraphs):
        for p in paragraphs:
            for regra in regras:
                de, para = regra['de'], regra['para']
                if de in p.text:
                    # Tenta substituir nos 'runs' para preservar formatação (negrito, cor, etc.)
                    for run in p.runs:
                        if de in run.text:
                            run.text = run.text.replace(de, para)
                    # Fallback: Se a tag foi dividida em múltiplos runs pelo Word,
                    # substitui no parágrafo inteiro (pode perder formatações específicas)
                    if de in p.text:
                        p.text = p.text.replace(de, para)

    # Varre parágrafos normais
    substituir_texto(doc.paragraphs)
    
    # Varre tabelas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                substituir_texto(cell.paragraphs)
                
    doc.save(caminho_saida)

def processar_powerpoint(caminho_entrada, caminho_saida, regras):
    """Aplica as substituições em apresentações PowerPoint (.pptx)."""
    prs = Presentation(caminho_entrada)
    
    for slide in prs.slides:
        for shape in slide.shapes:
            # Substituição em caixas de texto padrão
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        for regra in regras:
                            if regra['de'] in run.text:
                                run.text = run.text.replace(regra['de'], regra['para'])
            
            # Substituição dentro de tabelas
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                for regra in regras:
                                    if regra['de'] in run.text:
                                        run.text = run.text.replace(regra['de'], regra['para'])
                                        
    prs.save(caminho_saida)

def iniciar_automacao():
    """Inicia a interface de seleção e executa a lógica de automação."""
    # Instancia e oculta a janela principal do Tkinter para mostrar apenas o popup
    root = tk.Tk()
    root.withdraw()
    
    # 1. Solicita a seleção do arquivo de relatório
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Word ou PowerPoint para alterar",
        filetypes=[
            ("Documentos suportados", "*.docx *.pptx"), 
            ("Word", "*.docx"), 
            ("PowerPoint", "*.pptx")
        ]
    )
    
    if not caminho_arquivo:
        return # Execução silenciosamente encerrada se o usuário clicar em "Cancelar"
        
    try:
        # 2. Identifica o diretório (pasta) do arquivo selecionado
        diretorio_base = os.path.dirname(caminho_arquivo)
        nome_arquivo = os.path.basename(caminho_arquivo)
        extensao = nome_arquivo.lower()
        
        # 3. Busca obrigatoriamente o 'dados.xlsx' no mesmo diretório
        caminho_excel = os.path.join(diretorio_base, "dados.xlsx")
        
        if not os.path.exists(caminho_excel):
            messagebox.showwarning(
                "Aviso: Planilha não encontrada", 
                f"O arquivo 'dados.xlsx' não foi encontrado na pasta do relatório:\n\n{diretorio_base}\n\nPor favor, adicione o arquivo de dados e tente novamente."
            )
            return
            
        # 4. Lê as regras do Excel e define onde salvar o arquivo modificado
        regras = ler_regras_excel(caminho_excel)
        
        if not regras:
            messagebox.showinfo("Aviso", "A planilha 'dados.xlsx' foi encontrada, mas parece estar vazia ou sem regras válidas.")
            return

        caminho_saida = os.path.join(diretorio_base, f"Modificado_{nome_arquivo}")
        
        # 5. Processa de acordo com o formato
        if extensao.endswith('.docx'):
            processar_word(caminho_arquivo, caminho_saida, regras)
        elif extensao.endswith('.pptx'):
            processar_powerpoint(caminho_arquivo, caminho_saida, regras)
        else:
            messagebox.showerror("Erro", "Formato de arquivo não suportado. Escolha um .docx ou .pptx.")
            return
            
        # 6. Exibe mensagem de sucesso ao final
        messagebox.showinfo("Sucesso", f"Relatório modificado com sucesso!\n\nFoi gerado um novo arquivo na mesma pasta:\n{caminho_saida}")
        
    except PermissionError:
        messagebox.showerror(
            "Erro de Permissão", 
            "O arquivo original ou o arquivo 'dados.xlsx' está aberto em outro programa.\n\nFeche o Word, PowerPoint ou Excel e tente novamente."
        )
    except Exception as e:
        messagebox.showerror("Erro de Processamento", f"Ocorreu um erro durante a execução:\n\n{str(e)}")

if __name__ == "__main__":
    iniciar_automacao()