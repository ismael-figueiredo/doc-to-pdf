import os
import win32com.client

def doc_para_pdf(caminho_doc, caminho_pdf):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.visible = False
        doc = word.Documents.Open(caminho_doc)
        doc.SaveAs(caminho_pdf, FileFormat=17)
        doc.Close()
        word.Quit()
        print(f"Arquivo convertido: {caminho_pdf}")
    except Exception as e:
        print(f"Erro ao converter o arquivo {caminho_doc}: {str(e)}")

# Diretório onde estão os arquivos .doc
diretorio_raiz = 'C:\\Users\\User\\Desktop\\it'  # Atualize para seu caminho absoluto

# Cria uma pasta para os PDFs se ainda não existir
pasta_pdf = os.path.join(diretorio_raiz, 'pdf')
os.makedirs(pasta_pdf, exist_ok=True)

# Verifica e lista todos os arquivos .doc
arquivos_encontrados = os.listdir(diretorio_raiz)
print("Arquivos encontrados no diretório:", arquivos_encontrados)

for arquivo in arquivos_encontrados:
    if arquivo.endswith(".doc"):
        caminho_completo = os.path.join(diretorio_raiz, arquivo)
        nome_pdf = arquivo[:-4] + '.pdf'
        caminho_pdf = os.path.join(pasta_pdf, nome_pdf)
        doc_para_pdf(caminho_completo, caminho_pdf)

print("Conversão concluída! Verifique a pasta 'pdf' para os arquivos convertidos.")
