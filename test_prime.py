import requests
from bs4 import BeautifulSoup
import pandas as pd
from docx import Document
from tkinter import filedialog, Tk

# Função para extrair textos de um arquivo .docx
def extract_text_from_docx(file_path):
    doc = Document(file_path)
    texts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return texts

# Função para extrair textos de um arquivo .xlsx
def extract_text_from_excel(file_path):
    df = pd.read_excel(file_path, engine='openpyxl')
    texts = df.apply(lambda row: ' '.join(row.dropna().astype(str)), axis=1).tolist()
    return texts

# Função para buscar o conteúdo HTML de uma URL
def fetch_url_content(url):
    response = requests.get(url)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, 'html.parser')
    return soup

# Função para comparar os textos extraídos com o conteúdo da página
def compare_texts(texts, page_content):
    results = []
    for text in texts:
        # Usamos .get_text() para extrair o texto limpo da página, sem as tags HTML
        page_text = page_content.get_text(separator=' ', strip=True)  # Garantindo que as partes do texto dentro de <ul>/<li> sejam unidas corretamente
        
        # Verificamos se o texto procurado está presente no texto limpo da página
        if text in page_text:
            # Procuramos a tag HTML que contém o texto
            match = page_content.find(string=lambda t: t and text in t)
            if match:
                tag = match.find_parent()
                results.append({
                    "Texto": text,
                    "Presente na URL": "Sim",
                    "Tag HTML": tag.name
                })
        else:
            results.append({
                "Texto": text,
                "Presente na URL": "Não",
                "Tag HTML": "-"
            })
    return results


# Função para gerar um relatório em Excel
def generate_report(results, output_path):
    df = pd.DataFrame(results)
    df.to_excel(output_path, index=False)

# Função principal que recebe os caminhos do arquivo e da URL
def validate_texts(file_path, url, output_path):
    try:
        if file_path.endswith(".docx"):
            texts = extract_text_from_docx(file_path)
        elif file_path.endswith(".xlsx"):
            texts = extract_text_from_excel(file_path)
        else:
            raise ValueError("Formato de arquivo não suportado.")

        page_content = fetch_url_content(url)
        results = compare_texts(texts, page_content)
        generate_report(results, output_path)
        print(f"Validação concluída! Relatório gerado: {output_path}")
    except Exception as e:
        print(f"Ocorreu um erro: {str(e)}")

# Função para abrir a janela de seleção de pasta para salvar o arquivo de saída
def select_output_folder():
    root = Tk()
    root.withdraw()  # Não exibe a janela principal
    output_folder = filedialog.askdirectory(title="Escolha a pasta de saída")
    return output_folder

if __name__ == "__main__":
    file_path = input("Digite o caminho do arquivo (.docx ou .xlsx): ")
    url = input("Digite a URL a ser verificada: ")
    
    # Seleciona a pasta de saída
    output_folder = select_output_folder()
    if output_folder:
        output_path = f"{output_folder}/resultado_validacao.xlsx"
        validate_texts(file_path, url, output_path)
    else:
        print("Pasta de saída não selecionada.")
