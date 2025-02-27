import pdfplumber
import re
from pptx import Presentation
from pptx.util import Inches

# ------ PDF ----------
pdf_path = "arquivo.pdf"

# cabeça da matriz
dados_tabela = [["Código", "Especificação", "Valor Empenhado", "Valor Liquidado", "Valor Pago"]]  

with pdfplumber.open(pdf_path) as pdf:
    for page in pdf.pages:
        texto = page.extract_text()  
        if texto:

            #quando tiver um \n, ele vai dividir as strings
            linhas = texto.split("\n")  

            for linha in linhas:

                # filtra as linhas que tem código X.X.XX.XX.XX
                if re.match(r"^\d{1,2}\.\d{1,2}\.\d{1,2}\.\d{1,2}\.\d{1,2}", linha):

                    # regex para capturar as colunas
                    match = re.match(r"^(\d{1,2}\.\d{1,2}\.\d{1,2}\.\d{1,2}\.\d{1,2})\s+(.+?)\s+([\d\.,-]+)\s+([\d\.,-]+)\s+([\d\.,-]+)\*?$", linha)
                    
                    if match:
                        codigo = match.group(1)
                        especificacao = match.group(2).strip()  #strip tira espaços do texto
                        empenhado = match.group(3)
                        liquidado = match.group(4)
                        pago = match.group(5)

                        dados_tabela.append([codigo, especificacao, empenhado, liquidado, pago])

                    else:
                        # se o match nao capturar, ele entra nesse regex alternativo
                        partes = re.split(r"\s{2,}", linha)  # se tiver 2+ espaços, ele divide a linha
                        if len(partes) >= 5:
                            codigo = partes[0]
                            especificacao = " ".join(partes[1:-3])  # junta as partes da especificação
                            empenhado = partes[-3]
                            liquidado = partes[-2]
                            pago = partes[-1].replace("*", "")  # remove os asteriscos do texto

                            dados_tabela.append([codigo, especificacao, empenhado, liquidado, pago])

if len(dados_tabela) == 1:
    print("Nenhum dado válido foi encontrado no PDF.")
    exit()

# ------ PowerPoint ----------

# cria a apresentacao
apresentacao = Presentation()

# limite de linhas dentro da lâmina do slide
linhas_por_slide = 20  

def criar_slides():
    # Criar slides conforme necessário
    for i in range(0, len(dados_tabela), linhas_por_slide):
        # Adicionar slide
        slide = apresentacao.slides.add_slide(apresentacao.slide_layouts[5])

        rows = min(linhas_por_slide, len(dados_tabela) - i)
        cols = len(dados_tabela[0])

        # posição e tamanho
        left = Inches(0.5)
        top = Inches(1)
        width = Inches(9)
        height = Inches(5)

        # cria a tabela 
        table = slide.shapes.add_table(rows, cols, left, top, width, height).table

        # preenche a tabela com os dados da matriz
        for r, linha in enumerate(dados_tabela[i:i + linhas_por_slide]):
            for c, valor in enumerate(linha):
                table.cell(r, c).text = valor

criar_slides()

# salva a apresentacao
pptx_path = "tabela_pptx.pptx"
apresentacao.save(pptx_path)

print(f"Apresentação criada com sucesso! Arquivo salvo em: {pptx_path}")