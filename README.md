# Cartão de Reconhecimento Automatizado

Este projeto tem como objetivo gerar cartões personalizados de reconhecimento de desempenho para colaboradores, utilizando dados extraídos de uma planilha Excel e aplicando-os automaticamente em um modelo de slide PowerPoint.

## 🎯 Objetivo

Automatizar o processo de valorização interna dos funcionários, eliminando a necessidade de edição manual e promovendo agilidade, padronização e economia de recursos.

## 🛠️ Tecnologias Utilizadas

- Python 3.9+
- Bibliotecas:
  - `pandas` – para leitura e manipulação dos dados do Excel
  - `python-pptx` – para criação e edição dos slides no PowerPoint
- PowerPoint – como plataforma visual para os cartões

## 📁 Estrutura do Projeto

- `template.pptx` – modelo do cartão com campos marcados para substituição (`NOME`, `MENSAGEM`)
- `dados_funcis.xlsx` – planilha com os dados dos colaboradores
- `gerador_cartoes.py` – script Python que automatiza a geração dos cartões
- `README.md` – este arquivo de documentação

## 📊 Formato da Planilha

A planilha `dados_funcis.xlsx` deve conter as seguintes colunas:

| Nome           | Cargo              | Pontos | Mensagem personalizada  |
|----------------|--------------------|--------|-------------------------|
| Ana Souza      | Analista Financeiro| 9.2    | Excelente.              |

## ▶️ Como Executar

1. Instale as bibliotecas necessárias:
   ```bash
   pip install pandas python-pptx openpyxl

   import pandas as pd
from pptx import Presentation
from pptx.util import Inches

# Carregue o template do PowerPoint
pr = Presentation('template.pptx')

# Pegue o primeiro slide (o nosso template)
template_slide = pr.slides[0]

# Carregue os dados da planilha Excel
df = pd.read_excel('dados_funcis.xlsx')

# Iterar sobre cada linha (cada cartão) da planilha
for index, row in df.iterrows():
    # Crie um novo slide baseado no layout do template
    # Pegamos o layout do template_slide
    novo_slide = pr.slides.add_slide(template_slide.slide_layout)
    
    # Percorra todas as formas (caixas de texto, etc.) no slide
    for shape in novo_slide.shapes:
        if shape.has_text_frame:
            text_frame = shape.text_frame
            for paragraph in text_frame.paragraphs:
                # Substitua os placeholders com os dados da linha
                if '{{nome}}' in paragraph.text:
                    paragraph.text = paragraph.text.replace('{{nome}}', str(row['Nome']))
                if '{{cargo}}' in paragraph.text:
                    paragraph.text = paragraph.text.replace('{{cargo}}', str(row['Cargo']))
                if '{{pontos}}' in paragraph.text:
                    paragraph.text = paragraph.text.replace('{{pontos}}', str(row['Pontos']))
                if '{{mensagem}}' in paragraph.text:
                    paragraph.text = paragraph.text.replace('{{mensagem}}', str(row['Mensagem']))

# Remova o slide de template original
pr.slides._sldIdLst.remove(pr.slides._sldIdLst[0])

# Salve a nova apresentação
pr.save('cartoes_gerados.pptx')

print("Apresentação 'cartoes_gerados.pptx' criada com sucesso!")
