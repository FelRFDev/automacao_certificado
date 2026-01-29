from docxtpl import DocxTemplate
import pandas as pd
from datetime import datetime

# 1. Carregar a planilha de dados (pode ser .xlsx ou .csv)
# Certifique-se que os nomes das colunas na planilha batem com o código
df = pd.read_excel('dados_alunos.xlsx')

# 2. Carregar o modelo do certificado
doc = DocxTemplate("template.docx")

# 3. Iterar sobre cada aluno na planilha
for index, linha in df.iterrows():
    # Criar um dicionário com os dados que vão substituir as tags {{ }}
    # As chaves aqui (ex: 'nome_aluno') devem ser IGUAIS às que estão no Word
    contexto = {
        "nome_do_aluno": linha['nome_do_aluno'],
        "curso": linha['curso'],
        "professor": linha['professor'],
        "data": linha['data'],
        "horas": linha['horas'],
        "porcentagem": linha['porcentagem'],
        "data_dia": linha['data_dia'],
        "data_mes": linha['data_mes'],
        "nome_coordenador": linha['nome_coordenador'],
        "nome_diretor": linha['nome_diretor']
    }

    # Renderizar o documento com os dados deste aluno
    doc.render(contexto)

    # 4. Salvar o novo arquivo com um nome único
    nome_arquivo = f"Certificado_{linha['nome_do_aluno']}.docx"
    doc.save(nome_arquivo)
    
    print(f"Gerado: {nome_arquivo}")

print("Todos os certificados foram gerados com sucesso!")