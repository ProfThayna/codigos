import requests
import pandas as pd
from bs4 import BeautifulSoup
import os
from datetime import datetime


# Defina sua chave de API e o código do formulário
api_key = '0930cd99221b6e7303b04a15ef5b46ca'
form_id = '240654602222648'

def get_submissions(form_id, api_key, limit=1000, offset=0):
    url = f'https://api.jotform.com/form/{form_id}/submissions?apiKey={api_key}&limit={limit}&offset={offset}'
    response = requests.get(url)
    data = response.json()
    
    # Verifique se a chave 'content' está presente nos dados retornados
    if 'content' not in data:
        raise ValueError("A resposta da API não contém a chave 'content'.")
    
    return data['content']

def combine_pages(form_id, api_key, limit=1000):
    all_submissions = []
    offset = 0
    
    while True:
        submissions = get_submissions(form_id, api_key, limit, offset)
        if not submissions:
            break
        all_submissions.extend(submissions)
        offset += limit
    
    return all_submissions

# Obter todas as submissões
submissions = combine_pages(form_id, api_key)

# Função para limpar HTML
def clean_html(raw_html):
    if isinstance(raw_html, list):
        return ', '.join(raw_html)  # Concatenar lista com vírgulas
    elif isinstance(raw_html, str):
        return raw_html.strip()
    else:
        return ''


# Campos relevantes
campos_relevantes = {
    'created_at': 'Submission Date',
    '343': 'Patrocinador',
    '345': 'Cidade',
    '21': 'Turno:',
    '346': 'Local de execução',
    '385': 'Projeto',
    '20': 'Turma',
    '234': 'Quantidade de "canetas" aplicadas?',
    '235': 'Quantidade de "chapeus" aplicadas?',
    '236': 'Quantidade de "meia-luas" aplicadas?',
    '237': 'Quantidade de gols marcados?',
    '312': 'Selecione o tipo de encontro/ aula prevista',
    '190': 'Em relação a previsão a atividade foi:',
    '337': 'Selecione a aula Aplicada (alterada)',
}

# Extraindo os campos desejados
rows = []
for submission in submissions:
    answers = submission.get('answers', {})
    
    row = {}
    for campo_id, campo_nome in campos_relevantes.items():
        row[campo_nome] = clean_html(answers.get(campo_id, {}).get('answer', ''))

    # Debug: Verificar a linha extraída
    print(row)

    # Aplicar as condições especificadas
    tipo_encontro = row['Selecione o tipo de encontro/ aula prevista']
    aula_aplicada = row['Selecione a aula Aplicada (alterada)']
    atividade_prevista = row['Em relação a previsão a atividade foi:']

    if (tipo_encontro == 'FESTIVAL ESPORTIVO - FESTIVAL DE FUTEBOL DE RUA' and atividade_prevista == 'Aplicada') or \
       (atividade_prevista == 'Alterada' and aula_aplicada == 'FESTIVAL ESPORTIVO - FESTIVAL DE FUTEBOL DE RUA'):
        rows.append(row)

# Crie um DataFrame com os dados novos
df_new = pd.DataFrame(rows)

# Verificar se todas as colunas relevantes estão presentes
colunas_relevantes = list(campos_relevantes.values())
colunas_faltantes = [col for col in colunas_relevantes if col not in df_new.columns]
if colunas_faltantes:
    raise KeyError(f"As seguintes colunas estão faltando em df_new: {colunas_faltantes}")

# Selecionar apenas as colunas relevantes
df_new_relevante = df_new[colunas_relevantes]

# Caminho para salvar a nova planilha
output_file = r"C:\Users\FDR Thay\Downloads\tabelaFestivais.xlsx"

# Carregar a planilha existente se houver
if os.path.exists(output_file):
    df_existing = pd.read_excel(output_file)

    # Concatenar os dados existentes com os novos dados
    df_combined = pd.concat([df_existing, df_new_relevante])

    # Redefinir o índice para evitar problemas com rótulos duplicados
    df_combined.reset_index(drop=True, inplace=True)
else:
    df_combined = df_new_relevante

# Salvar a planilha combinada
df_combined.to_excel(output_file, index=False)
print(f"Nova planilha criada e salva em {output_file}")