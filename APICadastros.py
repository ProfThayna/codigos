import requests
import re
import pandas as pd
from bs4 import BeautifulSoup
import os
from datetime import datetime

# Defina sua chave de API e o código do formulário
api_key = ''
form_id = '232916171419659'


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


def clean_html(raw_html):
    if isinstance(raw_html, list):
        return ', '.join(raw_html)  # Concatenar lista com vírgulas
    elif isinstance(raw_html, str):
        return raw_html.strip()
    else:
        return ''

# Função para formatar a data
def format_submission_date(submission_date_str):
    # Converter a string para um objeto datetime
    submission_date = datetime.strptime(submission_date_str, '%Y-%m-%d %H:%M:%S')
    # Formatar a data no formato desejado "abr. 29, 2024"
    return submission_date.strftime('%b. %d, %Y')

# Extraindo os campos desejados
rows = []
for submission in submissions:
   # submission_date_formatted = format_submission_date(submission.get('created_at', ''))
   # update_date_formatted = format_submission_date(submission.get('updated_at', ''))
   # status_approval = submission.get('status')
    answers = submission.get('answers', {})

    # Tratamento das respostas para delimitá-las por vírgulas
    answer_list = answers.get('answer', [])
    if isinstance(answer_list, list):
        formatted_answers = clean_html(answer_list)  # Aplicar a função para listas
    else:
        formatted_answers = clean_html(answer_list)  # Se não for lista, limpe como string simples

    row = {
        'Submission Date': submission.get('created_at', ''),
        'Patrocinador': clean_html(answers.get('178', {}).get('answer', '')),
        'Approval Status': submission.get('status', ''),
        'Data da Última Atualização': submission.get('updated_at', ''),
        'Projeto': clean_html(answers.get('79', {}).get('answer', '')),                    
        'Nome Completo': clean_html(answers.get('80', {}).get('answer', '')),
        'Idade': clean_html(answers.get('86', {}).get('answer', '')),
        'Gênero': clean_html(answers.get('83', {}).get('answer', '')),
        'Período': clean_html(answers.get('165', {}).get('prettyFormat', '')),
        'Cidade': clean_html(answers.get('179', {}).get('answer', '')),
        'Local de Execução': clean_html(answers.get('180', {}).get('answer', '')),
        'COR/RAÇA/ETNIA': clean_html(answers.get('84', {}).get('answer', '')),
        'DATA DE INICIO DA CRIANÇA OU ADOLESCENTE NO PROJETO': clean_html(answers.get('101', {}).get('prettyFormat', '')),
        'DIA DA SEMANA': clean_html(answers.get('102', {}).get('prettyFormat', '')),
        'STATUS DO BENEFICIÁRIO': clean_html(answers.get('104', {}).get('answer', '')),
        'TURNO QUE O CADASTRADO ESTUDA': clean_html(answers.get('107', {}).get('answer', '')),
        'SÉRIE OU ANO ESCOLAR': clean_html(answers.get('108', {}).get('answer', '')),
    }
    rows.append(row)

# Crie um DataFrame com os dados novos
df_new = pd.DataFrame(rows)

# Colunas relevantes para a nova tabela
colunas_relevantes = ['Submission Date',
                      'Approval Status',
                      'Data da Última Atualização',
                      'Projeto',           
                      'Nome Completo',
                      'Idade',
                      'Gênero',
                      'Período',
                      'Patrocinador',
                      'Cidade',
                      'Local de Execução',
                      'COR/RAÇA/ETNIA',
                      'DATA DE INICIO DA CRIANÇA OU ADOLESCENTE NO PROJETO',
                      'DIA DA SEMANA',
                      'STATUS DO BENEFICIÁRIO',
                      'TURNO QUE O CADASTRADO ESTUDA',
                      'SÉRIE OU ANO ESCOLAR']

# Selecionar apenas as colunas relevantes
df_new_relevante = df_new[colunas_relevantes]

# Caminho para salvar a nova planilha
output_file = r"C:\Users\FDR Thay\Downloads\tabelaCadastros.xlsx"


# Carregar a planilha existente se houver
if os.path.exists(output_file):
    df_existing = pd.read_excel(output_file)

    # Debug: Mostrar o número de registros existentes
    print(f"Número de registros existentes: {len(df_existing)}")

    # Concatenar os dados existentes com os novos dados
    df_combined = pd.concat([df_existing, df_new_relevante])

    # Redefinir o índice para evitar problemas com rótulos duplicados
    df_combined.reset_index(drop=True, inplace=True)

    # Debug: Mostrar o número de registros combinados antes da remoção de duplicatas
    print(f"Número de registros combinados antes da remoção de duplicatas: {len(df_combined)}")

    # Remover duplicatas baseando-se nas colunas específicas
    df_combined = df_combined.drop_duplicates(subset=['Submission Date', 'Nome Completo', 'Idade'], keep='first')
    # Debug: Mostrar o número de registros combinados após a remoção de duplicatas
    print(f"Número de registros combinados após a remoção de duplicatas: {len(df_combined)}")
else:
    df_combined = df_new_relevante

# Salvar a nova tabela em um arquivo Excel
df_combined.to_excel(output_file, index=False)

print(f"Nova planilha criada e salva em {output_file}")
