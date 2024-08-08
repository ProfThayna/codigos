import requests
import pandas as pd
from bs4 import BeautifulSoup
import os
from datetime import datetime

# Defina sua chave de API e o código do formulário
api_key = ''
form_id = '240935142374657'

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
    submission_date_formatted = format_submission_date(submission.get('created_at', ''))
    answers = submission.get('answers', {})


    row = {
        'Submission Date': submission_date_formatted,
        'Patrocinador': clean_html(answers.get('439', {}).get('answer', '')),
        'Projeto': clean_html(answers.get('448', {}).get('answer', '')),
        'Cidade': clean_html(answers.get('440', {}).get('answer', '')),
        'Nome Completo': clean_html(answers.get('403', {}).get('answer', '')),
        'Turno': clean_html(answers.get('402', {}).get('answer', '')),
        'Idade': clean_html(answers.get('443', {}).get('answer', '')),
        'Local de execução': clean_html(answers.get('441', {}).get('answer', '')),
        '1 - Tem frequentado a escola diariamente?': clean_html(answers.get('386', {}).get('answer', '')),
        '3 - Em quais matérias/ disciplinas você considera ter mais dificuldade?  Assinale como VERDADEIRO apenas aquelas em que acredita que precisa melhorar:': clean_html(answers.get('452', {}).get('prettyFormat', '')),
        '4 - Como você considera seu relacionamento com os colegas de escola?': clean_html(answers.get('428', {}).get('answer', '')),
        '5 - Como o Projeto pode te ajudar a melhorar na escola? Marque as opções que você acha que se aplicam a você:': clean_html(answers.get('451', {}).get('prettyFormat', '')),
        '9 - Você consegue reconhecer o tipo de sentimento que está sentindo, como estar feliz, triste, com raiva ou com medo?': clean_html(answers.get('435', {}).get('answer', '')),
        '10 - Como é a sua atitude quando você tem um problema ou um conflito para resolver? Marque as atitudes que você tem em relação a isso': clean_html(answers.get('436', {}).get('prettyFormat', '')),

    }
    rows.append(row)

# Crie um DataFrame com os dados novos
df_new = pd.DataFrame(rows)

# Colunas relevantes para a nova tabela
colunas_relevantes = ['Submission Date',
                      'Patrocinador',
                      'Projeto',
                      'Cidade',
                      'Nome Completo',
                      'Turno',
                      'Idade',
                      'Local de execução',
                      '1 - Tem frequentado a escola diariamente?',
                      '3 - Em quais matérias/ disciplinas você considera ter mais dificuldade?  Assinale como VERDADEIRO apenas aquelas em que acredita que precisa melhorar:',
                      '4 - Como você considera seu relacionamento com os colegas de escola?',
                      '5 - Como o Projeto pode te ajudar a melhorar na escola? Marque as opções que você acha que se aplicam a você:',
                      '9 - Você consegue reconhecer o tipo de sentimento que está sentindo, como estar feliz, triste, com raiva ou com medo?',
                      '10 - Como é a sua atitude quando você tem um problema ou um conflito para resolver? Marque as atitudes que você tem em relação a isso',
                      ]

# Selecionar apenas as colunas relevantes
df_new_relevante = df_new[colunas_relevantes]

# Caminho para salvar a nova planilha
output_file = r"C:\Users\FDR Thay\Downloads\TabelaDiagnostico1.xlsx"


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
    df_combined = df_combined.drop_duplicates(subset=['Submission Date', 'Nome Completo', 'Idade', 'Cidade'], keep='first')
    # Debug: Mostrar o número de registros combinados após a remoção de duplicatas
    print(f"Número de registros combinados após a remoção de duplicatas: {len(df_combined)}")
else:
    df_combined = df_new_relevante


# Salvar a nova tabela em um arquivo Excel
df_combined.to_excel(output_file, index=False)

print(f"Nova planilha criada e salva em {output_file}")
