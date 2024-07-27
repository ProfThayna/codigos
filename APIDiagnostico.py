import requests
import pandas as pd
from bs4 import BeautifulSoup
import os
from datetime import datetime

# Defina sua chave de API e o código do formulário
api_key = '0930cd99221b6e7303b04a15ef5b46ca'
form_id = '241193203415648'


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

    # Tratamento das respostas para delimitá-las por vírgulas
    answer_list = answers.get('answer', [])
    if isinstance(answer_list, list):
        formatted_answers = clean_html(answer_list)  # Aplicar a função para listas
    else:
        formatted_answers = clean_html(answer_list)  # Se não for lista, limpe como string simples

    row = {
        'Submission Date': submission_date_formatted,
        'Patrocinador': clean_html(answers.get('439', {}).get('answer', '')),
        'Cidade': clean_html(answers.get('440', {}).get('answer', '')),
        'NOME COMPLETO': clean_html(answers.get('403', {}).get('answer', '')),
        'ESCOLA/NÚCLEO': clean_html(answers.get('455', {}).get('answer', '')),
        'TURNO': clean_html(answers.get('402', {}).get('answer', '')),
        'IDADE': clean_html(answers.get('453', {}).get('answer', '')),
        'Local de execução': clean_html(answers.get('441', {}).get('answer', '')),
        'Você gosta de praticar esportes?': clean_html(answers.get('458', {}).get('answer', '')),
        'Quais esportes você mais gosta?': clean_html(answers.get('452', {}).get('prettyFormat', '')),
        'Você considera que tem alguma dificuldade para praticar esportes?': clean_html(answers.get('428', {}).get('answer', '')),
        'Você sabe o que é Fair Play?': clean_html(answers.get('459', {}).get('answer', '')),
        'Marque as opções de Fair Play que você já pratica:': clean_html(answers.get('451', {}).get('prettyFormat', '')),
        'Você sabe o que é protagonismo?': clean_html(answers.get('460', {}).get('answer', '')),
        'O que você considera ser protagonismo?': clean_html(answers.get('461', {}).get('prettyFormat', '')),
        'Quais valores você tem praticado até o momento?': clean_html(answers.get('463', {}).get('prettyFormat', '')),
        'Você gosta de futebol?': clean_html(answers.get('514', {}).get('answer', '')),
        'Você conhece os dribles do futebol?': clean_html(answers.get('517', {}).get('prettyFormat', '')),
        'Quando você enfrenta dificuldades no dia a dia, como acha que pode superá-las? Marque as opções que você consegue fazer:': clean_html(answers.get('518', {}).get('prettyFormat', '')),
        'Como o Projeto Futebol de Rua Pela Educação pode ajudar no seu desenvolvimento? Marque as opções mais importantes para você:': clean_html(answers.get('519', {}).get('prettyFormat', '')),
        'Quais dos temas a seguir você já sabe ou já estudou antes de participar do projeto?': clean_html(answers.get('520', {}).get('prettyFormat', ''))
    }
    rows.append(row)

# Crie um DataFrame com os dados novos
df_new = pd.DataFrame(rows)

# Colunas relevantes para a nova tabela
colunas_relevantes = ['Submission Date',
                      'Patrocinador',
                      'Cidade',
                      'NOME COMPLETO',
                      'ESCOLA/NÚCLEO',
                      'TURNO',
                      'IDADE',
                      'Local de execução',
                      'Você gosta de praticar esportes?',
                      'Quais esportes você mais gosta?',
                      'Você considera que tem alguma dificuldade para praticar esportes?',
                      'Você sabe o que é Fair Play?',
                      'Marque as opções de Fair Play que você já pratica:',
                      'Você sabe o que é protagonismo?',
                      'O que você considera ser protagonismo?',
                      'Quais valores você tem praticado até o momento?',
                      'Você gosta de futebol?',
                      'Você conhece os dribles do futebol?',
                      'Quando você enfrenta dificuldades no dia a dia, como acha que pode superá-las? Marque as opções que você consegue fazer:',
                      'Como o Projeto Futebol de Rua Pela Educação pode ajudar no seu desenvolvimento? Marque as opções mais importantes para você:',
                      'Quais dos temas a seguir você já sabe ou já estudou antes de participar do projeto?']

# Selecionar apenas as colunas relevantes
df_new_relevante = df_new[colunas_relevantes]

# Caminho para salvar a nova planilha
output_file = r"C:\Users\FDR Thay\Downloads\diagnostico.xlsx"


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
    df_combined = df_combined.drop_duplicates(subset=['Submission Date', 'NOME COMPLETO', 'IDADE', 'Cidade'], keep='first')
    # Debug: Mostrar o número de registros combinados após a remoção de duplicatas
    print(f"Número de registros combinados após a remoção de duplicatas: {len(df_combined)}")
else:
    df_combined = df_new_relevante

# Criar a coluna 'Projeto' e mover o nome do projeto para essa coluna
df_combined['Projeto'] = df_combined.apply(lambda row: row['Patrocinador'] if row['Patrocinador'] in ['3TEMPOS', 'CULTURA EM CAMPO', 'JOGANDO JUNTOS', 'E-FDR', 'BIT MAKERS', 'JOVEM APRENDIZ'] else None, axis=1)
df_combined.loc[df_combined['Projeto'].notnull(), 'Patrocinador'] = None


# Salvar a nova tabela em um arquivo Excel
df_combined.to_excel(output_file, index=False)

print(f"Nova planilha criada e salva em {output_file}")
