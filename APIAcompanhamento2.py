import requests
import pandas as pd
import os
from datetime import datetime


# Configurações
form_id = '241572995863674'
api_key = '26bb906630c658694f0c3cec6961b9a5'

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


# Função para limpar HTML e formatar respostas
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
        'Projeto': clean_html(answers.get('539', {}).get('answer', '')),
        'Cidade': clean_html(answers.get('440', {}).get('answer', '')),
        'Nome Completo': clean_html(answers.get('403', {}).get('answer', '')),
        'Turno': clean_html(answers.get('402', {}).get('answer', '')),
        'Idade': clean_html(answers.get('538', {}).get('answer', '')),
        'Local de execução': clean_html(answers.get('441', {}).get('answer', '')),
        '01 - A partir da sua participação no projeto, você passou a praticar esporte com mais frequência?': clean_html(answers.get('386', {}).get('answer', '')),
        '02 - Você aprendeu sobre  Fair Play?': clean_html(answers.get('522', {}).get('answer', '')),
        '03 - Marque as opções de Fair Play que você passou a praticar com sua participação no projeto? Assinale até 3 opções que mais te identifica.': clean_html(answers.get('523', {}).get('prettyFormat', '')),
        '04 - Você aprendeu sobre protagonismo?': clean_html(answers.get('524', {}).get('answer', '')),
        '05 - O que você passou a entender sobre protagonismo? Assinale até 3 opções que mais te identificam.': clean_html(answers.get('525', {}).get('prettyFormat', '')),
        '06 - Quanto ao relacionamento com seus pais e responsáveis, mudou depois que você passou a participar do Projeto?': clean_html(answers.get('526', {}).get('prettyFormat', '')),
        '07 - Quais dos valores abaixo, você passou a praticar após sua participação no Projeto? Assinale até 3 opções que mais te identifica.': clean_html(answers.get('527', {}).get('prettyFormat', '')),
        '08 - Você passou a gostar de futebol?': clean_html(answers.get('514', {}).get('answer', '')),
        '09 - Você tem alguma dificuldade em praticar a modalidade Futebol de Rua?': clean_html(answers.get('536', {}).get('answer', '')),
        '10 - Você já participou de algum campeonato ou evento esportivo?': clean_html(answers.get('515', {}).get('answer', '')),
        '11 - Além de participar das atividades do Projeto Futebol de Rua pela Educação, você pratica outra modalidade esportiva? Assinale até 3 opções.': clean_html(answers.get('516', {}).get('prettyFormat', '')),
        '12 - Quais dos dribles abaixo você aprendeu no Projeto Futebol de Rua pela Educação?': clean_html(answers.get('517', {}).get('prettyFormat', '')),
        '13 - O Projeto Futebol de Rua pela Educação tem te ajudado a enfrentar alguma dificuldade no seu dia a dia?': clean_html(answers.get('537', {}).get('answer', '')),
        '14 - Assinale até 3 opções em que o Projeto tem te ajudado no seu dia a dia:': clean_html(answers.get('518', {}).get('prettyFormat', '')),
        '15 - Como o Projeto Futebol de Rua Pela Educação tem contribuído com o seu desenvolvimento? Marque 3 das opções mais importantes para você:': clean_html(answers.get('519', {}).get('prettyFormat', '')),
        '16 - Quais dos temas a seguir você aprendeu no Projeto até o momento?': clean_html(answers.get('520', {}).get('prettyFormat', '')),
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
                      '01 - A partir da sua participação no projeto, você passou a praticar esporte com mais frequência?',
                      '02 - Você aprendeu sobre  Fair Play?',
                      '03 - Marque as opções de Fair Play que você passou a praticar com sua participação no projeto? Assinale até 3 opções que mais te identifica.',
                      '04 - Você aprendeu sobre protagonismo?',
                      '05 - O que você passou a entender sobre protagonismo? Assinale até 3 opções que mais te identificam.',
                      '06 - Quanto ao relacionamento com seus pais e responsáveis, mudou depois que você passou a participar do Projeto?',
                      '07 - Quais dos valores abaixo, você passou a praticar após sua participação no Projeto? Assinale até 3 opções que mais te identifica.',
                      '08 - Você passou a gostar de futebol?',
                      '09 - Você tem alguma dificuldade em praticar a modalidade Futebol de Rua?',
                      '10 - Você já participou de algum campeonato ou evento esportivo?',
                      '11 - Além de participar das atividades do Projeto Futebol de Rua pela Educação, você pratica outra modalidade esportiva? Assinale até 3 opções.',
                      '12 - Quais dos dribles abaixo você aprendeu no Projeto Futebol de Rua pela Educação?',
                      '13 - O Projeto Futebol de Rua pela Educação tem te ajudado a enfrentar alguma dificuldade no seu dia a dia?',
                      '14 - Assinale até 3 opções em que o Projeto tem te ajudado no seu dia a dia:',
                      '15 - Como o Projeto Futebol de Rua Pela Educação tem contribuído com o seu desenvolvimento? Marque 3 das opções mais importantes para você:',
                      '16 - Quais dos temas a seguir você aprendeu no Projeto até o momento?',
                    ]

# Selecionar apenas as colunas relevantes
df_new_relevante = df_new[colunas_relevantes]

# Caminho para salvar a nova planilha
output_file = r"C:\Users\thays\Downloads\TabelaAcompanhamento2.xlsx"

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
    df_combined = df_combined.drop_duplicates(subset=['Submission Date', 'Nome Completo', 'Idade:', 'Cidade'], keep='first')
    # Debug: Mostrar o número de registros combinados após a remoção de duplicatas
    print(f"Número de registros combinados após a remoção de duplicatas: {len(df_combined)}")
else:
    df_combined = df_new_relevante

# Salvar a nova tabela em um arquivo Excel
df_combined.to_excel(output_file, index=False)

print(f"Nova planilha criada e salva em {output_file}")
