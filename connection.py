import requests

# Sua chave de API do JotForm
api_key = '26bb906630c658694f0c3cec6961b9a5'

# ID do formulário que você deseja buscar
form_id = '241193203415648'

url = f'https://api.jotform.com/form/{form_id}/submissions?apiKey={api_key}&limit=5000'

# Faz a requisição à API
response = requests.get(url)


# Verifica se a requisição foi bem-sucedida
if response.status_code == 200:
    data = response.json()
    submissions = data.get('content', [])
    
   # Conta o número total de submissões
    total_submissions = len(submissions)
    
    print(f"Número total de submissões: {total_submissions}")
else:
    print(f"Erro ao acessar a API: {response.status_code}")

