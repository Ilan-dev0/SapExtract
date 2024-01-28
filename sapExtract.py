import requests

odata_url = "https://<seu-servico-odata>/EntitySet"

username = "exemplo_usuario"
password = "exemplo_senha"

response = requests.get(odata_url, auth=(username, password))

if response.status_code == 200:

    data = response.json()
    print("Dados recebidos:", data)
else:

    print("Erro:", response.status_code, response.text)

## Extrair como XLSX e .DAT
