import json
import openpyxl
import requests
from datetime import datetime

# Leitura das informações da Planilha
planilha = openpyxl.load_workbook("./dados/Banco de Dados.xlsx")
paginaPlanilha = planilha["Dados"]

# Leitura das infomações do BOT
with open('./dados/config.json', 'r') as f:
    infos = json.load(f)

# Envio de mensagens:
def enviarMensagem(numero, cliente, data, preco, conteudo):
    image_url = ""
    download_link = ""

    url = f"https://graph.facebook.com/{infos['VERSION']}/{infos['PHONE_NUMBER_ID']}/messages"
    headers = {
        "Authorization": "Bearer " + infos['ACCESS_TOKEN'],
        "Content-Type": "application/json",
    }

    if conteudo == None:
        conteudo = ""

    data = {
        "messaging_product": "whatsapp",
        "recipient_type": "individual",
        "to": f"+55{numero}",
        "type": "template",
        "template": {"name": "cobranca", "language": {"code": "pt_BR"}, 
                     "components": [ 
                         {
                            "type": "header",
                            "parameters": [
                            {"type": "document", "document": {"link": download_link, "filename": "qrcode.pdf"}}
                        ]   
                         },
                         {"type": "body", 
                          "parameters": [
                            {"type": "text", "text": f"{cliente}"},
                            {"type": "text", "text": f"{data}"},   
                            {"type": "text", "text": f"{preco}"},
                            {"type": "text", "text": f"> Mensalidade" + f"{conteudo}"}  
                            ] 
                         },
                         {
                            "type": "button",
                            "sub_type": "quick_reply",
                            "index": "0",
                            "parameters": [
                                {
                                    "type": "payload",
                                    "payload": "chave_pix"  
                                }
                            ]
                         },
                         {
                            "type": "button",
                            "sub_type": "quick_reply",
                            "index": "1",
                            "parameters": [
                                {
                                    "type": "payload",
                                    "payload": "duvidas"  
                                }
                            ]
                         }
                       ]
                     }
    }

    response = requests.post(url, headers=headers, json=data)
    return response


# Enviar para os clientes:
for linha in paginaPlanilha.iter_rows(min_row=2, max_row=2):
    # clientes, número, data
    cliente = linha[0].value
    numero = linha[1].value
    data = linha[2].value
    data_formatada = data.strftime('%d/%m/%Y')
    preco = linha[3].value
    conteudo = linha[4].value
    response = enviarMensagem(numero, cliente, data_formatada, preco, conteudo)
    print(cliente, numero, data_formatada, preco, conteudo, f"Response status: {response.status_code}")
