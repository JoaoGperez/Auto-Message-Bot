"""
Fazer um bot que envia mensagens automaticas de cobrança para cada contato do whatsapp que esta na planilha.
Como automatizar esse processo?
    - Onde está feito (Versão web do Whatsapp)
* Tecnologias necessarias:
    - Bibliteca: webbrowser (Acesso ao site)
    - Biblioteca: openpyxl (Automatizar leitura dos dados da planilha)
    - Link do whatsapp
"""

# Ler a planilha e armazenar os dados. Nome, Telefone e data de Vencimento
# Criar links personalizados do whatsapp e enviar para cada cliente

import openpyxl
from urllib.parse import quote
import webbrowser


planilha = openpyxl.load_workbook('contatos.xlsx')
pagina_clientes = planilha ['Planilha1']

# Extrair, nome, telefone e data de vencimento
for linha in pagina_clientes.iter_rows(min_row = 2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    
    mensagem = f'Olá, {nome}. Seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. Favor pagar via pix.'
    link_mensagem = f'https://whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

