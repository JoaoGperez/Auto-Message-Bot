import time
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import quote

"""
Função para carregar a planilha e extrair os dados
"""
def carregar_planilha(caminho_arquivo, pagina):
    planilha = load_workbook(caminho_arquivo)
    
    # Verificar se a aba existe
    if pagina not in planilha.sheetnames:
        raise ValueError(f"A aba '{pagina}' não existe. As abas disponíveis são: {', '.join(planilha.sheetnames)}")

    pagina_clientes = planilha[pagina]
    contatos = []

    for linha in pagina_clientes.iter_rows(min_row=2):
        nome = linha[0].value
        telefone = linha[1].value
        vencimento = linha[2].value

        vencimento_formatado = (
            vencimento.strftime('%d/%m/%Y') if vencimento else "Data de vencimento não especificada."
        )

        contatos.append({
            "nome": nome,
            "telefone": telefone,
            "vencimento": vencimento_formatado
        })

    return contatos


"""
Função para configurar o WebDriver
Inicializa o Selenium e abre o WhatsApp Web no Chrome.
"""
def configurar_driver():
    driver = webdriver.Chrome()
    driver.get("https://web.whatsapp.com")

    print("Escaneie o código QR code.")
    time.sleep(30)
    return driver


""" 
Função para enviar mensagens
"""


def enviar_mensagens(driver, contatos, mensagem_template):
    for contato in contatos:
        mensagem = mensagem_template.format(**contato)
        telefone = str(contato['telefone'])

        # Gerar links do WhatsApp
        url = f"https://wa.me/{telefone}?text={quote(mensagem)}"
        driver.get(url)
        time.sleep(5)

        try:
            # Esperar até que o campo de mensagem esteja disponível
            campo_mensagem = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//div[@title="Mensagem"]'))
            )
            campo_mensagem.send_keys(Keys.ENTER)
            print(f"Mensagem enviada para {contato['nome']} ({telefone}).")
        except Exception as e:
            print(f"Erro ao enviar mensagem para {contato['nome']} ({telefone}): {e}")
"""
Main
"""
if __name__ == "__main__":
    print("Digite o caminho do arquivo.")
    print(r"(ex: C:\Users\seu-usuario\Desktop\contatos.xlsx)")
    caminho_do_arquivo = input("\n")

    print("Especifique a página da planilha onde estão os contatos.")
    print("ex: Pagina1 ou Sheet1")
    pagina = input("\n")

    mensagem_template = input("Digite a mensagem com placeholders ({nome}, {vencimento}): ")

    contatos = carregar_planilha(caminho_do_arquivo, pagina)
    driver = configurar_driver()

    try:
        enviar_mensagens(driver, contatos, mensagem_template)
    finally:
        driver.quit()