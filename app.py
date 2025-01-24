import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

# Carregar planilha e ler dados.
def carregar_planilha(caminho_arquivo, pagina):
    try:
        dados = pd.read_excel(caminho_arquivo, sheet_name=pagina)
        print(dados.columns)  # Imprime os nomes das colunas para verificação
        return dados[['Nome', 'Telefone']]  # Ajuste para os nomes corretos
    except Exception as e:
        print(f"Erro ao carregar planilha: {e}")
        return None


# Configurar e inicializar o Selenium WebDriver.
def iniciar_webdriver(caminho_driver):
    options = Options()
    options.add_argument("--start-maximized")

    service = Service(caminho_driver)

    driver = webdriver.Chrome(service=service, options=options)
    return driver

# Função para logar no WhatsApp Web
def logar_whatsapp(driver):
    driver.get("https://web.whatsapp.com")
    print("Escaneie o QR Code para fazer Login.")
    time.sleep(30)

# Localizar contato ou número.
def localizar_contato(driver, numero):
    """
    Acessa diretamente o chat do número no WhatsApp Web e aguarda o chat carregar.
    
    :param driver: Instância do WebDriver.
    :param numero: Número de telefone do contato (formato internacional).
    """
    link = f"https://web.whatsapp.com/send?phone={numero}"  # Link direto para o WhatsApp Web
    driver.get(link)
    time.sleep(10)  # Espera o WhatsApp Web carregar
    
    try:
        # Espera o campo de mensagem ser carregado antes de enviar a mensagem
        campo_mensagem = driver.find_element(By.XPATH, "//div[@contenteditable='true']")
        campo_mensagem.click()  # Clica no campo de mensagem para garantir que o foco está lá
        time.sleep(2)
    except Exception as e:
        print(f"Erro ao localizar o número {numero}: {e}")


def enviar_mensagem(driver, mensagem):
    """
    Envia uma mensagem para o contato no WhatsApp Web.
    
    :param driver: Instância do WebDriver.
    :param mensagem: A mensagem a ser enviada.
    """
    try:
        # Localiza o campo de mensagem (onde o texto é digitado)
        campo_mensagem = driver.find_element(By.XPATH, "//div[@contenteditable='true']")
        campo_mensagem.send_keys(mensagem)  # Digita a mensagem
        time.sleep(2)  # Pausa para garantir que a mensagem foi digitada corretamente
        
        # Localiza o botão de enviar e clica nele
        botao_enviar = driver.find_element(By.XPATH, "//button[@data-testid='send']")
        botao_enviar.click()  # Envia a mensagem
        time.sleep(2)  # Pausa após o envio da mensagem
    except Exception as e:
        print(f"Erro ao enviar a mensagem: {e}")

# Função principal para executar a automação
def main():
    # Caminho da planilha e do ChromeDriver
    caminho_planilha = 'contatos.xlsx'
    caminho_driver = r'drivers\chromedriver.exe'

    # Carregar os dados da planilha
    dados = carregar_planilha(caminho_planilha, "Planilha1")  # Corrigido para 'caminho_planilha'
    if dados is None:
        return

    # Inicializar o WebDriver
    driver = iniciar_webdriver(caminho_driver)

    # Logar no WhatsApp Web
    logar_whatsapp(driver)

    # Iterar pelos contatos e enviar mensagens
    for _, linha in dados.iterrows():
        numero = linha['Telefone']
        mensagem = f"Olá {linha['Nome']}, sua mensagem personalizada aqui!"  # Ajuste conforme necessário
        localizar_contato(driver, numero)
        enviar_mensagem(driver, mensagem)

    # Finalizar WebDriver
    driver.quit()

# Executar o programa
if __name__ == "__main__":
    main()
