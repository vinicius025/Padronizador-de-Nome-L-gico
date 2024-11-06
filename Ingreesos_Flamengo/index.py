from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Caminho para o OperaDriver (ChromeDriver compatível com o Opera)
opera_driver_path = r'C:\Users\contv\OneDrive\Documentos\OperaWebDriver\operadriver_win64\operadriver.exe'

# Configurar as opções do Chrome para usar o Opera
opera_options = webdriver.ChromeOptions()
opera_options.binary_location = r'C:\Users\contv\AppData\Local\Programs\Opera GX\launcher.exe'  # Caminho para o executável do Opera

# Inicializar o driver do Chrome com as opções configuradas
driver = webdriver.Chrome(executable_path=opera_driver_path, options=opera_options)

# Navegar até o site do Flamengo Ingressos
driver.get('https://flamengo.superingresso.com.br/#!/home')

# Encontrar e clicar no botão para comprar o ingresso
comprar_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/main/div[2]/div[2]/div/section/div[3]/div[1]/div[3]/button')))
comprar_button.click()

# Adicionar ao carrinho
# (você precisará inspecionar o elemento na página para encontrar o ID, classe, etc.)
# por exemplo:
# carrinho_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'id_do_botao_do_carrinho')))
# carrinho_button.click()

# Agora, você pode continuar interagindo com o site conforme necessário para completar a compra
# Lembre-se de sempre tratar exceções e erros possíveis durante a execução do seu script

