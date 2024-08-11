from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, ElementNotVisibleException, ElementNotSelectableException
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
from datetime import datetime, timedelta
from documentos import *
import schedule
import os

cores = {
    'azul': '\033[94m',
    'amarelo': '\033[93m',
    'verde': '\033[92m',
    'ciano': '\033[96m',
    'resetar': '\033[0m'
}

def init_driver():
    chrome_options = Options()

    arguments = ['--lang=pt-BR', '--window-size=1200,1400', '--incognito','--log-level=3', '--disable-extensions']

    for argument in arguments:
        chrome_options.add_argument(argument)

    user_profile = os.getenv('USERPROFILE')
    path_to_download = os.path.join(user_profile, 'Desktop')

    chrome_options.add_experimental_option("prefs", {
        'download.default_directory': path_to_download,
        'download.directory_upgrade': True,
        'download.prompt_for_download': False,
        "profile.default_content_setting_values.notifications": 2, 
        "profile.default_content_setting_values.automatic_downloads": 1,
    })

    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 10, poll_frequency=1,
                        ignored_exceptions=[NoSuchElementException, ElementNotVisibleException, ElementNotSelectableException])
    
    return driver, wait

def main():
    driver, wait = init_driver()
    
    os.system('cls' if os.name == 'nt' else 'clear')
    print('-------------------------------------')
    print('Inicializando Bot de Cotação do Dolar')
    print(f'------------------------------------')
    print(f'{cores['amarelo']}-> Abrindo o Site de Cotação!{cores['resetar']}')
    sleep(2)

    # Consulta o site:
    site = 'https://www.melhorcambio.com/dolar-hoje'
    driver.get(site)

    print(f'{cores['amarelo']}---> Site Aberto Com Sucesso!{cores['resetar']}')
    sleep(1)

    print(f'{cores['verde']}-> Cotação Em Andamento!{cores['resetar']}')
    sleep(1)
    # Coletar o valor do dólar para o dia atual
    elemento = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="operacao"]/div/div[3]/div/input')))
    dolar = elemento.get_attribute('value')

    print(f'{cores['verde']}---> Cotação Realizada Com Sucesso!{cores['resetar']}')
    sleep(2)

    # Data da cotação
    data_atual = datetime.now().strftime('%Y-%m-%d')
    data_objeto = datetime.strptime(data_atual, '%Y-%m-%d')
    data_formatada = data_objeto.strftime('%d/%m/%Y')

    print(f'{cores['ciano']}-> Tirando Print da Cotação!{cores['resetar']}')
    sleep(1)

    # ○ Print(imagem) do site onde a cotação foi realizada
    user_profile = os.getenv('USERPROFILE')
    desktop_path = os.path.join(user_profile, 'Desktop')
    caminho_imagem = os.path.join(desktop_path, f'cotacao_{data_atual}.png')

    #caminho_imagem = f'cotacao_{data_atual}.png'
    driver.save_screenshot(caminho_imagem)

    print(f'{cores['ciano']}---> Print da Cotação Realizado Com Sucesso!{cores['resetar']}')
    sleep(1)

    # Criação de arquivo word
    gerar_word(dolar, site, data_formatada, caminho_imagem)

    # Transforme em um PDF
    converte_pdf()

# Agende a função para ser executada todos os dias às 9h00 AM
schedule.every().day.at("09:00").do(main)

main()

print('-> O Bot está programado para rodar todos os dias as 09H00 AM')

# Loop principal que verifica e executa as tarefas agendadas
while True:
    schedule.run_pending()  
  
    agora = datetime.now()
    proximo_dia = agora + timedelta(days=1)
    data_futura = datetime(proximo_dia.year, proximo_dia.month, proximo_dia.day, 9, 0)
    diferenca = data_futura - agora

    horas_restantes = diferenca.days * 24 + diferenca.seconds // 3600
    minutos_restantes = (diferenca.seconds % 3600) // 60

    print(f'---> A Próxima Cotação Será Realizada Novamente Daqui {horas_restantes} horas e {minutos_restantes} minutos')
    sleep(3600) 