import os
import time
from datetime import date, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
from dotenv import load_dotenv

load_dotenv()

def get_periodos_do_mes_atual():
    hoje = date.today()
    mes = hoje.month
    ano = hoje.year
    periodo1 = (date(ano, mes, 1), date(ano, mes, 10))
    periodo2 = (date(ano, mes, 11), date(ano, mes, 20))
    proximo_mes = date(ano, mes, 28) + timedelta(days=4)
    ultimo_dia = proximo_mes - timedelta(days=proximo_mes.day)
    periodo3 = (date(ano, mes, 21), ultimo_dia)
    return [periodo1, periodo2, periodo3]

def solicitar_e_baixar_xml(driver, wait, empresa, data_inicio, data_fim):
    print(f"--- Iniciando processo para a empresa: {empresa} ---")
    print(f"Período: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Novo']"))).click()
    seletor_empresa = Select(wait.until(EC.visibility_of_element_located((By.XPATH, "//select[contains(@id, 'cboEmpresa')]"))))
    seletor_empresa.select_by_visible_text(empresa)
    seletor_origem = Select(driver.find_element(By.XPATH, "//select[contains(@id, 'cboOrigemDocumento')]"))
    seletor_origem.select_by_visible_text("Terceiro")
    formato_data = "%d/%m/%Y"
    driver.find_element(By.XPATH, "//input[contains(@id, 'txtDataEmissaoInicial')]").send_keys(data_inicio.strftime(formato_data))
    driver.find_element(By.XPATH, "//input[contains(@id, 'txtDataEmissaoFinal')]").send_keys(data_fim.strftime(formato_data))
    driver.find_element(By.XPATH, "//input[contains(@id, 'btnSolicitar')]").click()
    print("Solicitação enviada. Aguardando processamento...")
    timeout_processamento = 300
    inicio_espera = time.time()
    download_realizado = False
    while time.time() - inicio_espera < timeout_processamento:
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Pesquisar']"))).click()
            time.sleep(15)
            primeira_linha = wait.until(EC.visibility_of_element_located((By.XPATH, "//table[contains(@id, 'grdDownload')]/tbody/tr[1]")))
            status_col = primeira_linha.find_elements(By.TAG_NAME, "td")[4]
            status_atual = status_col.text.strip()
            print(f"Status atual: {status_atual}")
            if status_atual == "Processado":
                print("Status: Processado! Clicando para baixar...")
                primeira_linha.find_element(By.XPATH, ".//a[contains(@id, 'hplDownload')]").click()
                download_realizado = True
                break
        except Exception:
            print(f"Aguardando o processamento...")
    if not download_realizado:
        raise Exception(f"O processamento para a empresa {empresa} não foi concluído no tempo esperado.")
    print("Aguardando o arquivo ser salvo no disco...")
    time.sleep(10)

def painel_target_automacao_completa():
    download_path = os.path.abspath("downloads_xml")
    os.makedirs(download_path, exist_ok=True)
    options = webdriver.FirefoxOptions()
    options.set_preference("browser.download.folderList", 2)
    options.set_preference("browser.download.dir", download_path)
    options.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/zip, application/octet-stream, application/x-zip-compressed")
    driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 30)
    driver.maximize_window()
    try:
        print("Acessando o Painel Target com Firefox...")
        driver.get("https://painel.targetmob.com.br/frmLogin.aspx")

        try:
            wait.until(EC.element_to_be_clickable((By.ID, "ctl00_cphPrincipal_uscLogin_ButtonCookie"))).click()
            print("Botão de cookies aceito.")
        except Exception:
             print("Banner de cookies não encontrado (normal).")
        
        wait.until(EC.element_to_be_clickable((By.ID, "ctl00_cphPrincipal_uscLogin_btnEntrar")))
        driver.find_element(By.ID, "ctl00_cphPrincipal_uscLogin_txtLogin").send_keys(os.getenv("PAINEL_USUARIO"))
        driver.find_element(By.ID, "ctl00_cphPrincipal_uscLogin_txtSenha").send_keys(os.getenv("PAINEL_SENHA"))
        driver.find_element(By.ID, "ctl00_cphPrincipal_uscLogin_btnEntrar").click()
        print("Login realizado com sucesso.")

        seletor_produto_element = wait.until(EC.presence_of_element_located((By.ID, "ctl00_uscUsuarioLogado1_ddlProduto")))
        select_object = Select(seletor_produto_element)
        select_object.select_by_value('4')
        print("TARGET DFE selecionado.")
        
        # --- CORREÇÃO FINAL E DEFINITIVA ---
        print("Aguardando o menu lateral esquerdo recarregar...")
        menu_documentos = wait.until(EC.element_to_be_clickable((By.XPATH, "//ul[@class='sidebar-menu']//span[text()='DOCUMENTOS']")))
        print("Menu recarregado. Clicando em 'DOCUMENTOS'...")
        menu_documentos.click()
        
        print("Clicando em 'Download em Massa'...")
        submenu_download = wait.until(EC.element_to_be_clickable((By.XPATH, "//ul[contains(@class, 'menu-open')]//a[contains(., 'Download em Massa')]")))
        submenu_download.click()
        
        print("Sucesso! Chegamos na página de Download em Massa.")
        # --- FIM DA CORREÇÃO ---
        
        empresas_para_baixar = ["PROSPER ALI", "PROSPER MED", "PROSPER ES"]
        periodos = get_periodos_do_mes_atual()
        
        for empresa in empresas_para_baixar:
            for periodo in periodos:
                data_inicio, data_fim = periodo
                solicitar_e_baixar_xml(driver, wait, empresa, data_inicio, data_fim)
                print("-" * 40)
        
        print("AUTOMAÇÃO DO SITE CONCLUÍDA COM SUCESSO!")

    except Exception as e:
        print(f"Ocorreu um erro fatal durante a automação: {e}")
        driver.save_screenshot("erro_fatal_screenshot.png")
    finally:
        print("Fechando o navegador em 15 segundos...")
        time.sleep(15)
        driver.quit()

if __name__ == '__main__':
    painel_target_automacao_completa()