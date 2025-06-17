import logging
import time
import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

# Configura o logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def configurar_navegador():
    try:
        logging.info("Configurando o navegador Chrome.")
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-notifications")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])  # Desativa logs do Chrome
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        return driver
    except Exception as e:
        logging.error(f"Erro ao configurar o navegador: {e}")
        return None


def realizar_login(driver):
    try:
        logging.info("Acessando o portal do professor.")
        driver.get("https://guairaca.jacad.com.br/prof/professor.efetuaLogin.logic")

        wait = WebDriverWait(driver, 15)
        
        logging.info("Preenchendo login e senha.")
        wait.until(EC.presence_of_element_located((By.ID, "login_prof"))).send_keys("123123")
        driver.find_element(By.ID, "senha").send_keys("123123")

        logging.info("Clicando no botão de login.")
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn-success"))).click()

        # Verificação mais robusta do login
        wait.until(lambda d: "efetuaLogin" not in d.current_url)
        logging.info(f"Login realizado com sucesso. URL atual: {driver.current_url}")
        return True

    except TimeoutException:
        logging.error("Tempo limite excedido durante o login.")
        return False
    except Exception as e:
        logging.error(f"Erro inesperado durante o login: {e}")
        return False


def navegar_para_tentativas(driver):
    try:
        wait = WebDriverWait(driver, 15)
        

        try:
            logging.info("Clicando em 'Digital Uniguairacá'.")
            botao = driver.find_element(By.XPATH, "//button[contains(., 'Digital UniGuairacá')]")
            botao.click()

        except:
            logging.info("Clicando em 'Digital Uniguairacá'.")
            botao = driver.find_element(By.XPATH, "//button[.//small[contains(text(), 'Digital UniGuairacá')]]")
            botao.click()

            pass

        
        # Espera pela nova aba e muda o foco
        wait.until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(driver.window_handles[-1])
        
        logging.info("Aguardando carregamento da página do Moodle.")
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
        
        # Tentativa mais robusta de localizar a matéria
        try:
            logging.info("Procurando a matéria 'Teste'.")
            materia_xpath = '//a[contains(@href, "course/view.php") and contains(., "Teste")]'
            materia = wait.until(EC.element_to_be_clickable((By.XPATH, materia_xpath)))
            materia.click()
        except TimeoutException:
            logging.warning("Não encontrou a matéria pelo nome, tentando pelo ID...")
            materia_xpath = '//*[@id="course-info-container-7414-6"]/div/div/a'
            wait.until(EC.element_to_be_clickable((By.XPATH, materia_xpath))).click()
        
        # Click no popup
        
        time.sleep(15)
        logging.info("Clicando no popup'.")
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="page-course-view-tiles"]/div[7]/div/div/div[3]/button[1]'))).click()
        
        logging.info("Clicando no botão 'Tarefas'.")
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sectionlink-4"]'))).click()
        
        # Tentativa mais robusta de localizar a tarefa/atividade
        try:
            logging.info("Procurando a atividade pelo nome 'Questionário 01 - Unidade 01'")
            atividade_xpaths = [
                '//a[contains(@href, "mod/quiz/view.php") and contains(., "Questionário 01 - Unidade 01")]',  # Pelo nome completo
                '//a[contains(@href, "mod/quiz/view.php") and contains(., "Questionário 01")]',  # Pelo nome parcial
                '//a[contains(@href, "mod/quiz/view.php")]',  # Qualquer link de quiz
                '//a[contains(@class, "completeonmanual")]'  # Pela classe específica
            ]
            
            atividade = None
            for xpath in atividade_xpaths:
                try:
                    atividade = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                    logging.info(f"Atividade encontrada usando XPath: {xpath}")
                    break
                except TimeoutException:
                    continue
            
            if atividade:
                # Rolando até o elemento para garantir visibilidade
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", atividade)
                time.sleep(0.5)  # Pequena pausa para o scroll
                atividade.click()
            else:
                raise NoSuchElementException("Nenhum XPath encontrou a atividade")
                
        except Exception as e:
            logging.warning(f"Não encontrou a atividade pelos XPaths padrões: {e}")
            logging.info("Tentando pelo ID específico...")
            try:
                atividade = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, '//a[contains(@href, "mod/quiz/view.php") and contains(@href, "id=196947")]')))
                atividade.click()
            except:
                logging.error("Falha ao localizar a atividade mesmo pelo ID")
                raise
        
        logging.info("Acessando tentativas")
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="region-main"]/div[2]/div[2]/a'))).click()
        
        return True
        
    except TimeoutException:
        logging.error("Tempo limite excedido durante a navegação.")
        return False
    except NoSuchElementException:
        logging.error("Elemento não encontrado durante a navegação.")
        return False
    except Exception as e:
        logging.error(f"Erro inesperado durante a navegação: {e}")
        return False


def extrair_dados_tabela(driver):
    try:
        wait = WebDriverWait(driver, 20)
        logging.info("Localizando a tabela de tentativas.")
        
        # Espera até que a tabela esteja completamente carregada
        tabela = wait.until(EC.presence_of_element_located((By.ID, "attempts")))
        wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "#attempts tbody tr.gradedattempt")) > 0)
        
        # Extrair cabeçalhos (ignorando colunas de seleção e imagem)
        cabecalhos = []
        ths = tabela.find_elements(By.CSS_SELECTOR, "thead th.header")[2:]  # Pula as duas primeiras colunas
        for th in ths:
            texto = th.text.replace("Ordenar por", "").replace("Ascendente", "").strip()
            if texto and "Selecionar tudo" not in texto:
                cabecalhos.append(texto.split('\n')[0])
        
        # Adicionar cabeçalho para Nome Completo
        cabecalhos = ["Nome Completo"] + cabecalhos
        
        # Extrair linhas de dados
        linhas = []
        trs = tabela.find_elements(By.CSS_SELECTOR, "tbody tr.gradedattempt")
        
        for tr in trs:
            # Extrair nome completo (coluna c2)
            nome_completo = tr.find_element(By.CSS_SELECTOR, "td.cell.c2.bold").text.split('\n')[0].strip()
            
            # Extrair os demais dados (ignorando as duas primeiras colunas)
            tds = tr.find_elements(By.CSS_SELECTOR, "td.cell")[2:]  # Pula checkbox e imagem
            linha = [nome_completo] + [td.text.strip().replace('\n', ' ') for td in tds if td.text.strip()]
            
            if len(linha) == len(cabecalhos):
                linhas.append(linha)
        
        if not linhas:
            logging.warning("Nenhum dado de tentativa encontrado na tabela.")
            return None
        
        return pd.DataFrame(linhas, columns=cabecalhos)
        
    except Exception as e:
        logging.error(f"Erro ao extrair dados da tabela: {e}")
        return None

def salvar_dados_excel(df, excel_path):
    try:
        if df is None or df.empty:
            logging.warning("Nenhum dado para salvar.")
            return False
        
        # Selecionar e renomear colunas relevantes
        colunas_map = {
            'Nome Completo': 'Nome_Completo',
            'Endereço de email': 'Email',
            'Estado': 'Estado',
            'Iniciado em': 'Iniciado_em',
            'Completo': 'Completo',
            'Tempo utilizado': 'Tempo_utilizado',
            'Avaliar/10,00': 'Nota',
            'Q. 1 /10,00': 'Questao_1'
        }
        
        # Filtrar colunas existentes
        colunas_finais = {k: v for k, v in colunas_map.items() if k in df.columns}
        df = df[list(colunas_finais.keys())].rename(columns=colunas_finais)
        
        # Salvar no Excel
        if os.path.exists(excel_path):
            with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name='DADOS', index=False)
        else:
            df.to_excel(excel_path, sheet_name='DADOS', index=False)
        
        logging.info(f"Dados salvos com sucesso em {excel_path}")
        return True
        
    except Exception as e:
        logging.error(f"Erro ao salvar dados no Excel: {e}")
        return False
    
if __name__ == "__main__":
    driver = configurar_navegador()
    if not driver:
        exit(1)
    
    try:
        if realizar_login(driver) and navegar_para_tentativas(driver):
            excel_path = r'C:\Users\enzog\OneDrive\Área de Trabalho\Projeto\validacao_atividades.xlsx'
            dados = extrair_dados_tabela(driver)
            if dados is not None:
                salvar_dados_excel(dados, excel_path)
    except Exception as e:
        logging.error(f"Erro fatal durante a execução: {e}")
    finally:
        time.sleep(2)
        driver.quit()
        logging.info("Navegador fechado.")
