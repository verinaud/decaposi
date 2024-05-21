from tempo_aleatorio import TempoAleatorio
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support import expected_conditions as EC
import sys

class SelecionaUnidade:
    def __init__(self, navegador, unidade):
        self.navegador = navegador
        self.unidade = unidade
        unidade_atual = navegador.find_element(By.XPATH,'/html/body/div[1]/nav/div/div[3]/div[2]/div[3]/div/a')
        if unidade_atual.text == unidade:
            print("Já está na unidade correta")
        else:
            self.seleciona_unidade()

    def seleciona_unidade(self):
        tempo = TempoAleatorio()
        try:
            WebDriverWait(self.navegador, 10).until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'div.nav-item:nth-child(3) > div:nth-child(1) > a:nth-child(1)'))).click()
            time.sleep(tempo.tempo_aleatorio())  # Simula comportamento humano

            WebDriverWait(self.navegador, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="selInfraOrgaoUnidade"]'))).send_keys(
                'MGI')  # Escolhe Unidade
            time.sleep(tempo.tempo_aleatorio())  # Simula comportamento humano

        except TimeoutException:
            print("Elemento não encontrado ou não clicável dentro do tempo limite.")

        rows = self.navegador.find_elements(By.XPATH, '/html/body/div[1]/div/div[2]/form/div[3]/table/tbody/tr')

        found = False

        for row in rows[1:]:
            try:

                second_column_text = row.find_element(By.XPATH, "./td[2]").text

                if self.unidade in second_column_text:
                    print(f"Texto '{self.unidade}' encontrado na linha!")

                    first_column_radiobutton = row.find_element(By.XPATH, "./td[1]//input[@type='radio']")
                    is_checked = first_column_radiobutton.get_attribute('checked')

                    if is_checked:

                        self.navegador.find_element(By.XPATH,
                                                    '/html/body/div[1]/nav/div/div[3]/div[2]/div[4]/a/img').click()
                    else:

                        self.navegador.execute_script("arguments[0].click();", first_column_radiobutton)

                    found = True
                    break
                else:
                    pass

            except NoSuchElementException as e:
                print("Erro ao processar uma linha: ", str(e))
                continue

        if not found:
            print(f"Texto '{self.unidade}' não foi encontrado em nenhuma linha!")
        time.sleep(tempo.tempo_aleatorio())  # Simula comportamento humano