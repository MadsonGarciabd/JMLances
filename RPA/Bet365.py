from selenium import webdriver
from RPA.Global import Utilitarios as utils

class Rb_Bet365:
    def __init__(self):
        pass

    def busca_dados (self):
        self.driver = webdriver.Chrome(executable_path=f"{utils.caminho_local()}\\chromedriver.exe")
        self.driver.get('https://www.bet365.com/#/AC/B1/C1/D56/E0/F2/J0/Q1/F^24/')
        nomes_times_tot = self.driver.find_elements_by_class_name('gll-Market')
        colunas_valores_apostas = self.driver.find_elements_by_class_name('ufm-MarketC5OddsSwitchNoHeightGrow')
        for times, coluna in zip (nomes_times_tot, colunas_valores_apostas):

            print(coluna.text)
            print(times.text)





        pass