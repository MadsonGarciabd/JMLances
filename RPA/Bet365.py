from selenium import webdriver
from RPA.Global import Utilitarios as utils

class Rb_Bet365:
    def __init__(self):
        pass

    def busca_dados (self):
        self.driver = webdriver.Chrome(executable_path=f"{utils.caminho_local()}\\arqs\\chromedriver.exe")
        self.driver.get('https://www.bet365.com/#/HO/')


        pass