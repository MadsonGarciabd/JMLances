from selenium import webdriver

class Rb_Bet365:
    def __init__(self):
        pass

    def busca_dados (self):
        self.driver = webdriver.Chrome(chrome_options=chromeOptions, executable_path=f"{utils.caminho_local()}\\arqs\\chromedriver.exe")

        pass