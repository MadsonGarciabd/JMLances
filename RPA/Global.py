import datetime
import getpass
from datetime import datetime as DateTime
import os
#import win32com.client as win32
import re
#import openpyxl
#from openpyxl.drawing.image import Image
import time
import shutil
import sys
import traceback



class Utilitarios:
    nome_aplicacao = "JM Lances"
    titulo = 'JMLances'
    versao = "1.2.5"
    executando = False

    # Retira um Mês da data informada
    @classmethod
    def subtrai_mes(cls, data_tratar=datetime.date.today()):
        try:
            if data_tratar.month == 1:
                data_auxiliar = f'{data_tratar.day}/12/{data_tratar.year - 1}'
            else:
                data_auxiliar = f'{data_tratar.day}/{data_tratar.month - 1}/{data_tratar.year}'

            data_tratar = DateTime.strptime(data_auxiliar, '%d/%m/%Y').date()

            return data_tratar
        except:
            pass

    # Soma um Mês da data informada
    @classmethod
    def soma_mes(cls, data_tratar=datetime.date.today()):
        try:
            if data_tratar.month == 12:
                data_auxiliar = f'{data_tratar.day}/01/{data_tratar.year + 1}'
            else:
                data_auxiliar = f'{data_tratar.day}/{data_tratar.month + 1}/{data_tratar.year}'

            data_tratar = DateTime.strptime(data_auxiliar, '%d/%m/%Y').date()

            return data_tratar
        except:
            pass

    # Retorna o ultimo dia do mês informado.
    @classmethod
    def ultimo_dia_mes(cls, data_tratar=datetime.date.today()):
        try:
            if data_tratar.month > 9:
                data_auxiliar = f"01/{data_tratar.month + 1}/{data_tratar.year}"
            else:
                data_auxiliar = f"01/0{data_tratar.month + 1}/{data_tratar.year}"

            data_auxiliar = DateTime.strptime(data_auxiliar, '%d/%m/%Y').date()
            data_auxiliar = data_auxiliar + datetime.timedelta(days=-1)
            return data_auxiliar
        except:
            pass

    # Retorna o primeiro dia da data informada.
    @classmethod
    def primeiro_dia_mes(cls, data_tratar=datetime.date.today()):
        try:
            if data_tratar.month > 9:
                data_auxiliar = f"01/{data_tratar.month} /{data_tratar.year}"
            else:
                data_auxiliar = f"01/0{data_tratar.month}/{data_tratar.year}"

            data_auxiliar = DateTime.strptime(data_auxiliar, '%d/%m/%Y').date()

            return data_auxiliar
        except:
            pass

    # Cria e alimenta um arquivo de log do sistema
    @classmethod
    def log_sys(cls, mensagem):
        # usuario = getpass.getuser()
        # caminho_logs = f'C:\\Users\\{usuario}\\Documents\\logs\\'
        caminho_logs = f'{cls.caminho_local()}\\logs'
        try:
            os.makedirs(caminho_logs)
        except:
            pass

        try:
            cls.nome_arquivo_log = f'{caminho_logs}\\{getpass.getuser()}-{datetime.date.today()}_{str(cls.nome_aplicacao).strip()}.log'
            # Abre o arquivo (escrita) e adiciona uma linha
            mensagem += '\n'
            arquivo = open(cls.nome_arquivo_log, 'a')
            arquivo.writelines(mensagem)
            arquivo.close()
            del arquivo
        except BaseException as e:
            print(f"Erro na criação do arquivo de log!! {e}")

    # Exclui arquivos na pasta de log que tenham mais de 7 dias
    @classmethod
    def exclui_log(cls):
        try:
            usuario = getpass.getuser()
            caminho_logs = f'C:\\Users\\{usuario}\\Documents\\logs\\'
            for nome in os.listdir(caminho_logs):
                mtime = datetime.datetime.fromtimestamp(os.path.getmtime(f'{caminho_logs}{nome}'))
                if mtime.date() < (datetime.date.today() - datetime.timedelta(days=7)):
                    os.remove(f'{caminho_logs}{nome}')
        except:
            pass

    # Cria pasta no caminho informado
    @classmethod
    def cria_pasta(cls, caminho):
        try:
            os.makedirs(caminho)
        except:
            pass

    @classmethod
    def exclui_pasta(cls, caminho):
        try:
            shutil.rmtree(caminho, True)
        except Exception as e:
            print(e)
            cls.grava_erro()
            pass

    # Retorna a extenção do arquivo informado
    @classmethod
    def retorna_extencao(cls, nome):
        try:
            explod_file = nome.split('.')
            extension = explod_file[-1]
            return extension.upper()
        except:
            pass

    # Retorna o nome do Mês referente a uma data
    @classmethod
    def retorna_nome_mes(cls, data=datetime.date.today(), idioma='ES'):
        if idioma.upper() == 'ES':
            mes_ext = {1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
                       9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'}

        return mes_ext[int(data.month)]

    # Cria um email no outlook e envia para os destinatários
    # @classmethod
    # def envia_email(cls, destinatarios, mensagem, assunto, caminho_anexo):
    #
    #     try:
    #         outlook = win32.Dispatch('outlook.application')
    #         mail = outlook.CreateItem(0)
    #         mail.To = destinatarios
    #         mail.Subject = assunto
    #         mail.Body = mensagem
    #         # mail.HTMLBody = '<h2>HTML Message body</h2>'  # this field is optional
    #
    #         # To attach a file to the email (optional):
    #         # attachment = "Path to the attachment"
    #         mail.Attachments.Add(caminho_anexo)
    #         # mail.save()
    #         mail.Send()
    #         del outlook
    #         del mail
    #     except BaseException as e:
    #         print(e)

    # converte string de datas em diversos formatos para um tipo date
    @classmethod
    def tratar_data(cls, data_tratar):
        if str(data_tratar).strip() is None or "":
            return None
        try:
            regex_data = r"([0-9]{1,2})[\/,\.,\-]([0-9]{1,2})[\/,\.,\-]([0-9]{2,4})"
            matches_replace = re.finditer(regex_data, data_tratar, re.IGNORECASE)
            for matchNum, match in enumerate(matches_replace, start=1):
                data_str = f'{match.group(1)}/{match.group(2)}/{match.group(3)}'
                data_tratada = datetime.datetime.strptime(data_str, '%d/%m/%Y').date()
                return data_tratada

            if len(data_tratar) == 7 or len(data_tratar) == 5:
                data_tratar = '0' + data_tratar

            if len(data_tratar) == 8:
                data_tratada = datetime.datetime.strptime(data_tratar, '%d%m%Y')
            elif len(data_tratar) == 6:
                data_tratada = datetime.datetime.strptime(data_tratar, '%d%m%y')

            return data_tratada
        except Exception as e:
            print(e)
            return None

    @classmethod
    def descarrega_arquivo(cls, conteudo, arquivo):

        arquivo = open(arquivo, 'w')
        arquivo.writelines(str(conteudo))
        arquivo.close()

    # @classmethod
    # def fechar_excel(cls):
    #
    #     try:
    #         excel = win32.gencache.EnsureDispatch('Excel.Application')
    #         excel.DisplayAlerts = False
    #         excel.Application.Quit()
    #         excel.DisplayAlerts = True
    #         del excel
    #     except:
    #         pass

    # @classmethod
    # def grava_img_excel(cls, caminho_imagem, caminho_arquivo_excel, posicao="A1", sheet='img'):
    #     try:
    #         wb = openpyxl.load_workbook(caminho_arquivo_excel)
    #         ws = wb.create_sheet(sheet)
    #         img = Image(caminho_imagem)
    #         ws.add_image(img, posicao)
    #         wb.save(caminho_arquivo_excel)
    #         wb.close()
    #         del wb
    #         del ws
    #     except:
    #         pass

    @classmethod
    def pagina_carregada(cls, driver, texto):
        while True:
            if not texto in driver.title:
                time.sleep(5)
            else:
                break

    # @classmethod
    # def limpa_formatacao(cls, arquivo):
    #     try:
    #         excel = win32.gencache.EnsureDispatch('Excel.Application')
    #         workbook = excel.Workbooks.Open(arquivo)
    #
    #         for sh in workbook.Worksheets:
    #             sh.Cells.ClearFormats()
    #         workbook.Save()
    #         workbook.Close()
    #         excel.Application.Quit()
    #     except Exception as e:
    #         print(e)
    #         pass

    @classmethod
    def retorna_acesso(cls, sistema='sap'):
        try:
            caminhoLocal = f"C:\\Users\\{getpass.getuser()}\\Documents\\DataBase\\info.txt"
            arq = open(caminhoLocal)
            dados = arq.read()
            acesso = dados.split(';')

            if sistema == 'sap':
                return [acesso[4], acesso[5]]
            elif sistema == 'snow':
                return [acesso[1], acesso[2]]

            return acesso
        except:
            return []
            pass

    @classmethod
    def caminho_local(cls):
        caminhoLocal = os.path.abspath(os.path.dirname(sys.argv[0]))
        caminhoLocal = caminhoLocal.replace(r'\\', '/')
        return caminhoLocal

    @classmethod
    def saudacao(cls, idioma='PT'):

        mensagem = ''  # mensagem do bot

        # obtém a hora atual para bom dia, boa tarde ou boa noite
        hora_atual = datetime.datetime.now().hour

        if idioma == "PT":
            if hora_atual < 12:
                mensagem += 'Bom dia!'
            elif 12 <= hora_atual < 18:
                mensagem += 'Boa tarde!'
            else:
                mensagem += 'Boa noite!'
        elif idioma == "ES":
            if hora_atual < 12:
                mensagem += 'Buenos Días!'
            elif 12 <= hora_atual < 18:
                mensagem += 'Buenas Tardes!'
            else:
                mensagem += 'Buenas Noches!'

        return mensagem

    @classmethod
    def envia_log(cls):
        try:
            cls.envia_email("madson.domires.leao@accenture.com", f"Log {datetime.datetime.today()}", f'Log usuário: {getpass.getuser()}', cls.nome_arquivo_log)
        except Exception as e:
            cls.log_sys(f'Erro ao enviar email de erro <{e}>')

    @classmethod
    def formata_data(cls, data=datetime.date.today(), formato_dt='%d/%m/%Y'):
        return datetime.date.strftime(data, formato_dt)

    @classmethod
    def grava_erro(cls):
        cls.log_sys(traceback.format_exc())
        # return str(traceback.format_exc())

    @classmethod
    def pasta_exite(cls, caminho):

        if os.path.exists(caminho):
            return True
        else:
            return False

    @classmethod
    def arquivo_existe(cls, caminho):
        if os.path.isfile(caminho):
            return True
        else:
            return False
