from distutils.command.config import config
import pyautogui
import os
import time 
import configTRK
from datetime import datetime, timedelta
from botcity.core import DesktopBot
# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *

class Bot(DesktopBot):
    def __init__(self):
        super().__init__()
        self.dia_anterior = datetime.now() - timedelta (1)
        self.data_pasta = self.dia_anterior.strftime('%d-%m-%Y')
        self.diaAnteriorSemFormatacao = self.dia_anterior.strftime('%d%m%Y')
        self.mesAnterior = configTRK.mesAnteriorCompleto.strftime('%m/%Y')
        self.fiveDays = (datetime.now() + timedelta(5)).strftime('%d%m%Y')
        self.primeiroDiaMes = datetime.today().strftime('01/%m/%Y')
        self.maxDiaMes = configTRK.maxDiaMes
        self.cria_diretorio()
        
    def cria_diretorio(self):
        caminho = configTRK.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta
        if not os.path.isdir(caminho):
            os.mkdir(caminho)

    def aguarda_download(self):
        seconds = 1
        caminhoDownload = configTRK.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta
        time.sleep(seconds)
        dl_wait = True
        while dl_wait and seconds < 60:
            time.sleep(2)
            dl_wait = False
            for fname in os.listdir(caminhoDownload):
                if fname.endswith('.crdownload'):
                    dl_wait = True
            seconds += 1
        return seconds

    def renomar_arquivo(self, original, alterar):
        #self.convert_to_parquet(original)
        self.aguarda_download()
        os.rename(original, alterar)


    def incluiIndexCSV(self):

        caminhoArquivoOrigem = (configTRK.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta + '\\' + 'previsao_pagamento_proprietario.csv' )
        print(caminhoArquivoOrigem)
        caminhoArquivoDestino = (configTRK.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta + '\\' + 'previsao_pagamento_proprietario.csv' )
        print(caminhoArquivoDestino)

        arquivo = open(caminhoArquivoOrigem , 'r')
        arquivo_saida = open(caminhoArquivoDestino, 'w')

        count = 0

        for i in range(0, len(arquivo)):
            if(arquivo[i].split(';')[0] == 'lsImovel_Detail'):
                count+=1
                arquivo[i] = str(count) + ';' + arquivo[i]
            else:
                arquivo[i] = str(count) + ';' + arquivo[i]

        for i in arquivo:
            arquivo_saida.write(i)

        arquivo_saida.close()

    def prepara_ambiente(self):
        print('Preparando ambiente para evitar erros...')

        #Minimiza todos os programas e move o mause para o canto da tela
        pyautogui.keyDown('win')
        pyautogui.press('d')
        pyautogui.keyUp('win')
        pyautogui.moveTo(1,1)
    
    def loguin_imobiliar(self):
        print('Realizando Loguin...')

        #Abre o imobiliar
        self.execute(configTRK.imobiliar)

        contador = 0
        while contador < 1:
            #insere senha
            if not self.find( "senha", matching=0.97, waiting_time=30000):
                print('Aguardando programa abrir...')
            else:
                self.click_relative(73, 16)
                contador = 1

        #insere a senha
        self.paste(configTRK.senha)

        #clica em ok
        if not self.find( "ok_loguin", matching=0.97, waiting_time=30000):
            self.not_found("ok_loguin")
        self.click_relative(50, 15)

        time.sleep(2)

        print('Loguin realizado com sucesso!')

    
    def extracao_pagamento_proprietarios_analitico(self):
        print('extraindo relatório de pagamento de proprietários, visualização ANAlÍTICA...')

        #clica em locação
        if not self.find_text( "locacao", threshold=230, waiting_time=30000):    
            self.not_found("locacao")
        self.click()
         
        #clica em geracao_pagamento_proprietario
        if not self.find_text( "geracao_pagamento_proprietario", threshold=230, waiting_time=30000):    
            self.not_found("geracao_pagamento_proprietario")
        self.click()

        #modifica a competência para o mês anterior
        if not self.find_text( "competencia", threshold=230, waiting_time=30000):    
            self.not_found("competencia")
        self.double_click_relative(74, 12)

        
        self.delete()
        self.control_home()
        self.paste(self.mesAnterior)

        #seleção de proprietários: Todos
        if not self.find( "todos", matching=0.97, waiting_time=30000):
            self.not_found("todos")
        self.click_relative(8, 8)

        #Proprietários com a data de pagamento: até a data/5 dias a frente
             #clica em ate data
        if not self.find( "ate_data", matching=0.97, waiting_time=30000):
            self.not_found("ate_data")
        self.click_relative(8, 7)        
        
            #clica na data
        if not self.find( "data_pagamento", matching=0.97, waiting_time=30000):
            self.not_found("data_pagamento")
        self.click_relative(8, 14)      
        
        #insere data de 5 dias a frente
        self.control_home()
        self.paste(self.fiveDays)

        #Desmarca caixa de devedores
        if not self.find( "devedor", matching=0.97, waiting_time=30000):
            self.not_found("devedor")
        self.click_relative(9, 7)

        #todos os filtros ok, chamando função de exportar arquivo
        self.exporta_arquivo('previsao_saldo_proprietario_analitico')

    
    def extracao_pagamento_proprietarios_sintetico(self):
        print('extraindo relatório pagamento_proprietarios_sintetico...')

        #modificando de analítico para sintético
        if not self.find( "sintetico", matching=0.97, waiting_time=30000):
            self.not_found("sintetico")
        self.click_relative(9, 8)

        self.exporta_arquivo('previsao_saldo_proprietario_sintetico')

        #fecha propriedades_pagamento
        if not self.find( "fechar_propriedade", matching=0.97, waiting_time=30000):
            self.not_found("fechar_propriedade")
        self.click_relative(100, 12)
        

    def Extrato_conta_corrente_data(self):
        print('Extraindo relatório extrato de conta corrente por data...')
        
        #clica em locação
        if not self.find_text( "locacao", threshold=230, waiting_time=30000):    
            self.not_found("locacao")
        self.click()

        #clica em relatórios
        if not self.find( "relatorios", matching=0.97, waiting_time=30000):
            self.not_found("relatorios")
        self.click()

        #clica em Extrato de conta corrente
        if not self.find_text( "extrato_cc", threshold=230, waiting_time=30000):    
            self.not_found("extrato_cc")
        self.click_relative(44, 9)

        #Inserindo período do mes na data
         #primeiro dia
        if not self.find( "data_inicial", matching=0.97, waiting_time=10000):
            self.not_found("data_inicial")
        self.click_relative(14, 29)

        self.control_home()
        self.paste(self.primeiroDiaMes)
        
         #max dia mês
        if not self.find( "data_final", matching=0.97, waiting_time=10000):
            self.not_found("data_final")
        self.click_relative(12, 28)
        
        self.control_home()
        self.paste(self.maxDiaMes)
        
        #inserindo seleção de proprietários = todos
        if not self.find( "todos_extrato", matching=0.97, waiting_time=30000):
            self.not_found("todos_extrato")
        self.click_relative(19, 29)
        
        #incluindo prorpietários inativos
        if not self.find( "incluir_inativos", matching=0.97, waiting_time=30000):
            self.not_found("incluir_inativos")
        self.click_relative(15, 13)
         
        #escolhendo tipo de relatório por data
        if not self.find( "por_data", matching=0.97, waiting_time=30000):
            self.not_found("por_data")
        self.click_relative(12, 9)
        
        #habilitar caixa mostrar saldo por lançamento
        if not self.find( "mostrar_lancamentos", matching=0.97, waiting_time=30000):
            self.not_found("mostrar_lancamentos")
        self.click_relative(7, 7)

        self.exporta_arquivo('extrato_conta_corrente_proprietario_por_data')

        #clica para fechar o relatorio
        if not self.find( "fechar_extrato", matching=0.97, waiting_time=10000):
            self.not_found("fechar_extrato")
        self.click_relative(102, 13)


    def extracao_diario_conta_corrente(self):
        print('Extraindo relatório diário de conta corrente...')

        #clica em locação
        if not self.find_text( "locacao", threshold=230, waiting_time=30000):    
            self.not_found("locacao")
        self.click()

        #clica em relatórios
        if not self.find( "relatorios", matching=0.97, waiting_time=30000):
            self.not_found("relatorios")
        self.click()

        #clica em diario de conta corrente
        if not self.find( "diario_cc", matching=0.97, waiting_time=10000):
            self.not_found("diario_cc")
        self.click_relative(39, 11)
        
        #clica em data inicial
        if not self.find( "data_inicial_diario", matching=0.97, waiting_time=10000):
            self.not_found("data_inicial_diario")
        self.click_relative(89, 13)
        
        self.control_home()
        self.paste(self.diaAnteriorSemFormatacao)

        #clica em data final
        if not self.find( "data_final_diario", matching=0.97, waiting_time=10000):
            self.not_found("data_final_diario")
        self.click_relative(77, 11)

        self.control_home()
        self.paste(self.diaAnteriorSemFormatacao)
        
        #clica em resumo por taxa
        if not self.find( "resumo_taxa", matching=0.97, waiting_time=10000):
            self.not_found("resumo_taxa")
        self.click_relative(8, 11)
        
        self.exporta_arquivo('diario_conta_corrente_locacao')

        #clica em fechar relatório
        if not self.find( "fechar_diario", matching=0.97, waiting_time=10000):
            self.not_found("fechar_diario")
        self.click_relative(107, 10)
        
        
    def extracao_previsao_saldos(self):
        print('Extraindo Relatório de Previsão de salvos...')

        #clica em locação
        if not self.find_text( "locacao", threshold=230, waiting_time=30000):    
            self.not_found("locacao")
        self.click()
         
        #clica em geracao_pagamento_proprietario
        if not self.find_text( "geracao_pagamento_proprietario", threshold=230, waiting_time=30000):    
            self.not_found("geracao_pagamento_proprietario")
        self.click()
        
        #modifica a competência para o mês anterior
        if not self.find_text( "competencia", threshold=230, waiting_time=30000):    
            self.not_found("competencia")
        self.double_click_relative(74, 12)

        self.delete()
        self.control_home()
        self.paste(self.mesAnterior)

        #seleção de proprietários: Todos
        if not self.find( "todos", matching=0.97, waiting_time=30000):
            self.not_found("todos")
        self.click_relative(8, 8)

            #clica em ate data
        if not self.find( "ate_data", matching=0.97, waiting_time=30000):
            self.not_found("ate_data")
        self.click_relative(8, 7)   

        #clica na data
        if not self.find( "data_pagamento", matching=0.97, waiting_time=30000):
            self.not_found("data_pagamento")
        self.click_relative(8, 14)      
        
        #insere data maxima do mês
        self.control_home()
        self.paste(self.maxDiaMes)

        self.exporta_arquivo('previsao_pagamento_proprietario')

        #fecha propriedades_pagamento
        if not self.find( "fechar_propriedade", matching=0.97, waiting_time=30000):
            self.not_found("fechar_propriedade")
        self.click_relative(100, 12)

        

        self.incluiIndexCSV()

    def fecha_imobiliar(self):
        print('Todos relatórios baixados, fechando imobiliar...')
        
        #fecha imobiliar
        if not self.find( "fecha_imobiliar", matching=0.97, waiting_time=10000):
            self.not_found("fecha_imobiliar")
        self.click_relative(103, 10)
         
        #confirma acao de fechar
        if not self.find( "sim_fechar_final", matching=0.97, waiting_time=10000):
            self.not_found("sim_fechar_final")
        self.click()
        
         

    def exporta_arquivo(self, nome_arquivo):
        print('Salvando arquivo...')
        
            #clica em visualizar
        if not self.find( "visualizar", matching=0.97, waiting_time=30000):
            print('não achei o botão visualizar')
        self.click()
        
        #clica em exportar para excel
        contador = 0
        while contador < 1:
            if not self.find( "exporta_excel", matching=0.97, waiting_time=30000):
                print('Aguardando a janela de extração abrir')            
            else:
                self.click_relative(10, 9)
                contador = 1
        
        #clica no caminho
        if not self.find( "clica_inserir_caminho", matching=0.97, waiting_time=30000):
            self.not_found("clica_inserir_caminho")
        self.click_relative(6, 40)

        #insere caminho
        self.control_a()
        self.delete()
        self.paste(configTRK.DIRETORIO_ARQUIVOS + '\\' + self.data_pasta + '\\' + nome_arquivo)

        #clica em salvar
        if not self.find_text( "salvar", threshold=230, waiting_time=30000):    
            self.not_found("salvar")
        self.click()

        #clica para fechar relatorio extraido
        if not self.find( "fecar_relatorio", matching=0.97, waiting_time=30000):
            self.not_found("fecar_relatorio")
        self.click_relative(16, 10)

        print(f'relatório: {nome_arquivo} baixado com sucesso!!!')
        
       


    def action(self, execution=None):
        self.cria_diretorio()
        self.prepara_ambiente()
        self.loguin_imobiliar()
        #self.extracao_pagamento_proprietarios_analitico()
        #self.extracao_pagamento_proprietarios_sintetico()
        #self.Extrato_conta_corrente_data()
        #self.extracao_diario_conta_corrente()
        self.extracao_previsao_saldos()
        self.fecha_imobiliar()

    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()





