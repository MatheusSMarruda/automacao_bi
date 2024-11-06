import pyautogui
import time
import keyboard
import pyperclip
import os

#definições de proporções de tele.
screenWidth, screenHeight = pyautogui.size() #obtem o tamanho do monitor principal
screenWidth, screenHeight
(1920, 1080)

#Definições de eixos
currentMouseX, currentMouseY = pyautogui.position() #obtem a posição do mouse.
currentMouseX, currentMouseY
(1314, 345)

#caso o programa falhe isso serve para que pare ao posicionar o mouse no canto superior esquerdo da tela o programa.
pyautogui.FailSafeException
pyautogui.FAILSAFE = True

#declaração do caminho do explorador de arquivos
caminho = input("Informe o caminho onde se está localizado os arquivos: Base de dados e Power Bi")

#DEFININDO E DECLARANDO LOGICAS#
def abertura_da_base_de_dados():
    os.startfile('explorer.exe') #abre o explorador de arquivos
    time.sleep(1)
    pyautogui.hotkey('alt', 'space')
    time.sleep(1)
    pyautogui.press('x') #redimensiona em tela cheia
    pyautogui.moveTo(756, 67, duration=0.2) #posiciona na barra de endereços
    time.sleep(1)
    pyautogui.click()
    pyautogui.write(caminho) #cola o valor do endereço inputado
    pyautogui.press("enter")
    time.sleep(1)
    posicao_cursor = pyautogui.locateCenterOnScreen("img_first_arq.png") #posiciona o cursor na respectiva imagem 'imagem da pasta da base de dados'
    pyautogui.click(posicao_cursor, clicks=2) #abre a base de dados
    time.sleep(30) #aguarda a planilha 'base de dados abrir'
    pyautogui.press("enter") #fecha o msgbox do exel 'msg de vinculos de uma ou mais pastas externas'
    pyautogui.hotkey('alt', 'space')
    time.sleep(1)
    pyautogui.press('x') #maximiza o planilha excel
    time.sleep(1)
    pyautogui.moveTo(76, 971, duration=1)
    time.sleep(1)
    pyautogui.rightClick()
    try:
        time.sleep(1)
        posicao_cursor_planilha_pipedrive = pyautogui.locateCenterOnScreen("img_planilha_pipedrive.png") #pesquisa pela imagem do print de boleta
        if posicao_cursor_planilha_pipedrive is not None:
                pyautogui.click(posicao_cursor_planilha_pipedrive, clicks=2) #posiciona e clica em 'boletas'
        else:
            posicao_cursor_planilha_pipedrive_select = pyautogui.locateCenterOnScreen("img_planilha_pipedrive_select.png") #pesquisa pela imagem do print de boleta
            if posicao_cursor_planilha_pipedrive_select is not None:
                pyautogui.click(posicao_cursor_planilha_pipedrive_select, clicks=2) #posiciona e clica em 'boletas'
            else:
                print('não encontrado')
    except Exception as e:
        posicao_cursor_planilha_pipedrive_select = pyautogui.locateCenterOnScreen("img_planilha_pipedrive_select.png") #pesquisa pela imagem do print de boleta
        pyautogui.click(posicao_cursor_planilha_pipedrive_select, clicks=2) #posiciona e clica em 'boletas'
    time.sleep(2)
    pyautogui.hotkey('ctrl', 't') #seleciona todas as celulas da planilha
    time.sleep(1.5)
    pyautogui.press('delete') #deleta todas as celulas antes selecionadas
    time.sleep(1.5)
    pyautogui.hotkey('ctrl', 'left') #sai do select all 'seta para direita'
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'up') #volta para a celula A1

def baixando_base_pipedrive():
    # #baixando as bases de dados#
    #base pypedrive
    pyautogui.hotkey('win', 'r')
    pyautogui.sleep(1)
    pyautogui.write('msedge')
    pyautogui.sleep(0.3)
    pyautogui.press('enter')
    pyautogui.sleep(1)
    pyautogui.hotkey('alt', 'space')
    time.sleep(2)
    pyautogui.press('x') #transforma em tela cheia
    time.sleep(2)
    pyautogui.write(r"https://tempoenergia.pipedrive.com/deal/2965") #cola o valor do site do pipedrive
    time.sleep(1)
    pyautogui.press('enter') #busca o site pipedrive
    time.sleep(5)
    pyautogui.moveTo(38, 311, duration=1) 
    time.sleep(0.3)
    pyautogui.click() #clica no icone de 'negocio'
    pyautogui.sleep(2)
    #baixa o relatorio negocio pipedrive e abre o arquivo
    pyautogui.moveTo(195, 244, duration=0.5) #move para lista
    pyautogui.click()
    pyautogui.sleep(1)
    pyautogui.moveTo(1869, 241, duration=0.5) #move para 3 pontos
    pyautogui.click()
    pyautogui.sleep(1)
    pyautogui.moveTo(1700, 294, duration=0.5) #move para exportar dados do filtro
    pyautogui.click()
    pyautogui.sleep(1)
    pyautogui.moveTo(1187, 483, duration=0.5) #move para exportar
    pyautogui.click()
    pyautogui.sleep(2)
    pyautogui.moveTo(980, 275, duration=0.5) #move para baixar 
    pyautogui.click()
    pyautogui.sleep(10)
    pyautogui.moveTo(1451, 172, duration=0.5) #move para abrir o arquivo
    pyautogui.click()
    pyautogui.sleep(10)

def atribui_dados_base_pipedrive():
    pyautogui.hotkey('ctrl', 't') #seleciona todos os itens da planilha
    pyautogui.sleep(0.3)
    pyautogui.hotkey('ctrl', 'c') #copia todos os itens da planilha
    pyautogui.sleep(1)
    pyautogui.hotkey('alt', 'tab') #retorna para a base de dados 'trocando de janela'
    pyautogui.sleep(1)
    pyautogui.hotkey('ctrl', 'v') #cola os valores copiados da planilha baixada no pipedrive
    pyautogui.sleep(5)
    pyautogui.hotkey('alt', 'space')
    pyautogui.sleep(0.3)
    pyautogui.hotkey('N') #minimiza a planilha 'base de dados'
    time.sleep(2)
    pyautogui.moveTo(1026, 562, duration=0.5)
    pyautogui.click()
    time.sleep(2)
    pyautogui.moveTo(1895, 30) #fecha planilha baixada do pipedrive
    pyautogui.click()
    time.sleep(2)

def baixando_base_listar_boletas_zeus():
    pyautogui.hotkey('win', 'r') #abre o executar
    pyautogui.sleep(1)
    pyautogui.write('msedge') 
    pyautogui.sleep(0.3)
    pyautogui.press('enter') #executa o navegador edge
    time.sleep(2)
    pyautogui.write(r"https://tempo.zeusenergia.com.br/resources/index.html#/app/financeiro/dashboard")
    pyautogui.sleep(2)
    pyautogui.press('enter') #entra no site zeus
    pyautogui.sleep(2)
    pyautogui.moveTo(965, 674, duration=0.3)
    pyautogui.sleep(3)
    pyautogui.click() #loga no usuario zeus
    time.sleep(15)
    posicao_cursor_boleta_zeus = pyautogui.locateCenterOnScreen("img_boletas.png") #pesquisa pela imagem do print de boleta
    pyautogui.click(posicao_cursor_boleta_zeus, clicks=1) #posiciona e clica em 'boletas'
    time.sleep(1)
    posicao_cursor_listar_boletas = pyautogui.locateCenterOnScreen("img_boletas_listar.png") #pesquisa pela imagem do print de listar boletas
    pyautogui.click(posicao_cursor_listar_boletas, clicks=1) #posiciona e clica em 'listar boletas'
    pyautogui.moveTo(423, 366, duration= 0.3) #posiciona e clica em 'data de operação'
    pyautogui.sleep(0.5)
    pyautogui.click()
    pyautogui.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a') #seleciona todos os itens do input de 'data de operação'
    pyautogui.sleep(0.5)
    pyautogui.press('delete') #apaga a data do filtro
    time.sleep(0.5)
    posicao_cursor_pesquisa_avancada = pyautogui.locateCenterOnScreen("img_pesquisa_avancada.png") #rastreia 'lista avançada'
    pyautogui.click(posicao_cursor_pesquisa_avancada, clicks=1) #posiciona e clica em 'lista avançada'
    time.sleep(0.5)
    posicao_cursor_procura_carteira = pyautogui.locateCenterOnScreen("procura_carteira.png") #rastreia 'carteiras'
    pyautogui.click(posicao_cursor_procura_carteira, clicks=1) #posiciona e clica em 'carteiras'
    pyautogui.sleep(0.5)
    #######escrevem as respectivas carteiras#######
    pyautogui.write(r'Desconto Garantido - Atacadista')
    pyautogui.sleep(0.5)
    pyautogui.press('enter')
    pyautogui.write(r'Varejista -')
    pyautogui.sleep(0.5)
    pyautogui.press('enter')
    pyautogui.write(r'DESCONTO GARANTIDO - VAREJISTA')
    pyautogui.sleep(0.5)
    pyautogui.press('enter')
    #######escrevem as respectivas carteiras#######
    time.sleep(2)
    posicao_cursor_buscar_boletas = pyautogui.locateCenterOnScreen("img_buscar_boletas.png") #rastreia 'buscar'
    pyautogui.click(posicao_cursor_buscar_boletas, clicks=1) #posiciona e clica em 'buscar'
    time.sleep(15)
    pyautogui.scroll(-600) #scroll para baixo
    time.sleep(0.5)
    posicao_cursor_excel = pyautogui.locateCenterOnScreen("img_excel.png") #rastreia 'excel'
    pyautogui.click(posicao_cursor_excel, clicks=1) #posiciona e clica em 'Excel'
    time.sleep(80)

def atribui_dados_base_listarboletas_zeus():
    pyautogui.moveTo(1430, 176) #posiciona em abrir o arquivo excel baixado do zeus
    pyautogui.click()
    pyautogui.sleep(15)
    pyautogui.hotkey('alt', 'space') #abre as configurações de exibição de janela
    time.sleep(1)
    pyautogui.press('x') #maximiza a pagina
    time.sleep(3)
    pyautogui.hotkey('alt', 'tab') #retorna para a base de dados 'trocando de janela'
    time.sleep(2)
    pyautogui.moveTo(73, 975, duration=1) #posiciona em 'seta de troca de planilha'
    time.sleep(1)
    pyautogui.rightClick() #abre as subplanilhas existentes
    time.sleep(2) 
    posicao_cursor_planilha_zeus = pyautogui.locateCenterOnScreen("img_planilha_zeus.png") #rastreia 'Zeus'
    pyautogui.click(posicao_cursor_planilha_zeus, clicks=2) #posiciona e clica em 'Zeus'
    time.sleep(3)
    pyautogui.hotkey('ctrl', 't') #seleciona todas as celulas
    time.sleep(1)
    pyautogui.press('delete') #deleta a base de dados existentes
    time.sleep(1)
    pyautogui.hotkey('alt', 'tab') #retorna para a base de dados baixada anteriormenete pelo Zeus 'trocando de janela'
    time.sleep(1)
    pyautogui.hotkey('ctrl', 't') #seleciona todas as celulas da planilha baixada do Zeus
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c') #copia todas as celulas da planilha baixada do Zeus
    time.sleep(2)
    pyautogui.hotkey('alt', 'tab') #retorna para a base de dados do BI 'trocando de janela'
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'left')
    time.sleep(0.3)
    pyautogui.hotkey('ctrl', 'up') #volta para a celula A1
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'v') #cola os valores copiados na base de dados do BI

def baixando_base_listar_mensal_zeus():
    pyautogui.hotkey('alt', 'tab') #troca de janela (voltando para pagina web Zeus)
    time.sleep(1)
    pyautogui.moveTo(1888, 27, duration=1)
    time.sleep(0.5)
    pyautogui.click()
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.hotkey('alt', 'tab')
    time.sleep(1)
    posicao_cursor_mensal_listar = pyautogui.locateCenterOnScreen("img_mensal_listar.png") #rastreia 'Listar mensal'
    pyautogui.click(posicao_cursor_mensal_listar, clicks=2) #posiciona e clica em 'Listar mensal'
    time.sleep(2)
    pyautogui.moveTo(388, 363, duration=0.5) #posiciona em 'fornecimento de'
    time.sleep(1)
    pyautogui.click(clicks=3)
    time.sleep(1)
    pyautogui.press('delete') #deleta data de filtro de 'fornecimento de'
    time.sleep(1)
    pyautogui.moveTo(841, 366, duration=0.5) #posiciona em 'fornecimento até'
    pyautogui.click(clicks=3)
    time.sleep(1)
    pyautogui.press('delete') #deleta data de filtro de 'fornecimento até'
    time.sleep(2)
    posicao_cursor_buscar_boletas = pyautogui.locateCenterOnScreen("img_buscar_boletas.png") #rastreia 'buscar'
    pyautogui.click(posicao_cursor_buscar_boletas, clicks=1) #posiciona e clica em 'buscar'
    time.sleep(10)
    pyautogui.scroll(-600) #scroll para baixo
    time.sleep(2)
    posicao_cursor_excel = pyautogui.locateCenterOnScreen("img_excel.png") #rastreia 'excel'
    pyautogui.click(posicao_cursor_excel, clicks=1) #posiciona e clica em 'Excel' (baixando base de dados listar mensal do Zeus)
    time.sleep(65)
    pyautogui.moveTo(1435, 171, duration=0.5) #posiciona em abrir relatorio baixado
    pyautogui.click() #abre relatorio

def abertura_base_inicio_fim_contratos():
    time.sleep(4)
    os.startfile('explorer.exe') #abre o explorador de arquivos
    time.sleep(1)
    pyautogui.hotkey('alt', 'space')
    time.sleep(1)
    pyautogui.press('x') #transforma em tela cheia
    time.sleep(2)
    pyautogui.moveTo(756, 67, duration=1) #posiciona na barra de endereços
    time.sleep(1)
    pyautogui.click()
    pyautogui.write(caminho) #cola o valor do endereço
    pyautogui.press("enter")
    time.sleep(2)
    posicao_cursor_base_contratos = pyautogui.locateCenterOnScreen("img_base_contratos.png") #busca a pasta de inicio e fim de contratos
    pyautogui.click(posicao_cursor_base_contratos, clicks=2) #abre a base de dados de inicio e fim de contratos
    time.sleep(55)

def atribui_dados_base_listarmensal_zeus():
    pyautogui.moveTo(75, 973)
    pyautogui.rightClick()
    try:
        time.sleep(2)
        posicao_cursor_zeus_no_select = pyautogui.locateCenterOnScreen("img_zeus_no_select.png") #pesquisa pela imagem do print de boleta
        if posicao_cursor_zeus_no_select is not None:
                pyautogui.click(posicao_cursor_zeus_no_select, clicks=2) #posiciona e clica em 'boletas'
        else:
            posicao_cursor_zeus_select = pyautogui.locateCenterOnScreen("img_zeus_select.png") #pesquisa pela imagem do print de boleta
            if posicao_cursor_zeus_select is not None:
                pyautogui.click(posicao_cursor_zeus_select, clicks=2) #posiciona e clica em 'boletas'
            else:
                print('não encontrado')
    except Exception as e:
        posicao_cursor_zeus_select = pyautogui.locateCenterOnScreen("img_zeus_select.png") #pesquisa pela imagem do print de boleta
        pyautogui.click(posicao_cursor_zeus_select, clicks=2) #posiciona e clica em 'boletas'
    pyautogui.hotkey('ctrl', 't')
    time.sleep(1)
    pyautogui.press('delete')
    time.sleep(1)
    pyautogui.hotkey('alt', 'tab')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 't')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)
    pyautogui.hotkey('alt', 'tab')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.hotkey('alt', 'tab')
    time.sleep(1)
    pyautogui.moveTo(1888, 28, duration=0.5)
    time.sleep(1)
    pyautogui.click()
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.moveTo(76, 972, duration=0.5)
    time.sleep(1)
    pyautogui.rightClick()
    time.sleep(1)
    posicao_cursor_contratos = pyautogui.locateCenterOnScreen("img_contratos.png") #busca a pasta de inicio e fim de contratos
    pyautogui.click(posicao_cursor_contratos, clicks=2) #abre a base de dados de inicio e fim de contratos
    time.sleep(1)

def rodar_macro_vba():
    posicao_cursor_atualizar_macro = pyautogui.locateCenterOnScreen("img_atualizar_macro.png") #busca a pasta de inicio e fim de contratos
    pyautogui.click(posicao_cursor_atualizar_macro, clicks=2) #abre a base de dados de inicio e fim de contratos
    time.sleep(45)

def abertura_powerbi():
    pyautogui.moveTo(1894, 26)
    pyautogui.click()
    time.sleep(20)
    pyautogui.click()
    time.sleep(20)
    time.sleep(2)
    posicao_cursor_bi = pyautogui.locateCenterOnScreen("img_dash_bi.png") #busca a pasta de inicio e fim de contratos
    pyautogui.click(posicao_cursor_bi, clicks=2) #abre a base de dados de inicio e fim de contratos
    time.sleep(15)
    pyautogui.hotkey('alt', 'space')
    time.sleep(1)
    pyautogui.hotkey('x')
    time.sleep(25)

def atualizar_publicar_powerbi():
    posicao_cursor_atualizar_bi = pyautogui.locateCenterOnScreen("img_atualizar.png") #busca a pasta de inicio e fim de contratos
    pyautogui.click(posicao_cursor_atualizar_bi, clicks=2) #abre a base de dados de inicio e fim de contratos
    time.sleep(20)
    pyautogui.moveTo(962, 625)
    pyautogui.click()
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')
    time.sleep(2)
    posicao_cursor_img_publicar = pyautogui.locateCenterOnScreen("img_publicar.png") #busca a pasta de inicio e fim de contratos
    pyautogui.click(posicao_cursor_img_publicar, clicks=2) #abre a bases de dados de inicio e fim de contratos
    time.sleep(1.5)
    posicao_cursor_salvar_bi = pyautogui.locateCenterOnScreen("img_salvar_bi.png") #buscar a img de salvar
    pyautogui.click(posicao_cursor_salvar_bi, clicks=1) #posiciona e clica em salvar bi
    time.sleep(4)
    pyautogui.moveTo(871, 453, duration=1)
    time.sleep(0.5)
    pyautogui.click(clicks=2)
    time.sleep(6)
    pyautogui.press('enter')
    time.sleep(6)
    pyautogui.hotkey('tab')
    time.sleep(0.5)
    pyautogui.hotkey('tab')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(5)
    pyautogui.moveTo(264, 543)
    time.sleep(1)
    pyautogui.click()
    time.sleep(2)
    time.sleep(3)
    posicao_cursor_back = list(pyautogui.locateAllOnScreen("img_back.png"))  # converte para lista
    time.sleep(1)
    if posicao_cursor_back:  # verifica se encontrou pelo menos uma posição
        pyautogui.click(posicao_cursor_back[0])  # clica na primeira posição encontrada
        pyautogui.alert('Relatório finalizado, tenha um bom trabalho!')

abertura_da_base_de_dados()
baixando_base_pipedrive()
atribui_dados_base_pipedrive()
baixando_base_listar_boletas_zeus()
atribui_dados_base_listarboletas_zeus()
baixando_base_listar_mensal_zeus()
abertura_base_inicio_fim_contratos()
atribui_dados_base_listarmensal_zeus()
rodar_macro_vba()
abertura_powerbi()
atualizar_publicar_powerbi()