#get_ipython().system('pip install opencv-python ')
#get_ipython().system('pip install pytesseract')
#get_ipython().system('pip install pyautogui')
#get_ipython().system('pip install pygetwindow')


#------------------------------------------------------------------------------------------------------------------------------

from time import sleep
from datetime import datetime
from tabula import read_pdf
from tabulate import tabulate


import os
import cv2
import math
import keyboard
import pyautogui
import pytesseract
import pandas as pd
import pygetwindow as gw



ARQUIVO_HOD = 'W:\\CGSIE\\PLACA_AMARELA\\ARQUIVO_HOD\\'
PASTA_ARQUIOS_HOD = 'C:\\Users\\Administrador\\Desktop\\script_placa_amarela\\TELA_HOD\\'
PASTA_ARQUIVOS_CHECKLIST = 'W:\\CGSIE\\PLACA_AMARELA\\PDF_CHECKLIST\\'
PASTA_ARQUIVOS_PROCESSADOS = 'W:\\CGSIE\\PLACA_AMARELA\\PROCESSADOS\\'

HOD_FILE = 'hodcivws.jsp'
APP_HOD = 'Terminal 3270 - A'
CMD = 'Administrador: Prompt de Comando'
WD_ON_DEMAND = 'Host On-Demand'
APP_CONTROL_PANEL = 'Painel de Controle'
EXTENSAO_PDF = '.pdf'


LB_TELA = 'TELA 002'
LB_MENU_SISTEMAS = 'MENU DE SISTEMAS'
LB_RENAVAM = 'RENAVAM'


tela901, tela903, tela911, tela110 = [], [], [], []


PASSO00 = '00 - SCRIPT INICIALIZADO:'
PASSO01 = '01 - ABRIR HOD'
PASSO02 = '02 - ABRIU CMD'
PASSO03 = '03 - FECHOU CMD'
PASSO04 = '04 - DIGITOU RENAVAM'
PASSO05 = '05 - HOD ABERTO'
PASSO06 = '06 - LEU PDF'
PASSO07 = '07 - DADOS EXTRAÍDOS'
PASSO08 = '08 - TELA 901 PREENCHIDA'
PASSO09 = '09 - TELA 901 COLETATA'
PASSO10 = '10 - TELA 903 PREENCHIDA'
PASSO11 = '11 - TELA 903 COLETATA'
PASSO12 = '12 - TELA 911 PREENCHIDA'
PASSO13 = '13 - TELA 911 COLETATA'
PASSO14 = '14 - DADOS PREENCHIDOS COM SUCESSO'



#------------------------------------------------------------------------------------------------------------------------------

def salvarTela(caminho_arquivo):
    pyautogui.screenshot(caminho_arquivo, region=(390, 190, 1150, 770))

def getDataAtual():
    data_hora_atual = datetime.now()
    return data_hora_atual.strftime('"%m/%d/%Y, %H:%M:%S"')

def salvarTela(caminho_arquivo):
    pyautogui.screenshot(caminho_arquivo, region=(390, 190, 1150, 770))

def fecharJanelaAtiva(janela):
    sleep(2)
    app = gw.getWindowsWithTitle(janela)[0]
    app.close()

def aguardarJanelaAtiva(janela, timer=0):
    aplicacoes = gw.getAllTitles()
    i = 0
    while janela not in aplicacoes:
        sleep(1)
        aplicacoes = gw.getAllTitles()
        i += 1
        if timer != 0 and i == timer:
            break

def redimensionarImagem(largura: int, altura: int, taxa=0.15) -> tuple:
    largura +=  (largura * taxa)
    altura += (altura * taxa)
    return (int(largura), int(altura))

def getGrayscale(image):
    return cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

def thresholding(image):
    return cv2.threshold(image, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

def limparTexto(texto):
    texto = list(filter(lambda linha: linha.strip(), texto.splitlines()))
    return texto

def lerImagemHOD(caminho_imagem: str):
    # extrairInformacoesVisionAPI(caminho_imagem)
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    imagem = cv2.imread(caminho_imagem)
    dim = redimensionarImagem(imagem.shape[1], imagem.shape[0], taxa=1)
    imagem = cv2.resize(imagem, dim, interpolation = cv2.INTER_AREA) 
    gray_image = getGrayscale(imagem)
    threshold_img = thresholding(gray_image)
    texto = pytesseract.image_to_string(threshold_img, lang="por")
    texto = limparTexto(texto)
    return texto

def verificarHODAbertoSemErro():
    sleep(5)
    hod = gw.getWindowsWithTitle(APP_HOD)[0]
    hod.restore()
    hod.activate()
    hod.maximize()
    sleep(5)
    salvarTela(f'{PASTA_ARQUIOS_HOD}temp.png')
    sleep(4)
    dc = lerImagemHOD(f'{PASTA_ARQUIOS_HOD}temp.png')
    sleep(6)
    return (len(list(filter(lambda item: LB_TELA in item or LB_MENU_SISTEMAS in item or LB_RENAVAM in item, dc))) >0)

def formatarCondicaoVIN(condicao):
    if condicao == 'NORMAL':
        return condicao[0]
    elif condicao == 'ORIGINAL':
        return condicao[5]
    else:
        return condicao[0]

def formatarProcedencia(procedencia):
    return procedencia.split()[0][0]
        
# def formatarMontagem(montagem):
#     return montagem[1]

def formatarTipoFaturado(faturado):
    return faturado.split()[0][0]

def formatarNumeroFaturado(faturado):
    return faturado.split()[-1]

def formatarQuantidadeDeEixos(eixos):
    return eixos[:]

def formatarNumeroMotor(motor):
    return motor[:]

def formatarCodigoTipo(tipo):
    return tipo.split()[0][:]

def formatarCodigoMarcaModelo(marca_modelo):
    return marca_modelo[:]

def formatarCodigoEspecie(especie):
    return especie.split()[0][1]

def formatarCodigoCarroceria(carroceria):
    return carroceria.split()[0][:]

def formatarCodigoCor(cor):
    return cor.split()[0][:].rjust(2, "0")

def formatarCodigoCombustivel(combustivel):
    return combustivel.split()[0][:].rjust(2, "0")

def formatarLotacao(lotacao):
    return lotacao.split()[0][:].rjust(3, "0")

def formatarPotencia(potencia):
    if '/' in potencia:
        return potencia.split('/')[0].split()[0][:].rjust(3, "0")
    else:
        return potencia.split()[0][:].rjust(3, "0")

def formatarCilindrada(cilindrada):
    if '/' in cilindrada:
        return cilindrada.split('/')[1].split()[0][:].ljust(4," ")
    else:
        return cilindrada.split()[1][:].ljust(4, " ")

# OBRIGATÓRIO PARA ESPÉCIE 02 (CARGA)
def formatarCapacidadeDeCarga(carga):
    if len(carga) == 1:
        return carga.rjust('2') + ',00'
    elif len(carga) == 2:
        return carga.rjust('1') + ',00'
    elif len(carga) == 3:
        return "{:.2f}".format(int(carga[:])).replace('.',',')
    else:
        return carga[:3] + ',' + carga[3] + '0'  


#------------------------------------------------------------------------------------------------------------------------------
# PRIMEIRA PARTE - ABRIR HOD

def executarHOD():
    pyautogui.press('winleft')
    sleep(2)
    pyautogui.write('cmd')
    print(PASSO02)
    sleep(2)
    pyautogui.press('enter')
    sleep(2)
    pyautogui.write(f'javaws W:\CGSIE\PLACA_AMARELA\ARQUIVO_HOD\{HOD_FILE}')
    sleep(2)
    pyautogui.press('enter')
    sleep(2)
    fecharJanelaAtiva(CMD)
    print(PASSO03)
    aguardarJanelaAtiva(WD_ON_DEMAND, timer=5)
    pyautogui.press('enter')
    aguardarJanelaAtiva(APP_HOD, timer=6)
    sleep(10)
    try:
        if verificarHODAbertoSemErro(): 
            pyautogui.write('RENAVAM')
            print(PASSO04)
            sleep(2)
            pyautogui.press('enter')
            sleep(2)
        else:
            print('NÃO ACESSOU')
            fecharJanelaAtiva(APP_HOD)
            fecharJanelaAtiva(APP_CONTROL_PANEL)
            raise Exception('SESSÃO EXPIRADA NO HOD')
    except IndexError:
        raise Exception('Arquivo hodcivws.jsp não localizado. Favor realizar upload do arquivo antes de executar o script.')
    except Exception as ex:
        raise Exception(ex)

#------------------------------------------------------------------------------------------------------------------------------
# SEGUNDA PARTE - LER ARQUIVO PDF


def lerArquivoPDF(arquivo):
    tables = read_pdf(PASTA_ARQUIVOS_CHECKLIST + arquivo, pages="all", lattice=True, encoding='ISO 8859-1')
    #df = pd.concat(tables[0], tables[1])
    #df = tables[0].drop(["Unnamed: 1", "Unnamed: 0", "Unnamed: 2"], axis=1)
    df = tables[0][["Item", "Análise"]]
    print(PASSO06)
    return df

def extrairDadosDaTabela(tabela):
    print('-- EXTRAINDO DADOS DA TABELA --')
    checklist = list()
    print(PASSO07)
    for _, row in tabela.iterrows():
        checklist.append(row["Análise"])
    return checklist


def preencherTelas(checklist):
    for i, valor in enumerate(checklist):
        if i == 9:
            tela901.append(str(valor).strip())
            tela911.append(str(valor).strip())
            tela110.append(str(valor).strip())  
        if i == 20:
            tela903.append(str(valor).strip())
        if i == 17:
            tela911.append(str(valor).strip())
        if i not in [9]:
            tela110.append(str(valor).strip())
    print ([list((i, tela110[i])) for i in range(len(tela110))])
          

#------------------------------------------------------------------------------------------------------------------------------
# TERCEIRA PARTE - FAZER AS CONSULTAS DAS TELAS 901 / 903 / 911
       
def consultaDeTelasNoHOD():
    sleep(5)
    print("-- CONSULTA VEICULO POR CHASSI/VIN --")
    pyautogui.write("901")
    sleep(1)
    pyautogui.write(tela901[0])
    print(PASSO08)
    sleep(1)
    pyautogui.press("enter")
    sleep(2)
    pyautogui.click(x=618, y=66) # salvar tela
    print(PASSO09)
    sleep(1)
    pyautogui.click(x=744, y=870) # clicar "entre com o comando"
    sleep(1)
    pyautogui.press("delete", presses=3)
    sleep(1)
    # -------------------------------------------
    print("-- CONSULTA VEICULO POR MOTOR --")
    pyautogui.write("903")
    sleep(1)
    pyautogui.press('delete', presses=21)
    sleep(1)
    pyautogui.write(tela903[0])
    print(PASSO10)
    sleep(1)
    pyautogui.press("enter")
    sleep(2)
    pyautogui.click(x=618, y=66)
    print(PASSO11)
    sleep(1)
    pyautogui.click(x=744, y=870)
    sleep(1)
    pyautogui.press('delete', presses=3)
    sleep(1)
    # -------------------------------------------
    print("-- CONSULTA CHASSI NA BASE ESTADUAL --")
    pyautogui.write("911")
    sleep(1)
    pyautogui.press('delete', presses=21)
    sleep(1)
    pyautogui.press("enter")
    sleep(1)
    pyautogui.write(tela911[1])
    sleep(1)
    pyautogui.write(tela901[0])
    print(PASSO12)
    sleep(1)
    pyautogui.press("enter")
    sleep(7)
    pyautogui.click(x=618, y=66)
    print(PASSO13)
    sleep(2)

#------------------------------------------------------------------------------------------------------------------------------
# QUARTA PARTE - CADASTRAR VEÍCULO 

def cadastrarChassiNoHOD():
    sleep(1)
    pyautogui.click(x=744, y=870)
    sleep(1)
    pyautogui.press('delete', presses=3)
    sleep(1)
    pyautogui.write("110")
    sleep(1)
    pyautogui.press('delete', presses=21)
    sleep(1)
    pyautogui.press("enter")
    sleep(1)
    # PODE TER 02 OPÇÕES AQUI (INCLUSAO E ALTERACAO)
    pyautogui.press("I")
    sleep(1)
    pyautogui.write(tela901[0])
    sleep(1)
    pyautogui.press("enter")
    print("-- TELA DE CADASTRO DE VEÍCULO --")
    # -- DENTRO DA TELA DE CADASTRO --
    sleep(1)
    pyautogui.press("tab")
    sleep(1)
    pyautogui.write(tela110[12])
    sleep(1)
    pyautogui.write(tela110[13])
    sleep(1)
    # MUDOU O CHECKLIST AGORA É 'O' E TEM Q COLOCAR PARA CONTINUAR ESCREVENDO 'N'.
    pyautogui.write(formatarCondicaoVIN(tela110[14]))
    sleep(1)
    pyautogui.write(formatarProcedencia(tela110[15]))
    sleep(1)
    pyautogui.press("tab")
    sleep(1)
    # MUDOU O LAYOUT DO CHECKLIST E AGORA É SEMPRE 01
    pyautogui.write('1')
    sleep(1)
    pyautogui.press("tab", presses=2)
    sleep(1)
    pyautogui.write(tela110[17])
    sleep(1)
    pyautogui.write(formatarTipoFaturado(tela110[18]))
    sleep(1)
    pyautogui.write(formatarNumeroFaturado(tela110[18]))
    sleep(1)
    pyautogui.press("tab", presses=14)
    sleep(1)
    pyautogui.write(formatarQuantidadeDeEixos(tela110[19]))
    sleep(1)
    pyautogui.press("tab", presses=2)
    sleep(1)
    pyautogui.write(formatarNumeroMotor(tela110[20]))
    sleep(1)
    pyautogui.press("tab", presses=2, interval=0.25)
    sleep(1)
    pyautogui.write(formatarCodigoTipo(tela110[21]))
    sleep(1)
    pyautogui.write(formatarCodigoMarcaModelo(tela110[11]))
    sleep(1)
    pyautogui.write(formatarCodigoEspecie(tela110[22]))
    sleep(1)
    pyautogui.write(formatarCodigoCarroceria(tela110[23]))
    sleep(1)
    pyautogui.write(formatarCodigoCor(tela110[24]))
    sleep(1)
    pyautogui.write(formatarCodigoCombustivel(tela110[25]))
    sleep(1)
    pyautogui.write(formatarLotacao(tela110[26]))
    sleep(2)  
    if 'CV' in tela110[27]:
        pyautogui.write(formatarPotencia(tela110[27]))
    else:
        pyautogui.press('tab')
    sleep(2)
    if 'CC' in tela110[27]:
        pyautogui.write(formatarCilindrada(tela110[27]))
    else:
        pyautogui.press('tab')
    print(PASSO14)

#------------------------------------------------------------------------------------------------------------------------------
# QUINTA PARTE - SALVAR AS TELAS EM FORMATO PDF

def salvarTelasColetadas():
    print('-- SALVANDO TELAS COLETADAS --')
    sleep(2)
    # CLIQUE - PROCESSAR COLETA
    pyautogui.click(x=656, y=65)
    sleep(5)
    # CLICQUE - SELECIONAR TODOS
    pyautogui.click(x=634, y=686)
    sleep(1)
    pyautogui.click(x=634, y=686)
    sleep(5)
    # CLIQUE - IMPRIMIR E EXCLUIR SELECIONADOS 
    pyautogui.click(x=908, y=685)
    sleep(5)
    # CLIQUE - SELECIONAR CAIXA 
    pyautogui.click(x=337, y=121)
    sleep(5)
    # CLIQUE - SELECIONAR OPÇÃO 'Microsoft Print to PDF'
    pyautogui.click(x=284, y=150)
    sleep(5)
    # CLIQUE - BOTÃO 'OK'
    pyautogui.click(x=361, y=360)
    sleep(5)
    pyautogui.click(x=645, y=350)
    sleep(5)
    pyautogui.write('W:\CGSIE\PLACA_AMARELA\PROCESSADOS')
    sleep(1)
    pyautogui.press('enter')
    sleep(5)
    pyautogui.click(x=688, y=735)
    sleep(5)
    # ESCREVER O NOME DO ARQUIVO = NUMERO DO PROCESSO
    print()

#------------------------------------------------------------------------------------------------------------------------------
# SEXTA PARTE - JUSTIFICATIVA DO CADASTRO

def justificarCadastroVINInvalido():
    print('-- JUSTIFICANDO CADASTRO --')
    sleep(2)
    pyautogui.press('F12')
    sleep(2)
    pyautogui.write('Placa Amarela - ${ARQUIVO}') # + número do processo.
    sleep(2)
    pyautogui.press('enter')
    sleep(2)
    print("-- CONSULTA VEICULO POR CHASSI/VIN --")
    pyautogui.write("901")
    sleep(2)
    pyautogui.write(tela901[0])
    print("TELA 901 PREENCHIDA")
    sleep(2)
    pyautogui.press("enter")
    sleep(2)
    print()

#------------------------------------------------------------------------------------------------------------------------------
# SÉTIMA PARTE - ESCREVER NOME DO ARQUIVO PDF = NÚMERO DO PROCESSO

def escreverArquivoFinal():
    arquivos_leitura = os.listdir(PASTA_ARQUIVOS_CHECKLIST)
    arquivos_processados = os.listdir(PASTA_ARQUIVOS_PROCESSADOS)
    arquivos = list(arquivos_leitura - arquivos_processados)

#------------------------------------------------------------------------------------------------------------------------------
# OITAVA PARTE - ESCREVER EMAIL DE RESPOSTA PARA DETRAN

def escreverEmail():
    sleep(2)
    pyautogui.click(x=114, y=333)
    pyautogui.write('Pre-cadastro veículo placa ' + tela110[1] + ' - Processo ' + tela110[0] + ' - CONCLUIDO')
    pyautogui.click(x=71, y=431)
    pyautogui.write('Prezados,')
    pyautogui.press('enter', presses=2)
    pyautogui.write('em atencao ao processo em epigrafe, o qual solicita o pre-cadastro do veículo duas alfas placa ' + tela110[1] + ' , chassi ')
    pyautogui.press('enter')
    pyautogui.write( tela110[9] + ' , informamos a conclusao do mesmo estando o veiculo pre-cadastrado e pronto para o emplacamento')
    pyautogui.press('enter')
    pyautogui.write('conforme extrato em anexo.')
    pyautogui.press('enter', presses=2)
    pyautogui.write('Atenciosamente,')
    pyautogui.press('enter', presses=2)
    pyautogui.write('Coordenacao-Geral de Sistemas, Informacao e Estatistica - CGSIE')
    pyautogui.press('enter')
    pyautogui.write('Departamento de Gestao da Politica de Transito - DGPT')
    pyautogui.press('enter')
    pyautogui.write('Secretaria Nacional de Transito - SENATRAN')
    pyautogui.press('enter')
    pyautogui.write('Ministerio da Infraestrutura - MINFRA')

#------------------------------------------------------------------------------------------------------------------------------

def pesquisarPlacaEmTodosOsEstados():

    estados = ["AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO"]

    print("-- CONSULTA PLACA EM TODOS OS ESTADOS --")

    for estado in estados:
        pyautogui.write("912")
        sleep(2)
        pyautogui.press("enter")
        sleep(2)
        pyautogui.write(estado)
        sleep(2)
        pyautogui.write("AJ9562")
        sleep(2)
        pyautogui.press("enter")
        sleep(10)
        pyautogui.click(x=618, y=66)
        sleep(2)
        pyautogui.click(x=746, y=876)


    




#------------------------------------------------------------------------------------------------------------------------------

def main():
    sleep(5)
    print(PASSO00 + getDataAtual())

    # pesquisarPlacaEmTodosOsEstados()

    # print(PASSO01)
    # executarHOD()
    # print(PASSO05)
    
    
    ARQUIVO = os.listdir(PASTA_ARQUIVOS_CHECKLIST)
    tabela = lerArquivoPDF(ARQUIVO[0])

 
    lista = extrairDadosDaTabela(tabela)
    
    
    preencherTelas(lista)

    consultaDeTelasNoHOD()
    cadastrarChassiNoHOD()

    # escreverEmail()

    # justificarCadastroVINInvalido(ARQUIVO)
    # escreverArquivoFinal()
    # salvarTelasColetadas()

    print('SCRIPT FINALIZADO: ' + getDataAtual())



main()

