import PySimpleGUI as sg
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.graphics.charts.linecharts import HorizontalLineChart
import os
from pprint import pprint
from pytz import timezone
from datetime import datetime
import threading
import time
#import datetime
import pyperclip
#from time import gmtime, strftime
#criar as janelas e estilos (layout)
def formulario():
    sg.theme('DarkAmber')
    layout = [
        [sg.Text('Formulário')],
        [sg.Multiline(size=(90, 20), background_color='black', text_color='white', key='for')],
        #, font='courier 10'
        [sg.Button('Cancelar'), sg.Button('Limpar'), sg.Button('Colar'), sg.Button('Imprimir'), sg.Text('', size=(50,1), key='output')]
    ]
    return sg.Window('CADASTRO AMBIENTAL RURAL DO ESTADO DO ACRE', icon='img\icone.ico',layout=layout, finalize=True)
   
#criar as janelas iniciais
janela1 = formulario()
#criar um loop de leitura de eventos

def impressao(VAR_DESCRICAO_SERVICO, VAR_DESCRICAO_DO_SERVIÇO_outros_, VAR_DESCRICAO_DO_SERVIÇO_outros_complemento_, VAR_NOME_PROPRIETARIO, VAR_CPF, VAR_ENDERECO_PROPRIETARIO, VAR_N_PROPRIETARIO, VAR_TELEFONE, VAR_MUNICIPIO_PROPRIETARIO, VAR_CONTATO, VAR_EMAIL, VAR_N_CAR, VAR_NOME_IMOVEL, VAR_ENDERECO_DO_IMOVEL, VAR_N_DO_IMOVEL, VAR_MUNICIPIO_DO_IMOVEL, VAR_BAIRRO_PROPRIETARIO, VAR_BAIRRO_DO_IMOVEL):
    janela1['output'].update('Imprimindo...')
    #di = str(123456789101112131415161718192021222324252627282930313233343536373839404142434445464748495051)
    nonPDF1 = VAR_CPF + '_'
    data_e_hora_atuais = datetime.now()
    #data_e_hora_atuais = datetime.date.today()
    fuso_horario = timezone('America/Bogota')
    data_e_hora_acre = data_e_hora_atuais.astimezone(fuso_horario)
    data_e_hora_acre_em_texto = data_e_hora_acre.strftime('%d-%m-%Y_%H;%M-%S')
    
    nonPDF2 = data_e_hora_acre_em_texto 
   
    nome_pdf = nonPDF1 + nonPDF2
    desktop = os.path.expandvars("%userprofile%\desktop\pdf\\")
    if not os.path.exists(desktop):
        print("Criando diretorio na area de trabalho...")
        os.mkdir(desktop)
        
    path_salvar = os.path.join(desktop, nome_pdf)
    pdf = canvas.Canvas(f'{path_salvar}.pdf')
    
    pdf.drawImage("img\logo.jpg", 60, 740, width=60,height=80)
    
    pdf.setFont("Courier", 12)
    
    varY = 468
    #VAR_NOME_PROPRIETARIO = 'Raylan Oliveira'
    pdf.drawString(100,varY - 28, VAR_NOME_PROPRIETARIO) #max 31 caractere
	
	#VAR_CPF = '020.444.722-55'
    pdf.drawString(428,varY - 28, VAR_CPF)
	
	#VAR_ENDERECO_PROPRIETARIO = 'rua fulano da silva'
    pdf.drawString(117,varY - 49, VAR_ENDERECO_PROPRIETARIO) #max 43 caractere
	
	#VAR_N_PROPRIETARIO = 'completo Endereço 1'
    pdf.drawString(84,varY - 70, VAR_N_PROPRIETARIO) # max 48 car
    
    #VAR_BAIRRO_PROPRIETARIO = 'QQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQ'
    pdf.drawString(213,varY - 233, VAR_BAIRRO_PROPRIETARIO)
	
	#VAR_TELEFONE = '68 99666-9999'
    pdf.drawString(113,varY - 91, VAR_TELEFONE)
	
	#VAR_MUNICIPIO_PROPRIETARIO = 'Acrelandia'
    pdf.drawString(120,varY - 112, VAR_MUNICIPIO_PROPRIETARIO) #max 43 car
	
	#VAR_CONTATO = '68 99999-9999'
    pdf.drawString(390,varY - 91, VAR_CONTATO)
	    
    VAR_EMAIL = 'email@gmail.com'
    #pdf.drawString(319,varY - 112, VAR_EMAIL)
	    
    #VAR_N_CAR = 'AC-1200013-E5388324573843A0A937C5BF62DE221B'
    pdf.drawString(153,varY - 170, VAR_N_CAR)
	    
    #VAR_NOME_IMOVEL = 'Nome do Imovel Completo'
    pdf.drawString(180,varY - 191, VAR_NOME_IMOVEL) #max 37 car
	    
    #VAR_ENDERECO_DO_IMOVEL = 'ENDERECO ENDERECO'
    pdf.drawString(117,varY - 212, VAR_ENDERECO_DO_IMOVEL) #max 43 carc
	    
    #VAR_N_DO_IMOVEL = 'COMPLEMENTO END 2'
    
    pdf.drawString(84,varY - 233, VAR_N_DO_IMOVEL)	# max 48 car

    
    pdf.drawString(213,varY - 70, VAR_BAIRRO_DO_IMOVEL) 
	    
    #VAR_MUNICIPIO_DO_IMOVEL = 'MUNICIPIO 2'
    pdf.drawString(120,varY - 254, VAR_MUNICIPIO_DO_IMOVEL) #max 43 car
    
    pdf.setFont("Times-Bold", 13)
    pdf.drawString(125,790, 'GOVERNO DO') 
    pdf.drawString(125,775, 'ESTADO DO ACRE')
    
    pdf.setFont("Courier", 12)
    pdf.drawString(125,760, 'www.ac.gov.br')
    
    pdf.setFont("Helvetica-Oblique", 10)
    pdf.drawString(265,800, 'SECRETARIA DE ESTADO DO MEIO AMBIENTE')                            
    pdf.drawString(265,788, 'E DAS POLÍTICAS INDÍGENAS - SEMAPI')
    pdf.drawString(265,776, 'INSTITUTO DE MEIO AMBIENTE DO ACRE – IMAC')
    pdf.drawString(265,764, 'ESCRITÓRIO TÉCNICO DE GESTÃO DO CAR E PRA - ACRE')
    
    pdf.drawImage("img\line.jpg", 60, 735, width=500,height=5)
    
    pdf.setFont("Times-Bold", 12)
    pdf.drawString(85,715, 'Requerimento para Solicitação de Serviços no âmbito do Cadastro Ambiental Rural e')
    pdf.drawString(150,700, 'Programa de Regularização Ambiental do Estado do Acre')
    
    
    pdf.drawString(60,670, '1. Descrição do serviço:')
    
    pdf.setFont("Times-Roman", 11)
    
    varY = 639
    txt = VAR_DESCRICAO_SERVICO
    
    suporteX = "SUPORTE" in txt
    if suporteX:
        pdf.drawString(68,varY, 'X') 
    
    prioritaria_X = "PRIORITÁRIA" in txt        
    if prioritaria_X:
        pdf.drawString(68,varY - 20, 'X')
        
    adesaoX = "ADESÃO" in txt
    if adesaoX:
        pdf.drawString(68,varY - 40, 'X')
        
    insercaoX = "INSERÇÃO" in txt
    if insercaoX:
        pdf.drawString(68,varY - 60, 'X')
        
    notificacaoX = "NOTIFICAÇÃO" in txt
    if notificacaoX:
        pdf.drawString(68,varY - 80, 'X')
        
    cancelamentoX = "CANCELAMENTO" in txt
    if cancelamentoX:
        pdf.drawString(68,varY - 100, 'X')
        
    outrosX = "Outros" in txt
    if outrosX:
        pdf.drawString(68,varY - 120, 'X')        
        pdf.setFont("Courier", 12)
        #pdf.drawString(130,540 - 20, digt)
        pdf.drawString(130,540 - 20, VAR_DESCRICAO_DO_SERVIÇO_outros_)#maximo 42 caractere
        pdf.drawString(65,540 - 42, VAR_DESCRICAO_DO_SERVIÇO_outros_complemento_)#maximo 48 caractere
        
    pdf.setFont("Courier-Bold", 12)
    pdf.line(85,653,60,653) #-        
    pdf.line(85,632,60,632) #-
    pdf.line(85,612,60,612) #-
    pdf.line(85,592,60,592) #-
    pdf.line(85,572,60,572) #-
    pdf.line(85,552,60,552) #-
    pdf.line(60,532,85,532) #-
    
    pdf.line(60,653,60,532 - 21) #||
    pdf.line(85,653,85,532 - 21) #||
    
    pdf.drawString(90,640, 'SUPORTE TÉCNICO PARA RETIFICAÇÃO DO CAR')
    pdf.drawString(90,620, 'ANÁLISE PRIORITÁRIA DO CAR')
    pdf.drawString(90,600, 'ADESÃO AO PROGRAMA DE REGULARIZAÇÃO AMBIENTAL DO ESTADO DO ACRE')
    pdf.drawString(90,580, 'INSERÇÃO DE PERIMETRAL')
    pdf.drawString(90,560, 'NOTIFICAÇÃO DA ANÁLISE')
    pdf.drawString(90,540, 'CANCELAMENTO DO CAR')    
    
    pdf.setFont("Times-Bold", 11)        
    pdf.drawString(90,540 - 20, 'Outros: ')        
            
    pdf.line(60,532 - 21,550,532 - 21) #__________        
    pdf.line(60,532 - 42,550,532 - 42) #__________        
    
    varX = 60
    varY = 468
    
    pdf.drawString(60,varY, '2. Identificação do Proprietário/Possuidor')
    
    #pdf.line(60,495,550,495) #__________
    pdf.line(60,varY - 15,550,varY - 15) #__________
    pdf.drawString(65,varY - 28, 'Nome: ')
    
    pdf.line(60,varY - 36,550,varY - 36) #__________
    pdf.line(395,varY - 15,395,varY - 35) #|
    pdf.drawString(400,varY - 28, 'CPF: ')
    
    pdf.drawString(65,varY - 49, 'Endereço: ')
    

    
    pdf.line(60,varY - 57,550,varY - 57) #__________
    pdf.drawString(65,varY - 70, 'Nº: ')
    
    pdf.drawString(175,varY - 70, 'Bairro: ')    
      
    pdf.line(170,varY-57,170,varY-78) #|
    pdf.line(60,varY - 78,550,varY - 78) #__________
    pdf.drawString(65,varY - 91, 'Telefone: ')
    
    pdf.drawString(65,varY - 112, 'Município: ')  
    
    pdf.line(60,varY - 99,550,varY - 99) #__________
    pdf.drawString(280,varY - 91, 'Telefone para contato: ')
    
    #pdf.drawString(280,varY - 112, 'E-mail: ')

    
    pdf.line(60,varY - 120,550,varY - 120) #__________
    pdf.line(275,varY-78,275,varY-99) #|
    pdf.line(60,varY-15,60,varY - 120) #||
    pdf.line(550,varY - 15,550,varY - 120) #||
    
    pdf.drawString(60,varY - 142, '3. Identificação do Imóvel Rural')
    
    pdf.line(60,varY - 157,550,varY - 157) #__________
    pdf.drawString(65,varY - 170, 'Registro no CAR: ')

    
    pdf.line(60,varY - 178,550,varY - 178) #__________
    pdf.drawString(65,varY - 191, 'Nome do Imóvel Rural: ')

    
    pdf.line(60,varY - 199,550,varY - 199) #__________
    pdf.drawString(65,varY - 212, 'Endereço: ')

    pdf.drawString(65,varY - 233, 'Nº: ')
    
    pdf.drawString(175,varY - 233, 'Bairro: ')
     
    pdf.line(170,varY-220,170,varY-241) #|
    
    pdf.line(60,varY - 220,550,varY - 220) #__________

    
    pdf.line(60,varY - 241,550,varY - 241) #__________
    
    pdf.drawString(65,varY - 254, 'Município: ')

    
    pdf.line(60,varY - 262,550,varY - 262) #__________
    pdf.line(60,varY - 157,60,varY - 262) #||
    pdf.line(550,varY - 157,550,varY - 262) #||
    
    varY = 470
    #strftime("%d_%b_%Y_%I-%M-%S", gmtime())
    Meses=('janeiro','fevereiro','mar','abril','maio','junho',
       'julho','agosto','setembro','outubro','novembro','dezembro')
       
    agora = datetime.now()
    
    mesNUM = (agora.month-1)
    mes = Meses[mesNUM]
    
    dia = agora.strftime('%d')
    ano = agora.strftime('%Y')
    pdf.drawString(60,varY - 290, 'Rio Branco – Ac, '+ str(dia) +' de '+ str(mes) +' de '+ str(ano) +'.')
    
    pdf.line(370,varY - 280,550,varY - 280) #_____
    pdf.line(370,varY - 280,370,varY - 400) #|
    pdf.line(550,varY - 280,550,varY - 400) #|
    pdf.setFont("Courier-Bold", 11)
    pdf.drawString(429,varY - 295, 'Protocolo')
    
    pdf.setFont("Times-Bold", 11)
    
    mes = agora.strftime('%m')
    
    pdf.drawString(422,varY - 320, 'Em: '+str(dia)+'/'+str(mes)+'/'+str(ano))
    hora = data_e_hora_acre.strftime('%H:%M')
    pdf.drawString(425,varY - 339, 'Às '+hora+' horas')
    pdf.line(385,varY - 375,535,varY - 375) #_____
    pdf.setFont("Times-Bold", 10)
    pdf.drawString(419,varY - 388, 'Atendimento IMAC')
    pdf.line(370,varY - 400,550,varY - 400) #_____
    
    
    
    pdf.line(60,varY - 360,300,varY - 360) #______
    pdf.drawString(115,varY - 373, 'Assinatura do Requerente')        
    
    varY = 500
    
    pdf.setFont("Times-Roman", 8)
    pdf.drawString(175,varY - 460, 'Rua Rui Barbosa, 135, Centro - CEP: 69.900-084 – Rio Branco - Acre - Brasil')
    pdf.drawString(235,varY - 470, 'Fone: +55 (68) 3224-5497/3223-2789')
    pdf.drawString(200,varY - 480, 'E-mail: car.acre@ac.gov.br // Homepage: www.car.ac.gov.br')
    
    
    pdf.save()
    #print('{}.pdf criado com sucesso!'.format(nome_pdf))
    #os.startfile(format(nome_pdf)+".pdf")
    os.startfile(format(path_salvar)+".pdf")
    #sg.popup('{}.pdf criado com sucesso!'.format(nome_pdf))
    janela1['output'].update("PDF criado com sucesso!")

while True:
    window,event,values = sg.read_all_windows()
    #quando janela for fechada
    if window == janela1 and event == sg.WIN_CLOSED:
        break
        
    if window == janela1 and event == 'Cancelar':
        break
        
    if window == janela1 and event == 'Limpar':
        janela1['for'].update('')
        janela1['output'].update('')
        #sg.popup(texto)
        continue
        
    if window == janela1 and event == 'Colar':
        janela1['output'].update('')
        janela1['for'].update('')
        janela1['for'].update(pyperclip.paste())
        continue
        
    if window == janela1 and event == 'Imprimir':
        janela1['output'].update('Imprimindo...')
        texto = values['for']
        texto_cortado = texto.split("\t")
        se_imprimir = texto.count("\t") #contar os tab
        pprint("Quantidade de dados (\t): " + str(se_imprimir))
        #if se_imprimir >= 20 and se_imprimir <= 23 and texto_cortado[6] != '' and texto_cortado[20].isnumeric() and texto_cortado[19].isnumeric():
        if se_imprimir >= 20:
            janela1['output'].update('Imprimindo...')
            if texto_cortado[2] == '':
                #sg.popup("vazio")
                janela1['output'].update('DESCRIÇÃO DO SERVIÇO VAZIO!')            
            elif texto_cortado[5] == '':
                janela1['output'].update('NOME DO IMÓVEL VAZIO!')            
            
            elif texto_cortado[6] == '':
                janela1['output'].update('NOME DO PROPRIETÁRIO VAZIO!') 
            
            elif texto_cortado[7] == '':
                janela1['output'].update('CPF VAZIO!')
                
            elif not texto_cortado[20].isnumeric():
                janela1['output'].update('Nº DO IMÓVEL INCORRETO!')  
             
            elif texto_cortado[12] == '':
                janela1['output'].update('MUNICÍPIO PROPRIETÁRIO VAZIO!')
                
            elif texto_cortado[11] == '':
                janela1['output'].update('ENDEREÇO PROPRIETÁRIO VAZIO!')
            
            elif texto_cortado[10] == '':
                janela1['output'].update('NÚMERO DE REGISTRO CAR VAZIO!')
            
            elif texto_cortado[13] == '':
                janela1['output'].update('ENDEREÇO DO IMÓVEL VAZIO!')
                
            elif texto_cortado[14] == '':
                janela1['output'].update('MUNICÍPIO DO IMÓVEL VAZIO!')
                
             
            else:
            
                VAR_NOME_PROPRIETARIO = texto_cortado[6]
                pprint('NOME: ' + VAR_NOME_PROPRIETARIO)
                
                VAR_CPF = texto_cortado[7]
                pprint('CPF: ' + VAR_CPF)
                
                VAR_TELEFONE = texto_cortado[8]
                pprint('TELEFONE PRINCIPAL: ' + VAR_TELEFONE)
                
                VAR_CONTATO = texto_cortado[9]
                pprint('TELEFONE SECUNDÁRIO: ' + VAR_CONTATO)
                
                VAR_N_CAR = texto_cortado[10]
                pprint('NÚMERO DE REGISTRO CAR: ' + VAR_N_CAR)
                
                
                VAR_ENDERECO_PROPRIETARIO = texto_cortado[11]
                pprint('ENDEREÇO PROPRIETÁRIO: ' + VAR_ENDERECO_PROPRIETARIO)
                
                VAR_N_PROPRIETARIO = str(texto_cortado[19])
                pprint('Nº PROPRIETÁRIO: ' + VAR_N_PROPRIETARIO)                
                
                
                VAR_ENDERECO_DO_IMOVEL = texto_cortado[13]
                pprint('ENDEREÇO DO IMÓVEL: ' + VAR_ENDERECO_DO_IMOVEL)
                
                nImovel = int(texto_cortado[20])
                VAR_N_DO_IMOVEL = str(nImovel)
                pprint('Nº DO IMÓVEL: ' + VAR_N_DO_IMOVEL)
                
                VAR_MUNICIPIO_DO_IMOVEL = texto_cortado[14]
                pprint('MUNICÍPIO DO IMÓVEL: ' + VAR_MUNICIPIO_DO_IMOVEL)
                
                VAR_DESCRICAO_DO_SERVIÇO_outros_ = texto_cortado[15]
                pprint('DESCRIÇÃO DO SERVIÇO (outros): ' + VAR_DESCRICAO_DO_SERVIÇO_outros_)
                
                VAR_DESCRICAO_DO_SERVIÇO_outros_complemento_ = texto_cortado[16]
                pprint('DESCRIÇÃO DO SERVIÇO (outros complemento): ' + VAR_DESCRICAO_DO_SERVIÇO_outros_complemento_)
                
                VAR_EMAIL = ''
                pprint('Email: ' + VAR_EMAIL)
                
                VAR_BAIRRO_PROPRIETARIO = texto_cortado[18]
                pprint('BAIRRO DO IMÓVEL: ' + VAR_BAIRRO_PROPRIETARIO)
                
                VAR_BAIRRO_DO_IMOVEL = texto_cortado[17]
                pprint('BAIRRO PROPRIETÁRIO: ' + VAR_BAIRRO_DO_IMOVEL)                
            
            
                VAR_NOME_IMOVEL = texto_cortado[5]
                pprint('NOME DO IMÓVEL: ' + VAR_NOME_IMOVEL)
                
                VAR_DESCRICAO_SERVICO = texto_cortado[2]
                pprint('DESCRIÇÃO DO SERVIÇO: '+VAR_DESCRICAO_SERVICO)
                
                VAR_MUNICIPIO_PROPRIETARIO = texto_cortado[12]
                pprint('MUNICÍPIO PROPRIETÁRIO: ' + VAR_MUNICIPIO_PROPRIETARIO)                
                
                janela1['output'].update('Imprimindo...')
                threading.Thread(target=impressao, args=(VAR_DESCRICAO_SERVICO, VAR_DESCRICAO_DO_SERVIÇO_outros_[:58], VAR_DESCRICAO_DO_SERVIÇO_outros_complemento_[:67], VAR_NOME_PROPRIETARIO[:40], VAR_CPF[:14], VAR_ENDERECO_PROPRIETARIO[:60], VAR_N_PROPRIETARIO[:11], VAR_TELEFONE[:13], VAR_MUNICIPIO_PROPRIETARIO[:59], VAR_CONTATO[:13], VAR_EMAIL, VAR_N_CAR[:43], VAR_NOME_IMOVEL[:51], VAR_ENDERECO_DO_IMOVEL[:60], VAR_N_DO_IMOVEL[:11], VAR_MUNICIPIO_DO_IMOVEL[:59], VAR_BAIRRO_PROPRIETARIO[:46], VAR_BAIRRO_DO_IMOVEL[:46]), daemon=True).start() 
                #sg.popup('{}.pdf criado com sucesso!'.format(nome_pdf))
                
        else:
            janela1['output'].update('Falta dados para impressão ou dados incorretos!') #if se_imprimir >= 20 and se_imprimir <= 24
            
        