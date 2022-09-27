from time import sleep
import PySimpleGUI as sg
from playwright.sync_api import sync_playwright
import os
from datetime import datetime
import smtplib
from openpyxl import Workbook, load_workbook
from datetime import date


cod = ''
codLista = []
credencialEmail = open('credencialEmail_AUT.txt', 'r')
loginEmail = []

for linhas in credencialEmail:
    linhas = linhas.strip()
    loginEmail.append(linhas)
usuario_email = loginEmail[0][17:-1]
senha_email = loginEmail[1][15:-1]
s = smtplib.SMTP('smtp.gmail.com: 587')
s.starttls()
s.login(usuario_email, senha_email)

btCalendario = b'iVBORw0KGgoAAAANSUhEUgAAABQAAAAUCAYAAACNiR0NAAAACXBIWXMAAA3XAAAN1wFCKJt4AAAAIGNIUk0AAHolAACAgwAA+f8AAIDpAAB1MAAA6mAAADqYAAAXb5JfxUYAAAGuSURBVHjarNS/b85RFMfx11N9tKiGikTSwSKxSAWxdDb4AyQmg/9BiJh0sFjE1jC0xKSLwSCpkjCQdFASosEgpEFoY6BazTH0PM1Jk8ojnpPcnPf9fu8993zO/SEilDYVEd/Wfftbm4uI6fqt26odw2EMowfntGd7sp3FDO6JiPPRORtpRERgFifxCH042GaGM+kPYRxDLcl38AxnsA/P2wx4CV9z7gSGGhHxGw9xEXNYwt42A77HJgxm3Y83ImIeO3TGFhoRsZS1GM30e7EtuYl+zCMwgO9Yxi78wM/k0zjanZOmcf0/s9uP4a7sLKcP3Cj8KvlX9uXiLb5aeBFrB3tr+lFMFZ5NvpylgLEMCpPYnNy3msaq3f6H67ZRG4+IaEn+VGReKfwg+S0+J98tMkcKf6ySW/4x3hV+UerWKsvL3Hn4gKfJzSp5ogOSx6rkL+lncaHwzeT7eJJ8rWxWtZUqdbE8R9sLDyTvXttFdua/9bZUAw6m7y8DKg8VPrHBwe6BRkTM5Yq3sJCLNPNKdeXAloLezGQFW0qwJk5hXkQciIjXHXhc30TEkT8DAFILwAACEvTGAAAAAElFTkSuQmCC'

passwords = open('credenciais.txt', 'r')
login = []

for linhas in passwords:
    linhas = linhas.strip()
    login.append(linhas)
usuario_me = login[0][14:-1]
senha_me = login[1][12:-1]
site = login[2][8:-1]

filiais_caminho = open('filiais.txt', 'r', encoding="UTF-8")
filiais = []

for linhas in filiais_caminho:
    linhas = linhas.strip()
    filiais.append(linhas)

listaTipo = []
dicioTipo = {}
with open("TipoRequisicao.txt", encoding="UTF-8") as dicionarioTipoReq:
    for line in dicionarioTipoReq:
       (k, v) = line.split(';')
       dicioTipo[str(k)] = v
    for chave in dicioTipo.keys():
        listaTipo.append(chave)

cod_caminho = open('codigos.txt', 'r', encoding="UTF-8")
codigos = []

for linhas in cod_caminho:
    linhas = linhas.strip()
    codigos.append(linhas)

cc_caminho = open('centrocustos.txt', 'r', encoding="UTF-8")
centroCustos = []

for linhas in cc_caminho:
    linhas = linhas.strip()
    centroCustos.append(linhas)

categorias_caminho = open('categorias.txt', 'r', encoding="UTF-8")
categorias = []

for linhas in categorias_caminho:
    linhas = linhas.strip()
    categorias.append(linhas)

tabela = load_workbook('notas.xlsm', data_only=True)
aba_ativa = tabela['REQUISIÇÕES PENDENTES']
ultimaLinha = 'B' + str(len(aba_ativa['B'])+1)

menu_def=[['Arquivos', ['Itens', 'Categorias', 'Centro de custos','Tipo requisição', 'Filiais','---','Credenciais ME','Monitorar requisições']]]

layout = [
    [sg.Menu(menu_def, pad=(10,10)),sg.Push()],
    [sg.Checkbox('Abre nav', default=False, key="abrirNav",enable_events=True),sg.Text('      Título da requisição'), sg.Push(), sg.Checkbox('Monitor', default=False, key="monitorReq",enable_events=True) ,sg.Push()],
    [sg.Push(), sg.Input(key='titulo_requisicao', enable_events=True), sg.Push()],
    [sg.Text('      Tipo de requisição'), sg.Push()],
    [sg.Push(), sg.Combo(listaTipo, key='tipoRequisicao', enable_events=True, size=(43,1), readonly=True), sg.Push()],
    [sg.Text('      Item'), sg.Push(), sg.Text('            Valor unitário'), sg.Push()],
    [sg.Push(), sg.Combo(codigos, size=(18, 1), key='item', enable_events=True), sg.Push(), sg.Input(key='valorun', size=(20, 1), enable_events=True), sg.Push()],
    [sg.Text('      Quantidade'), sg.Push(), sg.Text('    Data esperada'), sg.Push(), ],
    [sg.Push(), sg.Input(key='quant', size=(20, 1), enable_events=True), sg.Push(), sg.Input(key='data_esperada', size=(19, 1), enable_events=True),sg.CalendarButton('',close_when_date_chosen=True,  target='data_esperada', no_titlebar=False, format='%d/%m/%Y', image_data=btCalendario)],
    [sg.Text('      Selecione a categoria'),sg.Text('                 Centro de custo'), sg.Push()],
    [sg.Push() ,sg.Combo(categorias, key='catPedido', size=(23,1), readonly=True, enable_events=True),sg.Combo(centroCustos, size=(16, 1), key='centrocusto', readonly=True, enable_events=True), sg.Push()],
    [sg.Text('      Selecione a filial'), sg.Push()],
    [sg.Push(), sg.Combo(filiais, key='filial', size=(43,1), readonly=True, enable_events=True), sg.Push()],
    [sg.Text('      Comentario:')],
    [sg.Push(), sg.Input(key='comentario'),sg.Push()],
    [sg.Text("      Selecionar arquivo: (Apenas pedido de regularização)", key='txtSelecionaArquivo', visible=False)], 
    [sg.Push(), sg.Input(size=(35,1), key='inputCaminhoArquivo', visible=False, enable_events=True), sg.FilesBrowse('Procurar', key="caminhoArquivo", visible=False), sg.Text('',size=(2,1), key='espacamentoCaminho', visible=False)],
    [sg.Push(), sg.Text('', key='mensagemCampoVazio'), sg.Push(),],
    [sg.Push(), sg.Button('Criar requisição', size=(17, 1),key='botaoCriar', disabled=True), sg.Push(), sg.Button('Cancelar', size=(17, 1)),sg.Button('Limpar', size=(17, 1), key='limpar'), sg.Push()],
    [sg.Push(), sg.ProgressBar(max_value=7, orientation='h', size=(20, 20), key='progress', visible=False),sg.Push(), sg.Text('', visible=False, key='espacamento',size=(8,1))],
    [sg.Push(), sg.Text('', key='mensagemProgresso'), sg.Push(),],
    [sg.Push(), sg.Text('', key='mensagem2'), sg.Push(),],
    [sg.Push(), sg.Text('', key='mensagem3'), sg.Push(), sg.Input('', key='mensagem4',size=(18,1),disabled=True, visible=False), sg.Push(), sg.Text('', visible=False, key='espacamento2',size=(8,1))]
    ]

window = sg.Window('Abertura de requisições', size=(400, 600), layout = layout)
progress_bar = window['progress']
while True:
    event, values = window.read()

    if event == None:
        break
    
    def monitorME():
        with sync_playwright() as p:
            browser = p.chromium.launch(channel="chrome", headless=False)
            page = browser.new_page()
            page.goto(site)

            # LOGIN ME
            page.locator('xpath=//*[@id="LoginName"]').fill(usuario_me)
            page.locator('xpath=//*[@id="RAWSenha"]').fill(senha_me)
            page.locator('xpath=//*[@id="SubmitAuth"]').click()
            page.wait_for_timeout(1)

            # ANALISA STATUS DA REQUISIÇÃO E ATUALIZA PLANILHA
            for celula in aba_ativa['I']:
                linha = celula.row
                if celula.value == 'Pendente' and aba_ativa[f'E{linha}'].value == None:
                    cnpj = aba_ativa[f'D{linha}'].value
                    reqPendente = aba_ativa[f'B{linha}'].value
                    page.goto(f'https://www.me.com.br/DO/Request/Home.mvc/Show/{reqPendente}')
                    statusRequisicao = page.locator('//*[@id="formRequest"]/div/div[2]/div[2]/p[2]/span[2]').inner_html().strip()

                    if statusRequisicao == 'APROVADO':
                    #CRIAR PRE-PEDIDO
                        sg.popup(f'Requisição {reqPendente} aprovada!\nCriando pré-pedido')
                        # page.locator('xpath=//*[@id="btnEmergency"]').click()
                        # page.locator('xpath=/html/body/div[1]/div[3]/div/button[1]/span').click()
                        # page.locator('xpath=//*[@id="MEComponentManager_MEButton_2"]').click()
                        # page.locator('xpath=//*[@id="CGC"]').fill(cnpj)
                        # page.keyboard.press('Enter')
                        # page.locator('xpath=//*[@id="grid"]/div[2]/table/tbody/tr/td[1]/div/input').click()
                        # page.locator('xpath=//*[@id="btnSalvarSelecao"]').click()
                        # page.locator('xpath=//*[@id="btnVoltarPrePedEmergencial"]').click()
                        # page.locator('xpath=//*[@id="Resumo"]').fill(values['data_esperada'])
                        # filiaisPrePedido = page.locator('//select[@name="LocalCobranca"]').inner_html().split('\n')
                        # indice = [i for i, s in enumerate(filiaisPrePedido) if nome_filial in s][0]
                        # page.locator('//select[@name="LocalCobranca"]').select_option(index=indice-1)
                        # page.locator('xpath=//*[@id="DataEntrega"]').fill(values['titulo_requisicao'])
                        # page.locator('xpath=//*[@id="MEComponentManager_MEButton_3"]').click()
                        # page.locator('xpath=/html/body/main/form[2]/table[3]/tbody/tr[1]/td/input[1]').click()
                        # page.locator('xpath=//*[@id="MEComponentManager_MEButton_2"]').click()
                        # page.locator('xpath=//*[@id="MEComponentManager_MEButton_2"]').click()
                        # page.locator('xpath=//*[@id="formItemStatusHistory"]/div/b[1]/a').click()
                        # numPrePedido = page.locator('xpath=/html/body/main/div/div[1]/div[1]/p').inner_html().strip()
                        # statusPrePedido = page.locator('xpath=/html/body/main/div/div[1]/div[2]/div[2]/p[1]/span[2]').inner_html().strip()
                        # aba_ativa[f'F{linha}'] = date.today().strftime('%d/%m/%Y')
                        # aba_ativa[f'E{linha}'] = numPrePedido
                    
                if celula.value == 'Pendente' and aba_ativa[f'E{linha}'].value != None:
                    
                    prePedidoPendente = aba_ativa[f'E{linha}'].value
                    page.goto(f'https://www.me.com.br/VerPrePedidoWF.asp?Pedido={prePedidoPendente}&SuperCleanPage=false&Origin=home')
                    statusPrePedido = page.locator('xpath=/html/body/main/div/div[1]/div[2]/div[2]/p[1]/span[2]').inner_html().strip()[:8]
                    if statusPrePedido == 'APROVADO':
                        numPedidoSAP = page.locator('xpath=/html/body/main/div/div[1]/div[1]/p[1]').inner_html().strip()
                        aba_ativa[f'G{linha}'] = numPedidoSAP
                        sg.popup(f'Pré-Pedido {prePedidoPendente} aprovado!\nO número do seu pedido é {numPedidoSAP}')

            tabela.save('Tabelateste.xlsx')
    
    def limpaCampos(): # Limpa todos os campos
        window['titulo_requisicao'].update('')
        values['titulo_requisicao'] = ''
        window['item'].update('')
        values['item'] = ''
        window['valorun'].update('')
        values['valorun'] = ''
        window['quant'].update('')
        values['quant'] = ''
        window['data_esperada'].update('')
        values['data_esperada'] = ''
        window['centrocusto'].update('')
        values['centrocusto'] = ''
        window['filial'].update('')
        values['filial'] = ''
        window['inputCaminhoArquivo'].update('')
        values['inputCaminhoArquivo'] = ''
        window['catPedido'].update('')
        values['catPedido'] = ''
        window['comentario'].update('')
        values['comentario'] = ''
        window['tipoRequisicao'].update('')
        values['tipoRequisicao'] = ''
        cod = ''
        validaPedido()
        validacao()
        window['botaoCriar'].update(disabled=True)

    def validacao():   # Verifica se todos campos estão preenchidos
        if values['titulo_requisicao'] and values['catPedido'] and values['item'] and values['valorun'] and values['quant'] and values['data_esperada'] and values['centrocusto'] and values['filial'] != '':
            if values['catPedido'] == 'PEDIDO REGULARIZACAO' and values['inputCaminhoArquivo'] == '':
                window['botaoCriar'].update(disabled=True)
            else:
                window['botaoCriar'].update(disabled=False)

    def validaPedido(): # Oculta campo de inserir arquivo
        if values['catPedido'] == 'PEDIDO REGULARIZACAO':
            window['txtSelecionaArquivo'].update(visible = True)
            window['inputCaminhoArquivo'].update(visible = True)
            window['caminhoArquivo'].update(visible = True)
            window['espacamentoCaminho'].update(visible = True)
        else:
            window['espacamentoCaminho'].update(visible = False)
            window['txtSelecionaArquivo'].update(visible = False)
            window['inputCaminhoArquivo'].update(visible = False)
            window['caminhoArquivo'].update(visible = False)
    while values['monitorReq'] == True:
        sleep(300)
        monitorME()
    validacao()
    match(event):
        case 'tipoRequisicao':
            cod = dicioTipo[values["tipoRequisicao"]] + ';' + cod
            window['item'].update(cod)
            codLista.append(dicioTipo[values["tipoRequisicao"]])
            
        case 'catPedido':
            validacao()
            validaPedido()
            window['inputCaminhoArquivo'].update('')
            values['inputCaminhoArquivo'] = ''
        case 'limpar':
            limpaCampos()
            cod = ''
        case 'Itens':
            os.system('codigos.txt')
        case 'Monitorar requisições':
            monitorME()
        case 'Categorias':
            os.system('categorias.txt')
        case 'Centro de custos':
            os.system('centrocustos.txt')
        case 'Filiais':
            os.system('filiais.txt')
        case 'Credenciais ME':
            os.system('credenciais.txt')
        case 'Tipo requisição':
            os.system('TipoRequisicao.txt')

        case sg.WIN_CLOSED:
            break
        case None:
            break
        case 'Cancelar':
            break
        case 'botaoCriar': 
            comentario = values['comentario']
            caminho_arquivo = str(values["caminhoArquivo"]).split(';')
            centro_custo = values['centrocusto']
            cat_Pedido = values['catPedido']
            titulo_requisicao = values['titulo_requisicao']
            item = str(values['item']).replace(' ','').split(";")
            valorun = str(values['valorun']).replace(',','.').split(";")
            quant = str(values['quant']).split(";")
            data_esperada = values['data_esperada']
            filial = values['filial']
            nome_filial = filial.split('-',1)[1][1:]
            
            
            window['mensagem2'].update('')
            window['mensagem3'].update('')
            window['mensagem4'].update(visible=False)
            window['espacamento2'].update(visible = False)
            window['mensagem4'].update('')

            progress_bar.update(visible = True)
            window['espacamento'].update(visible = True)

            with sync_playwright() as p:
                
                # valorTotal = str(float(valorun) * int(quant))
                if values['abrirNav']:
                    browser = p.chromium.launch(channel="chrome",headless=False)
                else:
                    browser = p.chromium.launch(channel="chrome")
                page = browser.new_page()
                page.goto(site)
                progress_bar.update(visible = True)
                window['espacamento'].update(visible = True)

                # LOGIN ME
                window['mensagemProgresso'].update('Efetuando login no ME')      
                progress_bar.update_bar(1) 
                page.locator('xpath=//*[@id="LoginName"]').fill(usuario_me)
                page.locator('xpath=//*[@id="RAWSenha"]').fill(senha_me)
                page.locator('xpath=//*[@id="SubmitAuth"]').click()

                # CONFIGURAÇÃO DA REQUISIÇÃO
                window['mensagemProgresso'].update('Configurando a requisição')      
                progress_bar.update_bar(2)
                page.locator('xpath=//*[@id="__layout"]/div/main/div/div/div[2]/div/div[1]/section/div/div/div/button').click()
                page.locator('xpath=//*[@id="__layout"]/div/main/div/div/div[2]/div/div[1]/section/div[1]/div[2]/ul/li[1]/a').click()
                page.locator('xpath=//*[@id="__layout"]/div/main/div/div/div[2]/div/div[1]/section/div[1]/div[2]/ul[2]/li[1]/a').click()
                frame = page.frame_locator('#PopUpConfiguration-if')
                page.wait_for_timeout(1000)
                frame.locator('xpath=//*[@id="select2-Categoria_Value-container"]').click()
                page.wait_for_timeout(500)
                frame.locator('xpath=/html/body/span/span/span[1]/input').fill(cat_Pedido)
                page.wait_for_timeout(500)
                frame.locator('xpath=/html/body/span/span/span[1]/input').press('Enter')
                page.wait_for_timeout(500)
                frame.locator('xpath=//*[@id="BOrgs_1__BorgDescription"]').press('Tab')
                page.wait_for_timeout(500)
                frame.locator('xpath=//*[@id="BOrgs_1__BorgDescription"]').fill(filial)
                page.wait_for_timeout(500)
                frame.locator('xpath=//*[@id="BOrgs_1__BorgDescription"]').press('Tab')
                frame.locator('xpath=//*[@id="btnSave"]').click()

                # SELECIONA ITENS E QUANTIDADES
                window['mensagemProgresso'].update('Adicionando itens e quantidades')      
                progress_bar.update_bar(3)
                for i in range(len(quant)):
                    page.wait_for_timeout(1000)
                    page.locator('xpath=//*[@id="Valor"]').fill(item[i])
                    page.locator('xpath=//*[@id="btnSearchSimple"]').click()
                    page.locator('.icon-shopping-cart').click()
                    page.wait_for_timeout(1000)
                    page.keyboard.press('Control+A')
                    page.wait_for_timeout(1000)
                    page.keyboard.type(quant[i])
                    page.wait_for_timeout(1000)
                    page.keyboard.press('Tab')
                    page.wait_for_timeout(1000)
                page.locator('xpath=//*[@id="btnAvancar"]').click()

                # TELA CONDIÇÕES GERAIS
                window['mensagemProgresso'].update('Ajustando condições gerais')      
                progress_bar.update_bar(4)
                page.locator('xpath=//*[@id="Titulo_Value"]').fill(titulo_requisicao)
                page.locator('xpath=//*[@id="DataEsperada_Value"]').fill(data_esperada)
                page.wait_for_timeout(500)
                page.locator('xpath=//*[@id="select2-LocalEntrega_Value-container"]').click()
                page.locator('xpath=/html/body/span/span/span[1]/input').fill(nome_filial)
                page.locator('xpath=/html/body/span/span/span[1]/input').press('Enter')
                page.locator('xpath=//*[@id="CentroCusto_Text"]').fill(centro_custo[:4])
                page.wait_for_timeout(1000)
                page.locator('xpath=//*[@id="ui-id-2"]').click()
                page.locator('xpath=//*[@id="select2-LocalFaturamento_Value-container"]').click()
                page.locator('xpath=/html/body/span/span/span[1]/input').fill(nome_filial)
                page.locator('xpath=/html/body/span/span/span[1]/input').press('Enter')
                page.locator('xpath=//*[@id="Observacao_Value"]').fill(comentario)
                page.locator('xpath=//*[@id="btnAvancar"]').click()
                
                #            TELA DETALHES DOS ITENS
                window['mensagemProgresso'].update('Ajustando detalhes dos itens')      
                progress_bar.update_bar(5)
                for i in range(len(quant)):
                    valorTotal = str(float(valorun[i]) * int(quant[i]))
                    page.locator(f'xpath=//*[@id="Itens_{i}__PrecoEstimado_Value"]').fill(valorun[i].replace(".",","))
                    page.locator(f'xpath=//*[@id="select2-Itens_{i}__CategoriaContabil_Value-container"]').click()
                    page.locator(f'xpath=//*[@id="select2-Itens_{i}__CategoriaContabil_Value-container"]').press('Enter')
                    page.wait_for_timeout(1000)
                    if cat_Pedido == 'PEDIDO REGULARIZACAO':
                        page.locator(f'xpath=//*[@id="Itens_{i}__Attributes_0__valor"]').fill(valorTotal.replace(".",","))
                    else:
                        page.locator(f'xpath=//*[@id="Itens_{i}__Attributes_1__valor"]').fill(titulo_requisicao)
                        page.locator(f'xpath=//*[@id="Itens_{i}__Attributes_0__valor"]').fill(valorTotal.replace(".",","))
                page.locator('xpath=//*[@id="btnAvancar"]').click()
                page.wait_for_timeout(1000)


                # FINALIZAR REQUISIÇÃO
                window['mensagemProgresso'].update('Finalizando a requisição')      
                progress_bar.update_bar(6)
                if cat_Pedido == 'PEDIDO REGULARIZACAO':
                    with page.expect_popup() as popup_info:
                        page.locator('xpath=//*[@id="anexoReq_link"]').click()
                        popup = popup_info.value
                        popup.wait_for_load_state()
                        popup.locator('xpath=//*[@id="fuArquivo"]').set_input_files(caminho_arquivo)
                        popup.locator('xpath=//*[@id="ctl00_conteudo_formUpload_btn_ctl00_conteudo_formUpload_btnEnviar"]').click()
                        popup.wait_for_load_state()
                        popup.close()

                page.locator('xpath=//*[@id="btnAvancar"]').click()
                requisicao = page.locator('.badge-code ')
                page.wait_for_timeout(1000)
                progress_bar.update_bar(7)
                window['mensagemProgresso'].update('########### REQUISIÇÃO FINALIZADA ###########')
                window['mensagem2'].update(titulo_requisicao)
                window['mensagem3'].update('Sua requisição é: ')
                window['mensagem4'].update(visible=True)
                window['espacamento2'].update(visible = True)
                window['mensagem4'].update(requisicao.inner_html().strip()[4:])
                if cat_Pedido == 'PEDIDO REGULARIZACAO':
                    aba_ativa[ultimaLinha] = requisicao