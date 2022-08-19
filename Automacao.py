from datetime import date
import PySimpleGUI as sg
from playwright.sync_api import sync_playwright
import os

passwords = open('C:/DEV/PYTHON/ProjetoAutomacao/credenciais.txt', 'r')
login = []

for linhas in passwords:
    linhas = linhas.strip()
    login.append(linhas)
usuario_me = login[0][14:-1]
senha_me = login[1][12:-1]
site = login[2][8:-1]

filiais_caminho = open('C:/DEV/PYTHON/ProjetoAutomacao/filiais.txt', 'r')
filiais = []

for linhas in filiais_caminho:
    linhas = linhas.strip()
    filiais.append(linhas)

cod_caminho = open('C:/DEV/PYTHON/ProjetoAutomacao/codigos.txt', 'r')
codigos = []

for linhas in cod_caminho:
    linhas = linhas.strip()
    codigos.append(linhas)

cc_caminho = open('C:/DEV/PYTHON/ProjetoAutomacao/centrocustos.txt', 'r')
centroCustos = []

for linhas in cc_caminho:
    linhas = linhas.strip()
    centroCustos.append(linhas)

categorias_caminho = open('C:/DEV/PYTHON/ProjetoAutomacao/categorias.txt', 'r')
categorias = []

for linhas in categorias_caminho:
    linhas = linhas.strip()
    categorias.append(linhas)

today = date.today()

menu_def=[['Arquivos', ['Itens', 'Categorias', 'Centro de custos', 'Filiais','---','Credenciais ME']]]
layout = [
    [sg.Menu(menu_def, pad=(10,10))],
    [sg.Text('      Título da requisição'), sg.Push()],
    [sg.Push(), sg.Input(key='titulo_requisicao'), sg.Push()],
    [sg.Text('      Item'), sg.Push(), sg.Text('            Valor unitário'), sg.Push()],
    [sg.Push(), sg.Combo(codigos, size=(18, 1), key='item', default_value='SRVT00190'), sg.Push(), sg.Input(key='valorun', size=(20, 1)), sg.Push()],
    [sg.Text('      Quantidade'), sg.Push(), sg.Text('    Data esperada'), sg.Push(), ],
    [sg.Push(), sg.Input(key='quant', size=(20, 1)), sg.Push(), sg.Input(today.strftime("%d/%m/%Y"), key='data_esperada', size=(20, 1)), sg.Push(), ],
    [sg.Text('      Selecione a categoria'),sg.Text('                 Centro de custo'), sg.Push()],
    [sg.Push() ,sg.Combo(categorias, key='catPedido', size=(23,1), readonly=True, default_value='PEDIDO COMPRA PADRAO'),sg.Combo(centroCustos, size=(16, 1), key='centrocusto', default_value='0312 - Supply Chain Estoque', readonly=True), sg.Push()],
    [sg.Text('      Selecione a filial'), sg.Push()],
    [sg.Push(), sg.Combo(filiais, key='filial', size=(43,1), readonly=True, default_value='86 - VERO SANTO ANTONIO DA PATRULHA II'), sg.Push()],
    [sg.Text('      Comentario:')],
    [sg.Push(), sg.Input(key='comentario'),sg.Push()],
    [sg.Text("      Selecionar arquivo: (Apenas pedido de regularização)", key='txtSelecionaArquivo')], 
    [sg.Push(),sg.Input(size=(35,1)), sg.FileBrowse('Procurar',key="caminhoArquivo"), sg.Push()],
    [sg.Push(), sg.Button('Criar requisição', size=(17, 1)), sg.Push(), sg.Button('Cancelar', size=(17, 1)), sg.Push()],
    [sg.Push(), sg.ProgressBar(max_value=7, orientation='h', size=(20, 20), key='progress', visible=False),sg.Push(), sg.Text('', visible=False, key='espacamento',size=(8,1))],
    [sg.Push(), sg.Text('', key='mensagemProgresso'), sg.Push(),],
    [sg.Push(), sg.Text('', key='mensagem2'), sg.Push(),],
    [sg.Push(), sg.Text('', key='mensagem3'), sg.Push(), sg.Input('', key='mensagem4',size=(18,1),disabled=True, visible=False), sg.Push(), sg.Text('', visible=False, key='espacamento2',size=(8,1))]
    ]

window = sg.Window('Abertura de requisições', size=(400, 600), layout = layout)
progress_bar = window['progress']

while True:
    event, values = window.read()

    match(event):
        case 'Itens':
            os.system('C:/DEV/PYTHON/ProjetoAutomacao/codigos.txt')
        case 'Categorias':
            os.system('C:/DEV/PYTHON/ProjetoAutomacao/categorias.txt')
        case 'Centro de custos':
            os.system('C:/DEV/PYTHON/ProjetoAutomacao/centrocustos.txt')
        case 'Filiais':
            os.system('C:/DEV/PYTHON/ProjetoAutomacao/filiais.txt')
        case 'Credenciais ME':
            os.system('C:/DEV/PYTHON/ProjetoAutomacao/passwords.txt')

        case None:
            break
        case 'Cancelar':
            break
        case 'Criar requisição': 
            window['mensagem2'].update('')
            window['mensagem3'].update('')
            window['mensagem4'].update(visible=False)
            window['espacamento2'].update(visible = False)
            window['mensagem4'].update('')

            comentario = values['comentario']
            caminho_arquivo = values["caminhoArquivo"]
            centro_custo = values['centrocusto']
            cat_Pedido = values['catPedido']
            titulo_requisicao = values['titulo_requisicao']
            item = values['item']
            valorun = str(values['valorun']).replace(',','.')
            quant = str(values['quant'])
            data_esperada = values['data_esperada']
            filial = values['filial']
            progress_bar.update(visible = True)
            window['espacamento'].update(visible = True)

            with sync_playwright() as p:
                valorTotal = str(float(valorun) * int(quant))
                browser = p.chromium.launch(channel="chrome",headless=False)
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
                page.wait_for_timeout(1000)
                page.locator('xpath=//*[@id="Valor"]').fill(item)
                page.locator('xpath=//*[@id="btnSearchSimple"]').click()
                page.locator('.icon-shopping-cart').click()
                page.wait_for_timeout(1000)
                page.locator('[name=\"item-quantity\"]').fill(quant)
                page.wait_for_timeout(1000)
                page.locator('[name=\"item-quantity\"]').press('Tab')
                page.wait_for_timeout(1000)
                page.locator('xpath=//*[@id="btnAvancar"]').click()

                # TELA CONDIÇÕES GERAIS
                window['mensagemProgresso'].update('Ajustando condições gerais')      
                progress_bar.update_bar(4)
                page.locator('xpath=//*[@id="Titulo_Value"]').fill(titulo_requisicao)
                page.locator('xpath=//*[@id="DataEsperada_Value"]').fill(data_esperada)
                page.wait_for_timeout(500)
                page.locator('xpath=//*[@id="select2-LocalEntrega_Value-container"]').click()
                page.locator('xpath=/html/body/span/span/span[1]/input').fill(filial[5:])
                page.locator('xpath=/html/body/span/span/span[1]/input').press('Enter')
                page.locator('xpath=//*[@id="CentroCusto_Text"]').fill(centro_custo[:4])
                page.wait_for_timeout(1000)
                page.locator('xpath=//*[@id="ui-id-2"]').click()
                page.locator('xpath=//*[@id="select2-LocalFaturamento_Value-container"]').click()
                page.locator('xpath=/html/body/span/span/span[1]/input').fill(filial[5:])
                page.locator('xpath=/html/body/span/span/span[1]/input').press('Enter')
                page.locator('xpath=//*[@id="Observacao_Value"]').fill(comentario)
                page.locator('xpath=//*[@id="btnAvancar"]').click()
                
                #            TELA DETALHES DOS ITENS
                window['mensagemProgresso'].update('Ajustando detalhes dos itens')      
                progress_bar.update_bar(5)
                page.locator('xpath=//*[@id="Itens_0__PrecoEstimado_Value"]').fill(valorun.replace(".",","))
                page.locator('xpath=//*[@id="select2-Itens_0__CategoriaContabil_Value-container"]').click()
                page.locator('xpath=//*[@id="select2-Itens_0__CategoriaContabil_Value-container"]').press('Enter')
                page.wait_for_timeout(1000)
                if cat_Pedido == 'PEDIDO REGULARIZAÇÃO':
                    page.locator('xpath=//*[@id="Itens_0__Attributes_0__valor"]').fill(valorTotal.replace(".",","))
                else:
                    page.locator('xpath=//*[@id="Itens_0__Attributes_1__valor"]').fill(titulo_requisicao)
                    page.locator('xpath=//*[@id="Itens_0__Attributes_0__valor"]').fill(valorTotal.replace(".",","))
                page.locator('xpath=//*[@id="btnAvancar"]').click()
                page.wait_for_timeout(1000)


                # FINALIZAR REQUISIÇÃO
                window['mensagemProgresso'].update('Finalizando a requisição')      
                progress_bar.update_bar(6)
                if cat_Pedido == 'PEDIDO REGULARIZAÇÃO':
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
                window['mensagem4'].update(requisicao.inner_html().replace(' ',''))