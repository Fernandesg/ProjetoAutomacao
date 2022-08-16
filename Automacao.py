from datetime import date
import time
import PySimpleGUI as sg
from playwright.sync_api import sync_playwright
from passwords import *
import sys

today = date.today()


layout = [
    [sg.Text('      Título da requisição'), sg.Push()],
    [sg.Push(), sg.Input(key='titulo_requisicao'), sg.Push()],
    [sg.Text('      Item'), sg.Push(), sg.Text('            Valor unitário'), sg.Push()],
    [sg.Push(), sg.Input('SRVT00190', size=(20, 1), key='item'), sg.Push(), sg.Input(key='valorun', size=(20, 1)), sg.Push()],
    [sg.Text('      Quantidade'), sg.Push(), sg.Text('    Data esperada'), sg.Push(), ],
    [sg.Push(), sg.Input(key='quant', size=(20, 1)), sg.Push(), sg.Input(today.strftime("%d/%m/%Y"), key='data_esperada', size=(20, 1)), sg.Push(), ],
    [sg.Text('      Selecione a filial'), sg.Push()],
    [sg.Push(), sg.Combo(['75 - VERO SANTO ANTONIO DA PATRULHA','86 - VERO SANTO ANTONIO DA PATRULHA II'], key='filial', size=(43,1)), sg.Push()],
    [sg.Push(), sg.Button('Criar requisição', size=(17, 1)), sg.Push(), sg.Button('Cancelar', size=(17, 1)), sg.Push()],
    [sg.Push(), sg.ProgressBar(max_value=8, orientation='h', size=(20, 20), key='progress', visible=False),sg.Push(), sg.Text('', visible=False, key='espacamento',size=(8,1))],
    [sg.Push(), sg.Text('', key='mensagemProgresso'), sg.Push(),],
    [sg.Push(), sg.Text('', key='mensagem2'), sg.Push(),],
    [sg.Push(), sg.Text('', key='mensagem3'), sg.Push(), sg.Input('', key='mensagem4',size=(18,1),disabled=True, visible=False), sg.Push(), sg.Text('', visible=False, key='espacamento2',size=(8,1))]
    ]

window = sg.Window('Abertura de requisições', size=(400, 400), layout = layout)
progress_bar = window['progress']

while True:
    
    event, values = window.read()
    print(event, values)
    match(event):
        case None:
            break
        case 'Cancelar':
            break
        case 'Criar requisição': 
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
                browser = p.chromium.launch()
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
                frame.locator('//*[@id="BOrgs_1__BorgDescription"]').fill(filial)
                page.wait_for_timeout(1000)
                frame.locator('//*[@id="BOrgs_1__BorgDescription"]').press('Tab')
                page.wait_for_timeout(1000)
                frame.locator('//*[@id="btnSave"]').click()

                # SELECIONA ITENS E QUANTIDADES
                window['mensagemProgresso'].update('Adicionando itens e quantidades')      
                progress_bar.update_bar(3)
                page.locator('xpath=//*[@id="Valor"]').fill(item)
                page.locator('xpath=//*[@id="btnSearchSimple"]').click()
                page.locator('xpath=//*[@id="149419566"]').click()
                page.wait_for_timeout(2000)
                page.locator('//*[@id="input_qtde_149419566"]').fill(quant)
                page.wait_for_timeout(1000)
                page.locator('//*[@id="input_qtde_149419566"]').press('Tab')
                page.wait_for_timeout(2000)
                page.locator('xpath=//*[@id="btnAvancar"]').click()

                # TELA CONDIÇÕES GERAIS
                window['mensagemProgresso'].update('Ajustando condições gerais')      
                progress_bar.update_bar(4)
                page.locator('xpath=//*[@id="Titulo_Value"]').fill(titulo_requisicao)
                page.locator('xpath=//*[@id="select2-LocalEntrega_Value-container"]').click()
                page.locator('xpath=/html/body/span/span/span[1]/input').fill(filial[5:])
                page.locator('xpath=/html/body/span/span/span[1]/input').press('Enter')
                page.locator('xpath=//*[@id="CentroCusto_Text"]').fill('0312')
                page.wait_for_timeout(1000)
                page.locator('xpath=//*[@id="ui-id-2"]').click()
                page.locator('xpath=//*[@id="select2-LocalFaturamento_Value-container"]').click()
                page.locator('xpath=/html/body/span/span/span[1]/input').fill(filial[5:])
                page.locator('xpath=/html/body/span/span/span[1]/input').press('Enter')
                page.locator('xpath=//*[@id="btnAvancar"]').click()
                
                #            TELA DETALHES DOS ITENS
                window['mensagemProgresso'].update('Ajustando detalhes dos itens')      
                progress_bar.update_bar(5)
                page.locator('xpath=//*[@id="Itens_0__PrecoEstimado_Value"]').fill(valorun.replace(".",","))
                page.locator('xpath=//*[@id="select2-Itens_0__CategoriaContabil_Value-container"]').click()
                page.locator('xpath=//*[@id="select2-Itens_0__CategoriaContabil_Value-container"]').press('Enter')
                page.wait_for_timeout(1000)
                page.locator('xpath=//*[@id="Itens_0__Attributes_0__valor"]').fill(valorTotal.replace(".",","))
                page.locator('xpath=//*[@id="Itens_0__Attributes_1__valor"]').fill(titulo_requisicao)
                page.locator('xpath=//*[@id="btnAvancar"]').click()

                # FINALIZAR REQUISIÇÃO
                window['mensagemProgresso'].update('Finalizando a requisição')      
                progress_bar.update_bar(6)
                page.locator('xpath=//*[@id="btnAvancar"]').click()
                requisicao = page.locator('.badge-code ')
                page.wait_for_timeout(1000)
                window['mensagemProgresso'].update('########### REQUISIÇÃO FINALIZADA ###########')
                window['mensagem2'].update(titulo_requisicao)
                window['mensagem3'].update('Sua requisição é: ')
                window['mensagem4'].update(visible=True)
                window['espacamento2'].update(visible = True)
                window['mensagem4'].update(requisicao.inner_html().replace(' ',''))

