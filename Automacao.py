from playwright.sync_api import sync_playwright
import time
from passwords import *

with sync_playwright() as p:
    item = int(input('Selecione o item: \n' + '1: SRVT00190\n' + '2: Inserir outro\n' + 'Selecione a opção desejada: '))
    if item == 1:
        item = 'SRVT00190'
    elif item == 2:
        item = input('Insira o código do item: ').upper()
    else:
        print('|-----OPÇÃO INVÁLIDA-----|')
        exit()

    titulo_requisicao = input('Insira o título da requisição: ')
    valorun = str(input('Insira o valor unitário: ')).replace(',','.')
    quant = str(input('Insira a quantidade: '))
    data_esperada = input('Insira a data: Ex: DD/MM/AAAA: ')

    filial = str(input('FILIAIS:\n' + '1: 75 - VERO SANTO ANTONIO DA PATRULHA\n' + '2: 86 - VERO SANTO ANTONIO DA PATRULHA II\n'+ 
    'Seleciona a opção: '))

    if filial == '1':
        filial = '75 - VERO SANTO ANTONIO DA PATRULHA'
    elif filial == '2':
        filial = '86 - VERO SANTO ANTONIO DA PATRULHA II'
    else:
        print('|-----OPÇÃO INVÁLIDA-----|')
        exit()

    print('-=' * 15 + ' RESUMO DA REQUISIÇÃO ' + '=-' * 15)
    print('Item: ' + item)
    print('Título da requisição: ' + titulo_requisicao)
    print('Valor un: R$' + valorun.replace('.',','))
    print('Quantidade: ' + quant)
    print('Data: ' + data_esperada)
    print('Filial: ' + filial)
    continuar = input('Deseja continuar? [S]/[N]: ').upper()
    if continuar == 'N':
        exit()


    valorTotal = str(float(valorun) * int(quant))
    browser = p.chromium.launch()
    page = browser.new_page()
    page.goto(site)

    # LOGIN ME

    page.locator('xpath=//*[@id="LoginName"]').fill(usuario_me)
    page.locator('xpath=//*[@id="RAWSenha"]').fill(senha_me)
    page.locator('xpath=//*[@id="SubmitAuth"]').click()

    # CONFIGURAÇÃO DA REQUISIÇÃO

    page.locator('xpath=//*[@id="__layout"]/div/main/div/div/div[2]/div/div[1]/section/div/div/div/button').click()
    page.locator('xpath=//*[@id="__layout"]/div/main/div/div/div[2]/div/div[1]/section/div[1]/div[2]/ul/li[1]/a').click()
    page.locator('xpath=//*[@id="__layout"]/div/main/div/div/div[2]/div/div[1]/section/div[1]/div[2]/ul[2]/li[1]/a').click()
    frame = page.frame_locator('#PopUpConfiguration-if')
    frame.locator('//*[@id="BOrgs_1__BorgDescription"]').fill(filial)
    time.sleep(1)
    frame.locator('//*[@id="BOrgs_1__BorgDescription"]').press('Tab')
    time.sleep(1)
    frame.locator('//*[@id="btnSave"]').click()

    # SELECIONA ITENS E QUANTIDADES

    page.locator('xpath=//*[@id="Valor"]').fill(item)
    page.locator('xpath=//*[@id="btnSearchSimple"]').click()
    page.locator('xpath=//*[@id="149419566"]').click()
    time.sleep(2)
    page.locator('//*[@id="input_qtde_149419566"]').fill(quant)
    time.sleep(2)
    page.locator('xpath=//*[@id="btnAvancar"]').click()

     # TELA CONDIÇÕES GERAIS

    page.locator('xpath=//*[@id="Titulo_Value"]').fill(titulo_requisicao)
    page.locator('xpath=//*[@id="select2-LocalEntrega_Value-container"]').click()
    page.locator('xpath=/html/body/span/span/span[1]/input').fill(filial[5:])
    page.locator('xpath=/html/body/span/span/span[1]/input').press('Enter')
    page.locator('xpath=//*[@id="CentroCusto_Text"]').fill('0312')
    time.sleep(1)
    page.locator('xpath=//*[@id="ui-id-2"]').click()
    page.locator('xpath=//*[@id="select2-LocalFaturamento_Value-container"]').click()
    page.locator('xpath=/html/body/span/span/span[1]/input').fill(filial[5:])
    page.locator('xpath=/html/body/span/span/span[1]/input').press('Enter')
    page.locator('xpath=//*[@id="btnAvancar"]').click()
    
    #            TELA DETALHES DOS ITENS

    page.locator('xpath=//*[@id="Itens_0__PrecoEstimado_Value"]').fill(valorun.replace(".",","))
    page.locator('xpath=//*[@id="select2-Itens_0__CategoriaContabil_Value-container"]').click()
    page.locator('xpath=//*[@id="select2-Itens_0__CategoriaContabil_Value-container"]').press('Enter')
    time.sleep(1)
    page.locator('xpath=//*[@id="Itens_0__Attributes_0__valor"]').fill(valorTotal.replace(".",","))
    page.locator('xpath=//*[@id="Itens_0__Attributes_1__valor"]').fill(titulo_requisicao)
    page.locator('xpath=//*[@id="btnAvancar"]').click()

    # FINALIZAR REQUISIÇÃO

    page.locator('xpath=//*[@id="btnAvancar"]').click()
    requisicao = page.locator('.badge-code ')
    print('########### REQUISIÇÃO FINALIZADA ###########\n')
    print(titulo_requisicao)
    print("\nSua requisição é: " + requisicao.inner_html().replace(' ',''))