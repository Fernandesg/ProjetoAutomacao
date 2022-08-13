from playwright.sync_api import sync_playwright
import time
from passwords import *

with sync_playwright() as p:
    item = 'SRVT00190'
    titulo_requisicao = 'AJUSTE IGPM FATURA NF 191382'
    valorun = '7.42'
    quant = '11'
    data_esperada = '12/08/2022'
    filialSAPII =  "86 - VERO SANTO ANTONIO DA PATRULHA II"
    filialSAP = "75 - VERO SANTO ANTONIO DA PATRULHA"
    filial = filialSAPII

    valorTotal = str(float(valorun) * int(quant))
    usuario_me = usuario_me
    senha_me = senha_me
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