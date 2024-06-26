import openpyxl 
import pyperclip
import pyautogui
from time import sleep
# Entrar na planilha 
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_prdodutos = workbook['Produtos']

# Copiar informaçao de um capo e colar no seu campo correspondente 

for linha in sheet_prdodutos.iter_rows(min_row=2):

    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1220,320, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1220,419, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    
    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1220,589, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1220,700, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(1220,808, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1220,918, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

# Clicando botão de próximo
    pyautogui.click(1220,985,duration=0.5)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(1220,347, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    qtd_estoque = linha[7].value
    pyperclip.copy(qtd_estoque)
    pyautogui.click(1220,454, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.click(1220,562, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(1220,672, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyautogui.click(1220,780, duration=0.5)
    if tamanho == 'Pequeno':
        pyautogui.click(1260,820, duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1260,845, duration=1)
    else:
        pyautogui.click(1260,870, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(1220,884, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

# Clicando botão de próximo
    pyautogui.click(1220,969,duration=0.5)
    sleep(2)


    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1220,372, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1220,482, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    obsercoes = linha[14].value
    pyperclip.copy(obsercoes)
    pyautogui.click(1220,583, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1220,751, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1220,859, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

# Clicando botão de CONCLUIR
    pyautogui.click(1265,938,duration=1)
    sleep(2)

# Clicando botão de OK -1
    pyautogui.click(1735,240,duration=1)
    sleep(2)

# Clicando botão de OK -2
    pyautogui.click(1735,240,duration=1)
    sleep(1)

# Clicando botão de Adicionar mais produtos
    pyautogui.click(1538,658,duration=1)
    sleep(2)





