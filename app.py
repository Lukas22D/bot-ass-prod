import openpyxl
import pyperclip
import pyautogui
from time import sleep

#Entra no arquivo excel
wb = openpyxl.load_workbook('produtos_ficticios.xlsx')
#Seleciona a aba Produtos
sheet = wb['Produtos']
#Pega o valor da celulas e armazena em variaveis
for line in sheet.iter_rows(min_row=2, max_row= 5):
    nome_produto = line[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1367,371, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    descricao = line[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1363,459, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    categoria = line[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1365,586, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    codigo_produto = line[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1385,676, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    peso = line[4].value
    pyperclip.copy(peso)
    pyautogui.click(1363,760, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    dimension = line[5].value
    pyperclip.copy(dimension)
    pyautogui.click(1368,853, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.click(1389,902, duration=1)
    sleep(3)
    preco = line[6].value
    pyperclip.copy(preco)
    pyautogui.click(1364,395, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    quantidade_em_estoque = line[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(1365,477, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    data_de_validade = line[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(1369,557, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    cor = line[9].value
    pyperclip.copy(cor)
    pyautogui.click(1362,652, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    tamanho = line[10].value
    pyautogui.click(1372,741, duration=1)
    if(tamanho == "Grande"):
        pyautogui.click(1363,805, duration=1)
    elif(tamanho == "Médio"):
        pyautogui.click(1385,787, duration=1)
    elif(tamanho == "Pequeno"):
        pyautogui.click(1387,767, duration=1)
    material = line[11].value
    pyperclip.copy(material)
    pyautogui.click(1364,823, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.click(1373,885, duration=1)
    sleep(3)
    fabricante = line[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(1362,412, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    pais_origem = line[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1370,501, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    observacoes = line[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(1364,578, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    codigo_de_barras = line[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(1366,716, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    localizacao_armazem = line[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1364,809, duration=1)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.click(1385,870, duration=1)
    pyautogui.click(1786,197, duration=1)
    pyautogui.click(1578,634, duration=1)
    sleep(3)



