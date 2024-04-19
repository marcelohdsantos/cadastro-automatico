import openpyxl
import pyperclip
import pyautogui
from time import sleep

# abrir a planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# copiar informações de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(3324, 420, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(3277, 508, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(3227, 645, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    codigo_do_produto = linha[3].value
    pyperclip.copy(codigo_do_produto)
    pyautogui.click(3227, 730, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(3248, 818, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(3266, 901, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    # clicar no botão "Próximo"
    pyautogui.click(3246, 965, duration=0.5)
    sleep(1)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(3300, 449, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    quantidade_estoque = linha[7].value
    pyperclip.copy(quantidade_estoque)
    pyautogui.click(3243, 531, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    data_validade = linha[8].value
    pyperclip.copy(data_validade)
    pyautogui.click(3227, 622, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(3253, 709, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    # clicar no botão "Tamanho"
    pyautogui.click(3279, 791, duration=0.5)

    tamanho = linha[10].value
    if tamanho == "Pequeno":
        pyautogui.click(3304, 824, duration=0.5)
    elif tamanho == "Médio":
        pyautogui.click(3276, 844, duration=0.5)
    else:
        pyautogui.click(3265, 866, duration=0.5)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(3308, 878, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    # clicar no botão "Próximo"
    pyautogui.click(3240, 933, duration=0.5)
    sleep(1)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(3270, 468, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value,
    str_pais_origem = str(pais_origem)
    pyperclip.copy(str_pais_origem)
    pyautogui.click(3254, 551, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(3256, 637, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    codigo_barras = linha[15].value
    pyperclip.copy(codigo_barras)
    pyautogui.click(3233, 777, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    localizao_estoque = linha[16].value
    pyperclip.copy(localizao_estoque)
    pyautogui.click(3249, 859, duration=0.5)
    pyautogui.hotkey('ctrl', 'v')

    # clicar no botão "Concluir"
    pyautogui.click(3237, 925, duration=0.5)
    sleep(1)

    # popup 1 "produto salvo no banco de dados: "
    pyautogui.click(3636, 233, duration=0.5)
    sleep(1)

    # Produto cadastrado com sucesso
    pyautogui.click(3460, 688, duration=0.5)
    sleep(1)
