import openpyxl #utilizado para importar arquivos .xlsx (excel) 
import pyperclip #biblioteca de caracteres do python que mantém acentos para não ter problemas de formatação
import pyautogui  #utilizado para automatizar tarefas usando mouse e teclas de atalho
from time import sleep


# Acessar a planilha
workbook = openpyxl.load_workbook('produtos.xlsx')
sheet_produtos = workbook['Produtos']

# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=3):
    
# ---------- Automatização da página 1 ----------
    
    nome_produto = linha[1].value # Guardar a informação da tabela (0 em python = 1 / 1 = 2 / 2 = 3 e assim por diante)
    pyperclip.copy(nome_produto) #copia a informação da variável conforme endereço de linha
    pyautogui.click(1400,349,duration=1) #clica no campo que deve ser preenchido com o nome do produto
    pyautogui.hotkey("ctrl", "v")  #cola a informação da variável que foi copiada de LINHA no respectivo campo 

    descricao = linha[2].value
    pyperclip.copy(descricao)
    pyautogui.click(1392,437,duration=1)
    pyautogui.hotkey("ctrl", "v")

    categoria = linha[3].value
    pyperclip.copy(categoria)
    pyautogui.click(1403,569,duration=1)
    pyautogui.hotkey("ctrl", "v")

    codigo = linha[4].value
    pyperclip.copy(codigo)
    pyautogui.click(1390,654,duration=1)
    pyautogui.hotkey("ctrl", "v")

    peso = linha[5].value
    pyperclip.copy(peso)
    pyautogui.click(1416,740,duration=1)
    pyautogui.hotkey("ctrl", "v")

    dimensoes = linha[6].value
    pyperclip.copy(dimensoes)
    pyautogui.click(1392,826,duration=1)
    pyautogui.hotkey("ctrl", "v")

    pyautogui.click(1400,883,duration=1)
    sleep(3)

# ---------- Automatização da página 2 ----------
    
    preco = linha[7].value
    pyperclip.copy(preco)
    pyautogui.click(1391,378,duration=1)
    pyautogui.hotkey("ctrl", "v")

    estoque = linha[8].value
    pyperclip.copy(estoque)
    pyautogui.click(1401,460,duration=1)
    pyautogui.hotkey("ctrl", "v")

    validade = linha[9].value
    pyperclip.copy(validade)
    pyautogui.click(1397,544,duration=1)
    pyautogui.hotkey("ctrl", "v")

    cor = linha[10].value 
    pyperclip.copy(validade)
    pyautogui.click(1396,631,duration=1)
    pyautogui.hotkey("ctrl", "v")

 
    tamanho = linha[11].value
    pyautogui.click(1400,712,duration=1) #irá clicar no menu pulldown

    if tamanho == 'Pequeno': #se for pequeno na planilha, irá clicar na posição pequeno
      pyautogui.click(1405,725,duration=1)
    elif tamanho == 'Médio': #se for médio na planilha, irá clicar na posição médio
       pyautogui.click(1398,771,duration=1)
    else: #se for grande na planilha, irá clicar na posição grande
      pyautogui.click(1407,795,duration=1)
    
    material = linha[12].value
    pyperclip.copy(validade)
    pyautogui.click(1399,805,duration=1)
    pyautogui.hotkey("ctrl", "v")

    pyautogui.click(1405,861,duration=1)
    sleep(3)

# ---------- Automatização da página 3 ----------

    fabricante = linha[13].value
    pyperclip.copy(fabricante)
    pyautogui.click(1387,393,duration=1)
    pyautogui.hotkey("ctrl", "v")

    pais_origem = linha[14].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1390,477,duration=1)
    pyautogui.hotkey("ctrl", "v")

    observacoes = linha[15].value    
    pyperclip.copy(observacoes)
    pyautogui.click(1392,566,duration=1)
    pyautogui.hotkey("ctrl", "v")
    
    codigo_barras = linha[16].value
    pyperclip.copy(codigo_barras)
    pyautogui.click(1400,694,duration=1)
    pyautogui.hotkey("ctrl", "v")

    local_armazem = linha[17].value
    pyperclip.copy(local_armazem)
    pyautogui.click(1396,782,duration=1)
    pyautogui.hotkey("ctrl", "v")

    pyautogui.click(1403,845,duration=1) #concluir na etapa 3

    pyautogui.click(1795,190,duration=1) #clicar em OK no pop-up de confirmação de cadastro de produto
    sleep(3)

# ---------- Automatização da página 4 ----------
    
    pyautogui.click(1580,621,duration=1) #clicar em ADICIONAR MAIS UM para reiniciar o processo de cadastro de produto
    sleep(3)

# Repetir esses passos para outros campos até preencher toda os campos da página
# Clicar em próxima
# Repetir os mesmos passos e ir para a próxima página (página 3)
# Repetir os mesmos passos e finalizar o cadastro daquele produto clicando em concluir
# Clicar em ok para finalizar o processo
# Clicar no ok mais uma vez na mensagem de confirmação de salvamento de banco de dados
# Clicar em "adicionar mais um" e repetir o processo até finalizar toda a planilha" 