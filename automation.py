#install alternative in VScode- openpyxl, pandas, numpy and use pip sem ! 
import pandas as pd
import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

# Passo 1: Entrar no sistema da empresa(no caso aqui é o link do drive: https://drive.google.com/seudrive)
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")

#abrir uma nova aba no chrome, caso ja esteja aberto com uma aba em outro lugar
# pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/seudrive")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(6)

# Passo 2: Navegar no sistema e encontrar a base de vendar (no caso aqui é entrar na pasta exportar)
# pyautogui.click(x=304, y=253, clicks=2)
pyautogui.press("up")
time.sleep(0.5)
pyautogui.press("up")
pyautogui.press("enter")
pyautogui.press("up")


# Passo 3: Fazer o Download da base de vendas
pyautogui.click(x=1653, y=200, clicks=1)
time.sleep(0.5)
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
time.sleep(0.7)
pyautogui.press("enter")
time.sleep(1)
pyautogui.click(x=956, y=692, clicks=1)

# Passo 4: Importar a base de vendas no Python


tabela = pd.read_excel(r"A:\project\cloud\Vendas - Dez.xlsx")
# Se estiver usando Jupyter
# display(tabela)
print(tabela)
# Passo 5: Calcular os indicadores da empresa
faturamento = tabela["Valor Final"].sum()
print(faturamento)
quantidade = tabela["Quantidade"].sum()
print(quantidade)
# Passo 6: Enviar e-mail para o solicitante com os indicadores de vendas