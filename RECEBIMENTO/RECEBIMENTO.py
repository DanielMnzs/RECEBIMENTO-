import pyautogui
import time
import PIL
import pyscreeze
import pandas as pd 
from openpyxl import load_workbook


caminho_arquivo = r"//depo0903/gpa$/PAC/Daniel Menezes/Python/RECEBIMENTO/C OR R.xlsx"
ler = pd.read_excel(caminho_arquivo)
tabela = pd.DataFrame(ler)
print(tabela)
workbook = load_workbook(caminho_arquivo)
sheet = workbook.active
x = 0
y = 0
z = 2
 
pyautogui.write("DP01")
pyautogui.press("enter")
time.sleep(1)
pyautogui.write("1942")
pyautogui.write("CO")
pyautogui.press("enter")

encontrou = False

while True and x < len(tabela):
    pedido = tabela.iloc [x , y + 9]
    print(pedido)
    time.sleep(0.2)
    pyautogui.write("C")
    pyautogui.press("enter")
    pyautogui.write(str(pedido),interval=0.111)
    pyautogui.press("enter")
    pyautogui.moveTo(17, 319)
    pyautogui.dragTo(198, 317, duration= 1, button='left')
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.2)
    pyautogui.hotkey("F3")
    time.sleep(0.5)
    pyautogui.write("a")
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.2)
    pyautogui.write("x")
    for i in range (2):
        pyautogui.hotkey("F4")
    for i in range (5):
        try:
            img = pyautogui.locateOnScreen(r"//depo0903/gpa$/PAC/Daniel Menezes/Python/RECEBIMENTO/ARMAZENAGEM.png",confidence= 0.999)
            if img:
                sheet[f"K{z}"] = "CONGELADO"
                print("CONGELADO")
                time.sleep(0.5)
                for i in range (4): 
                    pyautogui.hotkey("F3")
                    time.sleep(1)
                encontrou = True


        except:
            print("IMAGEM NÃO ENCONTRADA...")
            time.sleep(0.5)
            print(f"VERIFICANDO...{i} de 5 TENTATIVAS")
    
    if not encontrou:
        for i in range (5):
            try:
                A = pyautogui.locateAllOnScreen(r"//depo0903/gpa$/PAC/Daniel Menezes/Python/RECEBIMENTO/A.png",confidence= 0.999)
                if A:
                    sheet[f"K{z}"] = "aaaaaa"
                    print ("RESFRIADO")
                    print("IMAGEM ENCONTRADA!")
                    for i in range (4):
                        pyautogui.hotkey("F3")
                        time.sleep(1)       
                break
            except:
                print("IMAGEM NÃO ENCONTRADA...")
                time.sleep(0.5)
                print(f"VERIFICANDO...{i} de 5 TENTATIVAS")


    workbook.save(caminho_arquivo)
    
    x = x + 1
    z = z + 1
