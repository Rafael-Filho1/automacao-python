#ler dados da planilha
# inserir cada célula de cada linha em um campo do sistema
import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_row=2):
    pyautogui.click(1591,239,duration=0.001)
    pyautogui.write(linha[0].value)
    pyautogui.click(1538,269,duration=0.001)
    pyautogui.write(linha[1].value)
    pyautogui.click(1545,301,duration=0.001)
    pyautogui.write(str(linha[2].value))
    pyautogui.click(1624,334,duration=0.001)
    pyautogui.write(linha[3].value)
    pyautogui.click(1458,364,duration=0.001)
    pyautogui.click(777, 571, duration=0.001)
