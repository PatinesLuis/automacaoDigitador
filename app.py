import openpyxl
import pyperclip
import pyautogui

#entrar na planilha
planilhaDados = openpyxl.load_workbook('planilha.xlsx')

#pegar a aba da planilha
abaPlanilha = planilhaDados['planilha']

#copiar informação de um campo/celula
for linha in abaPlanilha.iter_rows(min_row=2):
   promotor = linha[0].value
   pyperclip.copy(promotor)
   pyautogui.click(1762,271, duration=1)
   pyautogui.hotkey('ctrl','v')

   cliente = linha[1].value
   pyperclip.copy(cliente)
   pyautogui.click(1740,410,duration=1)
   pyautogui.hotkey('ctrl','v')

   telefone = linha[2].value
   pyperclip.copy(telefone)
   pyautogui.click(1741,563,duration=1)
   pyautogui.hotkey('ctrl','v')

   valor = linha[3].value
   pyperclip.copy(valor)
   pyautogui.click(1748,517,duration=1)
   pyautogui.hotkey('ctrl','v')
