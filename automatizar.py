import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1


def main():
    downloadData()
    enviarEmail(*(calcularMetricas()))


def downloadData():
    pyautogui.press("winleft")
    pyautogui.write("opera")
    pyautogui.press("enter")
    pyperclip.copy(
        "https://drive.google.com/drive/u/1/folders/1fzTXM8gWiSrVBswKjKZG6QoXqY2x3sDk"
    )
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")
    while not pyautogui.locateOnScreen("drivecarregado.PNG", confidence=0.9):
        time.sleep(1)
    time.sleep(1)
    pyautogui.click(x=406, y=305, clicks=2)
    time.sleep(1)
    pyautogui.click(x=398, y=390)
    time.sleep(1)
    pyautogui.click(x=1715, y=189)
    pyautogui.click(x=1555, y=530)
    time.sleep(5)


def calcularMetricas():
    tabela = pd.read_excel(r"C:\Users\breno\Downloads\Vendas - Dez.xlsx")
    quantidadeVendida = tabela["Quantidade"].sum()
    valorVendido = tabela["Valor Final"].sum()
    return (quantidadeVendida, valorVendido)


def enviarEmail(quantidadeVendidas, valorTotal):
    pyautogui.hotkey("ctrl", "t")
    pyperclip.copy("https://mail.google.com/mail/u/1/#inbox")
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")
    while not pyautogui.locateOnScreen("gmailcarregado.PNG", confidence=0.9):
        time.sleep(1)
    time.sleep(1)
    pyautogui.click(x=125, y=201)
    time.sleep(1)
    pyautogui.write("dummyemail.gmail.com")
    pyautogui.press("tab")
    pyautogui.press("tab")
    assunto = "Relatório de Vendas do mês de dezembro"
    pyperclip.copy(assunto)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("tab")
    corpoDoEmail = f"""
    Boa noite,
        Durante o mês de dezembro nós vendemos {quantidadeVendidas:,} produtos e obtivemos um faturamento de R${valorTotal:,.2f}
        Segue em anexo a planilha de dados
        
        Att
            Breno R. Dória"""
    pyperclip.copy(corpoDoEmail)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.click(x=1426, y=1006)
    pyautogui.click(x=1624, y=948)
    time.sleep(2)
    arquivo = r"C:\Users\breno\Downloads\Vendas - Dez.xlsx"
    pyperclip.copy(arquivo)
    pyautogui.hotkey("ctrl", "v")
    pyautogui.press("enter")
    pyautogui.click(x=1309, y=1005)


if __name__ == "__main__":
    main()
