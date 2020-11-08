# py -m pip install -r requirements.txt 
from openpyxl import Workbook, load_workbook
import datetime
import os

FILENAME = "Daten.xlsx"
DATEFORMAT = "%Y-%m-%d %X"

def clear(): os.system('cls') #on Windows System
clear()

try:
    wb = load_workbook(FILENAME)
except Exception:
    wb = Workbook()

# grab the active worksheet
ws = wb.active

def Eingabe(): 

    nameBuyer = input("Name des Käufers: ")
    nameReciever = input("Name des Empängers: ")
    jahrgang = input("Jahrgang des Empängers: ")
    grade = input("Klasse des Empängers: ")
    menge = input("Menge: ")
    anonym = input("Anonym: (J/N): ").lower().startswith("j")

    row = [datetime.datetime.now().strftime(DATEFORMAT), nameBuyer, nameReciever, jahrgang, grade, menge, anonym]

    
    print("\nRichtig?", row)
    if input("[J]a / [N]ein : ").lower().startswith("j"):
        ws.append(row)
        wb.save(FILENAME)


if __name__ == "__main__":
    try: 
        while True:
            Eingabe()
    except KeyboardInterrupt:
        pass