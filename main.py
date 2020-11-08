# py -m pip install -r requirements.txt 
from openpyxl import Workbook, load_workbook
import datetime
import os

FILENAME = "Daten.xlsx"
DATEFORMAT = "%Y-%m-%d %X"

def clear(): os.system('cls') #on Windows System
clear()

try:
    wb = load_workbook(FILENAME) # loads existing Workbook
    ws = wb.active # grab the active worksheet
except Exception:
    wb = Workbook() # creates new Workbook
    ws = wb.active  # grab the active worksheet
    ws.append(["Zeit", "Name des Käufers","Name des Empängers","Jahrgang des Empängers","Klasse des Empängers","Menge","Anonym"])




def Eingabe(): 

    nameBuyer = input("\nName des Käufers: ")
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
        clear()
        print("\nSaved")
    else:
        clear()
        print("\nCanceled")


if __name__ == "__main__":
    while True:
        try: 
            Eingabe()
        except KeyboardInterrupt:
            clear()
            print("\nCanceled")