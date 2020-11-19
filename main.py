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
    ws.append(["Zeit", "Name des Käufers","Name des Empängers","Jahrgang des Empängers","Klasse des Empängers","Vollmilch Menge", "Zartbitter Menge", "Anonym"])


def Eingabe():

    nameBuyer = input("\nName des Käufers: ")

    entrys = int(input("\nAnzahl der Käufe: "))
    for x in range (0, entrys):
        nameReciever = input("Name des Empängers "+ str(x+1) +" :")
        jahrgang = input("Jahrgang des Empängers "+ str(x+1) +" :")
        grade = input("Klasse des Empängers "+ str(x+1) +" :")

        vollmilch = input("Vollmilch groß Menge: ")
        zartbitter = input("Zartbitter klein Menge: ")

        anonym = input("Anonym: (J/N): ").lower().startswith("j")

        row = [datetime.datetime.now().strftime(DATEFORMAT), nameBuyer, nameReciever, jahrgang, grade, vollmilch, zartbitter, anonym]

        if input("Speichern: [J]a / [N]ein : ").lower().startswith("j"):
            ws.append(row)
            wb.save(FILENAME)
            if x == (entrys-1):
                clear()
                print("\nSaved")
            else:
                print("\nNext Entry")
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