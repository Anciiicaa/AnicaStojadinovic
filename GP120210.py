import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import pyodbc
import qrcode
import os
import io
import win32print
import win32ui
from PIL import Image, ImageWin
#import datetime

predlozi_za_part_number = []
# Funkcija za ažuriranje predloga za "part number" zavisno od izbora u drugim poljima
def azuriraj_part_number():
    global predlozi_za_part_number
    project = project_var.get()
    # Ovde implementirajte logiku za generisanje predloga za "part number" na osnovu stanice i project vrednosti
    # Na primer, koristite if-elif niz za različite kombinacije stanica i project-a
    if project == "PHEV":
        predlozi_za_part_number = ["4873027", "4873216", "4873484", "4872991", "4873400", "4873206", "8129738","4873479", "4873486", "4873035", "4873364", "4873207", "4873465",
                                   "4873493", "4873441", "4872989", "4873385", "4921485", "4873197", "4873224", "4873992", "4873403", "4873000", "4873482", "4873451", "4872994", 
                                   "4873020", "4873022", "4872996", "4873347", "4873014", "4873334", "4873261", "4873015", "4873241", "4873325", "4873011", "4873555", "4873557"]
    elif project == "G60":
        predlozi_za_part_number = ["4796915", "4796922", "4796953", "4796954", "4796979", "4797358", "4797358", "4797016", "4797036", "4797042", "4797044", "4797136", "4797177",
                                   "4797229", "4797234", "4797407", "4797412", "4797911", "4798125", "4798253", "Y0273015", "4922805", "5021654", "8012507", "4797098", 
                                   "8038771", "4797139", "8038774", "4797143", "8039756", "8129679", "8129680", "8129684", "8130841", "4922806", "8031329", "4797097", "8031328"]
    elif project == "MMA":
        predlozi_za_part_number = ["8119163", "8119164 ", "8119137", "8119150", "8119114", "8119139", "8119240", "8119141", "8119241", "8196766", "8119147", "8119156", "8119158",
                                   "8119160", "8119193", "8119223", "8119208", "8119210", "8119214", "8119218", "8119249", "8119237", "Y0408016", "8012507", "8119144", 
                                   "8119243", "8119146", "8119143", "8119245", "8119187", "8119189"]
    elif project == "EVA2":
        predlozi_za_part_number = ["EK22100302-211"]
    
    elif project == "G61":
        predlozi_za_part_number = ["4798396", "8012009", "4794690", "4798189", "4798731", "8003363", "4797658", "4797657", "4798452", "4873743", "YS8001686", "YS4873055", "YS4873303",
                                   "YS8053625", "YS4797860", "YS8012011", "8012032", "4797891", "4875135", "8049771", "8049773", "4798115", "8058512", "8049514", "8049516", 
                                   "8042661", "8042662", "8049359", "8042664", "4797042", "4797044", "4797229", "4797234", "4797911", "4798125", "4798253", "4797407", "8039756",
                                   "5021654", "8038771", "8038774", "4797177"]
  
    # Ažurirajte predloge za "part number" u comboboxu
    part_number_combobox['values'] = predlozi_za_part_number
    # Ažurirajte predloge za "part number"
    # part_number_combobox['values'] = predlozi_za_part_number
def azuriraj_defect():
    global predlozi_za_defect
    defect = defect_var.get()
    # Ovde implementirajte logiku za generisanje predloga za "part number" na osnovu stanice i project vrednosti
    # Na primer, koristite if-elif niz za različite kombinacije stanica i project-a
    if defect == "ESTETSKI":
        predlozi_za_defect = ["E1", "E2", "E3", "E4", "E5", "E6", "E7", "E8", "E9", "E10"]
    elif defect == "DIMENZIONI":
        predlozi_za_defect= ["D1", "D2","D3","D4","D5","D6","D7","D8","D9","D10"]
    elif defect == "FUNKCIONALNI":
        predlozi_za_defect = ["F1","F2","F3","F4","F5","F6","F7","F8","F9","F10"]
    elif defect == "NEMA DEFEKTA":
        predlozi_za_defect = [" "]
 
  
    # Ažurirajte predloge za "part number" u comboboxu
    ponudjene_vrednosti_var['values'] = predlozi_za_defect
def prikazi_predloge(*args):
    uneta_vrednost = part_number_var.get()
    predlozi = [predlog for predlog in predlozi_za_part_number if predlog.startswith(uneta_vrednost)]
    part_number_combobox['values'] = predlozi

# Kreiranje glavnog prozora
root = tk.Tk()
root.title("SHENCHI")
root.geometry("500x550")  # Promenite dimenzije prema vašim potrebama

# Postavljanje pozadinske boje na zelenu
root.configure(bg="white")

# Postavljanje stila za promenu veličine fonta
style = ttk.Style()
style.configure('TEntry', font=('Helvetica', 17, "bold"))  # Promenite veličinu fonta prema vašim potrebama

# Definisanje varijabli za različita polja
smena_var = tk.StringVar()
datum_var = tk.StringVar()
operater_var = tk.StringVar()
stanica_var = tk.StringVar()
project_var = tk.StringVar()
part_number_var = tk.StringVar()
batch_number_var = tk.IntVar()
kolicina_var = [tk.IntVar() for _ in range(3)]  # Koristimo IntVar za unos brojeva
defect_var = tk.StringVar()
ponudjene_vrednosti_var = tk.StringVar()
broj_delova_var = tk.IntVar()  # IntVar za unos broja delova
style.configure('TCombobox', borderwidth=2, relief="solid")
# Kreiranje i postavljanje labela i input polja
# Kreiranje i postavljanje labela i input polja
# Kreiranje i postavljanje labela i input polja
tk.Label(root, text="Smena:").grid(row=0, column=0)
# ... Sve ostale label widgete postavite na isti način

smena_combobox = ttk.Combobox(root, textvariable=smena_var, values=["I", "II", "III"])
smena_combobox.grid(row=0, column=1)
smena_combobox.grid(row=0, column=1, padx=7, pady=7)


tk.Label(root, text="Datum:").grid(row=1, column=0)
datum_entry = tk.Entry(root, textvariable=datum_var)
datum_entry.grid(row=1, column=1, padx=7, pady=7)

# Postavite okvir oko Entry polja
datum_entry.config(borderwidth=1, relief="solid")

# Postavite trenutni datum u polje
trenutni_datum = datetime.now().strftime("%Y-%m-%d")  # Formatiranje datuma kao YYYY-MM-DD
datum_var.set(trenutni_datum)


tk.Label(root, text="ID Operatera:").grid(row=2, column=0)
operater_entry = tk.Entry(root, textvariable=operater_var)
operater_entry.grid(row=2, column=1)
operater_entry.grid(row=2, column=1, padx=7, pady=7)

# Postavite okvir oko Entry polja
operater_entry.config(borderwidth=1, relief="solid")

tk.Label(root, text="Radna stanica:").grid(row=3, column=0)
stanica_combobox = ttk.Combobox(root, textvariable=stanica_var, values=["GP12", "WELDING", "STAMMPING", "WAREHOUSE"])
stanica_combobox.grid(row=3, column=1)
stanica_combobox.grid(row=3, column=1, padx=7, pady=7)


tk.Label(root, text="Projekat:").grid(row=4, column=0)
project_combobox = ttk.Combobox(root, textvariable=project_var, values=["G60", "G61", "PHEV", "MMA", "EVA2"])
project_combobox.grid(row=4, column=1)
project_combobox.grid(row=4, column=1, padx=7, pady=7)

tk.Label(root, text="Part Number:").grid(row=5, column=0)
part_number_combobox = ttk.Combobox(root, textvariable=part_number_var)
part_number_combobox.grid(row=5, column=1)
part_number_combobox.grid(row=5, column=1, padx=7, pady=7)

# Postavljanje akcije za ažuriranje predloga za "part number" kada se promeni stanica ili project
project_combobox.bind("<<ComboboxSelected>>", lambda event: azuriraj_part_number())

part_number_var.trace_add('write', prikazi_predloge)
tk.Label(root, text="Batch:").grid(row=6, column=0)
batch_number_var = tk.StringVar()  # Dodajte ovu liniju da biste kreirali StringVar za Batch broj
batch_number_entry = tk.Entry(root, textvariable=batch_number_var)  # Dodajte textvariable=batch_number_var ovde
batch_number_entry.grid(row=6, column=1)
batch_number_entry.grid(row=6, column=1, padx=7, pady=7)

# Postavite okvir oko Entry polja
batch_number_entry.config(borderwidth=1, relief="solid")

# Kreiranje i postavljanje labela i input polja za količine
kolicina_labels = ["Kolicina pregledanih delova", "Kolicina losih delova(skart)", "Kolicina OK"]
for i in range(3):
    tk.Label(root, text=f"{kolicina_labels[i]}:").grid(row=7+i, column=0)
    kolicina_entry = tk.Entry(root, textvariable=kolicina_var[i])
    kolicina_entry.grid(row=7+i, column=1)
    kolicina_entry.grid(row=7+i, column=1, padx=7, pady=7)

# Postavite okvir oko Entry polja
    kolicina_entry.config(borderwidth=1, relief="solid")

# Kreiranje i postavljanje labela i input polja za defect
tk.Label(root, text="Defect:").grid(row=10, column=0)
defect_combobox = ttk.Combobox(root, textvariable=defect_var, values=["ESTETSKI", "DIMENZIONI", "FUNKCIONALNI","NEMA DEFEKTA"])
defect_combobox.grid(row=10, column=1)
defect_combobox.bind("<<ComboboxSelected>>", lambda event: azuriraj_defect())
defect_combobox.grid(row=10, column=1, padx=7, pady=7)

# Postavite okvir oko Entry polja


tk.Label(root, text="Opis defecta:").grid(row=11, column=0)
ponudjene_vrednosti_var = ttk.Combobox(root, textvariable=ponudjene_vrednosti_var)
ponudjene_vrednosti_var.grid(row=11, column=1)
ponudjene_vrednosti_var.grid(row=11, column=1, padx=7, pady=7)
tk.Label(root, text="Broj delova po kutiji:").grid(row=12, column=0)
broj_delova_var_entry = tk.Entry(root, textvariable=broj_delova_var)
broj_delova_var_entry.grid(row=12, column=1)
broj_delova_var_entry.grid(row=12, column=1, padx=7, pady=7)

# Postavite okvir oko Entry polja
broj_delova_var_entry.config(borderwidth=1, relief="solid")
# Kreiranje i postavljanje

# Funkcija za čuvanje podataka
def sacuvaj_podatke():
  
# Postavite naziv servera i bazu podataka
    server = 'DESKTOP-M7QDFHB\MSSQLSERVER01'
    database = 'GP12'

    # Kreirajte konekciju sa bazom podataka koristeći Windows autentifikaciju
    conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';Trusted_Connection=yes;')
    kolicinadorada = 0
    # Kreirajte kursor za izvršavanje SQL upita
    cursor = conn.cursor()
    smena = smena_var.get()
    datum = datum_var.get()
    operater = operater_var.get()
    stanica = stanica_var.get()
    project = project_var.get()
    part_number = part_number_var.get()
    batch_number = batch_number_var.get()
    kolicina_pregledanih = kolicina_var[0].get()
    kolicina_nok = kolicina_var[1].get()
    kolicina_ok = kolicina_var[2].get()
    defect = defect_var.get()
    defect_description = ponudjene_vrednosti_var.get()
    kolicinadorada = kolicina_pregledanih - kolicina_ok - kolicina_nok
    sql_query = """INSERT INTO GP12Stanica (Smena, Datum, Operater, Workstation, Project, PartNumber, BatchNumber, QuantityInspected, QuantitySkart, QuantityDorada ,QuantityOK, Defect, DefectDescription)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""

    # Izvršavanje SQL upita sa vrednostima
    cursor.execute(sql_query, (smena, datum, operater, stanica, project, part_number, batch_number, kolicina_pregledanih, kolicina_nok,kolicinadorada ,kolicina_ok, defect, defect_description))

    # Potvrda promena u bazi
    conn.commit()


    # Zatvorite kursor i konekciju
    cursor.close()
    conn.close()
def praviqrKod():
    datum = datum_var.get()
    operater = operater_var.get()
    project = project_var.get()
    part_number = part_number_var.get()
    batch_number = batch_number_var.get()
    kolicina = broj_delova_var.get()
    ukupnakolicinaOk = kolicina_var[2].get()
    qr_content = f" Datum: {datum}\nID Operatera: {operater}\nProjekat: {project}\nPart Number: {part_number}\nBatch Number: {batch_number}\nKolicina: {kolicina}"
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(qr_content)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    desktop_directory = os.path.expanduser("~/Desktop/SI/VezbaPy/GP12/QRkodovi")
    brojkoda = ukupnakolicinaOk // kolicina

# Generišite ime fajla na osnovu part_number i dodajte ekstenziju ".png"
    file_name = f"{operater}_{datum}_{part_number}.png"
    qr_code_path = os.path.join(desktop_directory, file_name)
 
    img.save(qr_code_path)
    stampaj_qr_kod(qr_code_path, brojkoda)
def stampaj_qr_kod(fajl, brojkopija):
    PHYSICALWIDTH = 110
    PHYSICALHEIGHT = 110
    NUM_COPIES = brojkopija  # Promenite ovu vrednost na željeni broj kopija

    printer_name = win32print.GetDefaultPrinter()
    file_name = fajl

    hDC = win32ui.CreateDC()
    hDC.CreatePrinterDC(printer_name)
    printer_size = hDC.GetDeviceCaps(PHYSICALWIDTH), hDC.GetDeviceCaps(PHYSICALHEIGHT)

    bmp = Image.open(file_name)
    if bmp.size[0] < bmp.size[1]:
        bmp = bmp.rotate(90)

    hDC.StartDoc(file_name)

    for _ in range(NUM_COPIES):
        hDC.StartPage()

        dib = ImageWin.Dib(bmp)
        dib.draw(hDC.GetHandleOutput(), (0, 0, printer_size[0], printer_size[1]))

        hDC.EndPage()

    hDC.EndDoc()
    hDC.DeleteDC()

def dugme_sacuvaj_klik():
    smena = smena_var.get()
    datum = datum_var.get()
    operater = operater_var.get()
    stanica = stanica_var.get()
    project = project_var.get()
    part_number = part_number_var.get()
    batch_number = batch_number_var.get()
    kolicina_pregledanih = kolicina_var[0].get()
    kolicina_nok = kolicina_var[1].get()
    kolicina_ok = kolicina_var[2].get()
    defect = defect_var.get()
    defect_description = ponudjene_vrednosti_var.get()
    brojdelova = broj_delova_var.get()
      # Check if any of the required fields are empty
    if (
        not smena
        or not datum
        or not operater
        or not stanica
        or not project
        or not part_number
        or not batch_number
        or not kolicina_pregledanih
        or not defect
    ):
        messagebox.showerror("Greška", "Sva polja moraju biti popunjena.")
        return
    current_date = datetime.now().strftime("%Y-%m-%d")
    if datum != current_date:
        messagebox.showerror("Greška", "Datum mora biti jednak današnjem datumu.")
        return
    if len(str(batch_number)) < 5:
        messagebox.showerror("Greška", "Batch number mora sadržati barem pet cifara.")
        return  # Prekini izvršenje funkcije
    if defect in ["ESTETSKI", "DIMENZIONI", "FUNKCIONALNI"] and not defect_description:
        messagebox.showerror("Greška", "Unesite opis ako je defect označen kao 'ESTETSKI' ili 'DIMENZIONI' ili 'FUNKCIONALNI'.")
        return
    if kolicina_pregledanih < kolicina_ok + kolicina_nok:
        messagebox.showerror("Greška", "Proverite polje za kolicinu pregledanih delova.")
        return
    if brojdelova == 0:
        messagebox.showerror("Greška", "Broj delova po kutiji ne sme biti 0.")
        return
    sacuvaj_podatke()
    praviqrKod()
    odgovor = messagebox.askquestion("Dodaj novi unos", "Uspesno ste dodali! Da li želite da dodate novi unos?")
    if odgovor == "yes":
        # Resetujte polja za unos za novi unos
        smena_var.set("")
        datum_var.set(trenutni_datum)
        operater_var.set("")
        stanica_var.set("")
        project_var.set("")
        part_number_var.set("")
        batch_number_var.set(0)
        for i in range(3):
            kolicina_var[i].set(0)
        defect_var.set("")
        ponudjene_vrednosti_var.set("")
        broj_delova_var.set(0)
    else:
        root.quit()  # Zatvorite aplikaciju
    

    # Dugme za čuvanje podataka
# Kreiranje dugmeta za čuvanje podataka
sacuvaj_button = tk.Button(root, text="Sačuvaj", command=dugme_sacuvaj_klik)

# Postavljanje opcija stila za dugme
sacuvaj_button.configure(bg="green", fg="white", font=("Helvetica", 14), relief="raised", width=10, height=1)

# Postavljanje dugmeta na ekran
sacuvaj_button.grid(row=20, columnspan=2)


    # Pokretanje glavne petlje
root.mainloop()
