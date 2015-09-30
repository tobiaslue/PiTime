# -*- coding: utf-8 -*-
# Import
from Tkinter import *
import ttk
import tkMessageBox
import tkFileDialog
import nxppy
import sqlite3
import datetime
import time
import tkSimpleDialog
import xlsxwriter
import shutil
import tkFont
from subprocess import check_output


# ===== USER INTERFACE =====


# Programmfenster erstellen
root = Tk()
root.title('piTime - NFC Time Management')


# Programmfenster an Bildschirmgrösse anpassen
w = root.winfo_screenwidth()
h = root.winfo_screenheight() - root.winfo_screenheight()/7
root.geometry("%dx%d+0+0" % (w, h))

# Benutzerdefinierte Schrift erstellen
notification_font = tkFont.Font(family="Helvetica", size=20)

# Notebook mit 3 verschiedenen Tabs erstellen
notebook_styling = ttk.Style()
notebook_styling.configure('TNotebook.Tab', padding=(50,20))
tabs = ttk.Notebook(root)
tab1 = Frame(tabs)
tab2 = Frame(tabs)
tab3 = Frame(tabs)

tabs.add(tab1, text='Overview')
tabs.add(tab2, text='Show User')
tabs.add(tab3, text='Export')
tabs.pack(fill=BOTH)
root.title("piTime - Time Management")


# ===== FUNKTIONEN =====


# IP Adresse auslesen
ip_adress = check_output(['hostname', '-I'])

# Datenbankverbindung erstellen
db = sqlite3.connect('database.db')
cursor = db.cursor()

# Erstellung der Datenbanken
def create_tables():
    cursor.execute('''CREATE TABLE IF NOT EXISTS
    users(user_id INTEGER PRIMARY KEY, nfc_id TEXT, name TEXT, logged_in INTEGER)
    ''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS
    work_times(work_time_id INTEGER PRIMARY KEY, user_id INTEGER, time_start datetime, time_stop datetime, time_worked INTEGER)
    ''')


# Funktion ruft sich nach 1000ms selber auf und überprüft ob ein NFC-Tag vor dem Lesegerät ist

def get_nfc_id():

    mifare = nxppy.Mifare()

    # Wird ein NFC Tag erkennt, wird überprüft ob der Benutzer vorhanden ist.
    # Danach wird der Benutzer entweder ein bzw- ausgeloggt (write_to_database)
    # Ist der Benutzer nicht vorhanden wird die Dialogbox zum erstellen neuer Benutzer angezeigt

    #NFC Tag vorhanden
    try:
        nfc_id = mifare.select()
        check_user_existence_result = check_user_existence(nfc_id)
        searching_label.config(bg='green', text='Tag detected!')
        root.after(2000, reset_searching_label)

        if check_user_existence_result:
           write_to_database(nfc_id)
        else:
            create_user(nfc_id)

    # Kein NFC Tag vorhanden
    except nxppy.SelectError:
        pass
        
    root.after(1000, get_nfc_id) # Loop

# Existenz des Benutzers in Datenbank überprüfen
# Gibt Wert TRUE / FALSE aus
def check_user_existence(nfc_id):

    cursor.execute('''SELECT nfc_id FROM users''')
    for i in cursor:
                
        if i[0] == nfc_id:
            user_exists = True

            if user_exists:

                return True

            else:
                return False

# Zeiten werden in Datenbank eingetragen

def write_to_database(nfc_id):
    cursor.execute('''SELECT user_id FROM users WHERE nfc_id = ?''', (nfc_id,))
    user_id = cursor.fetchone()[0]
    
    current_datetime = datetime.datetime.now()
    current_time = current_datetime.strftime("%Y-%m-%d %H:%M")

    cursor.execute('''SELECT logged_in FROM users WHERE user_id = ?''', (user_id,))
    logged_in = cursor.fetchone()[0]

    # Zeit eintragen bei Login
    if logged_in == 0:
        cursor.execute('''INSERT INTO work_times(user_id, time_start)
            VALUES(?,?)''', (user_id, current_time))
        logged_in = 1
        cursor.execute('''UPDATE users SET logged_in = ? WHERE user_id = ?''', (logged_in, user_id))

        cursor.execute('''SELECT name FROM users WHERE nfc_id = ?''', (nfc_id,))
        name = cursor.fetchone()[0]
        statusbar.configure(text = name + ": Login at " + current_datetime.strftime("%d.%m.%Y %H:%M") + ".")


    # Zeit eintragen bei Logout
    else:
        
        cursor.execute('''UPDATE work_times SET time_stop = ? WHERE user_id = ? AND time_stop IS NULL''',
                       (current_time, user_id))
        logged_in = 0
        cursor.execute('''UPDATE users SET logged_in = ? WHERE user_id = ?''', (logged_in, user_id))

        cursor.execute('''SELECT name FROM users WHERE nfc_id = ?''', (nfc_id,))
        name = cursor.fetchone()[0]

        worktime = calculate_work_time(nfc_id)
        statusbar.configure(text =name + ": Logout at " + current_datetime.strftime("%d.%m.%Y %H:%M") + "\n"  +worktime)

    db.commit()

# Arbeitszeit berechnen

def calculate_work_time(nfc_id):
    cursor.execute('''SELECT user_id FROM users WHERE nfc_id = ?''', (nfc_id,))
    user_id = cursor.fetchone()[0]

    cursor.execute('''SELECT MAX(work_time_id) FROM work_times WHERE user_id = ? AND time_stop IS NOT NULL''',
                   (user_id,))
    work_time_id = cursor.fetchone()[0]
    
    cursor.execute('''SELECT time_start FROM work_times WHERE work_time_id = ?''', (work_time_id,))
    time_start = cursor.fetchone()[0]
    
    cursor.execute('''SELECT time_stop FROM work_times WHERE work_time_id = ?''', (work_time_id,))
    time_stop = cursor.fetchone()[0]

    time_worked = (datetime.datetime.strptime(time_stop, '%Y-%m-%d %H:%M') - datetime.datetime.strptime(time_start, '%Y-%m-%d %H:%M')).total_seconds()
    cursor.execute('''UPDATE work_times SET time_worked = ? WHERE work_time_id = ?''', (time_worked, work_time_id))

    return("You have worked " + time.strftime("%H:%M", time.gmtime(time_worked)) + " hours.")

    db.commit()

# Dialogfenster um neuen Benutzer zu erstellen
def create_user(nfc_id):
    result = tkMessageBox.askquestion("Register New User", "New NFC Tag detected! Would you like to create a new user?", icon='question')
    if result == 'yes':
        name = tkSimpleDialog.askstring("New User", "Enter new User Name", initialvalue = "John Doe")
        cursor.execute('''INSERT INTO users(nfc_id, name, logged_in)
            VALUES(?,?,?)''', (nfc_id, name, 0))
    
        db.commit()



# Searching Label auf Standard setzen
def reset_searching_label():
        searching_label.config(bg='red', text='Searching NFC Tag...')

# Benutzerübersicht in Tab 2 generieren
# Übersicht wird nach 2000ms gelöscht und neu generiert
def show_user_list():
    count = 0
    cursor.execute('''SELECT name,logged_in FROM users''')
    user_list_tmp = cursor.fetchall()
    inner_frame = Frame(tab2)
    for i in user_list_tmp:
        if user_list_tmp[count][1]:
            bgcolor = "green"
        else:
            bgcolor = "red"
        
        Label(inner_frame, text = user_list_tmp[count][0], bg = bgcolor, font=notification_font).pack(fill=X)
        count += 1
    inner_frame.pack()
    root.after(2000, show_user_list)
    root.after(2000, lambda: inner_frame.destroy())

# Statusleiste erstellen
statusbar = Label(root,bg="blue",fg="white", text='Nothing Special happening here',font=("Helvetica", 20), bd=1, relief=SUNKEN, anchor=W)
statusbar.pack(side=BOTTOM, fill=X)

# Xlsx Datei in ausgewähltem Zeitraum generieren
def export():

    # Zeitraum auslesen und in geeignetes Format für die Datenbank konvertieren
    start = start_date.get()
    start_strp = time.strptime(start, "%d.%m.%Y")
    start_strf = time.strftime("%Y-%m-%d", start_strp)
    
    end = end_date.get()
    end_strp = time.strptime(end, "%d.%m.%Y")
    end_strf = time.strftime("%Y-%m-%d", end_strp)

    # xlsx Datei erstellen
    workbook = xlsxwriter.Workbook('work_times.xlsx')

    # Formatvorlagen erstellen
    bold = workbook.add_format({'bold': True})
    title = workbook.add_format()
    title.set_font_size(30)
    title.set_align('center')
    title.set_bold()

    # Anzahl Benutzer zählen
    cursor.execute('''SELECT MAX(user_id) FROM work_times''')
    max_user_id = cursor.fetchone()[0]

    # Neues Worksheet jeden für Benutzer erstellen
    for i in range(0, max_user_id):
        time_worked = 0.0
        row = 2
        col = 0

        cursor.execute('''SELECT name FROM users WHERE user_id = ?''', (i + 1,))
        name = cursor.fetchone()[0]

        worksheet = workbook.add_worksheet(name)

        worksheet.merge_range('A1:E1', name, title)

        worksheet.write('A2', 'Zeit ID:', bold)
        worksheet.write('B2', 'User ID:', bold)
        worksheet.write('C2', 'Startzeit:', bold)
        worksheet.write('D2', 'Endzeit:', bold)
        worksheet.write('E2', 'Gearbeitet:', bold)

        worksheet.set_column('C:E', 30)

        cursor.execute('''SELECT * FROM work_times WHERE user_id = ? AND time_start BETWEEN ? AND ? AND time_stop IS NOT NULL''', (i + 1, start_strf, end_strf))
        # Datensatz in Worksheet schreiben bzw. konvertieren
        for j in cursor:
            
            for k in j:
                if col <= 1:
                    worksheet.write(row, col, k)
                    col += 1
                elif col >=2 and col <= 3:
                    time_formatted = datetime.datetime.strptime(k.encode("iso-8859-16"),"%Y-%m-%d %H:%M")
                    worksheet.write(row, col, time_formatted.strftime("%d.%m.%Y %H:%M"))
                    col += 1
                else:
                    convert_to_hours = round(float(k)/3600, 2)
                    worksheet.write(row, col, convert_to_hours)
                    col += 1
            time_worked += j[4]
            col = 0
            row += 1
        # Totale Arbeitszeit in Stunden
        time_worked_hours = round(time_worked/3600, 2)
        worksheet.write(row + 1, 3, 'Total', bold)
        worksheet.write(row + 1, 4, time_worked_hours)
    workbook.close()

    export_file_directory = tkFileDialog.askdirectory()
    export_file_directory = export_file_directory + "/work_times.xlsx"
    shutil.move("work_times.xlsx", export_file_directory)


# ===== WIDGETS =====


# Tab 1 Widgets
searching_label = Label(tab1, text='Searching NFC Tag...', bg='red', font=notification_font)
searching_label.pack(fill=X)
progressbar = ttk.Progressbar(tab1, orient=HORIZONTAL, mode='determinate')
progressbar.pack(fill=X)
progressbar.start()


# Tab 3 Widgets
# Zusätzlicher Rahmen erstellen um Grid Layout zu ermöglichen
grid_frame = Frame(tab3)
grid_frame.pack()
Label(grid_frame, text="Export all User to Excel Worksheet", font=notification_font).grid(row=0, column=0, columnspan=2)
start_date = Entry(grid_frame)
start_date.grid(row=1, column=1, sticky=W)
start_date.insert(0, "25.01.2015")
Label(grid_frame, text="Enter Startdate").grid(row=1, column=0, sticky=W)
Label(grid_frame, text="Enter Enddate").grid(row=2, column=0, sticky=W)
end_date = Entry(grid_frame)
end_date.grid(row=2, column=1, sticky=W)
end_date.insert(0, "25.02.2015")
Label(grid_frame, text = "Current IP:").grid(row=5, column=0, sticky=W)
Label(grid_frame, text = ip_adress).grid(row=5, column=1, sticky=W)
Label(grid_frame, text="Connect to Raspi", font=notification_font).grid(row=4, column=0, sticky=W)
exportButton = Button(grid_frame, text='Export Data', command = export)
exportButton.grid(row=3, column=0, columnspan=2, pady=(25,50))

# Main Funktion
def main():
    create_tables()
    get_nfc_id()
    show_user_list()


main()
root.mainloop()






