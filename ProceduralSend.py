import win32com.client as win32
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyautogui
import PySimpleGUI as sg
import time

#send email function, read an excel file and select the user who the e-mail will be sent
def send_email(technician_id):
    try:
        df = pd.read_excel(f'PATH OF THE EXCEL FILE WICH GOING TO BE READ')
        for i, contact in enumerate(df['USER']):

            name = df.loc[i, "NOME"]
            path_attach= f'PATH OF THE FILE WHICH WILL BE ATTACHED TO THE E-MAIL'
        
            file = path_attach + df.loc[i, "ANEXO"]
            date = df.loc[i, "DATA"]
            barcode = df.loc[i, "BARCODE"]
    
            # creation of outlook interaction
            outlook = win32.Dispatch('outlook.application')

            #create an e-mail
            email = outlook.CreateItem(0)

            #Configure e-mail parameters
            email.to = contact
            email.Subject = f"Termo de responsabilidade - {name}"

            email.HTMLBody = f'''
                <p> Olá, 
                <br>
                <br>
                <p> Segue em anexo neste e-mail o termo de responsabilidade, em arquivo PDF, referente ao 
                laptop que o usuário retirou no dia <strong> {date} </strong>. Máquina de barcode <strong> {barcode}.</strong>
                <br>
                <br>
                Atenciosamente; </p><br>

                <p><strong> <font color = "#006600">Time de TI – Catalão <br>
                IT I&O Edge Operations<br>
                John Deere Catalão - NW </strong></font><br>
                John Deere Catalão – Catalão, Goiás, Brazil<br>
                Rua Quadra 11, s/n, Eixo 3, LOTE 000
                </p>
            '''

            #Item attach
            email.Attachments.Add(file)

            #Send e-mail
            email.Send()

    #exception window, indicates an error when sending email
    except Exception as e:
        sg.theme('DarkRed1')
        layout = [[sg.Text(f'E-mail {i+1} has an error and will not be sent. \n Error: {e}')]]
        window = sg.Window('E-mail', layout)

        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == 'Cancel':
                break
            print('You entered ', values[0])

        window.close()

        print(e)
        pass

#update database function, it read the excel file and open the database site (sharepoint list) and update it.
def update_database(technician_id):
    try:

        df = pd.read_excel(f'PATH OF THE EXCEL FILE WICH GOING TO BE READ')
        for i, contact in enumerate(df['USER']):

            #separate first name and last name from a string (NOME)
            name = df.loc[i, "NOME"]
            name_list = name.split()
            name = name_list[0].upper()
            name_last= name_list[1].upper()
            
            barcode = df.loc[i, "BARCODE"]
           
           #open the database site (if you use a sharepoint list)
            browser = webdriver.Edge()
            browser.get('PUT THE SITE ADRESS')

            browser.find_element(By.XPATH, 'PUT THE XPATH OF THE OBJECT THAT YOU WANT TO SELECT').click()
            time.sleep(1.5)
            pyautogui.typewrite(f'{barcode}\t{name}\t{name_last}')
            time.sleep(1)
            pyautogui.press('tab',2)
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(1)
            browser.close()

    except Exception as e:
        print(e)
        sg.theme('DarkRed1')
        layout = [[sg.Text(f'Database update has an error!\nErro: {e}')]]
        window = sg.Window('Database error', layout)

        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED or event == 'Cancel':
                break
            print('You entered ', values[0])

        window.close()
        pass

def main():

    #Window to the user type his user ID, it will identify who will use the automation
    sg.theme('Dark')
    layout = [
        [sg.Text('Type the technician UserID: ')],
        [sg.Input(key='Technician UserID')],
        [sg.Button('OK')]
    ]

    window = sg.Window('Technician UserID', layout, size=(250, 100), finalize=True)

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == 'OK':
            break

    technician_id = values['Technician UserID']
    window.close()

    #call the functions 
    send_email(technician_id)
    update_database(technician_id)

    # confirmation window about the correct execution
    sg.theme('DarkGreen1')
    layout = [[sg.Text('E-mail and database have been succesfully sent/updated!')]]
    window = sg.Window('E-mail/Database confirmation', layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancel':
            break
        print('You entered ', values[0])

    window.close()

if __name__ == '__main__':
    main()
