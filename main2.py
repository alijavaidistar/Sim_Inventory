import PySimpleGUI as sg
import gspread
from Table import create_table
from oauth2client.service_account import ServiceAccountCredentials
import re
import json
import os
import pygame


# change harcode adress
# fix table layout more and color
# exception error added for view data if no internet

# GLOBAL VAR
# Get the current directory of your Python script
current_directory = os.path.dirname(__file__)
key = "readin_googlesheet.json"
home_image ="C:\\Users\\ali\\PycharmProjects\\Sim_Inventory\\geologo2_.png"
icon1 = "C:\\Users\\ali\\PycharmProjects\\Sim_Inventory\\geo_default_logo.ico"
CREDENTIALS_JSON_PATH = os.path.join(current_directory,key)


# Initialize pygame for sound
pygame.init()

# Load the error sound file (replace 'error_sound.wav' with the actual file path)
error_sound = pygame.mixer.Sound('duplicate_error.mp3')








# complete but confirm exception error
def write_sim_inventory():

    # Get the current directory of your Python script
    #CREDENTIALS_JSON_PATH = os.path.join(current_directory,key) # reading the file for API



    # Replace with the full URL of the Google Sheets document
    SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1Ki5T0J0j8IHIE59Tvev7h-fFWmc5SFMPC5Hk7vC5EF8/edit#gid=338593507'

    # Extract the spreadsheet ID from the URL
    spreadsheet_id = re.findall("/spreadsheets/d/([a-zA-Z0-9-_]+)", SPREADSHEET_URL)[0]

    # Create a scope and credentials
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name("readin_googlesheet.json", scope) # json directory


    client = gspread.authorize(credentials)

    # Open the spreadsheet by ID
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.sheet1

    # Customer Drop-Down

    # Shipment drop-down
    shipment_menu = ["DHL", "FEDEX", "UPS", "USPS", "NONE"]

    # status
    status_menu = ["Ordered", "Transit", "Delivered"]

    layout = [
        [sg.CalendarButton("Order date:", format='%m/%d/%Y',tooltip="Skip if ordered date is not available"),
         sg.Input(key='order_date', do_not_clear=True, pad=(22, 0),size=(20, 1),disabled=True,background_color="#b0b0b0",text_color="black",tooltip="Click the button to select date")],
        [sg.Text("Customer"), sg.Input(pad=(36, 0), key='customer', do_not_clear=True, size=(20, 1))],
        [sg.Text("Shipment"), sg.DropDown(shipment_menu, key="shipment", pad=(37, 0),size=(10,1),text_color="white")],
        [sg.Text("Tracking No:"), sg.Input(pad=(19, 0), key='tracking', do_not_clear=True, size=(20, 1))],
        [sg.Text("Qty"), sg.Input(pad=(71, 0), key='qty', do_not_clear=True, size=(10, 1))],
        [sg.Text("Status"), sg.DropDown(status_menu, key="status", pad=(54, 0),text_color="white")],
        [sg.CalendarButton("Deliver date:", format='%m/%d/%Y',size=(10,1)),
         sg.Input(key='deliver', do_not_clear=True, pad=(7, 0),size=(20, 1),disabled=True,background_color="#b0b0b0",text_color="black",tooltip="Click the button to select date")],
        [sg.Text("Comments"), sg.Input(key='comments', do_not_clear=True, size=(20, 1), pad=(28, 0),text_color="white")],
        [sg.Button('Submit')]
    ]

    window = sg.Window("Geometris LP SIM Inventory 2023 - Sim entry ", layout,icon="geo_default_logo.ico")

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED or event == 'Back':
            break
        elif event == 'Submit':

            order_date = values['order_date']
            customer = values['customer']
            shipment = values['shipment']
            tracking = values['tracking']
            qty = values['qty']
            status = values['status']
            deliver = values['deliver']
            comments = values['comments']


            # data validation
            # Check if shipment is a valid option
            if shipment not in ["DHL", "FEDEX", "UPS", "USPS", "NONE"]:
                error_sound.play()
                sg.popup_error("Invalid shipment method! Please select a valid option.",title="Error")

                continue

            if status not in ["Ordered", "Transit", "Delivered"]:
                error_sound.play()
                sg.popup_error("Invalid status! Please select a valid option.",title="Error")
                continue





            # check for duplicates
            previous_tracking_number_ = worksheet.col_values(4) # testing

            # Check if user_input already exists in the list of tracking numbers
            if tracking in previous_tracking_number_:
                error_sound.play()
                sg.popup_error("Duplicate! This tracking number already exists.", title="Error")
                continue

            else:

                # Append the data to the Google Sheets document
                worksheet.append_row([order_date, customer, shipment, tracking, qty, status, deliver, comments])

                sg.popup("Data saved!",title="Success!")
                break



    window.close()


# Scan Driver ID
def driver_id():


    # Initialize a list to store scanned serial numbers
    scanned_serial_numbers = []  # live duplicate catch





    # Replace with the full URL of the Google Sheets document
    SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1-CEUe6m4IP9TNY1mhP8b576FDAXqPV4URo9usjyR6GQ/edit#gid=0'

    # Extract the spreadsheet ID from the URL
    spreadsheet_id = re.findall("/spreadsheets/d/([a-zA-Z0-9-_]+)", SPREADSHEET_URL)[0]

    # Create a scope and credentials
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name("readin_googlesheet.json", scope) # json directory


    client = gspread.authorize(credentials)

    # Open the spreadsheet by ID
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.sheet1

    # Customer Drop-Down

    # Shipment drop-down


    # customer array
    customer_driver_id= ["CUSTOMER A", "CUSTOMER B", "CUSTOMER C"]



    # screen
    layout = [
        [sg.CalendarButton("Ship date:", format='%m/%d/%Y', tooltip=""), sg.Input(key="Driver_ID_ship_date")],
        [sg.Text("Customer:"), sg.DropDown(customer_driver_id, key="driverid_customer", size=(20, 0))],
        [sg.Text("Scan:"), sg.Input(key="driverid_barcode")],
        [sg.Button("Add", key="Add", bind_return_key=True, auto_size_button=True)],
        [sg.Multiline(key="barcode_saved", disabled=True,size=(50, 20))],
        [sg.Button("Submit"),sg.Button("Help")],  # Fixed missing closing bracket for the "Submit" button

    ]

    window = sg.Window("Geometris LP SIM Inventory 2023 - Sim entry ", layout,icon="geo_default_logo.ico")

    while True:
        event, values = window.read()




        if event == sg.WINDOW_CLOSED or event == 'Back':
            break

        elif event =="Add":
            # get the user input values
            each_id = values["driverid_barcode"]  # not working

            ############
            # serial_number = values['-INPUT-']
            if each_id:
                if each_id in scanned_serial_numbers:
                    sg.popup_error(f'Duplicate Serial Number: {each_id}')
                else:
                    scanned_serial_numbers.append(each_id)
                    # sg.popup(f'Successfully Scanned: {serial_number}')
                    # Clear the input field
                    window['driverid_barcode'].update('')

            # Update the output area with all scanned serial numbers

            window['barcode_saved'].update('\n'.join(scanned_serial_numbers))


        elif event == 'Submit':
            # get the user input values
            ship_date = values['Driver_ID_ship_date']
            customer = values['driverid_customer']
            driver_barcode_list = values['barcode_saved']

            #used for google sheets
            driver_barcode_list2 = driver_barcode_list.splitlines()



            # 2. if customer in sheet then append to that sheet else create a new sheet and append data to it.

            # Get a list of all sheet names in the spreadsheet
            sheet_names = [sheet.title for sheet in spreadsheet.worksheets()]
            print(sheet_names)

            if customer in sheet_names:

                try:
                    # Append the data to the Google Sheets document
                    worksheet = spreadsheet.worksheet(customer) # select the sheet where data is supposed to be inputted
                    worksheet.append_row(["",ship_date]) # append data to desired customer sheet


                    # add the list of code in here
                    for barcode in driver_barcode_list2:
                        worksheet.append_row(["","",barcode])


                    # always append the serial to the main sheet where all the IDS exist
                    #worksheet = spreadsheet.worksheet("DRIVERID_RECEIVED")
                    #worksheet.append_row([])
                    sg.popup("Data has been successfully entered!")

                except Exception as e:
                    sg.popup_error(f"Something went wrong: {e}")
            else:
                print("Sheet",customer," not found in the spreadsheet.")

















            # data validation
            # Check if shipment is a valid option
            #if shipment not in ["DHL", "FEDEX", "UPS", "USPS", "NONE"]:
               # error_sound.play()
               # sg.popup_error("Invalid shipment method! Please select a valid option.",title="Error")

               # continue

           # if status not in ["Ordered", "Transit", "Delivered"]:
               # error_sound.play()
               # sg.popup_error("Invalid status! Please select a valid option.",title="Error")
               # continue





            # check for duplicates
            previous_tracking_number_ = worksheet.col_values(4) # testing

            # Check if user_input already exists in the list of tracking numbers
            #if tracking in previous_tracking_number_:
                #error_sound.play()
                #sg.popup_error("Duplicate! This tracking number already exists.", title="Error")
                #continue

            #else:

                ## Append the data to the Google Sheets document
                #worksheet.append_row([order_date, customer, shipment, tracking, qty, status, deliver, comments])

               # sg.popup("Data saved!",title="Success!")
                #break



    window.close()





# Google Sheets authentication##################################################################
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]


creds = ServiceAccountCredentials.from_json_keyfile_name("readin_googlesheet.json", scope) # my pc

client = gspread.authorize(creds)





# Homescreen##################################################################################
sg.theme("darkblue")
sg.set_global_icon("geo_default_logo.ico")

layout = [

    [sg.Text("Geometris LP SIM Inventory 2023 - V1.0", text_color="white",),sg.Image(filename="geologo2_.png")],#sg.Image(filename="C:\\Users\\ali\\PycharmProjects\\Sim_Inventory\\geologo2_.png")], fix image
    [sg.Text("Geometris Support Projects",justification="right",expand_x=True)],
    [sg.Text("Select option:")],
    [sg.Button("View orders"),sg.Button("Sim entry:"), sg.Button("Scan Driver ID")],
    #[sg.InputText(key="-URL-")],
]





window = sg.Window("Geo Sim Inventory - V1.0 - September 2023", layout,titlebar_background_color="black",icon="geo_default_logo.ico")

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Back'):
        break

    # read the data from sheet
    elif event == "View orders":
        try:
            google_sheets_url = "https://docs.google.com/spreadsheets/d/1Ki5T0J0j8IHIE59Tvev7h-fFWmc5SFMPC5Hk7vC5EF8/edit#gid=338593507"  # sim inventory link

            # Extract spreadsheet key from URL
            spreadsheet_key = google_sheets_url.split("/")[5]
            sheet = client.open_by_key(spreadsheet_key).sheet1

            data_list = sheet.get_all_values()

            # Check if data_list is empty
            if not data_list:
                sg.popup_error("No data found in the Google Sheet.", title="Error")
            else:
                headers = data_list[0]
                data_array = data_list[1:]

                create_table(headers, data_array)
        except Exception as e:
            custom_error_message = "An error occurred. Please check your internet connection and try again."
            sg.popup_error(custom_error_message, title="Error")

    elif event == "Sim entry:":
        try:
            write_sim_inventory()

        except Exception as e:
            custom_error_message = "An error occurred. Please check your internet connection and try again."
            sg.popup_error(custom_error_message, title="Error")


    elif event =="Scan Driver ID":
        driver_id()




window.close()
