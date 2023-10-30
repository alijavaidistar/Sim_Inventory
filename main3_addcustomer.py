import PySimpleGUI as sg
import gspread
from Table import create_table, create_table_ble
from oauth2client.service_account import ServiceAccountCredentials
import re
import json
import os
import pygame
import openpyxl
import time
import sys
import matplotlib.pyplot as plt
import numpy as np
from collections import Counter

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
success_sound = pygame.mixer.Sound('saved.mp3')
menu_click = pygame.mixer.Sound("menuclick.mp3")

# json
g_api = "readin_googlesheet.json"








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

    window = sg.Window("Sim entry ", layout,icon="geo_default_logo.ico")

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

                success_sound.play()
                sg.popup("Data saved!",title="Success!")
                break



    window.close()


# Scan Driver ID
def driver_id(): # qouta 60 values in a minutes only

    ####################function to add row number
    def add_row_numbers(values):
        multiline_value = values['barcode_saved'].split('\n')
        output = ""
        for i, line in enumerate(multiline_value):
            output += f"{i + 1}. {line}\n"
        window['barcode_saved'].update(output)



    # Initialize a list to store scanned serial numbers
    scanned_serial_numbers = []  # live duplicate catch

    # customer drop-down
    customer_driver_id = []

    # Open the file and read its contents
    try:
        with open('customer_data_driverID.txt', 'r') as file:
            data = file.read().splitlines()
    except FileNotFoundError:
        data = []

    # Append the data to the customer_driver_id list
    customer_driver_id.extend(data)




    # Replace with the full URL of the Google Sheets document
    SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1IuVS4mbAziZLVrig5lEwO3K4b3ymvNUhcCSMdJclems/edit#gid=0'

    # Extract the spreadsheet ID from the URL
    spreadsheet_id = re.findall("/spreadsheets/d/([a-zA-Z0-9-_]+)", SPREADSHEET_URL)[0]

    # Create a scope and credentials
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name("readin_googlesheet.json", scope) # json directory


    client = gspread.authorize(credentials)

    # Open the spreadsheet by ID
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.sheet1



    # screen
    layout = [

        [sg.Text("Please scan a maximum of 60 per minute.",background_color="#247DBF",text_color="")],
        [sg.CalendarButton("Ship date:", format='%m/%d/%Y', tooltip=""), sg.Input(key="Driver_ID_ship_date",size=(21, 0))],
        [sg.Text("Customer:",pad=(7,0)), sg.DropDown(customer_driver_id, key="driverid_customer", size=(20, 0),)],
        [sg.Text("",pad=(12,0)),sg.Text("Scan:",justification="right"), sg.Input(key="driverid_barcode",size=(40,0))],
        [sg.Button("Add", key="Add", bind_return_key=True, auto_size_button=True),sg.Button("Count"),sg.Button("Samples"),sg.Button("Inventory"),],
        [sg.Multiline(key="barcode_saved", disabled=True,size=(50, 20),background_color="#D3D3D3",text_color="black")],
        [sg.Button("Submit"),sg.Button("Clear"),sg.Button("Timer"),sg.Button("Help")],  # Fixed missing closing bracket for the "Submit" button

    ]

    window = sg.Window("Driver ID entry ", layout,icon="ble_driver.ico")

    while True:
        event, values = window.read()




        if event == sg.WINDOW_CLOSED or event == 'Back':
            break

        ######################
        elif event =="Add":
            # get the user input values
            each_id = values["driverid_barcode"]  # not working

            ############
            # serial_number = values['-INPUT-']
            if each_id:
                if each_id in scanned_serial_numbers:
                    error_sound.play()
                    sg.popup_error(f'Duplicate Serial Number: {each_id}')
                else:
                    scanned_serial_numbers.append(each_id)
                    # sg.popup(f'Successfully Scanned: {serial_number}')
                    # Clear the input field
                    window['driverid_barcode'].update('')

            # Update the output area with all scanned serial numbers

            window['barcode_saved'].update('\n'.join(scanned_serial_numbers))

        ######################
        elif event == 'Submit':
            # get the user input values
            ship_date = values['Driver_ID_ship_date']
            customer = values['driverid_customer']
            driver_barcode_list = values['barcode_saved']

            #used for google sheets
            driver_barcode_list2 = driver_barcode_list.splitlines()






            # 2. if customer in sheet then append to that sheet else create a new sheet and append data to it.
            sheet_name = ""
            # Get a list of all sheet names in the spreadsheet
            sheet_names = [sheet.title for sheet in spreadsheet.worksheets()]
            print(sheet_names)



            ##################################### if no customer is selected
            if customer not in customer_driver_id:
                error_sound.play()
                sg.popup_error('Please select a customer',title="")
                continue


            ######################################  append to main sheet
            counter = 0  # Initialize the counter variable

            if "MAIN" in sheet_names:
                try:
                    main_worksheet = spreadsheet.worksheet("MAIN")
                    previous_serial_ = main_worksheet.col_values(1)

                    duplicates = []
                    for barcode in driver_barcode_list2:
                        if barcode in previous_serial_:
                            duplicates.append(barcode)

                    if duplicates:
                        error_sound.play()
                        duplicate_msg = "Duplicates found: " + ", ".join(duplicates)
                        sg.popup_scrolled(f"These serial numbers already exist\n: {duplicate_msg}", title="Error!",background_color="black",text_color="red")
                    else:
                        if counter + len(driver_barcode_list2) <= 60:
                            # Append all barcodes at once
                            values_to_append = [[barcode] for barcode in driver_barcode_list2]
                            main_worksheet.append_rows(values_to_append)

                            counter += len(driver_barcode_list2)  # Increment the counter by the number of inputs

                            if counter >= 60:
                                time.sleep(60)  # Add a 1-minute delay after 60 inputs
                                counter = 0  # Reset the counter

                            if customer in sheet_names:
                                worksheet = spreadsheet.worksheet(customer)
                                # Add an empty row before appending any data
                                worksheet.append_row(["SHIP-DATE", "QR CODE", "SERIAL", "MAC"])
                                worksheet.append_row([ship_date])
                                # Append all barcodes at once
                                worksheet.append_rows([["", barcode] for barcode in driver_barcode_list2])
                            else:
                                spreadsheet.add_worksheet(customer, rows=100, cols=20)
                                worksheet = spreadsheet.worksheet(customer)
                                worksheet.append_row(["SHIP-DATE", "QR CODE", "SERIAL", "MAC"])
                                time.sleep(2)
                                worksheet.append_row([ship_date])
                                worksheet.append_rows([["", barcode] for barcode in driver_barcode_list2])

                            success_sound.play()
                            sg.popup("Data has been successfully entered!",title='Success')
                            break
                        else:
                            error_sound.play()
                            sg.popup_error("Limit exceeded. 60 input allowed per minute",title="Wait")

                except gspread.exceptions.APIError as e:
                    error_sound.play()
                    if 'Quota exceeded' in str(e):
                        sg.popup_error("Limit exceeded. 60 input allowed per minute")
                    else:
                        sg.popup_error(f"Something went wrong: {e}")


        ######################
        elif event == "Clear":
            sg.popup_ok("Reopen the window.")
            break


        ######################
        elif event=="Help":
            sg.popup_ok("To add a new customer, please configure the Notepad file in the program directory.",title="BLE DRIVER ID")

        ######################
        elif event =="Count":
            add_row_numbers(values)

        ######################
        elif event =="Timer":
            def create_layout():
                layout = [
                    [sg.Text("Timer: ", size=(10, 1)), sg.Text("60", size=(5, 1), key="-TIMER-")],
                    [sg.Button("Start Timer", key="-START-"), sg.Button("Reset", key="-RESET-"), sg.Button("Exit")]
                ]
                return layout

            def start_timer(window):
                time_left = 60
                while time_left >= 0:
                    window["-TIMER-"].update(value=f"{time_left}")
                    event, values = window.read(timeout=1000)  # 1 second timeout
                    if event in (sg.WIN_CLOSED, "Exit"):
                        break
                    if event == "-RESET-":
                        time_left = 60
                    time_left -= 1

            #sg.theme("Black")
            layout = create_layout()
            window = sg.Window("Countdown Timer", layout)

            while True:
                event, values = window.read()
                if event == sg.WIN_CLOSED or event == "Exit":
                    break
                if event == "-START-":
                    start_timer(window)

            window.close()

        elif event =="Samples":
            main_worksheet = spreadsheet.worksheet("SAMPLE")

            # Get the data range of the worksheet
            data = main_worksheet.col_values(5)  # Assuming you want to analyze data from Column E

            # Count the occurrences of each value
            counts = dict(Counter(data))

            # Extract unique values and their corresponding counts
            unique_values = list(counts.keys())
            unique_counts = list(counts.values())

            # Create a bar graph
            fig, ax = plt.subplots()
            colors = plt.cm.Paired(np.linspace(0, 1, len(unique_values)))
            bars = ax.bar(unique_values, unique_counts, color=colors)
            ax.set_title('Total Samples')
            ax.set_xlabel('Customer')
            ax.set_ylabel('Quantity')
            plt.xticks(rotation=45, ha='right')
            for bar, count in zip(bars, unique_counts):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height(), count, ha='center', va='bottom')

            plt.tight_layout()

            plt.show()


        elif event =="Inventory":
            main_worksheet = spreadsheet.worksheet("INVENTORY")

            # Get the data range of the worksheet
            data = main_worksheet.get_all_values()

            # Calculate the total number of cells that have data
            num_rows = len(data)
            num_columns = len(data[0]) if data else 0
            total_cells = num_rows * num_columns

            sg.popup(f'Total in stock: {total_cells}', title='Statistics Result')

        """
        elif event =="Order":
            # Get all worksheet titles
            all_worksheets = [worksheet.title for worksheet in spreadsheet.worksheets()]

            # Exclude "MAIN" and "SAMPLE" from the list
            relevant_worksheets = [worksheet for worksheet in all_worksheets if worksheet not in ["MAIN", "SAMPLE"]]

            total_cells_list = []
            for worksheet_title in relevant_worksheets:
                worksheet = spreadsheet.worksheet(worksheet_title)
                data = worksheet.get_all_values()
                num_rows = len(data)
                num_columns = len(data[0]) if data else 0
                total_cells_list.append(num_rows * num_columns)

            # Create a bar graph
            fig, ax = plt.subplots()
            ax.bar(relevant_worksheets, total_cells_list, color='skyblue')
            ax.set_title('Total BLE Orders')
            ax.set_xlabel('Customers')
            ax.set_ylabel('Quantity')
            plt.xticks(rotation=45, ha='right')
            for i, v in enumerate(total_cells_list):
                ax.text(i, v + 3, str(v), ha='center')

            plt.tight_layout()
            plt.show()
            """




    window.close()


def tool_tag():

    ####################function to add row number
    def add_row_numbers(values):
        multiline_value = values['barcode_saved'].split('\n')
        output = ""
        for i, line in enumerate(multiline_value):
            output += f"{i + 1}. {line}\n"
        window['barcode_saved'].update(output)


    ####################################################
    # Initialize a list to store scanned serial numbers
    scanned_serial_numbers = []  # live duplicate catch

    # customer drop-down
    customer_tool_id = []

    # Open the file and read its contents
    try:
        with open('customer_data_tooltag.txt', 'r') as file:
            data = file.read().splitlines()
    except FileNotFoundError:
        data = []

    # Append the data to the customer_tool_id list
    customer_tool_id.extend(data)

    # Replace with the full URL of the Google Sheets document
    SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1-pyinVIiW4vFy0NCCbgRz8bAZhV-ZXAOiMf7RpObL6g/edit#gid=825221153'

    # Extract the spreadsheet ID from the URL
    spreadsheet_id = re.findall("/spreadsheets/d/([a-zA-Z0-9-_]+)", SPREADSHEET_URL)[0]

    # Create a scope and credentials
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name("readin_googlesheet.json",
                                                                   scope)  # json directory

    client = gspread.authorize(credentials)

    # Open the spreadsheet by ID
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.sheet1

    # screen
    layout = [

        [sg.Text("Please scan a maximum of 60 per minute.", background_color="#247DBF", text_color="")],
        [sg.CalendarButton("Ship date:", format='%m/%d/%Y', tooltip=""),
         sg.Input(key="tool_ID_ship_date", size=(21, 0))],
        [sg.Text("Customer:", pad=(7, 0)),
         sg.DropDown(customer_tool_id, key="toolid_customer", size=(20, 0), )],
        [sg.Text("", pad=(12, 0)), sg.Text("Scan:", justification="right"),
         sg.Input(key="toolid_barcode", size=(40, 0))],
        [sg.Button("Add", key="Add", bind_return_key=True, auto_size_button=True), sg.Button("Count"),
         sg.Button("Samples"), sg.Button("Inventory"), ],
        [sg.Multiline(key="barcode_saved", disabled=True, size=(50, 20), background_color="#D3D3D3",
                      text_color="black")],
        [sg.Button("Submit"), sg.Button("Clear"), sg.Button("Timer"), sg.Button("Help")],
        # Fixed missing closing bracket for the "Submit" button

    ]

    window = sg.Window("Tool ID entry ", layout, icon="ble_tool.ico")

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED or event == 'Back':
            break

        ######################
        elif event == "Add":
            # get the user input values
            each_id = values["toolid_barcode"]  # not working

            ############
            # serial_number = values['-INPUT-']
            if each_id:
                if each_id in scanned_serial_numbers:
                    error_sound.play()
                    sg.popup_error(f'Duplicate Serial Number: {each_id}')
                else:
                    scanned_serial_numbers.append(each_id)
                    # sg.popup(f'Successfully Scanned: {serial_number}')
                    # Clear the input field
                    window['toolid_barcode'].update('')

            # Update the output area with all scanned serial numbers

            window['barcode_saved'].update('\n'.join(scanned_serial_numbers))

        ######################
        elif event == 'Submit':
            # get the user input values
            ship_date = values['tool_ID_ship_date']
            customer = values['toolid_customer']
            tool_barcode_list = values['barcode_saved']

            # used for google sheets
            tool_barcode_list2 = tool_barcode_list.splitlines()

            # 2. if customer in sheet then append to that sheet else create a new sheet and append data to it.
            sheet_name = ""
            # Get a list of all sheet names in the spreadsheet
            sheet_names = [sheet.title for sheet in spreadsheet.worksheets()]
            print(sheet_names)

            ##################################### if no customer is selected
            if customer not in customer_tool_id:
                error_sound.play()
                sg.popup_error('Please select a customer', title="")
                continue

            ######################################  append to main sheet
            counter = 0  # Initialize the counter variable

            if "MAIN" in sheet_names:
                try:
                    main_worksheet = spreadsheet.worksheet("MAIN")
                    previous_serial_ = main_worksheet.col_values(1)

                    duplicates = []
                    for barcode in tool_barcode_list2:
                        if barcode in previous_serial_:
                            duplicates.append(barcode)

                    if duplicates:
                        error_sound.play()
                        duplicate_msg = "Duplicates found: " + ", ".join(duplicates)
                        sg.popup_scrolled(f"These serial numbers already exist\n: {duplicate_msg}", title="Error!",
                                          background_color="black", text_color="red")
                    else:
                        if counter + len(tool_barcode_list2) <= 60:
                            # Append all barcodes at once
                            values_to_append = [[barcode] for barcode in tool_barcode_list2]
                            main_worksheet.append_rows(values_to_append)

                            counter += len(tool_barcode_list2)  # Increment the counter by the number of inputs

                            if counter >= 60:
                                time.sleep(60)  # Add a 1-minute delay after 60 inputs
                                counter = 0  # Reset the counter

                            if customer in sheet_names:
                                worksheet = spreadsheet.worksheet(customer)
                                # Add an empty row before appending any data
                                worksheet.append_row(["SHIP-DATE", "QR CODE", "SERIAL", "MAC"])
                                worksheet.append_row([ship_date])
                                # Append all barcodes at once
                                worksheet.append_rows([["", barcode] for barcode in tool_barcode_list2])
                            else:
                                spreadsheet.add_worksheet(customer, rows=100, cols=20)
                                worksheet = spreadsheet.worksheet(customer)
                                worksheet.append_row(["SHIP-DATE", "QR CODE", "SERIAL", "MAC"])
                                time.sleep(2)
                                worksheet.append_row([ship_date])
                                worksheet.append_rows([["", barcode] for barcode in tool_barcode_list2])

                            success_sound.play()
                            sg.popup("Data has been successfully entered!", title='Success')
                            break
                        else:
                            error_sound.play()
                            sg.popup_error("Limit exceeded. 60 input allowed per minute", title="Wait")

                except gspread.exceptions.APIError as e:
                    error_sound.play()
                    if 'Quota exceeded' in str(e):
                        sg.popup_error("Limit exceeded. 60 input allowed per minute")
                    else:
                        sg.popup_error(f"Something went wrong: {e}")


        ######################
        elif event == "Clear":
            sg.popup_ok("Reopen the window.")
            break


        ######################
        elif event == "Help":
            sg.popup_ok("To add a new customer, please configure the Notepad file in the program directory.",
                        title="BLE DRIVER ID")

        ######################
        elif event == "Count":
            add_row_numbers(values)

        ######################
        elif event == "Timer":
            def create_layout():
                layout = [
                    [sg.Text("Timer: ", size=(10, 1)), sg.Text("60", size=(5, 1), key="-TIMER-")],
                    [sg.Button("Start Timer", key="-START-"), sg.Button("Reset", key="-RESET-"), sg.Button("Exit")]
                ]
                return layout

            def start_timer(window):
                time_left = 60
                while time_left >= 0:
                    window["-TIMER-"].update(value=f"{time_left}")
                    event, values = window.read(timeout=1000)  # 1 second timeout
                    if event in (sg.WIN_CLOSED, "Exit"):
                        break
                    if event == "-RESET-":
                        time_left = 60
                    time_left -= 1

            # sg.theme("Black")
            layout = create_layout()
            window = sg.Window("Countdown Timer", layout)

            while True:
                event, values = window.read()
                if event == sg.WIN_CLOSED or event == "Exit":
                    break
                if event == "-START-":
                    start_timer(window)

            window.close()

        elif event == "Samples":
            main_worksheet = spreadsheet.worksheet("SAMPLE")

            # Get the data range of the worksheet
            data = main_worksheet.col_values(5)  # Assuming you want to analyze data from Column E

            # Count the occurrences of each value
            counts = dict(Counter(data))

            # Extract unique values and their corresponding counts
            unique_values = list(counts.keys())
            unique_counts = list(counts.values())

            # Create a bar graph
            fig, ax = plt.subplots()
            colors = plt.cm.Paired(np.linspace(0, 1, len(unique_values)))
            bars = ax.bar(unique_values, unique_counts, color=colors)
            ax.set_title('Total Samples')
            ax.set_xlabel('Customer')
            ax.set_ylabel('Quantity')
            plt.xticks(rotation=45, ha='right')
            for bar, count in zip(bars, unique_counts):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height(), count, ha='center', va='bottom')

            plt.tight_layout()

            plt.show()


        elif event == "Inventory":
            main_worksheet = spreadsheet.worksheet("INVENTORY")

            # Get the data range of the worksheet
            data = main_worksheet.get_all_values()

            # Calculate the total number of cells that have data
            num_rows = len(data)
            num_columns = len(data[0]) if data else 0
            total_cells = num_rows * num_columns

            sg.popup(f'Total in stock: {total_cells}', title='Statistics Result')

        """
        elif event =="Order":
            # Get all worksheet titles
            all_worksheets = [worksheet.title for worksheet in spreadsheet.worksheets()]

            # Exclude "MAIN" and "SAMPLE" from the list
            relevant_worksheets = [worksheet for worksheet in all_worksheets if worksheet not in ["MAIN", "SAMPLE"]]

            total_cells_list = []
            for worksheet_title in relevant_worksheets:
                worksheet = spreadsheet.worksheet(worksheet_title)
                data = worksheet.get_all_values()
                num_rows = len(data)
                num_columns = len(data[0]) if data else 0
                total_cells_list.append(num_rows * num_columns)

            # Create a bar graph
            fig, ax = plt.subplots()
            ax.bar(relevant_worksheets, total_cells_list, color='skyblue')
            ax.set_title('Total BLE Orders')
            ax.set_xlabel('Customers')
            ax.set_ylabel('Quantity')
            plt.xticks(rotation=45, ha='right')
            for i, v in enumerate(total_cells_list):
                ax.text(i, v + 3, str(v), ha='center')

            plt.tight_layout()
            plt.show()
            """

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

    [sg.Text("Geometris LP SIM Inventory 2023 - V1.1", text_color="white",),sg.Image(filename="geologo2_.png")],#sg.Image(filename="C:\\Users\\ali\\PycharmProjects\\Sim_Inventory\\geologo2_.png")], fix image
    [sg.Text("Geometris Support Projects",justification="right",expand_x=True)],
    [sg.Text("Select option:")],
    [sg.Button("View orders"), sg.Button("Sim entry:"), sg.Button("Driver ID entry"),sg.Button("Tool Tag entry")],
    #[sg.InputText(key="-URL-")],
]





window = sg.Window("Geo Sim Inventory - V1.1 - November 2023", layout,titlebar_background_color="black",icon="geo_default_logo.ico")

# progress bar
progress_bar = sg.ProgressBar(100, orientation='h', size=(20, 20), key='progressbar')


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
            error_sound.play()
            custom_error_message = "Please verify your internet connection or the existence of the file "
            sg.popup_error(custom_error_message, title="Error")

    elif event == "Sim entry:":

        try:
            write_sim_inventory()

        except Exception as e:
            error_sound.play()
            custom_error_message = "Please verify your internet connection or the existence of the file "
            sg.popup_error(custom_error_message, title="Error")


    elif event =="Driver ID entry":
        try:

            driver_id()

        except Exception as e:
            error_sound.play()
            custom_error_message = "Please verify your internet connection or the existence of the file "
            sg.popup_error(custom_error_message, title="Error")

    elif event == "Tool Tag entry":
        try:

            tool_tag()

        except Exception as e:
            error_sound.play()
            custom_error_message = "Please verify your internet connection or the existence of the file "
            sg.popup_ok(custom_error_message, title="Error")






window.close()
