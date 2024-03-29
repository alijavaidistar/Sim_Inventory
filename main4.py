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

# 1.2

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
error_sound = pygame.mixer.Sound('errorsound.mp3')
success_sound = pygame.mixer.Sound('successsound.mp3')
menu_click = pygame.mixer.Sound("menuclick.mp3")
window_main ={}

# json
g_api = "readin_googlesheet.json"

# global for URL and json files





########################################################################################################################
# functions start


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
        [sg.Text("Customer"), sg.Input(pad=(36, 0), key='customer', do_not_clear=True, size=(20, 1),text_color="white")],
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

# sim count
def sim_count():
    import PySimpleGUI as sg

    def calculate_sim_count(user_input):
        one_sim = 0.09
        one_container = 19.93

        sim_weight = user_input - one_container
        sim_count_total = sim_weight / one_sim

        return sim_count_total

    layout = [
        [sg.Text("Input the total weight (Sim + Container):")],
        [sg.InputText(key="weight_input",size=(32,0))],
        [sg.Button("Calculate"), sg.Button("Exit")],
        [sg.Text("", size=(30, 1), key="output")]
    ]

    window = sg.Window("SIM Count", layout)

    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED or event == "Exit":
            break

        if event == "Calculate":
            try:
                user_input = float(values["weight_input"])
                sim_count = calculate_sim_count(user_input)
                window["output"].update(f"Estimated SIM Count: {int(sim_count)}")
            except ValueError:
                sg.popup_error("Invalid input. Please enter a valid number.")

    window.close()

def beacon():


    # Function to get customer names from the customer sheet
    def get_customer_names(sheet):
        customers = sheet.col_values(1)[1:]
        sorted_customers = sorted(customers)
        return sorted_customers

    ## Replace with the full URL of the Google Sheets document
    SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1NsC2hFneYMvdYZrIj8ZHLhENlK2BOBBi_rFR3yK_ke4/edit#gid=0'

    # Extract the spreadsheet ID from the URL
    spreadsheet_id = re.findall("/spreadsheets/d/([a-zA-Z0-9-_]+)", SPREADSHEET_URL)[0]

    # Create a scope and credentials
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name("readin_googlesheet.json", scope) # json directory

    client = gspread.authorize(credentials)

    # Open the spreadsheet by ID
    spreadsheet = client.open_by_key(spreadsheet_id)

    # Open Customer, Sample, and PO sheets
    customer_sheet = spreadsheet.worksheet("Customers")
    tool = spreadsheet.worksheet("Tool_PO")
    driver = spreadsheet.worksheet("Driver_PO")
    sample_sheet = spreadsheet.worksheet("Samples")
    table_data = []

    # facts total driver ID sold, and tool tag
    driver_total = len(driver.get_all_values())-1
    tool_total= len(tool.get_all_values())-1
    sample_total = len(sample_sheet.get_all_values()) - 1



    # for beacon table
    count_table_data = 0
    customer_name = 'None'

    # Layout for the GUI
    layout = [
        [sg.Text("A maximum of 60 entries are permitted within a minute",background_color='Red',text_color='Black',border_width=3)],
        [sg.CalendarButton("Order date:", format='%m/%d/%Y'),
         sg.Input(key='-DATE-', do_not_clear=True, pad=(22, 0), size=(18, 1), disabled=True, background_color="#b0b0b0",
                  text_color="black", tooltip="Click the button to select date")],
        [sg.Text('Customer:',justification='left'), sg.Combo(values=get_customer_names(customer_sheet), key='-CUSTOMER-',pad=(32,0),size=(16, 1),text_color='white')],
        [sg.Button('Add Customer', button_color=('white', '#3D3D3D'), border_width=2,)],
        [sg.Text('Order type:',justification='left'), sg.Radio('Samples', "RADIO1", key='-SAMPLES-'), sg.Radio('PO', "RADIO1", key='-PO-')],
        [sg.Text('Beacon type:',justification='left'),sg.Text(pad=(3,0)), sg.DropDown(values=['DriverID', 'Tool Tag'], key="device_type",size=(18, 1),text_color='white')],
        [sg.Text('Scan Beacon: ',justification='left'), sg.InputText(key='-BEACON-',size=(64,1))],
        [sg.Button('Submit', bind_return_key=True), sg.Button('Exit')],
        [sg.Table(values=[table_data], headings=['Date', 'Customer', 'Device Type', 'Beacon ID'], key='-TABLE-',background_color='white',alternating_row_color="#b0b0b0",text_color='black',max_col_width=(80),auto_size_columns=False,header_background_color="#247DBF",header_text_color='black',expand_x=True,justification='left',)],
        [sg.Text(f'         Customer:',font='any 10 bold'),sg.Text(f'{customer_name}',key='name_beacon')],
        [sg.Text(f' Total Scanned:',font='any 10 bold'),sg.Text(f'{count_table_data}',key='count_beacon')],
        [sg.Text(f'Total Driver IDs:',font='any 10 bold',text_color='Grey'),sg.Text(f'{driver_total}',text_color='Grey'),sg.Text(pad=(15,0)),sg.Text(f'Total Tool Tags:',font='any 10 bold',text_color='Gray'),sg.Text(f'{tool_total}',text_color='Grey'),sg.Text(pad=(15,0)),sg.Text(f'Total Samples:',font='any 10 bold',text_color='Gray'),sg.Text(f'{sample_total}',text_color='Grey')]


    ]

    window = sg.Window('Beacon Scanner', layout,icon='ble_driver.ico')

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED or event == 'Exit':
            break

        try:

            if event == 'Add Customer':
                new_customer = sg.popup_get_text('Enter new customer name:',title='New Customer')
                if new_customer:
                    customer_sheet.append_row([new_customer])
                    success_sound.play()
                    sg.popup('Customer added refreshing...',title='Success')
                    break

            if event == 'Submit' or (event == '-BEACON-' and values['-BEACON-']):
                # avoid duplicates, get values of all beacons first
                past_beacons_sample = sample_sheet.col_values(4)
                past_beacons_driver = driver.col_values(3)
                past_beacons_tool = tool.col_values(3)
                past_customers = customer_sheet.col_values(1)


                date = values['-DATE-']
                customer = values['-CUSTOMER-']
                beacon_id = values['-BEACON-']
                device_type = values['device_type']

                if not values['-SAMPLES-'] and not values['-PO-']:
                    error_sound.play()
                    sg.popup_error('Please select either Samples or PO',title='Error')
                elif not device_type:
                    error_sound.play()
                    sg.popup_error('Please select a device type',title='Error')
                elif not customer:
                    error_sound.play()
                    sg.popup_error('Please select a customer',title='Error')
                elif not beacon_id:
                    error_sound.play()
                    sg.popup_error('Beacon ID cannot be blank', title='Error')
                else:
                    if values['-SAMPLES-']:
                        if beacon_id in past_beacons_sample:
                            error_sound.play()
                            sg.popup_error('Beacon already exists',title='Error')
                        else:
                            sample_sheet.append_row([date, customer, device_type, beacon_id])
                            table_data.append([date, customer, device_type, beacon_id])
                    elif values['-PO-']:
                        if device_type == 'DriverID':
                            if beacon_id in past_beacons_driver:
                                error_sound.play()
                                sg.popup_error('Driver Id already exists',title='Error')
                            else:
                                driver.append_row([date, customer, beacon_id])
                                table_data.append([date, customer, device_type, beacon_id])
                        elif device_type == 'Tool Tag':
                            if beacon_id in past_beacons_tool:
                                error_sound.play()
                                sg.popup_error('Tool already exists',title='Error')
                            else:
                                tool.append_row([date, customer, beacon_id])
                                table_data.append([date, customer, device_type, beacon_id])
                    window['-BEACON-'].update("")

                    # Ensure only unique data is appended to table_data*****************fix duplicates should not append to able
                    #table_data.append([date, customer, device_type, beacon_id])

                    # table data information
                    count_table_data = len(table_data)
                    customer_name = customer

                    window['-TABLE-'].update(values=table_data)
                    window['count_beacon'].update(count_table_data)
                    window['name_beacon'].update(customer_name)





        except Exception as e:
            error_sound.play()
            custom_error_message = "You have reached entry limit, please wait for 60 seconds, else contact developer"
            sg.popup_error(custom_error_message, title="Error")

    window.close()











########################################################################################################################
# functions end



########################################################################################################################
# Google Sheets authentication##################################################################
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]


creds = ServiceAccountCredentials.from_json_keyfile_name("readin_googlesheet.json", scope) # my pc

client = gspread.authorize(creds)





# Homescreen##################################################################################
sg.theme("darkblue")
sg.set_global_icon("geo_default_logo.ico")
sg.set_options(keep_on_top=True)

layout = [

    [sg.Text("Geometris LP SIM Inventory 2024",relief='sunken', text_color="white",font=('Arial',15),border_width=5),sg.Image(filename="geologo2_.png")],#sg.Image(filename="C:\\Users\\ali\\PycharmProjects\\Sim_Inventory\\geologo2_.png")], fix image
    [sg.Text("Geometris Support Projects",justification="right",expand_x=True)],
    [sg.Text("")],
    [sg.Button("Sim Entry",border_width=3,),sg.Button("View Sim",border_width=3),sg.Button("Sim Count",border_width=3),sg.Button('Beacon',border_width=3)],

    #[sg.InputText(key="-URL-")],
]





window_main = sg.Window("Geo Sim Inventory - V1.1.2 - BUILD MARCH 2024", layout,titlebar_background_color="black",icon="geo_default_logo.ico")

# progress bar
progress_bar = sg.ProgressBar(100, orientation='h', size=(20, 20), key='progressbar')


while True:
    event, values = window_main.read()
    if event in (sg.WIN_CLOSED, 'Back'):
        break

    # read the data from sheet
    elif event == "View Sim":

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
            custom_error_message = "Please verify your internet connection or contact Developer"
            sg.popup_error(custom_error_message, title="Error")

    elif event == "Sim Entry":

        try:
            write_sim_inventory()

        except Exception as e:
            error_sound.play()
            custom_error_message = "Please verify your internet connection or contact Developer "
            sg.popup_error(custom_error_message, title="Error")

    elif event =="Sim Count":
        try:

            sim_count()

        except Exception as e:
            error_sound.play()
            custom_error_message = "Please verify your internet connection or contact Developer "
            sg.popup_error(custom_error_message, title="Error")

    elif event =='Beacon':
        try:   
            beacon()
        except Exception as e:
            error_sound.play()
            custom_error_message = "Please verify your internet connection or contact Developer"
            sg.popup_error(custom_error_message, title="Error")


window_main.close()
