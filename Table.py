import PySimpleGUI as sg
import os
current_directory = os.path.dirname(__file__)
key = "readin_googlesheet.json"
home_image ="geologo2_.png"
icon = "geo_default_logo"
import pygame
CREDENTIALS_JSON_PATH = os.path.join(current_directory,key)



# Create a right-click context menu to copy the selected row

#sg.theme("darkblue")
# table for view sim
def create_table(headings, data):

    print(headings)
    print(data)


    grade_information_window_layout = [
        [sg.Table(
            values=data,
            headings=headings,
            display_row_numbers=True,
            max_col_width=35,
            auto_size_columns=False,
            justification='left',
            num_rows=15,
            key='-TABLE-',
            alternating_row_color="#b0b0b0",
            header_background_color="#247DBF",
            text_color="black",
            background_color='lightgray',
            header_text_color="Black",
            row_height=35,
            enable_events=True,
            selected_row_colors=['white', '#0011ff'],




        )]
    ]

    grade_information_window = sg.Window("Geo Sim Inventory - September 2023 - View inventory", grade_information_window_layout, modal=True)

    while True:
        event, values = grade_information_window.read()
        if event == 'Exit' or event == sg.WIN_CLOSED:
            break



    grade_information_window.close()


# table for view BLE driver id:
def create_table_ble(headings, data):

    print(headings)
    print(data)


    grade_information_window_layout = [
        [sg.Table(
            values=data,
            headings=headings,
            display_row_numbers=True,
            max_col_width=35,
            auto_size_columns=False,
            justification='left',
            num_rows=15,
            key='-TABLE-',
            alternating_row_color="#b0b0b0",
            header_background_color="#247DBF",
            text_color="black",
            background_color='lightgray',
            header_text_color="Black",
            row_height=35,
            enable_events=True,
            selected_row_colors=['white', '#0011ff'],




        )]
    ]

    grade_information_window = sg.Window("Geo Sim Inventory - September 2023 - View inventory", grade_information_window_layout, modal=True)

    while True:
        event, values = grade_information_window.read()
        if event == 'Exit' or event == sg.WIN_CLOSED:
            break



    grade_information_window.close()
