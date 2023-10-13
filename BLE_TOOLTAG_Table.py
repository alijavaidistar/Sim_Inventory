# code workds

import PySimpleGUI as sg


def main():
    # Define the layout of the GUI
    layout = [
        [sg.Text('Scan Serial Number:')],
        [sg.InputText(key='-INPUT-')],
        [sg.Button('Add', key='-ADD-', bind_return_key=True, auto_size_button=True)],
        [sg.Output(size=(40, 10), key='-OUTPUT-')],
    ]

    # Create the window
    window = sg.Window('Serial Number Scanner', layout)

    # Initialize a list to store scanned serial numbers
    scanned_serial_numbers = []

    while True:
        event, values = window.read()

        if event in (sg.WIN_CLOSED, 'Exit'):
            break

        if event == '-ADD-':
            serial_number = values['-INPUT-']
            if serial_number:
                if serial_number in scanned_serial_numbers:
                    sg.popup_error(f'Duplicate Serial Number: {serial_number}')
                else:
                    scanned_serial_numbers.append(serial_number)
                    #sg.popup(f'Successfully Scanned: {serial_number}')
                    # Clear the input field
                    window['-INPUT-'].update('')

        # Update the output area with all scanned serial numbers
        window['-OUTPUT-'].update('\n'.join(scanned_serial_numbers))

    window.close()

if __name__ == '__main__':
    main()
