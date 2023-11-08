import PySimpleGUI as sg

# Initialize data list to store input data
data = []

# Define the layout
layout = [
    [sg.Text('Enter input:'), sg.InputText(key='-INPUT-'), sg.Button('Add')],
    [sg.Table(values=data, headings=['Input Data'], key='-TABLE-', auto_size_columns=False,
              display_row_numbers=True, col_widths=[30])]
]

# Create the window
window = sg.Window('Input Table', layout)

# Event loop
while True:
    event, values = window.read()

    # Close the window if the 'X' button is clicked
    if event == sg.WIN_CLOSED:
        break

    # Append the input data to the table
    if event == 'Add':
        input_data = values['-INPUT-']
        if input_data:
            data.append([input_data])
            window['-TABLE-'].update(values=data)

# Close the window
window.close()
