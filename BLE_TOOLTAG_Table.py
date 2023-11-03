import PySimpleGUI as sg

layout = [
    [sg.Text('Select an option:')],
    [sg.Combo(['Option 1', 'Option 2'], default_value='Option 1', key='combo')],
    [sg.Text(size=(40, 1), key='-OUTPUT-')]
]

window = sg.Window('Combo Example', layout)

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Cancel'):
        break
    selected_option = values['combo']
    window['-OUTPUT-'].update(f'You selected {selected_option}')

window.close()
