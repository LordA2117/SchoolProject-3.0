import PySimpleGUI as sg
from Fscan import checkClashes, viewFreeAndBusy, readWorkbook
from os import listdir
from sys import exit

sg.theme('reddit')


def setup() -> str:
    # Window layout
    layout = [
        [sg.Text('Select the folder containing the timetables')],
        [sg.In(size=(25, 1), enable_events=True,
               key='-FOLDER-'), sg.FolderBrowse()],
        [sg.Button('Ok'), sg.Button('Quit')]

    ]

    # Window
    window = sg.Window('Folder Select', layout=layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Quit':
            exit()

        if values['-FOLDER-'] == '':
            sg.popup('Please select a folder', title='Error')
            continue

        if event == 'Ok':
            if values['-FOLDER-'] != None:
                window.close()
                break

            else:
                sg.popup('Select a folder', title='Error')
                continue
    return values['-FOLDER-']


def main_window(list_of_files, folder) -> None:
    days_of_the_week = ['Monday', 'Tuesday',
                        'Wednesday', 'Thursday', 'Friday', 'Saturday']

    # layout
    layout = [
        [sg.Text('Timetables to compare')],
        [sg.Combo(list_of_files, size=(25, 1))],
        [sg.Combo(list_of_files, size=(25, 1))],
        [sg.Text('              ')],
        [sg.Listbox(days_of_the_week, size=(30, 5))],
        [sg.Radio('Check Clashes', 'Function'), sg.Radio(
            'Get free teachers', 'Function'), sg.Radio('Get Busy Teachers', 'Function')],
        [sg.Text('Enter Period'), sg.InputText()],
        [sg.Text('              ')],
        [sg.Button('Ok', size=(4, 1)), sg.Button('Quit', size=(4, 1))]
    ]

    window = sg.Window('App', layout=layout)

    # main loop
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Quit':
            break

        # fn1 = Check clashes, fn2 = get free teachers, fn3 = get busy teachers
        f1, f2, day, fn1, fn2, fn3, period = list(values.values())

        if f1 == f2:
            sg.popup('Please select 2 different timetables', title='Error')
            continue

        if event == 'Ok':
            try:
                if fn1 == True:
                    res = checkClashes(
                        f'{folder}/{f1}', f'{folder}/{f2}', str(day[0]))
            except KeyError:
                sg.popup(
                    'No periods on this day', title='Error')
                continue

            if fn2:
                try:
                    res = viewFreeAndBusy(
                        folder=folder, day=day[0], period_number=int(period), view_busy=False)
                except ValueError:
                    sg.popup('Enter a value for the period', title='Error')
                    continue
                except IndexError:
                    sg.popup(
                        'Enter a value for the period within the correct number of periods', title='Error')
                    continue
                except KeyError:
                    sg.popup(
                        'No periods on this day', title='Error')
                    continue

            if fn3:
                try:
                    res = viewFreeAndBusy(
                        folder=folder, day=day[0], period_number=int(period), view_busy=True)
                except ValueError:
                    sg.popup('Enter a correct value for the period',
                             title='Error')
                    continue
                except IndexError:
                    sg.popup(
                        'Enter a value for the period within the correct number of periods', title='Error')
                    continue
                except KeyError:
                    sg.popup(
                        'No periods on this day', title='Error')
                    continue

        if len(res) != 0:
            sg.popup('\n'.join(i.replace('.xlsx', '')
                     for i in res), title='Result')
        else:
            sg.popup('None', title='Result')


def main():
    k = setup()
    fileList = listdir(k)
    main_window(fileList, k)

# Scope: Compare timetable of one teacher to another for checking clashes. Most manipulations will be performed only on 2 timetables.


if __name__ == '__main__':
    main()
