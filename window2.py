import PySimpleGUI as sg
import pandas
import fpdf
import main

total_subjects = int(main.sub_no)

text_box1 = sg.Text(f"Enter subjects names : ", size=22,
                   text_color="Blue", font=("ArialBlack", 11, "bold")
                   , justification="c")

exit_button = sg.Button("Exit", key="exit", font=("ArialBlack", 11, "bold"))
next_button = sg.Button("Next", key="next", font=("ArialBlack", 11, "bold"))

un_formatted_layout = [[text_box1]]

for subject in range(total_subjects):
    text_box = sg.Text(f"Subject : {subject+1}", size=22,
            text_color="Black", font=("ArialBlack", 15, "bold")
            , justification="c")
    input_box = sg.InputText(tooltip=f" Enter subject {subject+1} name", key=f"subject{subject}",
                         size=(20, 9),  justification="c", font=("ArialBlack", 13, "bold"))
    un_formatted_layout.append([[text_box], [input_box]])

un_formatted_layout.append([exit_button, next_button])
window = sg.Window("Input subjects name", layout=un_formatted_layout, size=(200,450))
subject_list=[]
while True:
    event, values = window.read()
    match event:
        case 'exit':
            break
        case 'next':
            for sub_no in range(total_subjects):
                subject = values[f"subject{sub_no}"]
                if subject.strip() == "":
                    sg.popup("Please Enter all Fields")
                else:
                    subject_list.append(subject)
            break
        case sg.WINDOW_CLOSED:
            break


window.close()
