import PySimpleGUI as sg
import pandas
import fpdf
import functions
# sg.theme("Dark1")
# Making UI labels
label_semester_no = sg.Text("Semester no : ", size=22,
                            text_color="Black", font=("ArialBlack", 16, "bold")
                            , justification="c")

label_total_students = sg.Text("No of Students : ", size=22,
                               text_color="Black", font=("ArialBlack", 15, "bold")
                               , justification="c")

label_total_subjects = sg.Text("No of Subjects : ", size=22,
                               text_color="Black", font=("ArialBlack", 15, "bold")
                               ,justification="c")

label_batch_number = sg.Text("Enter Batch No : ", size=22,
                               text_color="Black", font=("ArialBlack", 15, "bold")
                               ,justification="c")

# Making UI input boxes
inputbox1 = sg.InputText(tooltip="Semester no 1-8", key="no_of_sem"
                         , size=(20, 9), justification="c", font=("ArialBlack", 13, "bold"))

inputbox2 = sg.InputText(tooltip="Number of students in semester", key="no_of_std",
                         size=(20, 9),  justification="c", font=("ArialBlack", 13, "bold"))

inputbox3 = sg.InputText(tooltip="Number of subjects in semester", key="no_of_sub",
                         size=(20, 9),  justification="c", font=("ArialBlack", 13, "bold"))

inputbox4 = sg.InputText(tooltip="Batch number only (excluding Roll no)", key="batch_no",
                         size=(20, 9),  justification="c", font=("ArialBlack", 13, "bold"))

# making buttons
exit_button = sg.Button("Exit", key="exit", font=("ArialBlack", 11, "bold"))
next_button = sg.Button("Next", key="next", font=("ArialBlack", 11, "bold"))

# creating layout
unformatted_layout = [[label_semester_no], [inputbox1],
            [label_total_students], [inputbox2],
            [label_total_subjects], [inputbox3],
            [label_batch_number], [inputbox4],
            [exit_button, next_button]]

# formatting layout to center
layout2 = [[sg.VPush()],
           [sg.Push(), sg.Column(unformatted_layout, justification="c"), sg.Push()],
            [sg.VPush()]]

# creating window with above layout
window = sg.Window("College Result Maker", layout=layout2, size=(250, 300))

# storing pressed button in event and values inside input boxes in values
while True:
    event, values = window.read()
    match event:
        case "exit":
            break
        case "next":
            sem_no = values["no_of_sem"]
            sub_no = values["no_of_sub"]
            std_no = values["no_of_std"]
            batch_no = values["batch_no"]
            if sem_no == "" or sub_no == "" or std_no == "" or batch_no == "":
                sg.popup("Enter all Fields")
            else:
                if 1 > int(sem_no) > 8:
                    sg.popup("Enter Semester in range 1-8")
                else:
                    break
                if 1 > int(sub_no) > 6:
                    sg.popup("Enter Subjects in range 1-6")
                else:
                    break


            # functions.open_subject_window()

        case sg.WINDOW_CLOSED:
            break




window.close()
