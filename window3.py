import PySimpleGUI as sg
import pandas
import fpdf
import os

import functions
import main
import window2
from PIL import Image

# loading images and resizing them
# image = Image.open("C:\\Users\\infall\\Desktop\\image.png")
# image = image.resize((50,50))

# accessing main and window2 data
sem_no = main.sem_no
sub_no = main.sub_no
std_no = main.std_no
batch_no = main.batch_no
subject_list = window2.subject_list

layout = []

# creating labels
sem_no_text = sg.Text("Semester no : "+sem_no,
                               text_color="Gold3", font=("ArialBlack", 15, "bold"))
sub_no_text = sg.Text("Subject no : "+sub_no,
                               text_color="Gold3", font=("ArialBlack", 15, "bold"))
std_no_text = sg.Text("Student no : "+std_no,
                               text_color="Gold3", font=("ArialBlack", 15, "bold"))

# row1 = [[sem_no_text, sub_no_text, std_no_text]]
# centered_row1 = sg.Column(row1, element_justification="c")

# adding labels to layout
layout.append([[sem_no_text, sub_no_text, std_no_text]])

# number of columns and rows
columns = 2
rows = int(main.sub_no)

# first row labels
first_row = [[sg.Text("Subject names", text_color="Black", font=("ArialBlack", 13, "bold")),
              sg.Text("   Obtained Marks", text_color="Black", font=("ArialBlack", 13, "bold")),
              sg.Text("   Credit Hours", text_color="Black", font=("ArialBlack", 13, "bold")),
              sg.Text("")]]

button_student_pdf = sg.Button("Create Student DMC", font=("ArialBlack", 11, "bold"), key="std_res")
button_update_clg_pdf = sg.Button("Update Total Results", font=("ArialBlack", 11, "bold"), key="std_clg_res")
button_create_pdf = sg.Button("  Create PDF Result  ", font=("ArialBlack", 11, "bold"), key="std_res_pdf")

layout.append(first_row)

col1 = sg.Column([[sg.Text(subject_list[row], text_color="Black", font=("TimesNewRoman", 13, "bold"), justification="c")]for row in range(rows)])

col2 = sg.Column([[sg.Input(tooltip="Numbers only", pad=((0,0),(0,0)), size=(17, 20), justification="c", key="Obtained marks", font=("ArialBlack", 14, "bold")) for column in range(columns)] for row in range(rows)])

col3 = sg.Column([[button_student_pdf], [button_update_clg_pdf], [button_create_pdf]])

layout.append([[col1, col2, col3]])

batch_text = sg.Text("Enter Roll Number : ", text_color="Black", font=("ArialRoundedMTBold", 13, "bold"))
button_send_mail = sg.Button("Send Result via Email", font=("ArialBlack", 11, "bold"), key="send_pdf")
close_button = sg.Button("Close", font=("ArialBlack", 11, "bold"), key="close")
roll_no_input = sg.Input(tooltip="Enter Roll no number", size=(17, 20), justification="c", key="roll_no", font=("ArialBlack", 14, "bold"))

layout.append([batch_text, roll_no_input, button_send_mail, close_button])

# cell1 = sg.Table(values=[[subject_list]])
window = sg.Window("Input data window", layout=layout)
while True:
    event, values = window.read()
    print(event)
    print(values)   

    match event:
        case 'std_res':
            gpa_list=[]
            total_obtained_marks = 0
            obtained_marks_lists = []
            credithours_list = []
            prefix = "Obtained marks"
            for i in range(rows*columns-1):
                if(i==0):
                    key = f"{prefix}"
                    obtained_marks_lists.append(values.get(key, ""))
                if(i%2==1):
                    key=f"{prefix}{i}"
                    obtained_marks_lists.append(values.get(key, ""))
                if(i%2==0):
                    key=f"{prefix}{i}"
                    credithours_list.append(values.get(key, ""))
            try:
                total_credit_hours = 0
                total_marks_list = []
                grades_list = []
                avg_list = []
                obtained_marks_lists = [float(item) for item in obtained_marks_lists]
                credithours_list = [float(item) for item in credithours_list]
                if not all(0<=item<=100 for item in obtained_marks_lists):
                    sg.popup("Please Enter marks in range 1-100")
                else:
                    if not all(item in [3.0, 4.0] for item in credithours_list):
                        sg.popup("Please Enter 3 or 4 as credit hours")
                    else:
                        if str(values["roll_no"]).strip() == "":
                            sg.popup("Please Enter Roll no")
                        else:
                            total_obtained_marks = functions.total_obtained_marks(obtained_marks_lists)
                            total_marks = functions.total_marks()
                            total_avg = functions.total_average(total_marks, total_obtained_marks)
                            functions.calc_gpa(obtained_marks_lists, gpa_list)
                            for obtained_marks in obtained_marks_lists:
                                avg_list.append(obtained_marks)
                            gpa_list = [float(item) for item in gpa_list]
                            for credit_hour in credithours_list:
                                total_credit_hours += credit_hour
                            no_of_subjects = len(subject_list)
                            total_marks_list = [100 for subject in range(no_of_subjects)]
                            SGPA = functions.calc_SGPA(gpa_list, credithours_list, no_of_subjects, total_credit_hours)
                            functions.calc_grades(obtained_marks_lists, grades_list)
                            SGPA=round(SGPA, 2)
                            # print(no_of_subjects)
                            # print(total_credit_hours)
                            # print(gpa_list)
                            # print(total_avg)
                            # print(total_marks)
                            # print(total_obtained_marks)
                            # print(obtained_marks_lists)
                            # print(credithours_list)
                        # Excel file data
                        # creating excel file
                        if not os.path.exists(f"files/{batch_no}{values['roll_no']}.csv"):
                            with open(f"files/{batch_no}{values['roll_no']}.csv", "w") as file:
                                pass
                        subject_list.append("Total : ")
                        obtained_marks_lists.append(total_obtained_marks)
                        total_marks_list.append(total_marks)
                        credithours_list.append(total_credit_hours)
                        avg_list.append(total_avg)
                        gpa_list.append(SGPA)
                        data = [subject_list, obtained_marks_lists, total_marks_list, avg_list, credithours_list, gpa_list,
                                grades_list]
                        dataframe = pandas.DataFrame(data).transpose()

                        # setting columns name
                        dataframe.columns = ["Subject", "Obtained Marks", "Total Marks", "Average Marks", "Credit Hours", "GPA",
                                             "Grades"]
                        dataframe.to_csv(f"files/{batch_no}{values['roll_no']}.csv", index=False)
                        print(data)

            except ValueError:
                sg.popup("Please Full all fields (numerical values only)")

        case "std_clg_res":
            gpa_list = []
            total_obtained_marks = 0
            obtained_marks_lists = []
            credithours_list = []
            prefix = "Obtained marks"
            for i in range(rows * columns - 1):
                if (i == 0):
                    key = f"{prefix}"
                    obtained_marks_lists.append(values.get(key, ""))
                if (i % 2 == 1):
                    key = f"{prefix}{i}"
                    obtained_marks_lists.append(values.get(key, ""))
                if (i % 2 == 0):
                    key = f"{prefix}{i}"
                    credithours_list.append(values.get(key, ""))
            try:
                total_credit_hours = 0
                total_marks_list = []
                grades_list = []
                obtained_marks_lists = [float(item) for item in obtained_marks_lists]
                credithours_list = [float(item) for item in credithours_list]
                if not all(0 <= item <= 100 for item in obtained_marks_lists):
                    sg.popup("Please Enter marks in range 1-100")
                else:
                    if not all(item in [3.0, 4.0] for item in credithours_list):
                        sg.popup("Please Enter 3 or 4 as credit hours")
                    else:
                        if str(values["roll_no"]).strip() == "":
                            sg.popup("Please Enter Roll no")
                        else:
                            total_obtained_marks = functions.total_obtained_marks(obtained_marks_lists)
                            total_marks = functions.total_marks()
                            total_avg = functions.total_average(total_marks, total_obtained_marks)
                            avg_list = obtained_marks_lists
                            functions.calc_gpa(obtained_marks_lists, gpa_list)
                            gpa_list = [float(item) for item in gpa_list]
                            for credit_hour in credithours_list:
                                total_credit_hours += credit_hour
                            no_of_subjects = len(subject_list)
                            total_marks_list = [100 for subject in range(no_of_subjects)]
                            SGPA = functions.calc_SGPA(gpa_list, credithours_list, no_of_subjects, total_credit_hours)
                            functions.calc_grades(obtained_marks_lists, grades_list)
                            SGPA = round(SGPA, 2)

                            roll_no_list = []
                            for roll in range(int(std_no)):
                                if(0<=roll<=9):
                                    roll_no_list.append(f"{batch_no}0{roll+1}")
                                else:
                                    roll_no_list.append(f"{batch_no}{roll}")
                            print(roll_no_list)


            except ValueError:
                sg.popup("Please Full all fields (numerical values only)")

            if not os.path.exists(f"files/Semester-{sem_no}.csv"):
                with open(f"files/Semester-{sem_no}.csv", "w") as file:
                    pass

            data = {
                "Obtained Marks": []
            }
            dataframe = pandas.DataFrame(data, index=roll_no_list)
            dataframe.index.name("Roll no")
            dataframe.to_csv(f"files/Semester-{sem_no}.csv")

        case "exit":
            break

        case sg.WINDOW_CLOSED:
            break

window.close()
print(subject_list)

