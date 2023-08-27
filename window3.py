import PySimpleGUI as sg
import numpy as np
import pandas
import os
import functions
import main
import window2
from fpdf import FPDF
# importin files for sending email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# accessing main and window2 data
sem_no = main.sem_no
sub_no = main.sub_no
std_no = main.std_no
batch_no = main.batch_no
subject_list = window2.subject_list

layout = []
# turn variable later use in creating college result
turn = 1

# creating labels
sem_no_text = sg.Text("Semester no : " + sem_no,
                      text_color="Gold3", font=("ArialBlack", 15, "bold"))
sub_no_text = sg.Text("Subject no : " + sub_no,
                      text_color="Gold3", font=("ArialBlack", 15, "bold"))
std_no_text = sg.Text("Student no : " + std_no,
                      text_color="Gold3", font=("ArialBlack", 15, "bold"))


# adding labels to layout
layout.append([[sem_no_text, sub_no_text, std_no_text]])

# first row labels
first_row = [[sg.Text("Subject names", text_color="Black", font=("ArialBlack", 13, "bold")),
              sg.Text("   Obtained Marks", text_color="Black", font=("ArialBlack", 13, "bold")),
              sg.Text("   Credit Hours", text_color="Black", font=("ArialBlack", 13, "bold")),
              sg.Text("")]]

button_student_pdf = sg.Button("Create Student DMC", font=("ArialBlack", 11, "bold"), key="std_res")
button_update_clg_pdf = sg.Button("Update Total Results", font=("ArialBlack", 11, "bold"), key="std_clg_res")
button_create_pdf = sg.Button("  Create PDF Result  ", font=("ArialBlack", 11, "bold"), key="std_res_pdf")

layout.append(first_row)

# number of columns and rows
columns = 2
rows = int(main.sub_no)

col1 = sg.Column(
    [[sg.Text(subject_list[row], text_color="Black", font=("TimesNewRoman", 13, "bold"), justification="c")] for row in
     range(rows)])

col2 = sg.Column([[sg.Input(tooltip="Numbers only", pad=((0, 0), (0, 0)), size=(17, 20), justification="c",
                            key="Obtained marks", font=("ArialBlack", 14, "bold")) for column in range(columns)] for row
                  in range(rows)])

col3 = sg.Column([[button_student_pdf], [button_update_clg_pdf], [button_create_pdf]])

layout.append([[col1, col2, col3]])

batch_text = sg.Text("Enter Roll Number : ", text_color="Black", font=("ArialRoundedMTBold", 13, "bold"))
button_send_mail = sg.Button("Send Result via Email", font=("ArialBlack", 11, "bold"), key="send_pdf")
close_button = sg.Button("Close", font=("ArialBlack", 11, "bold"), key="close")
roll_no_input = sg.Input(tooltip="Enter Roll no number", size=(17, 20), justification="c", key="roll_no",
                         font=("ArialBlack", 14, "bold"))

layout.append([batch_text, roll_no_input, button_send_mail, close_button])

# cell1 = sg.Table(values=[[subject_list]])
window = sg.Window("Input data window", layout=layout)
while True:
    event, values = window.read()

    match event:
        case 'std_res':
            # creating lists
            gpa_list = []
            obtained_marks_lists = []
            credithours_list = []
            total_obtained_marks = 0
            # getting all values from input boxes from a dictionary of values to lists (obtained marks list and credithours list
            prefix = "Obtained marks"
            for i in range(rows * columns - 1):
                if i == 0:
                    key = f"{prefix}"
                    obtained_marks_lists.append(values.get(key, ""))
                if i % 2 == 1:
                    key = f"{prefix}{i}"
                    obtained_marks_lists.append(values.get(key, ""))
                if i % 2 == 0:
                    key = f"{prefix}{i}"
                    credithours_list.append(values.get(key, ""))
            try:
                # declaring others necessary lists
                total_credit_hours = 0
                total_marks_list = []
                grades_list = []
                avg_list = []

                # converting string lists to float list
                obtained_marks_lists = [float(item) for item in obtained_marks_lists]
                credithours_list = [float(item) for item in credithours_list]

                # check if inputted marks are in range 1-100
                if not all(0 <= item <= 100 for item in obtained_marks_lists):
                    sg.popup("Please Enter marks in range 1-100")
                else:
                    # check if GPA is 3 or 4
                    if not all(item in [3.0, 4.0] for item in credithours_list):
                        sg.popup("Please Enter 3 or 4 as credit hours")
                    else:
                        # check if Roll_no field is not empty
                        if str(values["roll_no"]).strip() == "":
                            sg.popup("Please Enter Roll no")
                        else:
                            # calling functions from functions file to calculate necessary data
                            total_obtained_marks = functions.total_obtained_marks(obtained_marks_lists)
                            total_marks = functions.total_marks()
                            total_avg = functions.total_average(total_marks, total_obtained_marks)
                            functions.calc_gpa(obtained_marks_lists, gpa_list)
                            gpa_list = [float(item) for item in gpa_list]
                            # assigning every obtained marks to average marks list as both are same
                            for obtained_marks in obtained_marks_lists:
                                avg_list.append(obtained_marks)
                            # finding total credit hours
                            for credit_hour in credithours_list:
                                total_credit_hours += credit_hour
                            no_of_subjects = len(subject_list)
                            total_marks_list = [100 for subject in range(no_of_subjects)]
                            SGPA = functions.calc_SGPA(gpa_list, credithours_list, no_of_subjects, total_credit_hours)
                            functions.calc_grades(obtained_marks_lists, grades_list)
                            SGPA = round(SGPA, 2)

                        # Excel file data
                        # creating excel file if not already exists
                        if not os.path.exists(f"files/{batch_no}{values['roll_no']}.csv"):
                            with open(f"files/{batch_no}{values['roll_no']}.csv", "w") as file:
                                pass
                        # appending total results cell at last of every column
                        subject_list.append("Total : ")
                        obtained_marks_lists.append(total_obtained_marks)
                        total_marks_list.append(total_marks)
                        credithours_list.append(total_credit_hours)
                        total_avg = round(total_avg, 2)
                        avg_list.append(total_avg)
                        gpa_list.append(SGPA)
                        data = [subject_list, obtained_marks_lists, total_marks_list, avg_list, credithours_list,
                                gpa_list,
                                grades_list]
                        # Transposing row into columns
                        dataframe = pandas.DataFrame(data).transpose()

                        # setting columns name
                        dataframe.columns = ["Subject", "Obtained Marks", "Total Marks", "Average Marks",
                                             "Credit Hours", "GPA",
                                             "Grades"]
                        # writing data
                        dataframe.to_csv(f"files/{batch_no}{values['roll_no']}.csv", index=False)
                        print(data)

            except ValueError:
                sg.popup("Please Full all fields (numerical values only)")

        case "std_clg_res":
            # creating lists
            gpa_list = []
            total_obtained_marks = 0
            obtained_marks_lists = []
            credithours_list = []
            #getting all values from input boxes from a dictionary
            # of values to lists (obtained marks list and total marks list
            prefix = "Obtained marks"
            for i in range(rows * columns - 1):
                if i == 0:
                    key = f"{prefix}"
                    obtained_marks_lists.append(values.get(key, ""))
                if i % 2 == 1:
                    key = f"{prefix}{i}"
                    obtained_marks_lists.append(values.get(key, ""))
                if i % 2 == 0:
                    key = f"{prefix}{i}"
                    credithours_list.append(values.get(key, ""))
            try:
                # declaring other necessary lists
                total_credit_hours = 0
                total_marks_list = []
                grades_list = []

                # converting string lists to float lists
                obtained_marks_lists = [float(item) for item in obtained_marks_lists]
                credithours_list = [float(item) for item in credithours_list]
                # check if inputted marks are in range 1-100
                if not all(0 <= item <= 100 for item in obtained_marks_lists):
                    sg.popup("Please Enter marks in range 1-100")
                else:
                    # check if gpa is 3 or 4
                    if not all(item in [3.0, 4.0] for item in credithours_list):
                        sg.popup("Please Enter 3 or 4 as credit hours")
                    else:
                        # check if roll no field is not empty
                        if str(values["roll_no"]).strip() == "":
                            sg.popup("Please Enter Roll no")
                        else:
                            # calling functions from functions file to calculate necessary data
                            total_obtained_marks = functions.total_obtained_marks(obtained_marks_lists)
                            total_marks = functions.total_marks()
                            total_avg = functions.total_average(total_marks, total_obtained_marks)
                            avg_list = obtained_marks_lists
                            functions.calc_gpa(obtained_marks_lists, gpa_list)
                            gpa_list = [float(item) for item in gpa_list]
                            #finding total credit hours
                            for credit_hour in credithours_list:
                                total_credit_hours += credit_hour
                            no_of_subjects = len(subject_list)
                            total_marks_list = [100 for subject in range(no_of_subjects)]
                            SGPA = functions.calc_SGPA(gpa_list, credithours_list, no_of_subjects, total_credit_hours)
                            functions.calc_grades(obtained_marks_lists, grades_list)
                            SGPA = round(SGPA, 2)

                            # creating roll no list and assigning bactchno+rollno as rollno
                            roll_no_list = []
                            for roll in range(int(std_no)):
                                if 0 <= roll <= 8:
                                    roll_no_list.append(f"{batch_no}0{roll + 1}")
                                else:
                                    roll_no_list.append(f"{batch_no}{roll+1}")
                            print(roll_no_list)
            except ValueError:
                sg.popup("Please Full all fields (numerical values only)")

            # reading file path if exists
            filepath = f"files/Semester-{sem_no}.csv"
            if os.path.exists(filepath):
                dataframe = pandas.read_csv(filepath)
                dataframe.set_index("Roll No", inplace=True)
            else:
                # creating and writing file if not exists
                if not os.path.exists(filepath):
                    with open(filepath, "w") as file:
                        pass
                # declaring an empty list in order to give the column names
                empty_list = []
                # assigning nan to every field in empty list
                for n in range(int(std_no)):
                    empty_list.append(np.NaN)
                # making data layout
                data = [roll_no_list, empty_list, empty_list, empty_list, empty_list]
                # Transposing row into columns
                dataframe = pandas.DataFrame(data).transpose()

                # setting columns name
                dataframe.columns = ["Roll No", "Obtained Marks", "Total Marks", "Average Marks",
                                     "SGPA"]
                dataframe.set_index("Roll No", inplace=True)
                # writing data

            index_no = str(f"{batch_no}{values['roll_no']}")
            index_no = int(index_no)
            test = int(values['roll_no'])

            # turn variable is used to solve a bug when file writes for first time it writes at end
            # this code fix that
            if turn == 3:
                dataframe.loc[index_no, "Obtained Marks"] = total_obtained_marks
                dataframe.loc[index_no, "Total Marks"] = total_marks
                dataframe.loc[index_no, "Average Marks"] = total_avg
                dataframe.loc[index_no, "SGPA"] = SGPA

            elif turn == 2:
                dataframe.loc[index_no1, "Obtained Marks"] = total_obtained_marks1
                dataframe.loc[index_no1, "Total Marks"] = total_marks1
                dataframe.loc[index_no1, "Average Marks"] = total_avg1
                dataframe.loc[index_no1, "SGPA"] = SGPA1

                dataframe.loc[index_no, "Obtained Marks"] = total_obtained_marks
                dataframe.loc[index_no, "Total Marks"] = total_marks
                dataframe.loc[index_no, "Average Marks"] = total_avg
                dataframe.loc[index_no, "SGPA"] = SGPA
                # deleting the last extra column
                dataframe = dataframe.iloc[:-1]

                turn = 3

            elif turn == 1:
                dataframe.loc[index_no, "Obtained Marks"] = ""
                dataframe.loc[index_no, "Total Marks"] = ""
                dataframe.loc[index_no, "Average Marks"] = ""
                dataframe.loc[index_no, "SGPA"] = ""
                total_obtained_marks1 = total_obtained_marks
                total_marks1 = total_marks
                total_avg1 = total_avg
                SGPA1 = SGPA
                index_no1 = int(index_no)
                turn = 2
            # writing dataframe or adding new data to final csv file
            dataframe.to_csv(f"files/Semester-{sem_no}.csv")
        case "std_res_pdf":
            # checking if the DMC csv file exists
            filepath = f"files/{batch_no}{values['roll_no']}.csv"
            if os.path.exists(filepath):
                pdf = FPDF(orientation="L", unit="mm", format="A4")
                pdf.add_page()
                filename = f"{batch_no}{values['roll_no']}"
                print(filename)
            # setting text at start
            pdf.set_font(size=16, family="Times", style="B")
            pdf.set_text_color(0, 0, 0)  # Set text color to black
            pdf.cell(w=0, h=14, txt="Government Post Graduate College Mansehra", align="C", ln=1)
            pdf.set_font(size=12, family="Times", style="B")
            pdf.cell(w=110, h=14, txt=f"Roll no:{filename}", align="l")
            pdf.line(10, 21, 200, 21)

            pdf.set_font(size=12, family="Times", style="B")
            pdf.cell(w=50, h=14, txt="Degree:BS Computer-Sciene", align="C", ln=1)
            df = pandas.read_csv(filepath)

            pdf.set_font(size=10, family="Times", style="B")
            pdf.cell(w=25, h=12, border=1, align='C', txt='Subject')
            pdf.cell(w=27, h=12, border=1, align='C', txt='Obtained Marks')
            pdf.cell(w=25, h=12, border=1, align='C', txt='Total Marks')
            pdf.cell(w=25, h=12, border=1, align='C', txt='Average Marks')
            pdf.cell(w=25, h=12, border=1, align='C', txt='Credit Hours')
            pdf.cell(w=25, h=12, border=1, align='C', txt='GPA')
            pdf.cell(w=25, h=12, border=1, align='C', txt='Grades', ln=1)

            # creating cells of lists using for loop
            for index, row in df.iterrows():
                pdf.set_font(size=10, family="Times", style="B")
                pdf.set_text_color(80, 80, 80)

                pdf.cell(w=25, h=12, border=1, align='C', txt=str(row['Subject']))
                pdf.cell(w=27, h=12, border=1, align='C', txt=str(row['Obtained Marks']))
                pdf.cell(w=25, h=12, border=1, align='C', txt=str(row['Total Marks']))
                pdf.cell(w=25, h=12, border=1, align='C', txt=str(row['Average Marks']))
                pdf.cell(w=25, h=12, border=1, align='C', txt=str(row['Credit Hours']))
                pdf.cell(w=25, h=12, border=1, align='C', txt=str(row['GPA']))
                pdf.cell(w=25, h=12, border=1, align='C', txt=str(row['Grades']), ln=1)

            # creating final results
            pdf.output(f"files/{filename}.pdf")
        case "send_pdf":
            text = sg.Text("Please Enter the reciever email")
            reciever_mail = sg.InputText(tooltip="Enter reciever email here", size=(17, 20), justification="c", key="rec_mail",
                         font=("ArialBlack", 14, "bold"))
            send_button = sg.Button("Send", font=("ArialBlack", 11, "bold"), key="send")
            close_button = sg.Button("Cancel", font=("ArialBlack", 11, "bold"), key="cancel")
            layoutpdf = [[text], [reciever_mail], [send_button, close_button]]
            window_send = sg.Window("Sending mail", layout=layoutpdf, size=(250,100))
            while(True):
                    event_send, values_send = window_send.read()
                    pdf_file_path = f"files/{batch_no}{values['roll_no']}.pdf"
                    pdf_file_name = f"{batch_no}{values['roll_no']}.pdf"
                    match(event_send):
                        case 'send':
                            if os.path.exists(f"files/{pdf_file_name}"):
                                if values_send['rec_mail'].endswith("@gmail.com"):
                                    sender_email = "ebadinfalltraders@gmail.com"
                                    receiver_email = values_send['rec_mail']
                                    subject = "Result File"
                                    body = "Please find the attached CSV file."

                                    smtp_server = "smtp.gmail.com"
                                    smtp_port = 587
                                    username = "ebadinfalltraders@gmail.com"
                                    password = os.getenv("PASS")

                                    msg = MIMEMultipart()
                                    msg["From"] = sender_email
                                    msg["To"] = receiver_email
                                    msg["Subject"] = subject

                                    msg.attach(MIMEText(body, "plain"))

                                    # Attach the PDF file

                                    with open(pdf_file_path, "rb") as attachment:
                                        part = MIMEApplication(attachment.read(), Name=f"Result{batch_no}{values['roll_no']}.pdf")
                                        part["Content-Disposition"] = f'attachment; filename="{pdf_file_path}"'
                                        msg.attach(part)

                                    try:
                                        server = smtplib.SMTP(smtp_server, smtp_port)
                                        server.starttls()
                                        server.login(username, password)

                                        text = msg.as_string()
                                        text = text.encode('utf8')
                                        server.sendmail(sender_email, receiver_email, text)
                                        print("Email successfully sent!")

                                    except smtplib.SMTPAuthenticationError as e:
                                        print("Authentication error:")
                                        print(e)

                                    except smtplib.SMTPException as e:
                                        print("SMTP error:")
                                        print(e)

                                    except Exception as e:
                                        print("Error:")
                                        print(e)

                                    finally:
                                        if server:
                                            server.quit()
                                else:
                                    sg.popup("Please Enter a valid gmail address (ending with @gmail.com)")
                                    break
                            else:
                                sg.popup("Create PDF File First")
                        case 'cancel':
                            break
                        case sg.WINDOW_CLOSED:
                            break


            window.close()

        case "close":
            break

        case sg.WINDOW_CLOSED:
            break

window.close()
