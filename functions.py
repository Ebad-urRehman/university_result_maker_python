import os
import window2


def open_subject_window():
    os.system("window2.py")


subject_list = window2.subject_list
no_of_subjects = len(subject_list)


def total_obtained_marks(obtained_marks_list):
    total_obtained_marks = 0.0
    for obtained_marks in obtained_marks_list:
        print(obtained_marks)
        total_obtained_marks += float(obtained_marks)
    return total_obtained_marks


def total_marks():
    total_marks = no_of_subjects * 100
    return total_marks


def total_average(total_marks, total_obtained_marks):
    avg = (total_obtained_marks / total_marks) * 100
    return avg


def calc_gpa(obtained_marks_list, gpa_list):
    gpa_final = 0.0
    gpa_start = 0.0
    full_gpa = 0.0
    for obtained_marks in obtained_marks_list:
        if 80 <= obtained_marks <= 100:
            full_gpa = 4.0
            gpa_list.append(full_gpa)
        elif 70 <= obtained_marks <= 79:
            gpa_start = 3
            gpa_final = int(obtained_marks) % 10
            full_gpa = str(gpa_start) + "." + str(gpa_final)
            gpa_list.append(full_gpa)
        elif 60 <= obtained_marks <= 69:
            gpa_start = 2
            gpa_final = int(obtained_marks) % 10
            full_gpa = str(gpa_start) + "." + str(gpa_final)
            gpa_list.append(full_gpa)
        elif 50 <= obtained_marks <= 59:
            gpa_start = 1
            gpa_final = int(obtained_marks) % 10
            full_gpa = str(gpa_start) + "." + str(gpa_final)
            gpa_list.append(full_gpa)
        elif obtained_marks >= 60 >= obtained_marks:
            gpa_start = 2
            gpa_final = int(obtained_marks) % 10
            full_gpa = str(gpa_start) + "." + str(gpa_final)
            gpa_list.append(full_gpa)
        elif 0 <= obtained_marks <= 49:
            full_gpa = 0.0
            gpa_list.append(full_gpa)


def calc_SGPA(gpa_list, credit_hour_list, no_of_subjects, total_credit_hours):
    SGPA = 0.0
    for i in range(no_of_subjects):
        gpa = gpa_list[i] * credit_hour_list[i]
        SGPA += gpa

    SGPA = SGPA / total_credit_hours
    return SGPA


def calc_grades(obtained_marks_list, grades_list):
    for obtained_marks in obtained_marks_list:
        grades_dict = {
            (0, 49): "F",
            (50, 50): "D-",
            (51, 54): "D",
            (55, 59): "D+",
            (60, 60): "C-",
            (61, 64): "C",
            (65, 69): "C+",
            (70, 70): "B-",
            (71, 74): "B",
            (75, 79): "B+",
            (80, 85): "A-",
            (86, 90): "A",
            (91, 100): "A+"
        }
        for marks_range, grade in grades_dict.items():
            if marks_range[0] <= obtained_marks <= marks_range[1]:
                grades_list.append(grade)
