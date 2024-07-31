#! python3
"""
    twig.py -- TeacherWIse [timetable] Generator
    
    A python script to generate Teacherwise timetable from Classwise timetable
    (and vice versa -- yet to be implemented).
    
    The classwise timetable is read from CLASSWISE sheet and the generated teacherwise
    timetable is saved in TEACHERWISE sheet of the same input workbook.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    Written by Sunil Sangwal (sunil.sangwal@gmail.com)
    Date written: 20-Apr-2022
    Last Modified: 18-May-2023
"""
import argparse
import re
import time
import os           # splitext()
import sys
import shutil       # copy file

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

# configuration variables before running the script

expand_names = False    # set this to True to write full names of teachers

# filename = 'C:\\Users\\acer\\Downloads\\CLASSWISE TIMETABLE 2022-23.xlsx'
# filename = 'C:\\Users\\acer\\Documents\\classwise-timetable.xlsx'
# output_filename = 'C:\\Users\\acer\\Documents\\TEACHERWISE TIMETABLE-tmp2.xlsx'

# utility functions

def escape_special_chars(c):
    if c == '\n':
        c = '\\n'
    elif c == '\t':
        c = '\\t'
    return c

def expand_days(days):
    """
        Parameter
            days : eg. "1-2, 3, 4-6"
        
        Returns:
            [1, 2, 3, 4, 5, 6]
    """
    ret = []
    if days.find(',') >= 0:
        groups = days.split(',')
    else:
        groups = [days]

    for days in groups:
        if days.find('-') >= 0:
            start_day, end_day = days.split('-')
            start_day = int(start_day)
            end_day = int(end_day)

            # swap if in reverse order: 6-4 is same as 4-6
            if end_day < start_day:
                start_day, end_day = end_day, start_day

            for i in range(start_day, end_day+1):
                ret.append(i)
        else:
            ret.append(int(days))
    return ret

def compress_days(days):
    """
        Parameter:
            days -- a list containing days in expanded form eg [1,2,3,5,6]
        
        Returns:
            a string of the form "1-3, 5-6"
    """
    days = sorted(days)
    ret = []
    start = end = 0
    for i in range(len(days) - 1):
        if days[i + 1] - days[i] == 1:
            continue
        else:
            end = i
            if start == end:
                s = f"{days[start]}"
            else:
                s = f"{days[start]}-{days[end]}"
            ret.append(s)
            start = i + 1

    end = i+1
    if start == end:
        ret.append(f"{days[start]}")
    else:
        ret.append(f"{days[start]}-{days[end]}")

    retval =  ", ".join(ret)
    return retval

# print(compress_days([1, 2, 3, 4, 5, 6]))
# exit(0)

def count_days(days):
    return len(set(expand_days(days)))

def count_periods(teacher, timetable):
    period_count = {}
    for period_info in timetable[teacher]:
        column, class_name, days, subject = period_info
        # print(period_info)
        days = expand_days(days)
        if column not in period_count:
            period_count[column] = []

        period_count[column].extend(days)   # combine two lists

    total_periods = 0
    
    for column in period_count:
        periods = len(set(period_count[column]))   # remove duplicates
        total_periods += periods

    return total_periods

# print(expand_days("1-2, 3, 4-6"))
# exit(0)

def load_teacher_names(workbook):
    # the sheet "TEACHERS" contains data about teacher in format
    # TEACHER-CODE ; FULLNAME
    sheet = workbook['TEACHERS']
    
    teacher_names = {}
    
    row = 2
    while True:
        teacher_code = sheet.cell(row, 1).value
        if teacher_code == None:
            break

        if teacher_code in teacher_names:
            # teacher code has been repeated
            raise Exception(f"Teacher code '{teacher_code}' has been used more than once. Modify TEACHERS sheet to remove the error.")

        teacher_names[teacher_code] = sheet.cell(row, 2).value # full name
        row += 1

    return teacher_names

def get_class_number(_class):
    return _class[:len(_class) - 1] # remove section (for example, 'A' from '10A')

def highlight_clashes(sheet, context):
    """
        reads teacherwise timetable and highlights possible clashes
        by prepending **CLASH** to the offending cell
    """
    SEPARATOR = context['SEPARATOR']
    CLASH_MARK = '**CLASH** '
    total_clashes = 0

    # format of line is "CLASS (1-3,5-6) SUBJECT", e.g., 10A (1-2, 4) MATH
    p = re.compile(r'^(?P<class_name>[\w]+)\s*\((?P<days>.*)\)\s*(?P<subject>[\w \-.]+)$')
    
    row = 2
    while True:
        if not sheet.cell(row=row, column=1).value:
            break

        for column in range(2, 10):
            content = sheet.cell(row, column).value
            # skip empty cells in class timetable with a warning
            if not content:
                # cells in teacherwise timetable could be empty; just skip them
                # warnings += 1
                # print(f"Warning: Cell {get_column_letter(column)}{row} of teacherwise timetable is empty.")
                continue
            # content = content.replace('\n', ';')
            # lines = content.split(";")
            lines = content.split(SEPARATOR) # SEPARATOR is "\n" or ;
            
            entry = {}

            for line in lines:
                line = line.strip()
                if line == "":
                    # skip empty lines
                    continue

                m = p.match(line)
                if not m:
                    print(f"\nWarning: Cell {get_column_letter(column)}{row} in 'Teacherwise' timetable has formatting issue.")
                    print("    >>> ", line)
                    # warnings += 1
                    continue
                class_name, days, subject = m.groups()
                subject = subject.strip()
                # try:
                days = expand_days(days)
                # except:
                #     print(f"\nERROR: (row={row}, column={column}) (Cell {get_column_letter(column)}{row}) in 'Teacherwise' timetable has formatting issue")
                #     # print(e)
                #     exit(1)
                
                """
                    Ex 1:
                        10A (1-2) MATH
                        10B (2-3) MATH
                    
                    is not a clash; but

                    Ex 2:
                        10A (1-2) MATH
                        9B (2-3) MATH

                    is a clash.

                    2: [10, 9]
                    In Example 2 above, 2nd period: classes 10 and 9 simultaneously is a clash
                """

                for day in days:
                    if not day in entry:
                        entry[day] = []
                    entry[day].append(get_class_number(class_name)+ '-' + subject)    # Eg., '10-SCI' (from 10A (1-6) SCI)
                    # the above code now ensures that the case "7A (1) PE, 7B (1-4) MATH" is marked as a clash

            # after all lines in a cell have been processed
            clash_days = []
            for day in entry:
                entry[day] = set(entry[day])    # remove duplicates
                if len(entry[day]) > 1:
                    # possible clash
                    clash_days.append(day)

            # if there are clashes, write them
            if len(clash_days) > 0:
                total_clashes += len(clash_days)
                # converts list [1, 2, 5] into a string
                clash_days = repr(clash_days)
                sheet.cell(row=row, column=column).value = CLASH_MARK + f"{clash_days}:\n" + sheet.cell(row=row, column=column).value

        row += 1

    return total_clashes

def clear_sheet(sheet):
    # clear the sheet before starting writing...
    row = 2
    while True:
        if not sheet.cell(row=row, column=1).value:
            # we have reacher EOF
            break

        for column in range(2, 11):
            sheet.cell(row=row, column=column).value = ""

        row += 1

    return

def generate_teacherwise(workbook, context):
    SEPARATOR = context['SEPARATOR']

    if "CLASSWISE" in workbook:
        print("Reading 'CLASSWISE' sheet... ", end='')
        input_sheet = workbook["CLASSWISE"]
        print("done.")
    else:
        print("Sheet 'CLASSWISE' not found. Reading active sheet instead... ")
        input_sheet = workbook.active

    # if names are to be replaced with full names for teachers,
    # then we must have 'TEACHERS' sheet in the input file
    teacher_names = {}
    if "TEACHERS" in book:
        print("Reading teacher details from 'TEACHERS' sheet... ", end='')
        teacher_names = load_teacher_names(workbook)
        print("done.")

    timetable = {}  # variable to hold teacherwise timetable

    print("Processing timetable ...")
    # p = re.compile(r'^(?P<subject>[\w -.]+)\s*\((?P<days>.*)\)\s*(?P<teacher>\w+)$') # format "SUBJECT (1-3,5-6) TEACHER"
    p = re.compile(r'^(?P<subject>[\w \-.]+)\s*\((?P<days>[1-6,\- ]+)\)\s*(?P<teacher>[A-Z]+)$')

    warnings = 0
    row = 2
    while True:
        class_name = input_sheet.cell(row, 1).value
        if not class_name:
            break       # we have reached the end of CLASSWISE sheet, so stop further processing

        periods_assigned = {}   # subjectwise keep track of how many periods have been assigned

        print(f"Class: {class_name}... ", end="")
        for column in range(2, 10):
            content = input_sheet.cell(row, column).value
            # skip empty cells in class timetable with a warning
            if not content:
                warnings += 1
                print(f"Warning: Cell {get_column_letter(column)}{row} is empty.")
                continue

            lines = content.split(SEPARATOR) # SEPARATOR is "\n" or ;
            
            days_assigned = []
            for line in lines:
                line = line.strip()
                if line == '' or line.startswith('#'):  # ignore empty lines and the ones starting with '#' -- used as comment
                    continue

                m = p.match(line)
                if m is None:   # no match
                    # print(f"\nWarning: (row={row}, column={column}) (Cell {get_column_letter(column)}{row}) has some formatting issue")
                    print(f"\nWarning: Cell {get_column_letter(column)}{row} in CLASSWISE sheet has some formatting issue.")
                    print("    >>> ", line)
                    warnings += 1
                    continue

                subject, days, teacher = m.groups()
                subject = subject.strip()
                days_assigned.extend(expand_days(days))

                if subject not in periods_assigned:
                    periods_assigned[subject] = count_days(days)
                else:
                    periods_assigned[subject] += count_days(days)

                if teacher not in timetable:
                    timetable[teacher] = []
                
                period = column                     # column denotes "period"
                timetable[teacher].append((period, class_name, days, subject))

            # check if all days in a particular period have been assigned
            # if row == 3 and column == 3:
            #     print(days_assigned)
            #     exit(1)

            if set(days_assigned) != set([1, 2, 3, 4, 5, 6]):
                warnings += 1
                print(f"\nWarning: not all days have been assigned in cell {get_column_letter(column)}{row}")


        # calculate the number of periods assigned to different subjects
        # num_periods_assigned = calculate_subject_periods(periods_assigned)

        # calculate the number of periods assigned to different subjects
        periods_assigned = sorted(periods_assigned.items())

        period_list = []
        num_periods_assigned = 0
        for subject, periods in periods_assigned:
            num_periods_assigned += periods
            period_list.append(f"{subject}: {periods}")

        period_list.append(f"TOTAL: {num_periods_assigned}")
        
        # finally, write the 'SUBJECT: #periods' for all subjects to their respective classes
        input_sheet.cell(row=row, column=10).value = ", ".join(period_list)

        print("done.")
        # process next class
        row += 1
        # end while loop

    ars = context['ARGS']
    if args.keepstamp:  # don't update time stamp on the original timetable
        pass
    else:
        # source timetable update time
        input_sheet.cell(row, 2).value = "Last updated on " + time.ctime()

    # count total periods for each teacher
    total_periods = {}

    for teacher in timetable:
        total_periods[teacher] = count_periods(teacher, timetable)

    # everything has been read into the timetable
    # now write back to a new work book

    if 'TEACHERWISE' in book:
        output_sheet = book['TEACHERWISE']
    else:
        print("Creating TEACHERWISE sheet... ", end='')
        output_sheet = book.create_sheet(title='TEACHERWISE', index=1)
        print("done.")

    # change styles
    # alignment = Alignment(horizontal='general',
    #     vertical='top',
    #     text_rotation=0,
    #     wrap_text=True,
    #     shrink_to_fit=False,
    #     indent=0)

    # Clear the TEACHERWISE sheet before writing
    clear_sheet(output_sheet)

    # writing the teacherwise timetable to the TEACHERWISE sheet
    header = ["Name", 1, 2, 3, 4, 5, 6, 7, 8, "Periods"]
    for column in range(2, len(header) + 1):
        output_sheet.cell(row=1, column=column).value = header[column - 1]

    unsorted_teachers = timetable.keys()
    sorted_teachers = []

    # check if there all teacher codes have associated full names
    # and add teacher if his/her whose fullname is not written in the TEACHERS sheet.
    for teacher in teacher_names:   # for every teacher code in timetable ...
        if teacher in unsorted_teachers:
            sorted_teachers.append(teacher)

    # start writing in 2nd row and then move to the following rows
    row = 2         

    for teacher in sorted_teachers:
        
        periods = timetable[teacher]
        
        # sheet.cell(row, 1).value = teacher # teacher code
        if expand_names and (teacher in teacher_names):
            output_sheet.cell(row, 1).value = teacher_names[teacher] # full teacher name
        else:
            output_sheet.cell(row, 1).value = teacher # abbreviation as has been used in classwise timetable

        # sort day-wise
        periods = sorted(periods, key=lambda x:x[2])

        for period in periods:
            (column, class_name, days, subject) = period
            class_name = class_name.strip()
            if output_sheet.cell(row, column).value:
                output_sheet.cell(row, column).value += f"{SEPARATOR}{class_name} ({days}) {subject}"
            else:
                output_sheet.cell(row, column).value = f"{class_name} ({days}) {subject}"

        output_sheet.cell(row, 10).value = total_periods[teacher]

        row += 1                    # move to the next row
        # end for

    if args.keepstamp:  # keep timestamp
        # don't change the time stamp
        pass
    else:
        # timestamp
        row = len(sorted_teachers) + 2
        # print(f"Row is {row}")
        output_sheet.cell(row, 2).value = "Generated on " + time.ctime()
    
    # done writing to the TEACHERWISE sheet

    return warnings
    # end generate_teacherwise()

def generate_classwise(workbook):
    # 'sheet' contains Teacherwise timetable

    # **TODO**:
    print("Warning: This feature is under development.")
    print("Generating Classwise from Teacherwise timetable... Done.")

    # we should never write to the CLASSWISE sheet as it may corrupt our original timetable;
    # instead we write to a CLASSWISE-GEN sheet.

    CW = 'CLASSWISE-GEN'

    if CW in workbook:
        output_sheet = workbook[CW]
    else:
        output_sheet = workbook.create_sheet(title=CW, index=0)

    input_sheet = workbook['TEACHERWISE']       # there must be TEACHERWISE sheet

    # set up loops and process
    context = {
        'input_sheet': input_sheet,
        'output_sheet': output_sheet,
               # _class ( days ) subject
        'p' : r'^(?P<_class>[\w -.]+)\s*\((?P<days>.*)\)\s*(?P<subject>\w+)$' # format "CLASS (1-3,5-6) SUBJECT"
    }
    process_sheet(process_teacherwise_timetable_line, context)


    warnings = 0
    return warnings

    # end generate_classwise(filename)

def process_sheet(callback, context):
    input_sheet = context['input_sheet']
    output_sheet = context['output_sheet']
    p = context['p']

    # end process_sheet()

def process_teacherwise_timetable_line(context):
    ...
    # end process_teacherwise_timetable_line()

if __name__ == '__main__':
    ##########################################################
    #
    # process commmand line arguments
    #
    #
    parser = argparse.ArgumentParser(prog='twig.py', description='Generates teacherwise (or classwise) timetable from classwise (or teacherwise) timetable.')
    parser.version = '1.0'
    parser.add_argument('-f', '--fullname', action='store_true', help='replace short names with full names')
    parser.add_argument('-v', '--version', action='store_true', help='display version information')
    parser.add_argument('-b', '--backup', action='store_true', help='generate backup file')
    parser.add_argument('-k', '--keepstamp', action='store_true', help='keep time stamp intact')
    parser.add_argument('-s', '--separator', action='store', help='newline separator; default is \\n')
    parser.add_argument('-c', '--classwise', action='store_true', help='generate classwise timetable from the teacherwise timetable')

    parser.add_argument('filename', type=str, action='store', nargs='?', help='file containing timetable')

    # args = parser.parse_args(['Timetable-2023-24-Handcrafted.xlsx', 'Timetable-2023-24-Handcrafted-Teacherwise.xlsx'])
    args = parser.parse_args()
    # print(args)
    if args.version:
        print("twig.py: version 1.0.0")
        exit(0)

    expand_names = args.fullname    # True or False; default = True

    if not args.filename:
        filename = 'Timetable.xlsx'
    else:
        filename = args.filename

    if not args.separator:
        SEPARATOR = "\n"    # multi-line separator
    else:
        SEPARATOR = args.separator
        if SEPARATOR == '\\n':
            SEPARATOR = '\n'
        print(f"Using Separator '{escape_special_chars(SEPARATOR)}' ...")

    

    startTime = time.time()

    DEBUG = False
    if DEBUG:
        filename = "Class-Wise(19-07-2023).xlsx"
        SEPARATOR = ';'

    # save file as backup
    if args.backup:
        print(f"Generating a backup of the timetable in {filename}... ", end='')
        formatted_time = time.strftime("%Y-%m-%d-%H%M%S", time.localtime(time.time()))
        base_name, extension = os.path.splitext(filename)
        shutil.copy(filename, f'backup-{base_name}-{formatted_time}.xlsx')
        print("done.")
    else:
        print(f"Skipping backup generation of the timetable in {filename}.")

    print(f"Reading CLASSWISE timetable from '{filename}'... ", end="")
    book = openpyxl.load_workbook(filename)
    print("done.")

    context = {
        'SEPARATOR' : SEPARATOR,
        'ARGS' : args
    }
    if args.classwise:
        warnings = generate_classwise(book, context)
    else:
        warnings = generate_teacherwise(book, context)

    # generate_free_teacher_report()

    endTime = time.time()
    # Highlight possible clashes
    context = {'SEPARATOR': SEPARATOR}
    teacherwise_sheet = book['TEACHERWISE']
    
    # total_clashes = -1
    total_clashes = highlight_clashes(teacherwise_sheet, context)

    print(f"Saving to TEACHERWISE sheet of '{filename}'... ", end="")
    book.save(filename)
    print("done.")

    print(f"Clashes: {total_clashes}")
    print(f"Warnings: {warnings}")

    print("Finished processing in %.3f seconds." % (endTime - startTime))
    print("Have a nice day!\n")
