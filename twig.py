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
    Last Modified: 30-Aug-2024
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

def get_formatted_time():
    # t = time.localtime()
    # return f"{t.tm_year}{t.tm_mon:02d}{t.tm_mday:02d}{t.tm_hour:02d}{t.tm_min:02d}{t.tm_sec:02d}"
    return "Last updated on " + time.ctime()

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

def load_teacher_details(workbook):
    # the sheet "TEACHERS" contains data about teacher in format
    # TEACHER-CODE ; FULLNAME
    sheet = workbook['TEACHERS']
    
    teacher_details = {}
    
    row = 2
    while True:
        teacher_code = sheet.cell(row, 1).value
        if teacher_code == None:
            break

        if teacher_code in teacher_details:
            # teacher code has been repeated
            raise Exception(f"Teacher code '{teacher_code}' has been used more than once. Modify TEACHERS sheet to remove the error.")

        for col in range(1, 6):
            if teacher_code not in teacher_details:
                teacher_details[teacher_code] = {}
            teacher_details[teacher_code][sheet.cell(1, col).value] = sheet.cell(row, col).value

        row += 1

    return teacher_details

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

        for column in range(1, 11):
            sheet.cell(row=row, column=column).value = ""

        row += 1
        if not sheet.cell(row=row, column=1).value:
            # we have reacher EOF
            break

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
                    print(f"Warning: Cell {get_column_letter(column)}{row} in CLASSWISE sheet has some formatting issue.")
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

            if set(days_assigned) != set([1, 2, 3, 4, 5, 6]):
                warnings += 1
                print(f"Warning: not all days have been assigned in cell {get_column_letter(column)}{row}.")


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
        formatted_time = get_formatted_time()
        # input_sheet.cell(row, 2).value = "Last updated on " + time.ctime()
        input_sheet.cell(row, 2).value = formatted_time

    # count total periods for each teacher
    total_periods = {}

    for teacher in timetable:
        total_periods[teacher] = count_periods(teacher, timetable)

    # everything has been read into the timetable
    # now write back to the TEACHERWISE worksheet

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

    timetable_teachers = timetable.keys()
    sorted_teachers = []

    # check if there all teacher codes have associated full names
    # and add teacher if his/her whose fullname is not written in the TEACHERS sheet.
    for teacher in teacher_names:   # for every teacher in TEACHERS sheet ...
        if teacher in timetable_teachers:   # keep teachers in the timetable and remove any extra teacher in TEACHERWISE
            sorted_teachers.append(teacher)

    # if a teacher is not in the TEACHERS sheet but appears in the timetable,
    # append him to the `sorted_teachers' as well so that his timetable can be generated
    # ensure every teacher in the timetable is has been appended
    for teacher in timetable_teachers:   # for every teachers in the classwise timetable ...
        if teacher not in sorted_teachers:
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
        # output_sheet.cell(row, 2).value = "Generated on " + time.ctime()
        output_sheet.cell(row, 2).value = get_formatted_time()
    
    # done writing to the TEACHERWISE sheet

    return warnings
    # end generate_teacherwise()

def generate_classwise(input_book, outfile):

    master_sheet = None

    input_sheet = input_book['CLASSWISE']

    try:
        output_book = openpyxl.load_workbook(outfile) # Workbook()
    except:
        # create an empty book
        output_book = openpyxl.Workbook()
        
    if 'MASTER' not in output_book:
        master_sheet = output_book.create_sheet('MASTER')

        master_sheet['A1'] = 'GSSS AMARPURA (FAZILKA)'
        master_sheet['A4'] = 'Mon'
        master_sheet['A5'] = 'Tue'
        master_sheet['A6'] = 'Wed'
        master_sheet['A7'] = 'Thu'
        master_sheet['A8'] = 'Fri'
        master_sheet['A9'] = 'Sat'

        for col in range(2, 10):
            master_sheet.cell(3, col).value = col - 1   # periods 1 - 8
    else:
        master_sheet = output_book['MASTER']
    
    # read the names of incharges from the TEACHERS sheet
    teachers_sheet = input_book['TEACHERS']
    class_incharge = {}
    
    # some settings!!
    FULLNAME_COLUMN = 2
    GENDER_COLUMN = 5
    INCHARGE_COLUMN = 6
    
    row = 2
    while True:
        teacher_code = teachers_sheet.cell(row, 1).value
        if teacher_code is None or teacher_code == '':
            break
        klass = teachers_sheet.cell(row, INCHARGE_COLUMN).value
        if klass is not None:
            class_incharge[klass] = teacher_code

        row += 1

    # copy/create templates for each class
    row = 2
    while True:
        klass = input_sheet.cell(row, 1).value
        if klass is None or klass == '':
            break

        # the following code effectively clears the sheet before writing any data

        if klass in output_book:
            # delete old one
            del output_book[klass]

        # create new by copying from the master
        print(f"creating sheet {klass} ...")
        copy = output_book.copy_worksheet(master_sheet)
        copy.title = klass

        row += 1

    output_book.save(outfile)

    # set up loops and process
    p = re.compile(r'^(?P<subject>[\w \-.]+)\s*\((?P<days>[1-6,\- ]+)\)\s*(?P<teacher>[A-Z]+)$')

    teacher_details = load_teacher_details(input_book)
    # print(teacher_details)
    
    warnings = 0
    row = 2
    while True:
        # print(f"Input Sheet: {input_sheet.title} row={row}")
        class_name = input_sheet.cell(row, 1).value
        if not class_name:
            break       # we have reached the end of CLASSWISE sheet, so stop further processing
        
        sheet_name = class_name
        # write class name
        # print(output_book.worksheets)
        output_book[sheet_name].cell(2, 1).value = f"Class: {class_name}"

        # write name of the class in-charge as well
        output_book[sheet_name].cell(2, 6).value = "Incharge: "
        # if class_name in class_incharge:
        #     output_book[sheet_name].cell(2, 6).value += class_incharge[class_name]
        if class_name in class_incharge:
            title = 'Smt.' if teacher_details[class_incharge[class_name]]['GENDER'] == 'f' else 'Sh.'
            output_book[sheet_name].cell(2, 6).value = f"Incharge: {title} {teacher_details[class_incharge[class_name]]['NAME']}"
        else:
            output_book[sheet_name].cell(2, 6).value = "Incharge: Sh./Smt. "
        

        for column in range(2, 10):
            content = input_sheet.cell(row, column).value
            # skip empty cells in class timetable with a warning
            if not content:
                warnings += 1
                print(f"Warning: Cell {get_column_letter(column)}{row} is empty.")
                continue

            lines = content.split(SEPARATOR) # SEPARATOR is "\n" or ;
            
            
            for line in lines:
                line = line.strip()
                if line == '' or line.startswith('#'):  # ignore empty lines and the ones starting with '#' -- used as comment
                    continue

                m = p.match(line)
                if m is None:   # no match
                    # print(f"\nWarning: (row={row}, column={column}) (Cell {get_column_letter(column)}{row}) has some formatting issue")
                    print(f"Warning: Cell {get_column_letter(column)}{row} in CLASSWISE sheet has some formatting issue.")
                    print("    >>> ", line)
                    warnings += 1
                    continue

                subject, days, teacher = m.groups()
                subject = subject.strip()
                days = expand_days(days)

                # copy data to the respective classwise sheet
                for day in days:
                    r = day + 3
                    # print(f"output_book[{sheet_name}].cell({r}, {column}).value += {subject} ({teacher})")
                    if output_book[sheet_name].cell(r, column).value is None:
                        output_book[sheet_name].cell(r, column).value = ''
                    output_book[sheet_name].cell(r, column).value += f"{subject} ({teacher})\n"
        
        row += 1
        # end of while True loop

    # get the time stamp from the CLASSWISE sheet
    timestamp = input_sheet.cell(row, 2).value
    for ws in output_book:
        ws.cell(10, 2).value = timestamp
    # save everything to the file
    output_book.save(outfile)

    # end generate_classwise(filename)

def process_sheet(callback, context):
    input_sheet = context['input_sheet']
    output_sheet = context['output_sheet']
    p = context['p']

    # end process_sheet()

def process_teacherwise_timetable_line(context):
    ...
    # end process_teacherwise_timetable_line()

def show_differences(base, current):
    """
        Shows difference between base and current timetables

        base    -- filename of the base timetable (.xlsx)
        current -- filename of the current timetable (.xlsx)
    """

    # load the two  workbooks
    wb_base = openpyxl.load_workbook(base)
    wb_current = openpyxl.load_workbook(current)

    # ensure that these sheets exist
    ws_base = wb_base['CLASSWISE']
    ws_current = wb_current['CLASSWISE']

    
    differences = 0
    row = 2
    while True:
        class_name = ws_base.cell(row, 1).value
        if class_name is None:
            break

        for col in range(1, 10):
            print(f"checking ({row}, {col}) ...")
            if ws_base.cell(row, col).value != ws_current.cell(row, col).value:
                differences += 1
                print(f"Difference in ({row}, {col})")

        row += 1

    return differences

if __name__ == '__main__':
    ##########################################################
    #
    # process command line arguments
    #
    #
    parser = argparse.ArgumentParser(prog='twig.py', description='Generates teacherwise (or classwise) timetable from classwise (or teacherwise) timetable.')
    parser.version = '1.0'

    parser.add_argument('-f', '--fullname', action='store_true', help='replace short names with full names')
    parser.add_argument('-k', '--keepstamp', action='store_true', help='keep time stamp intact')
    parser.add_argument('-s', '--separator', action='store', help='newline separator; default is \\n')
    parser.add_argument('-v', '--version', action='store_true', help='display version information')

    # Create a subparsers object
    subparsers = parser.add_subparsers(dest="command", help="Subcommands")

    # Subcommand 'teacherwise'
    tw_parser = subparsers.add_parser("teacherwise", help="Generate teacherwise timetable")
    # start_parser.add_argument("-p", "--port", type=int, default=8080, help="Port to run the service on")
    tw_parser.add_argument("infile", type=str, action="store", help="File containing classwise timetable")

    # Subcommand 'classwise'
    cw_parser = subparsers.add_parser("classwise", help="Generate classwise timetable")
    # cw_parser.add_argument("-f", "--force", action="store_true", help="Force stop the service")
    cw_parser.add_argument("infile", type=str, action="store", help="File containing classwise timetable")
    cw_parser.add_argument("outfile", type=str, action="store", help="File to write classwise timetable")
    
    diff_parser = subparsers.add_parser("diff", help="compare two timetables")
    diff_parser.add_argument("base", type=str, action="store", help="base classwise timetable to compare against")
    diff_parser.add_argument("current", type=str, action="store", help="current timetable to be compared against base timetable")

    # Parse the arguments
    args = parser.parse_args()

    # print(args)
    if args.version:
        print("twig.py: version 240830")
        exit(0)

    expand_names = args.fullname    # True or False; default = False

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

    # # save file as backup
    # if args.backup:
    #     print(f"Generating a backup of the timetable in {filename}... ", end='')
    #     formatted_time = time.strftime("%Y-%m-%d-%H%M%S", time.localtime(time.time()))
    #     base_name, extension = os.path.splitext(filename)
    #     shutil.copy(filename, f'backup-{base_name}-{formatted_time}.xlsx')
    #     print("done.")
    # else:
    #     print(f"Skipping backup generation of the timetable in {filename}.")

    
    context = {
        'SEPARATOR' : SEPARATOR,
        'ARGS' : args
    }

    if args.command in ['teacherwise', 'classwise']:
        if not args.infile:
            filename = 'Timetable.xlsx'
        else:
            filename = args.infile

        print(f"Reading CLASSWISE timetable from '{filename}'... ", end="")
        book = openpyxl.load_workbook(filename)
        print("done.")

    if args.command == 'classwise':
        warnings = generate_classwise(book, args.outfile)
    elif args.command == 'teacherwise':
        warnings = generate_teacherwise(book, context)
                # Highlight possible clashes
        context = {'SEPARATOR': SEPARATOR}
        teacherwise_sheet = book['TEACHERWISE']
        
        total_clashes = highlight_clashes(teacherwise_sheet, context)

        print(f"Saving to TEACHERWISE sheet of '{filename}'... ", end="")
        book.save(filename)
        print("done.")

        print(f"Clashes: {total_clashes}")
        print(f"Warnings: {warnings}")
    elif args.command == 'diff':
        base = args.base
        current = args.current

        # compare "base" with "current"
        print(f"Comparing '{base}' with '{current}' ..." )
        differences = show_differences(base, current)
        print(f"Found {differences} differences between {base} and {current}.")
    else:
        print("You should not have seen this line!")

    endTime = time.time()
    print("Finished processing in %.3f seconds." % (endTime - startTime))
    print("Have a nice day!\n")
