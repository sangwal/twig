# Introduction

This file is information about twig.py, which generates teacherwise timetable from classwise timetable in an Excel workbook.

# Setting up Environment

We first need to install Python in order to run `twig.py`. Please follow online tutorial on how to install Python on a Windows system.

After installing Python, we need to install `openpyxl` module. This is the module that
is used to manipulate Excel workbooks. To install `openpyxl` module, run the following
command in the CMD window.

`pip install openpyxl`

Successful run of this command will install the required module in Python.

# Writing the classwise timetable to an Excel Workbook

We have to store the timetable in CLASSWISE (in UPPERCASE) sheet of a workbook (say, `Timetable.xlsx`).

## Format of the Timetable.xlsx

The entries in the timetable has to follow a set procedure as given below.

Rename a sheet to `CLASSWISE`. This sheet will contain the classwise timetable.

In the first column, from the second cell we write the names of classes, e.g., `6A, 6B, 7A, ...` etc. The class name `6A` will appear in cell B1 and so on.

In the top most row we write the periods as `1, 2, 3, ..., 8` starting from the column `2`. We assume that there are only `8` periods per day. You can safely jump to the next paragraph. In case, your school has more periods per day, you may need to make suitable changes to `twig.py` itself, which is beyond the scope of this document.

Now you have to allot subjects to teachers on specific days and periods by filling in cells of the `CLASSWISE` sheet. We follow a set format `SUBJECT (DAYS) TEACHER` (e.g., `MATH (1-3, 5) SK`) format. Multiple lines can be inserted in a cell using `ALT+ENTER` combination. Once the timetable has been entered for all the classes, you are ready to run `twig.py`.

# Running twig.py

The following lines show the usage syntax of `twig.py`.

```
usage: twig.py [-h] [-f] [-v] [-s SEPARATOR] [-c] [filename]

Generates teacherwise (or classwise) timetable from classwise (or teacherwise)
timetable.

positional arguments:
  filename              file containing timetable

options:
  -h, --help            show this help message and exit
  -f, --fullname        replace short names with full names
  -v, --version         display version information
  -s SEPARATOR, --separator SEPARATOR
                        newline separator; default is \n
  -c, --classwise       generate classwise timetable from the teacherwise
                        timetable
```

To generate teacherwise timetable, we enter the command

`python twig.py Timetable.xlsx`

This will show clashes and warnings in the timetable on successful execution. The generated teacherwise timetable is stored in `TEACHERWISE` sheet of `Timetable.xlsx`. You may search for `**CLASH**` in the `TEACHERWISE` sheet which shows the days in which there are clashes. For example, if you see `**CLASH** [1, 4]` in a cell, it mean on day 1 and 4, there are clashes. You now need to modify your classwise timetable to remove the clashes and re-run `twig.py` to re-generate teacherwise timetable.

<div class="alert alert-info">You will need to close `Timetable.xlsx` in Excel before running `twig.py` again. </div> 

