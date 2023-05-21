# twig.py

# Introduction

This file is information about `twig.py`, which generates teacherwise timetable from classwise timetable in an Excel workbook.

Disclaimer: The software is provided as is without any warranty, usefulness or suitability for any purpose. By using the software, you agree

## Getting the sources

Source for `twig.py` can be downloaded from [here (https://github.com/sangwal/twig)](https://github.com/sangwal/twig). A sample timetable in MS Excel xlsx format can also be downloaded. Please note that you should not change the format of the classwise timetable; otherwise, `twig.py` may not understand what is going on in the timetable.

# Setting up Environment

You first need to install Python in order to run `twig.py`. Please follow online tutorial on how to install Python on a Windows system.

After installing Python, install `openpyxl` module. This is the module that is used to manipulate Excel workbooks. To install `openpyxl` module, run the following
command in the CMD window.

`pip install openpyxl`

A successful run of this command will install the required module in Python.

# Writing the classwise timetable to an Excel Workbook

We have to store the timetable in CLASSWISE (in UPPERCASE) worksheet of a workbook (say, `Timetable.xlsx`).

## Format of the Timetable.xlsx

The entries in the timetable has to follow a set format as given below.

The classwise timetable should be in `CLASSWISE` worksheet. 
In the first column, from the second cell we write the names of classes, e.g., `6A, 6B, 7A, ...` etc. The class name `6A` will appear in cell B1 and so on.

In the top most row we write the periods as `1, 2, 3, ..., 8` starting from the column `2`. We assume that there are only `8` periods per day. You can safely jump to the next paragraph. In case, your school has more periods per day, you may need to make suitable changes to `twig.py` itself, which is beyond the scope of this document.

## Guidelines to enter classwise timetable

Now you have to allot subjects to teachers on specific days and periods by filling in cells of the `CLASSWISE` sheet. We follow a set format `SUBJECT (DAYS) TEACHER` (e.g., `MATH (1-3, 5) SK`) format. Multiple lines can be inserted in a cell using `ALT+ENTER` combination. The `SUBJECT` can contain alphabets and hyphens. The `TEACHER` is usually a two letter short name for a teacher. If you want to put two or more teachers in a class at the same time, you should enter one line for each teacher. Enter the whole classwise timetable following above guidelines. Now save the workbook by giving it a convenient name such as `Timetable.xlsx` and close the workbook.

Once the timetable has been entered for all the classes, you are ready to run `twig.py`.

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

Tip: Close the timetable worksheet before running `twig.py`.

Open the terminal window (run CMD in windows) and generate teacherwise timetable using the command:

`python twig.py Timetable.xlsx`

Its output will show clashes and warnings in the timetable on successful execution. The generated teacherwise timetable is stored in `TEACHERWISE` sheet of `Timetable.xlsx`. You may search for `**CLASH**` in the `TEACHERWISE` sheet which shows the days in which there are clashes. For example, if you see `**CLASH** [1, 4]` in a cell, it mean on day 1 and 4, there are clashes. You now need to modify your classwise timetable to remove the clashes. Now close `Timetable.xlsx` in Excel and re-run `twig.py` to re-generate teacherwise timetable.

Once all clashes have been removed, you may print the teacherwise timetable from `TEACHERWISE` sheet of `Timetable.xlsx` after suitable formatting, such as resizing rows and columns to properly show their contents.


# Thanks

In case you run into some issue, you may ask for help using my email [sunil.sangwal@gmail.com](mailto:sunil.sangwal@gmail.com).

If you find `twig.py` useful, please consider buying me a cup of coffee by paying to my UPI address sunil.sangwal@okhdfcbank (Indian users).

Thank you for using the software.
