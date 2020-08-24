# excel2txt

Convert Excel files to delimited text

## Synopsis

For usage, run `excel2txt --help`:

```
usage: excel2txt [-h] [-o str] [-d str] [-D] [-n] [--version] FILE [FILE ...]

Convert Excel files to delimited text

positional arguments:
  FILE                  Input Excel file(s)

optional arguments:
  -h, --help            show this help message and exit
  -o str, --outdir str  Output directory (default:
                        /Users/kyclark/work/python/excel2txt-py)
  -d str, --delimiter str
                        Delimiter for output file (default: )
  -D, --mkdirs          Create separate directories for output files (default:
                        False)
  -n, --normalize       Normalize headers (default: False)
  --version             show program's version number and exit
```

Given one or more Excel files as positional parameters, the program
will create an output text file in the given output directory (which 
defaults to the current working directory).

For example:

```
$ excel2txt tests/test1.xlsx
  1: tests/test1.xlsx
Done, see output in "/Users/kyclark/work/python/excel2txt-py".
```

Now you should have a file called "test1__sheet1.txt" in the current directory.
You could use the "csvchk" program to see the structure of this file:

```
$ csvchk test1__sheet1.txt
// ****** Record 1 ****** //
name          : Ed
rank          : Captain
serial_number : 12345
```

If you are processing multiple files, you might find the "--mkdirs" option useful to put all the sheets from each workbook into a separate directories:

```
$ ./excel2txt.py tests/*.xlsx --outdir out --mkdirs
  1: tests/test1.xlsx
  2: tests/test2.xlsx
Done, see output in "/Users/kyclark/work/python/excel2txt-py/out".
```

In the "out" directory, there will be "test1" and "test2" directories:

```
$ find out -type f
out/test1/test1__sheet1.txt
out/test2/test2__sheet1.txt
```

You can use the "--delimiter" option to change the output file delimiter.

## Column, file normalization

The "--normalize" option will alter the headers of each output file to lowercase values and remove non-alphanumeric characters or the underscore.
This will also break "CamelCase" values into "snake_case."

This same normalization will be used to create the output file names so as to avoid any possibility of creating output files with illegal or difficult characters.

## See also

csvkit, csvchk

## Author

Ken Youens-Clark <kyclark@gmail.com>
