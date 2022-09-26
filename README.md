# excel_read_test
Simple program that tests time needed for reading different excel files using pandas

## Usage
1) run <code>python main.py</code> for usual execution routine (using excel files created earlier in a *docs* directory)
2) run <code>python main.py -c</code> for extended execution routine (automatically creates all necessary excel files)

## Story
It was always very fascinating to find out pandas performance when it comes to reading excel files. One day I tried to solve this mistery by writing a simple test app.
I wanted to find out how read time depends on several parameters:
1) Number of sheets
2) Number of columns
3) Number of rows
4) Percentage of cells that are non-blank
As a solution, this script was born

## Method
The program creates a collection of diffetent excel files ranging from small single-sheet files up to large 10MB files with multiple sheets.
Then, it reads all these files into dataframes using pandas and it's standard xlsx library.
Finally, when we aquire all load times, a simple linear regression model is fitted and coefs are displayed.
