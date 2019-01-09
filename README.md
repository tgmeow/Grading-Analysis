# Analyze Grades
This Python 3 program sorts grades by grader to generate a box plot and calculate the 
CDF under the assumption that the data models an approximately normal 
hypergeometric distribution. This helps to identify biased graders by finding the probability of getting 
a more extreme sample from the overall (where extreme means having more samples greater than or equal 
to some cutoff).

## Introduction

What a hypergeometric distribution models: (from the scipy documentation)  
> Suppose we have a collection of 20 animals, of which 7 
are dogs. Then if we want to know the probability of finding a given number of dogs if we 
choose at random 12 of the 20 animals, we can initialize a frozen distribution and plot the 
probability mass function.

To learn more, see [Wikipedia](https://en.wikipedia.org/wiki/Hypergeometric_distribution) or the 
[SciPy documentation.](https://docs.scipy.org/doc/scipy-0.14.0/reference/generated/scipy.stats.hypergeom.html)

Although in this program, each grader is intended to be assigned a grading group for each project,
the math works under the assumption that each grader gets an approximately random sample-without-replacement of
all the grades. The cutoffs, determined by the 1st, 2nd, and 3rd quartiles of each grader, are used
to count the number of grades that are larger than or equal to the cutoff in both the total and the
sample. This number represents how many items of interest we have in total and in our selection. 
By combining these two numbers with the sample size and the total size, we approximately model a hypergeometric
distribution.

A few problems with this approach are that we assume that the graders get a random sample of the grades.
Since the graders do not actually get a random sample, any non-random ordering or grouping MAY
affect the results. Additionally, since we do not know the actual grades for a correctly graded 
assignment, a large number of incorrectly graded assignments may skew the results and hide possible biases.
  

## Requirements
- Python 3
- numpy
- pandas
- scipy
- matplotlib
- openpyxl
- Something to view/edit xlsx and csv files (preferably Microsoft Excel)

## Running
Run in terminal and look for any errors printed in the console.
```bash
python analyzeGrades.py
```

## Example Output and Analysis
Must have input files (dataFile and groupMapFile as specified in the config).  
```dataFile``` is a table of the grades by project for all students and groups. See 
[input_data_example1.csv](/examples/input_data_example1.csv) for an example.   
```groupMapFile``` is a table that maps grader name and project number to group number. See 
[input_groups_example1.csv](/examples/input_groups_example1.csv) for an example.   
```bash
$ python analyzeGrades.py
Done.
```
Upon successful exit (as shown above), output_prob_data.xlsx will be created. This file
contains **three spreadsheets**. See [output_analysis_data.xlsx](/examples/output_analysis_data.xlsx) for an example.
  
1. **Analysis Output**   
    - Table with whiskers, quartiles, (M N n x) for the CDF function, and the CDFs for
each quartile. The CDFs show the probability of getting a less extreme sample (sample with fewer grades >= the cutoff).
    - Conditional Formatting is utilized to highlight the numbers on the lower and higher end for each row.
    
2. **Raw Data By Grader**   
    - Table containing all the grades each grader gave. This can be useful for visual inspection and confirmation.
    - Conditional Formatting is utilized to help highlight any other trends such as a grader giving out free 100s.

3. **Box And Whisker**   
    - Table containing the low, mid, high values for the box chart and the actual box chart. The box chart
    can be used to visually compare the middle 50% for each grader and identify possible issues or trends.


## Configuration File
A config.json file is necessary to set colors, filenames, values, columns, and other things. 
It must be named config.json. I would recommend focusing on the variables listed here.
 

```dataFile, useAbove, dataHeaders['data'], groupMapFile, groupHeaders['group'], output['filename'], output['xlsxChart']['y_axis']```
 
### Options
Outlined here is the JSON file with variable names, expected type, and example values. 
Take a look at [config.json](/examples/config.json) for reference.
 
**Make sure the column and row labels (project# and group#) for the two files match!**
 
- ```dataFile : string``` - Filename of the input grades csv file  
example="input_data_example1.csv"
- ```convertBlanksTo : float``` - Treat blank values in the CSV as this value  
example=0
- ```useAbove : float``` - Ignore grades less than or equal to this value 
(helps remove no-submissions and other 'ungraded' items)  
example=0.2
- ```dataHeaders : Object``` - Specifies the columns of the ```dataFile``` to use
    - ```group : string``` - Name of the column that contains the group number of that row   
    example="Group"
    - ```data : Array<string>``` - List of the grades columns to use in the analysis  
    example=["Project #0","Project #1","Project #2","Project #3"]
- ```groupMapFile : string``` - Filename of the table to map grader and project to group  
example="input_groups_example1.csv"
- ```groupHeaders : Object``` - Specifies the columns of the ```groupMapFile``` to use
    - ```project : string``` - Name of the column that contains project number of that row   
    example="Project"
    - ```group : Array<string>``` - List of the group numbers to use. Ignores data for groups not in this list   
    example=["1", "2", "3", "4", "5", "6"]
- ```output : Object``` - Specifies the output file(s) and options for them
    - ```filename : string``` - Name of the output file  
    example="output_analysis_data"
    - ```format : string``` - Output data file format. **Currently only fully supports xlsx**
    example="xlsx"
    - ```xlsxChart : Object``` - Specifies the options for the chart in the xlsx file
        - ```y_axis : Object``` - Specifies the options for the chart's y-axis
            - ```min : float``` - Minimum for the y-axis range  
            example=0.6
            - ```max : float``` - Maximum for the y-axis range  
            example=1.05
            - ```unit : float``` - Interval between tics
        - ```width : float``` - width of the chart in the xlsx file  
        example=30
        - ```height : float``` - height of the chart in the xlsx file  
        example=15
    - ```pyplot : boolean``` - Save the box and whisker plot from matplotlib?
    (optional, additionally shows whiskers and outliers)  
    example=false