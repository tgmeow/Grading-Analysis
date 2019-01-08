#!/usr/bin/python

# AUTHOR: Tiger Mou
# EDITED: 12/27/2018
# DESCRIPTION: Reads in a CSV file of grades and config file
#              creates a box and whisker plot and hypergeometric distribution analysis

import csv
import json
import sys
import itertools
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
import numpy as np
import pandas as pd
from scipy.stats import hypergeom
import matplotlib.pyplot as plt


def read_config(filename):
    """
    Read in config json file
    :param filename: name of the config file
    :returns: a dictionary with the config
    """
    try:
        with open(filename, 'r') as json_config:
            config = json.load(json_config)
    except EnvironmentError:
        sys.exit('EXITING... Failed to open config file.')
    return config


def read_groups(filename, columns):
    """
    Read in group map file
    (although this COULD be avoided under the assumption that each grader moves on to the
    next group after each project)
    :param filename: name of the group map file
    :param columns: array of the expected columns
    :returns: pandas table of the group map
    """
    try:
        groups_map = pd.read_csv(filename, sep=',', usecols=columns, index_col=0)
    except ValueError:
        sys.exit('EXITING... Columns specified in config not found in file %s' % filename)
    return groups_map


def read_data(filename, columns, convert):
    """
    Read in the grades data from file
    :param filename: name of the file
    :param columns: array of the expected columns
    :param convert: dictionary of converter functions
    :return: pandas table of the data
    """
    try:
        data = pd.read_csv(filename, sep=',', usecols=columns, converters=convert)
    except ValueError:
        sys.exit('EXITING... Columns specified in config not found in file %s' % filename)
    return data


def load_files():
    """
    Loads the required files from disk and returns them
    :return: config as a dictionary, groups_map as a pandas table, data as a pandas table
    """
    # Read in config file
    config = read_config('config.json')

    # Expected columns
    cols = config['groupHeaders']['group']
    cols = cols + [config['groupHeaders']['project']]

    # Read in groups map
    groups_map = read_groups(config['groupMapFile'], cols)

    # Expected columns
    cols = config['dataHeaders']['data']
    group = config['dataHeaders']['group']

    # Function to convert percents to floats in the 'data' columns
    def perc2float(x):
        if x == "":
            return config['convertBlanksTo']
        return float(x.strip('%')) / 100

    # turn function into dictionary from header to function
    convert = {header: perc2float for header in cols}
    cols = cols + [group]

    # Read in data file
    data = read_data(config['dataFile'], cols, convert)
    return config, groups_map, data


def main():
    config, groups_map, data = load_files()

    # Convert the 'data' df into a 2d array of grades given by each grader
    grades = {}
    combined = "ALL"
    data_cols = config['dataHeaders']['data']
    group_col = config['dataHeaders']['group']
    use_threshold = config['useAbove']
    for index1, row in data.iterrows():
        group_num = str(int(row.at[group_col]))
        for index2, item in row.filter(items=data_cols).iteritems():
            grader = groups_map.at[index2, group_num]
            # print(group_num, "\t", index2, "\t", grader)
            if item > use_threshold:
                grades.setdefault(grader, []).append(item)
                grades.setdefault(combined, []).append(item)

    # Create box plot
    plt.figure(figsize=(12, 4))
    grades_values = list(grades.values())
    grades_keys = list(grades.keys())
    box_values = plt.boxplot(grades_values, labels=grades_keys)
    # Retrieve data from the box plot
    res = {key: [v.get_ydata() for v in value] for key, value in box_values.items()}

    # EXTRACT TABLE DATA IF WANTED
    whiskers_min = [min(item) for item in res['whiskers'][::2]]  # Lower Whisker
    whiskers_Q1 = [max(item) for item in res['whiskers'][::2]]  # Q1
    whiskers_Q3 = [min(item) for item in res['whiskers'][1::2]]  # Q3
    whiskers_max = [max(item) for item in res['whiskers'][1::2]]  # Upper Whisker
    medians = [item[0] for item in res['medians']]  # Q2

    # ALTERNATIVE METHODS, EQUIVALENT DATA
    # caps_lower = [item[0] for item in res['caps'][::2]] # Lower Whisker
    # caps_upper = [item[0] for item in res['caps'][1::2]] # Upper Whisker
    # boxes_lower = [min(item) for item in res['boxes']] # Q1
    # boxes_upper = [max(item) for item in res['boxes']] # Q3

    # print(res['fliers']) # outliers
    # print(res['means']) # EMPTY ARRAY!

    # Format and display box plot
    plt.grid(axis='y')
    plt.yticks(np.arange(use_threshold, 1.1, step=0.05))
    # TODO TRY CATCH
    plt.savefig('output_box_plot.png', dpi=200)
    # TODO CONFIG FLAG(S)
    # plt.show()

    # Calculate hypergeometric probabilities
    h_cuts = [whiskers_Q1, medians, whiskers_Q3]
    M = len(grades[combined])  # total overall
    N = [len(graded) for graded in grades_values]  # number the grader completed
    # CALCULATING n using >= cut value
    n = [[sum([1 for item in grades[combined] if item >= val]) for val in cut]
         for cut in h_cuts]  # number of total satisfying
    # CALCULATING x using >= cut value
    x = [[sum([1 for item in gr if item >= val])
          for val, gr in zip(cut, grades_values)]
         for cut in h_cuts]  # number of selected satisfying

    h_cuts_probs = [[hypergeom.cdf(c_x, M, c_n, c_N)
                     for c_x, c_n, c_N in zip(cut_x, cut_n, N)]
                    for cut, cut_n, cut_x in zip(h_cuts, n, x)]

    data_out = [
        ["Labels"] + grades_keys,
        ["whisker_min"] + whiskers_min,
        ["Q1"] + whiskers_Q1,
        ["Q2"] + medians,
        ["Q3"] + whiskers_Q3,
        ["whisker_max"] + whiskers_max,
        ["M"] + [M for _ in range(len(N))],
        ["N"] + N,
        ["n_Q1"] + n[0],
        ["n_Q2"] + n[1],
        ["n_Q3"] + n[2],
        ["x_Q1"] + x[0],
        ["x_Q2"] + x[1],
        ["x_Q3"] + x[2],
        ["Fx(x)_Q1"] + h_cuts_probs[0],
        ["Fx(x)_Q2"] + h_cuts_probs[1],
        ["Fx(x)_Q3"] + h_cuts_probs[2]
    ]

    # TODO TRY CATCH?
    output_fname = 'output_prob_data'
    if config['output']['format'] == 'xlsx':
        wb = Workbook()
        ws1 = wb.active
        ws1.title = 'Analysis Output'
        for row in data_out:
            ws1.append(row)
        rule_color_grad = ColorScaleRule(start_type='percentile', start_value=0, start_color='3333FF',
                              mid_type='percentile', mid_value=50, mid_color='FFFFFF',
                              end_type='percentile', end_value=100, end_color='FF3333')

        # Color the Q1, Q2, Q3 and the 3 CDFs, 1 based index
        rows = ['3', '4', '5', '15', '16', '17']
        last_col = get_column_letter(len(data_out[0]))

        for r in rows:
            cond_range = 'B'+r+':'+last_col+r
            ws1.conditional_formatting.add(cond_range, rule_color_grad)

        # Print Raw Grades by Grader data for inspection
        ws2 = wb.create_sheet(title='Raw Data By Grader')
        ws2.append(grades_keys)
        for row in [*itertools.zip_longest(*grades_values)]:
            ws2.append(row)

        wb.save(filename=output_fname+'.xlsx')

    else:
        with open(output_fname+'.csv', 'w+') as csv_outfile:
            csv_w = csv.writer(csv_outfile, delimiter=',',lineterminator='\n')
            csv_w.writerows(data_out)



if __name__ == '__main__':
    print('Starting program...')
    main()
    print('Done.')
