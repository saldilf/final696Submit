"""
auspectra696.py
696 project

Handles the primary functions
"""
# !/usr/bin/env python3


from __future__ import print_function
import sys
from argparse import ArgumentParser
import numpy as np
import matplotlib.pyplot as plt
import os
import numpy as np
import matplotlib.pyplot as plt
import xlrd
import pandas as pd
import collections as cl
import six

SUCCESS = 0
INVALID_DATA = 1
IO_ERROR = 2

__author__ = 'salwan'

DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')
DEFAULT_DATA_FILE_NAME = os.path.join(DATA_DIR, "LongerTest.xlsx")


def warning(*objs):
    """Writes a message to stderr."""
    print("WARNING: ", *objs, file=sys.stderr)


def parse_cmdline(argv):
    """
    Returns the parsed argument list and return code.
    `argv` is a list of arguments, or `None` for ``sys.argv[1:]``.
    """
    if argv is None:
        argv = sys.argv[1:]

    # initialize the parser object:
    parser = ArgumentParser(
        description='Reads data from xlsx file, analyzes it, plots it, and reports key parameters. '
                    'All program output is user-controlled'
                    '')
    parser.add_argument("-w", "--workbook", help="The location (directory and file name) of the Excel file with "
                                                 " The default file is {} ".format(DEFAULT_DATA_FILE_NAME),
                        default=DEFAULT_DATA_FILE_NAME)
    parser.add_argument("-p", "--plot", help="Plot data from Excel file (default is true). Goes to 'data' directory",
                        action='store_true')
    parser.add_argument("-n", "--normalize", help="Normalize data to max absorbance (default is true).",
                        action='store_true')
    parser.add_argument("-t", "--table",
                        help="Tabular summary of SampleID, Amax, lMax, size (default is true). Goes to 'data' directory",
                        action='store_true')
    args = None
    try:
        args = parser.parse_args(argv)
        # args.wb_data = xlrd.open_workbook(input("Enter a file name (LongerTest.xlsx for testing purposes):  "))
    except IOError as e:
        warning("Problems reading file:", e)
        parser.print_help()
        return args, IO_ERROR
    return vars(args), SUCCESS  # make sample input have a list of expected args and compare


def norm(rawAbs, Amax, x):
    absoNorm = [x / Amax for x in rawAbs]  # normalizes to Amax
    maxAbsAfterNorm = max(absoNorm)
    return absoNorm


def data_analysisNorm(data_file):
    # read data from xlsx
    wb_data = xlrd.open_workbook(data_file)
    sheet1 = wb_data.sheet_by_index(0)
    r = sheet1.nrows
    c = sheet1.ncols
    xLimLflt = sheet1.cell(1, 0).value  # must be less than 400
    xLimL = int(xLimLflt)

    # Upper limit
    xLimUflt = sheet1.cell(r - 1, 0).value
    xLimU = int(xLimUflt)

    # dict for extracted or calculated values from data
    data = cl.OrderedDict({'Sample ID': [],
                           'lambdaMax (nm)': [],
                           'Amax': [],
                           'Size (nm)': []
                           })
    for x in range(3, c):

        # 299 is the number of data points minus 2
        lambdas = sheet1.col_values(colx=0, start_rowx=(xLimU - 299) - xLimL, end_rowx=r)  # x-vals (wavelength)
        abso = sheet1.col_values(colx=x, start_rowx=(xLimU - 299) - xLimL, end_rowx=r)  # y-vals (absorbance)

        # 400 is the lambda we start plotting at
        sID = sheet1.cell(0, x).value  # name of each column
        lMax = abso.index(max(abso)) + 400  # find max lambda for a column
        Amax = max(abso)  # get max Abs to norm against it

        absoNorm = norm(abso, Amax, x)

        size = -0.02111514 * (lMax ** 2.0) + 24.6 * (lMax) - 7065.
        # J. Phys. Chem. C 2007, 111, 14664-14669

        # if lambdaMax outisde of this range then 1) particles aggregated and 2)outside of correlation
        if 518 < lMax < 570:
            data['Size (nm)'].append(int(size))

        else:
            data['Size (nm)'].append('>100')

        # added extracted/calculated items into dict 'data'
        data['Sample ID'].append(sID)
        data['lambdaMax (nm)'].append(lMax)
        data['Amax'].append(Amax)

        # plot each column (cycle in loop)
        plt.plot(lambdas, absoNorm, linewidth=2, label=sheet1.cell(0, x).value)
        axes = plt.gca()
        box = axes.get_position()
        axes.set_position([box.x0, box.y0, box.width * 0.983, box.height])
        plt.xlabel('Wavelength (nm)')
        plt.ylabel('Absorbance (Normalized to Amax)')

    plt.legend(loc='center left', bbox_to_anchor=(1, 0.5), fancybox=True, shadow=True)
    axes.set_xlim([400, 700])
    axes.set_ylim([0, 1.5])
    plt.savefig('data/AuPlotNorm.png', format='png', dpi=500)
    return data


def data_analysis(data_file, plot=True):
    # read data from xlsx
    wb_data = xlrd.open_workbook(data_file)
    sheet1 = wb_data.sheet_by_index(0)
    r = sheet1.nrows
    c = sheet1.ncols
    xLimLflt = sheet1.cell(1, 0).value  # must be less than 400
    xLimL = int(xLimLflt)

    # Upper limit
    xLimUflt = sheet1.cell(r - 1, 0).value
    xLimU = int(xLimUflt)

    # dict for extracted or calculated values from data
    data = cl.OrderedDict({'Sample ID': [],
                           'lambdaMax (nm)': [],
                           'Amax': [],
                           'Size (nm)': []
                           })
    for x in range(3, c):

        # 299 is the number of data points minus 2
        lambdas = sheet1.col_values(colx=0, start_rowx=(xLimU - 299) - xLimL, end_rowx=r)  # x-vals (wavelength)
        abso = sheet1.col_values(colx=x, start_rowx=(xLimU - 299) - xLimL, end_rowx=r)  # y-vals (absorbance)

        # 400 is the lambda we start plotting at
        sID = sheet1.cell(0, x).value  # name of each column
        lMax = abso.index(max(abso)) + 400  # find max lambda for a column
        Amax = max(abso)  # get max Abs to norm against it

        size = -0.02111514 * (lMax ** 2.0) + 24.6 * (lMax) - 7065.
        # J. Phys. Chem. C 2007, 111, 14664-14669

        # if lambdaMax outisde of this range then 1) particles aggregated and 2)outside of correlation
        if 518 < lMax < 570:
            data['Size (nm)'].append(int(size))

        else:
            data['Size (nm)'].append('>100')

        # added extracted/calculated items into dict 'data'
        data['Sample ID'].append(sID)
        data['lambdaMax (nm)'].append(lMax)
        data['Amax'].append(Amax)

        # plot each column (cycle in loop)
        if plot is True:
            plt.plot(lambdas, abso, linewidth=2, label=sheet1.cell(0, x).value)
            axes = plt.gca()
            box = axes.get_position()
            axes.set_position([box.x0, box.y0, box.width * 0.983, box.height])
            plt.xlabel('Wavelength (nm)')
            plt.ylabel('Absorbance')
            plt.legend(loc='center left', bbox_to_anchor=(1, 0.5), fancybox=True, shadow=True)
            axes.set_xlim([400, 700])
            axes.set_ylim([0, Amax + 0.05])

    # plt.show()
    plt.savefig('data/AuPlot.png', format='png', dpi=500)
    # plt.savefig('data/AuPlot_Correct.png', format='png', dpi=1000)

    # return data['Amax']
    return data


# output data frame as nice table
def render_mpl_table(data, col_width=4.0, row_height=0.625, font_size=14,
                     header_color='#40466e', row_colors=['#f1f1f2', 'w'], edge_color='w',
                     bbox=[0, 0, 1, 1], header_columns=0,
                     ax=None, **kwargs):
    if ax is None:
        size = (np.array(data.shape[::-1]) + np.array([0, 1])) * np.array([col_width, row_height])
        fig, ax = plt.subplots(figsize=size)
        ax.axis('off')

    mpl_table = ax.table(cellText=data.values, bbox=bbox, colLabels=data.columns, **kwargs)

    mpl_table.auto_set_font_size(False)
    mpl_table.set_fontsize(font_size)

    for k, cell in six.iteritems(mpl_table._cells):
        cell.set_edgecolor(edge_color)
        if k[0] == 0 or k[1] < header_columns:
            cell.set_text_props(weight='bold', color='w')
            cell.set_facecolor(header_color)
        else:
            cell.set_facecolor(row_colors[k[0] % len(row_colors)])

    plt.savefig('data/DataTable.png', format='png', dpi=250)
    return ax


def main(argv=None):
    args, ret = parse_cmdline(argv)
    if ret != SUCCESS:
        return ret

        # decide program output based on user input
    if args['plot'] is True:
        if args['normalize'] is True:
            if args['table'] is True:
                print("You've generated a normalized plot and a table")
                data = data_analysisNorm(args['workbook'])
                df = pd.DataFrame(data)
                render_mpl_table(df, header_columns=0, col_width=3.0)
            else:
                print("You've generated a normalized plot without a table")
                data_analysisNorm(args['workbook'])
        else:
            if args['table'] is True:
                print("You've generated a non-normalized plot and a table")
                data = data_analysis(args['workbook'])
                df = pd.DataFrame(data)
                render_mpl_table(df, header_columns=0, col_width=3.0)

            else:
                print("You've generated a non-normalized plot and no table")
                data_analysis(args['workbook'])
    else:
        if args['table'] is True:
            print("You've generated a table and no plot")
            data = data_analysis(args['workbook'], plot=False)
            df = pd.DataFrame(data)
            render_mpl_table(df, header_columns=0, col_width=3.0)

        else:
            print("Default output: No table nor plot generated in data directory.")
            print("Run the script with -h to see all of the available user input options.")

    return SUCCESS  # success


if __name__ == "__main__":
    status = main(sys.argv[1:])
    sys.exit(status)
