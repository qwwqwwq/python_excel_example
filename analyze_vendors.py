import csv
import os
import pandas
import numpy

NULL_REPRESENTATION = ''


def read_and_fix_excel(filename):
    """
    Fix excel spreadsheet so that


    | Administrative Assistant |                  |
    |                          | Onshore - Junior |

    becomes

    | Administrative Assistant | Onshore - Junior |
    """
    df = pandas.read_excel(filename)
    df.to_csv('temp.csv')

    with open('temp.csv') as f, open('temp_fixed.csv', 'w') as of:
        level_0_value = None
        writer = csv.writer(of)
        reader = csv.reader(f)
        writer.writerow(next(reader)) # write header row
        for line in csv.reader(f):
            if line[1] == NULL_REPRESENTATION:
                level_0_value = line[0]
                continue
            line[0] = level_0_value
            writer.writerow(line)

    df2 = pandas.read_csv('temp_fixed.csv', index_col=[0,1])
    df2.index = df2.index.set_names('Experience Level', level=1)

    os.remove('temp.csv') # delete temporary file
    os.remove('temp_fixed.csv')
    return df2


def clean_rate_data(rate):
    """
    Turn rates like "60+" and "100 - 110"
    into real numbers "60.0" and "105"
    """
    if isinstance(rate, float):
        return rate # values that are already floating-point numbers should be passed through
    try:
        if '+' in rate.strip(): # strip with no argument strips spaces
            return float(rate.strip().strip('+')) # float fuction returns a floating point number
        elif '-' in rate:
            numpy.mean([int(x.strip()) for x in rate.split('-')])
        else:
            return float(rate)
    except: # if bad data, return nan ("not a number")
        return numpy.nan


if __name__ == '__main__': # this means this section is executed when run from the command line
    df = read_and_fix_excel('vendorRates.xlsx') # fix up the excel spreadsheet (ugly sorry)
    stacked = df.stack() # stack lines up all the data into rows with 4 columns: job type, level, vendor, rate
    stacked = stacked.apply(clean_rate_data)
    avg_rate_by_vendor = stacked.groupby(level=2).mean() # level 2 is the vendor
    avg_rate_by_job = stacked.groupby(level=0).mean() # level 0 is the job type
    writer = pandas.ExcelWriter('analysis.xlsx')
    pandas.DataFrame({'Rate ($)': avg_rate_by_vendor}).to_excel(writer, 'Avg Rate by Vendor') # create one sheet
    pandas.DataFrame({'Rate ($)': avg_rate_by_job}).to_excel(writer, 'Avg Rate by Job') # create another sheet
    writer.save() # save excel file
