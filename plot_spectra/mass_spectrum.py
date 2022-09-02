'//////////////////////////////////////////////////////////////'
# Premagnet current - molecular mass calibration and mass spectrum
# plot maker.
# By Luis Gil Martín. 2022
'//////////////////////////////////////////////////////////////'

import os
import sys
import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit


'------FUNCTIONS------'

def check_lineskip(command):
    return command[:len(command)-1]


# Fit function f(x) = ax^2 + bx + c
def square_fit(x, a, b, c):

    return a*x*x + b*x + c


# Function that calibrates the magnet current - mass relation
def calibrate(file_in, file_out):

    # Open file with calibration data
    with open(file_in, 'r') as file:

        lines=file.readlines()

        x, y = np.array([]), np.array([])

        # Read lines (removing header)
        for line_count in range(1, len(lines)):

            line = lines[line_count]
            line.replace('\t', ' ') # Change tabs to spaces
            # Remove lineskip from strings
            line = line.strip().split()

            x = np.append(x, float(line[0]))
            y = np.append(y, float(line[1]))

        # Sort
        x, y = zip(*sorted(zip(x, y)))
        x = np.array(x)
        y = np.array(y)

        # Square the current data
        x = x*x

        # Fit to f(x) = ax^2 + bx + c
        popt, pcov = curve_fit(square_fit, x, y)
        a, b, c  = popt
        ua, ub, uc = np.sqrt(np.diag(pcov))

        # R^2 Pearson coefficient
        residuals = y - square_fit(x, *popt)
        ss_res = np.sum(residuals**2)
        ss_tot = np.sum((y-np.mean(y))**2)
        r_squared = 1 - (ss_res / ss_tot)


        # Plot fit
        ax = plt.subplot()
        ax.scatter(x, y)
        ax.plot(x, square_fit(x, a, b, c), c='#005AA0')

        ax.set_xlabel(r'$I^2 ~(A^2)$')
        ax.set_ylabel(r'$m ~(u)$')

        plt.show()

    # Write parameters in a separate file
    with open(file_out, 'w') as file:

        params_str = str(a) + ' ' + str(b) + ' ' + str(c) + ' ' + str(r_squared)
        file.write('Param_a   Param_b   Param_c   R^2\n')
        file.write(params_str)


# Function that extracts the data from the mass spectrum data file and the masses file
def read_raw_data(fdata, fmasses, a, b, c):

    # Read spectrum data from file
    with open(fdata, 'r') as file:

        lines=file.readlines()

        # Read only after DATA line
        DATA_index = lines.index('DATA\n')

        lines = lines[DATA_index+2:]

        x, y = np.array([]), np.array([])

        # Read lines (removing header)
        for line_count in range(1, len(lines)):

            line = lines[line_count]
            line.replace('\t', ' ') # Change tabs to spaces

            # Split by spaces and remove lineskips
            line = line.strip().split()

            x = np.append(x, float(line[0]))
            y = np.append(y, float(line[1]))

    # Sort
    x, y = zip(*sorted(zip(x, y)))
    x = np.array(x)
    y = np.array(y)

    y = np.abs(y)*10**9 # To positive nanoamps

    # Convert magnet current to mass and
    # and add tiny offset to Y values for logscale plot (FC current)
    x = x*x
    mass_exp = a*x*x + b*x + c
    i_fc = y + 10**-12

    # Read molecular masses and name for every substance in the given file
    with open(fmasses, 'r') as file:

        lines=file.readlines()

        mass_tab, name = np.array([]), np.array([])

        # Read lines (removing header)
        for line_count in range(1, len(lines)):

            line = lines[line_count]
            line.replace('\t', ' ') # Change tabs to spaces

            # Split by spaces and remove lineskips
            line = line.strip().split()

            mass_tab = np.append(mass_tab, float(line[0]))
            name = np.append(name, line[1])
    
    return mass_exp, i_fc, mass_tab, name


# Function that makes the mass spectrum plot
def make_plot(title, x, y, m, n):

    fig = plt.figure()
    ax = fig.add_subplot(111)

    # Plot spectrum
    ax.plot(x, y, lw=0.5, label='Data')
    # Plot vertical lines with identified substances
    ax.vlines(m, ymin=-1000, ymax=1000, colors='black', lw=0.3, linestyles='dashed')
    # Plot horizontal line in 1 nA reference
    ax.axhline(y=1, lw=0.3, ls='dashed', color = 'r')

    # Show name of every identified substance
    for i, x in enumerate(m):
        if n[i] == '?':
            plt.text(x-1, 1000, s='?', size=10, rotation=0, verticalalignment='center', color='red')
        else:
            plt.text(x-1, 1000, s=n[i], size=8, rotation=90, verticalalignment='center')

    # Format
    ax.set_yscale('log')
    ax.set_ylim(0.1,)

    ax.set_ylabel(r'$I_{FC} ~~[nA]$')
    ax.set_xlabel(r'Mass $m ~~[u]$')

    ax.set_xticks(m)
    plt.xticks(rotation=-45)

    # Title
    plt.suptitle(title, fontname = 'Arial Rounded MT Bold')

    plt.show()



'------MAIN------'

# Looks for magnet current - mass calibration parameters file.
# If not found, asks for the data to make the calibration and creates
# the parameters file.

print('Please enter the name of the directory (absolute path of folder): ')
dir = sys.stdin.readline()
dir = check_lineskip(dir)

param_fname = 'fit_param.asc'

print(os.listdir(dir))

if param_fname not in os.listdir(dir):
    print('Enter the name of the calibration file with extension: ')
    calib_fname = sys.stdin.readline()
    calib_fname = check_lineskip(calib_fname)

    calib_fname = dir + '\\' + calib_fname
    param_fname = dir + '\\' + param_fname

    calibrate(calib_fname, param_fname)
    
# Reads the fit parameters from the file
param_fname = 'fit_param.asc'
param_fname = dir + '\\' + param_fname
with open(param_fname, 'r') as file:

    next(file) # Skip first line with labels

    lines = file.readlines()
    line = lines[0]

    line = line.split()

    a, b, c = float(line[0]), float(line[1]), float(line[2])

# Asks for the spectral data file and the list of substances with masses

print('Enter the name of the data file with extension: ')
data_fname = sys.stdin.readline()
data_fname = check_lineskip(data_fname)

plot_title = data_fname.split('.', 1)[0] # Remove the file extension for the plot title

masses_fname = 'masses_' + data_fname

data_fname = dir + '\\' + data_fname
masses_fname = dir + '\\' + masses_fname

# Extracts the data from both files

mass_exp, i_fc, mass_tab, name = read_raw_data(data_fname, masses_fname, a, b, c)

# Makes a plot

make_plot(plot_title, mass_exp, i_fc, mass_tab, name)