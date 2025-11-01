# Date: 19-10-2025
# Author: saruccio.culmone@yahoo.it
#
# Reading MSFS simulation variables from active simulation saving values into
# Excel and CSV tables
#

import pandas as pd
from pprint import pprint
from SimConnect import SimConnect, AircraftRequests
import time
import os
from glob import glob
import argparse
import psutil
import sys

# To bypass the simulation running check, set this flag to True.
DEBUG = True

# Load variable files from a directory and for each variable try to retrieve
# the value set for the current airplane value.

PROG_NAME = "MSFS Sim Vars Scanner"
PROG_DESCR = """
Read simulation valiables from XLS or CSV files an try to retrieve their values
from current simulation.
Retrieved values are saved into the 'value' column of the output file.
"""

def is_msfs_running():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and 'FlightSimulator' in proc.info['name']:
            return True
    return False


def load_sim_vars_table(tabfile: str):
    """Load simulation variable names form filepath passed as parameter and return
    the corresponding Pandas DataFrame.
    Only CSV or XLSX file table are loaded.

    In error case None is returned
    """
    if os.path.exists(tabfile) is False:
        print(f"ERROR: File '{tabfile}' not found ")
        return None

    # Try to load the table file
    print(f"Loading: {tabfile}")
    filename, extention = os.path.splitext(tabfile)
    basename = os.path.basename(filename)
    ext = extention.lower().strip('.')
    df = None
    if ext == 'csv':
        df = pd.read_csv(tabfile, sep=';')
    elif (ext == 'xlsx') or (ext == 'xls'):
        df = pd.read_excel(tabfile)
    else:
        print("Not CVS or Excel file. Skipped")

    return df


def query_sim_var(aq: AircraftRequests, sim_var:str):
    try:
        var_reading = aq.get(sim_var)
    except:
        var_reading = 'na'
    time.sleep(0.2)
    return var_reading


# MAIN -----------------------------------------------------------------------
def main():
    # Main arguments
    parser = argparse.ArgumentParser(prog=PROG_NAME, description=PROG_DESCR)
    parser.add_argument('filetab', help="File containig a table of simulation variables")
    parser.add_argument('indexes', nargs='*', type=int, help="List of indexes for indexed variables")
    parser.add_argument('-s', '--simname', help="Name of current simulation", default="simvars")
    parser.add_argument('-o', '--outdir', help="Output directory", default='.')
    args = parser.parse_args()

    # Loading variables
    simvardf = load_sim_vars_table(args.filetab)
    if simvardf is None:
        print(f"ERROR: No Sim Variable files found in '{args.filetab}'")
        return 1
    # Add column for values
    simvardf['value'] = ""

    # Verify if MSFS is running
    if not DEBUG:
        if not is_msfs_running():
            msg = """Microsoft Flight Simulator is not currently running.
Please launch it, then rerun this script once the simulation has started."""
            print()
            print(msg)
            return 1

    # Now connect to the simulator
    sm = SimConnect()
    aq = AircraftRequests(sm, _time=200)

    # Manage indexes
    indexes = args.indexes
    if indexes == []:
        indexes = [1]

    # Get values from simulator
    num_rows = simvardf.shape[0]
    print(f"Reading {num_rows} variables")
    for idx, row in simvardf.iterrows():
        sim_var = row['Simulation Variable']
        index_pos = sim_var.find("index")
        if index_pos == -1:
            print(f"{idx:04}/{num_rows:04} - {sim_var}= ", end="")
            try:
                value = query_sim_var(aq, sim_var)
                print(f"{value}")
            except:
                value = 'na'
                print(f"EXCEPT: {value}")
        else:
            # Compose indexed sim variable name and get its value
            print(f"{idx:04}/{num_rows:04} - {sim_var}= ", end="")
            values = []
            none_counter = 0
            base_name_var = sim_var[:index_pos]
            for index in args.indexes:
                var_idx_name = f"{base_name_var}{index}"
                value = query_sim_var(aq, var_idx_name)
                if value is None:
                    none_counter += 1
                values.append(f"{value}")
            # if the list contains all None a single None is returned
            if none_counter == args.indexes:
                value = None
            else:
                value = values
            print(f"{value}")
        simvardf.at[idx, 'value'] = f"{value}"

    # Reorder columns
    new_cols_order = ['topic', 'Simulation Variable', 'value', 'Units', 'Description']
    simvardf = simvardf.reindex(columns=new_cols_order)

    # Create output namefile and save it
    basename = os.path.basename(args.filetab)
    noext_basename = os.path.splitext(basename)[0]
    xls_namefile = f"{noext_basename}-{args.simname}.xlsx"
    csv_namefile = f"{noext_basename}-{args.simname}.csv"
    if os.path.exists(args.outdir):
        outdir = args.outdir
    else:
        print(f"ERROR: Output path '{args.outdir}' not found. Default '.'")
        outdir = '.'
    xls_fpath = os.path.join(outdir, xls_namefile)
    csv_fpath = os.path.join(outdir, csv_namefile)

    # Save XLS and CSV tables
    print("Saving files:")
    print(f"- '{xls_namefile}'")
    print(f"- '{csv_namefile}'")

    try:
        simvardf.to_excel(xls_fpath, na_rep='na', index=False)
        simvardf.to_csv(csv_fpath, sep=';', na_rep='na', index=False)
        print(f"in '{outdir}'")
    except:
        print(f"Failed! Reason: {sys.exc_info()[1]}")

    sm.exit()
    return 0


if __name__ == "__main__":
    main()