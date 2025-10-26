# Date: 19-10-2025
# Author: saruccio.culmone@yahoo.it
#
# Extracting MSFS simulation variables into Excel or CSV tables from online
# documentation
#

import pandas as pd
import random
import time
import argparse
import os


PROG_NAME = "MSFS Docs To Var TABs"
PROG_DESCR = """
Download simulation variables fom documentation HTML pages
"""

class UrlToTable():
    """This class scans the URL list received as input and gathers all found
    HTML tables into a single DataFrame.
    The DataFrame could be used as is or saved in CSV or XLS format
    """
    def __init__(self, url:str, page:str=""):
        """
        """
        self.tables = pd.read_html(url)
        self.df = None

        # Foreach table, only colums
        # ['Simulation Variable', 'Description', 'Units'] will be
        # retained
        self.tabs = []
        for tab in self.tables:
            try:
                self.tabs.append(tab[['Simulation Variable', 'Description', 'Units']])
            except:
                pass
        # Concatenate all tables into one
        self.df = pd.concat(self.tabs, ignore_index=True)
        # Add topic column
        self.df['topic'] = page
        # Reoder columns
        new_cols_order = ['topic', 'Simulation Variable', 'Units', 'Description']
        self.df = self.df.reindex(columns=new_cols_order)
        self.df['Simulation Variable'] = self.df['Simulation Variable'].str.replace(' ', '_')
        self.df['Simulation Variable'] = self.df['Simulation Variable'].str.strip(' \n\t\v')


    def to_csv(self, filepath: str):
        sep=";"
        na_rep="NA"
        self.df.to_csv(filepath, sep=sep, na_rep=na_rep, index=False)

    def to_xls(self, filepath: str):
        na_rep="NA"
        self.df.to_excel(filepath, na_rep=na_rep, index=False)

sdk_url_pages = [
{'page': "Aircraft_AutopilotAssistant_Variables",
 'url': "https://docs.flightsimulator.com/html/Programming_Tools/SimVars/Aircraft_SimVars/Aircraft_AutopilotAssistant_Variables.htm"},
{'page': "Aircraft_Brake_Landing_Gear_Variables",
 'url': "https://docs.flightsimulator.com/html/Programming_Tools/SimVars/Aircraft_SimVars/Aircraft_Brake_Landing_Gear_Variables.htm"},
{'page': "Aircraft_Control_Variables",
 'url': "https://docs.flightsimulator.com/html/Programming_Tools/SimVars/Aircraft_SimVars/Aircraft_Control_Variables.htm"},
{'page': "Aircraft_Electrics_Variables",
 'url': "https://docs.flightsimulator.com/html/Programming_Tools/SimVars/Aircraft_SimVars/Aircraft_Electrics_Variables.htm"},
{'page': "Aircraft_Engine_Variables",
 'url': "https://docs.flightsimulator.com/html/Programming_Tools/SimVars/Aircraft_SimVars/Aircraft_Engine_Variables.htm"},
{'page': "Aircraft_FlightModel_Variables",
 'url': "https://docs.flightsimulator.com/html/Programming_Tools/SimVars/Aircraft_SimVars/Aircraft_FlightModel_Variables.htm"},
{'page': "Aircraft_Fuel_Variables",
 'url': "https://docs.flightsimulator.com/html/Programming_Tools/SimVars/Aircraft_SimVars/Aircraft_Fuel_Variables.htm"},
{'page': "Aircraft_Misc_Variables",
 'url': "https://docs.flightsimulator.com/html/Programming_Tools/SimVars/Aircraft_SimVars/Aircraft_Misc_Variables.htm"},
{'page': "Aircraft_RadioNavigation_Variables",
 'url': "https://docs.flightsimulator.com/html/Programming_Tools/SimVars/Aircraft_SimVars/Aircraft_RadioNavigation_Variables.htm"},
{'page': "Aircraft_System_Variables",
 'url': "https://docs.flightsimulator.com/html/Programming_Tools/SimVars/Aircraft_SimVars/Aircraft_System_Variables.htm"},
{'page': "Helicopter_Variables",
 'url': "https://docs.flightsimulator.com/html/Programming_Tools/SimVars/Helicopter_Variables.htm"},
{'page': "Camera_Variables",
 'url': "https://docs.flightsimulator.com/html/Programming_Tools/SimVars/Camera_Variables.htm"},
{'page': "Miscellaneous_Variables",
 'url': "https://docs.flightsimulator.com/html/Programming_Tools/SimVars/Miscellaneous_Variables.htm"},
{'page': "Services_Variables",
 'url': "https://docs.flightsimulator.com/html/Programming_Tools/SimVars/Services_Variables.htm"},
# {'page': ""
#  'url': ""},

]


# MAIN -----------------------------------------------------------------------
def main():
    print(f"{PROG_NAME} started")
    parser = argparse.ArgumentParser(prog=PROG_NAME, description=PROG_DESCR)
    parser.add_argument('outdir', help="Ouput directory for sim variable files.")
    args = parser.parse_args()

    if os.path.exists(args.outdir) is False:
        print(f"ERROR: Output folder not found '{args.outdir}'")
        return 1

    # Set output directory as working dir
    os.chdir(args.outdir)
    print(f"Download Start in '{args.outdir}'\n")

    dflist = []
    for pu in sdk_url_pages:
        print(f"{pu['page']}")
        pagetab = UrlToTable(pu['url'], pu['page'])
        dflist.append(pagetab.df)
        tabname = pu['page'] + ".xlsx"
        pagetab.to_xls(tabname)
        time.sleep(1 + random.uniform(0, 1.3))

    # Concatenate all table in a dingle one
    dff = pd.concat(dflist)
    dff = dff.reset_index(drop=True)
    dff_name = "msfs2020_sim_vars.xlsx"
    dff.to_excel(dff_name, na_rep="NA", index=False)
    print(f"All sim vars in {dff_name}")

    print("\nEnd")
    return 0

if __name__ == "__main__":
    main()
