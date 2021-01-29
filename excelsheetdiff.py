"""
    Author: Sujayyendhiren RS
    Description: An attempt at summarizing difference between versions of 
                 two excel files
"""

import json

import pandas as pd
from jsondiff import diff

class ExcelDiff():

    def __init__(self, filename1, filename2, sheetname):
        """
            Parameters:
                @filename1: First version of file
                @filename2: Edited version of file
                @sheetname: Assumption in this snippet is that both files
                           have same sheetname. It's easy to accomodate
                           different names by taking two input sheetnames.
        """
        # Read excel files
        self.exdf1 = pd.read_excel (filename1, sheet_name=sheetname)
        self.exdf2 = pd.read_excel (filename2, sheet_name=sheetname)

    def diff_output(self, outpath, write=False): 
        """
            Parameters: 
                @outpath: Path where output result must be written
        """

        ## Convert dataframes into json data
        json1 = json.loads(self.exdf1.to_json(orient='records'))
        json2 = json.loads(self.exdf2.to_json(orient='records'))

        ## Calculate diff
        diffoutput = diff(json1, json2, syntax='symmetric')

        ## Decide on howto output
        if write:
            with open(f"{outpath}/exceldiff.json", 'w') as fd:
                fd.write(json.dumps(diffoutput, indent=4))
        else:
            print(json.dumps(diffoutput, indent=4))


if __name__ == '__main__':

    xldiff = ExcelDiff('datainput1.xlsx', 'datainput2.xlsx', 'Sample')
    xldiff.diff_output('/tmp/')
