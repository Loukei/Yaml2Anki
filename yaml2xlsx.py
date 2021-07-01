'''
引入沙拉查詞產生的yaml檔案，再轉換成anki可以接受的xlsx格式檔案
'''
from argparse import ArgumentParser, RawTextHelpFormatter
from pathlib import Path as plib_Path
import yaml
from yaml.loader import SafeLoader as yaml_SafeLoader
from openpyxl import Workbook

def ini_ArgParser()->ArgumentParser:
    '''
    Setup the argparser.
    ## Required parameter
    `-i`,`--input`: The input card yaml file path.
    '''
    parser = ArgumentParser(description=
    '''Turn your Yaml list file into Anki xlsx file.\nSee Spreadsheet Import Plus for Anki.\nhttps://github.com/HelenFoster/AnkiSpreadsheetImportPlus''',
    formatter_class=RawTextHelpFormatter)
    # require
    parser.add_argument("-i","--input",required=True,help="The input card file(.yaml) path.",type=plib_Path,dest="filename",metavar="File")
    # optional
    parser.add_argument("-o","--output",help="Assign the output file(.xlsx) path.",type=plib_Path,dest="outputfile",metavar="File")
    return parser

def yaml_to_anki_xlsx(infile:str,outfile:str)->None:
    '''
    Open infile(with .yaml format),parse the card structure,then trun into anki xlsx format.
    '''
    with open(infile,'r+',encoding='utf-8') as inputFile:
        cards:list = yaml.load(inputFile,Loader=yaml_SafeLoader) # turn yaml to list(dict)
        wb = Workbook() # create new xlsx workbook
        ws = wb.worksheets[0] # get worksheet 0
        ws['A1'] = "SpreadsheetImportPlus v1" # row 1 must be "SpreadsheetImportPlus v1"
        fields:list = list(cards[0].keys()) + ["_tags"]
        ws.append(fields)
        ws.append(["text"]*len(fields)) # all field use text format
        for card in cards:
            ws.append(list(card.values()))
        ws.insert_rows(4,1) # row 4 must be empty
        wb.save(outfile)
    return

def main():
    '''
    Setup the arg parse, then turn yaml file into anki xlsx file
    '''
    parser = ini_ArgParser()
    args = parser.parse_args()
    inputpath:plib_Path = args.filename
    if(inputpath.exists() and inputpath.suffix == ".yaml"):
        print(f"Receive arg: {inputpath}")
        outputfile = inputpath.with_suffix(".xlsx")
        print(f"Prepare output file: {outputfile}")
        yaml_to_anki_xlsx(infile=str(inputpath),outfile=str(outputfile))
    else:
        print("[ERROR] Check the file suffix or file is not exist. Input file must be .yaml")
        print(f"Receive arg: {inputpath}")
    pass

if __name__ == "__main__":
    main()
