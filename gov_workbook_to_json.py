from dataclasses import dataclass
from openpyxl import load_workbook
from itertools import product
from functools import partial
import pandas as pd

@dataclass
class GovWorkbookSpecs:    
    fileName: str
    sheetName: str
    valueStartRow: int
    maxRow = None
    maxCol = None

# def getValue(mergedBounds : list, row : int, col: int):
#     valueCoord = next((
#         (row_min, col_min) 
#         for (col_min, row_min, col_max, row_max) 
#         in mergedBounds 
#         if col_min <= col <= col_max and row_min <= row <= row_max
#         ), (row, col))
#     return valueCoord

def read(spec: GovWorkbookSpecs, mergedBounds):
    worksheet = pd.read_excel(spec.fileName, spec.sheetName)
    print(worksheet.head())


def readMergedbounds(spec : GovWorkbookSpecs):
    workbook = load_workbook(spec.fileName)
    worksheet = workbook[spec.sheetName]
    mergedBounds = [r.bounds for r in worksheet.merged_cells.ranges]
    return mergedBounds
    # valueGetter = partial(getValue, mergedBounds)

    # maxColumn = max([col for (_, _, col, _) in mergedBounds])

    # maxRow = next((i
    #     for i in range(spec.valueStartRow, worksheet.max_row)
    #     if worksheet.cell(i,1).value is None
    # ), worksheet.max_row)

    # return worksheet, valueGetter, maxRow, maxColumn

def buildPrefixes(spec : GovWorkbookSpecs, worksheet):
    pass

if __name__ == '__main__':
    spec = GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 1', 8)
    worksheet, valueGetter, spec.maxRow, spec.maxCol = read(spec)
    prefixes = buildPrefixes(spec, worksheet)
    