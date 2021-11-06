from dataclasses import dataclass, field
from openpyxl import load_workbook
from openpyxl.workbook import workbook
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd
from pandas.core.frame import DataFrame
from functools import lru_cache
import logging
logging.basicConfig(level=logging.DEBUG)
import os

@dataclass
class GovWorkbookSpecs:    
    fileName: str
    sheetName: str
    # Pandas index
    valueStartRow: int
    droppedRows: list[int] = field(default_factory=list)

@lru_cache
def getWorksheet(fileName, sheetName):
    logging.info(f'Loading {fileName}, {sheetName} (lite)')
    return pd.read_excel(fileName, sheetName, header=None)

@lru_cache
def loadWorkbook(fileName):
    logging.info(f'Loading {fileName}')
    return load_workbook(fileName)

def readInfo(spec : GovWorkbookSpecs):
    workbook = loadWorkbook(spec.fileName)
    worksheet = workbook[spec.sheetName]
    mergedBounds = [(a-1,b-1,c-1,d-1) for (a,b,c,d) in map(lambda r: r.bounds, worksheet.merged_cells.ranges)]
    groupedCols = getGroupedCols(worksheet)
    return mergedBounds, groupedCols

def buildPrefixes(spec : GovWorkbookSpecs, worksheet : DataFrame):
    keyArea = worksheet.iloc[0:spec.valueStartRow]
    prefixes = keyArea.fillna('').apply(lambda col: '->'.join(map(lambda x: str(x), filter(None,col))))
    return prefixes

def fill(worksheet, mergedBound):
    col_min, row_min, col_max, row_max = mergedBound
    forwardedValue = worksheet.iloc[row_min][col_min]
    worksheet.loc[row_min:row_max, col_min:col_max] = forwardedValue

@dataclass
class GroupedCol:
    min: int
    max: int

def merge(outlines: list[GroupedCol]):
    result : list[GroupedCol] = []
    for o in outlines:
        mismatched = result == [] or result[-1].max+1 < o.min
        if mismatched: result.append(o)
        else: result[-1] = GroupedCol(result[-1].min, o.max)
    return result

def getGroupedCols(ws: Worksheet):
    dim = ws.column_dimensions
    bounds = [GroupedCol(d.min, d.max) for d in dim.values() if d.outlineLevel == 1]
    groups = merge(bounds)
    groups = [GroupedCol(g.min-1, g.max-1) for g in groups]
    return groups

def addGroupedColsToPrefixes(prefixes, groupedCols):
    _prefixes = []
    for i,p in enumerate(prefixes):
        id = next((id
            for id,g in enumerate(groupedCols)
            if g.min<=i<=g.max
        ))
        _p = f"{id}=>{p}"
        _prefixes.append(_p)
    return _prefixes

def allGroupedCols(groupedCols : list, i = 0):
    if not groupedCols: return []
    g = groupedCols[0]
    return [GroupedCol(i,i) for i in range(i, g.min)] + [g] + allGroupedCols(groupedCols[1:], g.max+1)

def toWorksheet(spec : GovWorkbookSpecs):
    logging.info(f'Converting {spec.fileName}/{spec.sheetName}')

    _spec = spec
    mergedBounds, groupedCols = readInfo(_spec)
    worksheet : DataFrame = getWorksheet(_spec.fileName, _spec.sheetName)    
    worksheet.replace(r'\n|\\n',' ', regex=True, inplace=True) 
    for bound in mergedBounds: fill(worksheet, bound)
    worksheet.drop(spec.droppedRows, inplace=True)
    _spec.valueStartRow -= len(_spec.droppedRows)
    prefixes = buildPrefixes(_spec, worksheet)
    groupedCols.append(GroupedCol(worksheet.shape[1], worksheet.shape[1]))
    groupedCols = allGroupedCols(groupedCols)
    prefixes = addGroupedColsToPrefixes(prefixes, groupedCols)
    worksheet = worksheet[_spec.valueStartRow:]
    worksheet.columns = prefixes
    return worksheet

def write(fileName, content):
    os.makedirs(os.path.dirname(fileName), exist_ok=True)
    logging.info(f'Writing to {fileName} {hash(content)}')
    with open(fileName,'w') as f:
            f.write(json)
            f.close()

if __name__ == '__main__':
    specs = [
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 1', valueStartRow=7, droppedRows=[0,1,2,3,5,6]),
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 2', valueStartRow=7, droppedRows=[0,1,3,5,6]),
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 3', valueStartRow=7, droppedRows=[0,1,2,3,5,6]),
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 4', valueStartRow=7, droppedRows=[0,1,3,5,6]),
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 5', valueStartRow=7, droppedRows=[0,1,3,5,6]),
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 6', valueStartRow=9, droppedRows=[0,1,2,5,7,8]),
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 7', valueStartRow=8, droppedRows=[0,1,2,4,6,7]),
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 8', valueStartRow=7, droppedRows=[0,1,2,3,5,6]),
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 9', valueStartRow=8, droppedRows=[0,1,2,4,6,7]),
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 10', valueStartRow=8, droppedRows=[0,1,2,4,6,7]),
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 11', valueStartRow=8, droppedRows=[0,1,2,4,6,7]),
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 12', valueStartRow=8, droppedRows=[0,1,2,4,6,7]),
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 13', valueStartRow=8, droppedRows=[0,1,2,4,6,7])
    ]
    for spec in specs:
        worksheet = toWorksheet(spec)
        json = worksheet.to_json(orient='records')
        write(f'./result/{spec.sheetName}.json', json)