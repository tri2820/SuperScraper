from dataclasses import dataclass
from openpyxl import load_workbook
from openpyxl.workbook import workbook
from openpyxl.worksheet.worksheet import Worksheet
import pandas as pd
from pandas.core.frame import DataFrame
from functools import lru_cache

@dataclass(eq=True, frozen=True)
class GovWorkbookSpecs:    
    fileName: str
    sheetName: str
    # Pandas index
    valueStartRow: int

@lru_cache
def getWorksheet(spec: GovWorkbookSpecs):
    worksheet = pd.read_excel(spec.fileName, spec.sheetName, header=None)
    return worksheet

@lru_cache
def loadWorkbook(fileName):
    return load_workbook(fileName)

def readInfo(spec : GovWorkbookSpecs):
    workbook = loadWorkbook(spec.fileName)
    worksheet = workbook[spec.sheetName]
    mergedBounds = [r.bounds for r in worksheet.merged_cells.ranges]
    groupedCols = getGroupedCols(worksheet)
    return mergedBounds, groupedCols

def buildPrefixes(spec : GovWorkbookSpecs, worksheet : DataFrame):
    keyArea = worksheet.iloc[0:spec.valueStartRow]
    keyArea = keyArea.fillna('').astype(str).applymap(lambda s: s+'->')
    prefixes = keyArea.sum()
    return prefixes

def fill(worksheet, mergedBound):
    col_min, row_min, col_max, row_max = mergedBound
    area = worksheet.iloc[row_min-1:row_max,col_min-1:col_max]
    area.ffill(axis=1, inplace=True)
    area.ffill(axis=0, inplace=True)

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

def toWorksheet(spec):
    mergedBounds, groupedCols = readInfo(spec)
    worksheet = getWorksheet(spec)    
    worksheet.replace(r'\n|\\n',' ', regex=True, inplace=True) 
    for bound in mergedBounds: fill(worksheet, bound)
    prefixes = buildPrefixes(spec, worksheet)
    groupedCols.append(GroupedCol(worksheet.shape[1], worksheet.shape[1]))
    groupedCols = allGroupedCols(groupedCols)
    prefixes = addGroupedColsToPrefixes(prefixes, groupedCols)
    worksheet = worksheet[spec.valueStartRow:]
    worksheet.columns = prefixes
    return worksheet

if __name__ == '__main__':
    specs = [
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 1', 7),
        GovWorkbookSpecs('workbook/Annual fund-level superannuation statistics June 2020.xlsx', 'Table 3', 7)
    ]
    for spec in specs:
        worksheet = toWorksheet(spec)
        json = worksheet.to_json(orient='records')
        print(json)
    