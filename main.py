import openpyxl as xl
import os
import copy
# import numpy as np

# Forests where all districts are inactive as well - double check these later
INACTIVE_FORESTS = [ 110108, 110105, 110112, 110405, 110608, 110611, 110620, ]
RESULTS_XL = 'results.xlsx'
INPUT_XL = r'Units Comparison as of 2023-11-09 3_58 PM ET (1).xlsx'
PALS = "PALS Prod"
DB_SHEETS = [ 'CARA Test', 'CARA Prod2', 'DataMart Test', 'DataMart Prod', ]
NOPE = 'NOPE'
WORKING = 'WORKING'
VERIFIED = 'Verified'
CHECK = 'Check Further'
UPDATE = 'Update These'
CLEAR_SHEETS = [ VERIFIED, CHECK, UPDATE ]

def check_worksheet(pals, wb, results):
    db = dict()
    for sheet in DB_SHEETS:
    # sheet = DB_SHEETS[0]
        ws = wb[sheet]
        db_sheet = results[sheet]
        db_sheet.delete_rows(1, 5000)
        db_sheet.append(["Unit Number", "Problem", "Current Name", "CORRECT Name", "isActive", "CORRECT isActive"])
        db_remove_me = [ ]
        for rownum in range(2, ws.max_row+1):
            u = ws.cell(rownum, 1).value
            n = ws.cell(rownum, 2).value
            a = True if ws.cell(rownum, 3).value == 1 else False
            db[u] = [n,a]
        unit_missing = pals.keys() - db.keys()
        for unit in db:
            if unit in pals.keys():
                pals_name = pals[unit][0]
                pals_active = pals[unit][1]
                db_name = db[unit][0]
                db_active = db[unit][1] == 1
                if pals_name == db_name and pals_active == db_active:
                    results[VERIFIED].append([unit, pals_name, pals_active])
                else:
                    info = [''] * 6
                    info[0] = unit
                    if pals_name != db_name:
                        info[1] = 'Name '
                        info[2] = db_name
                        info[3] = pals_name
                    if pals_active != db_active:
                        info[1] += 'Active'
                        info[4] = db_active
                        info[5] = pals_active
                    results[ws.title].append(info)
                #
                # del db[unit]
            else:
                db_remove_me.append([unit, "Not in PALS", db[unit][0], db[unit][1]])
                # del db[unit]
                # print(f"================================> { unit }  *db[unit] ")
        for line in db_remove_me:
            results[ws.title].append(line)
        for unit in unit_missing:
            results[ws.title].append([unit, "In PALS, not here", pals[unit][0], pals[unit][1]])
    results.save(RESULTS_XL)

# Generate a dict of ret[unit_num] = (name, active)
def pals_setup(ws):
    pals = dict()
    for rownum in range(6, ws.max_row + 1):
        for c in [2, 6, 8]: #unit num columns for PALS Prod
            unit = f'{ws.cell(rownum, c).value}'
            name = ws.cell(rownum, c+1).value
            active = ws.cell(rownum, 10).value is None
            if unit and unit in INACTIVE_FORESTS:
                active = False
            pals[unit] = [name, active]
    return pals

def main():
    wb = xl.load_workbook(INPUT_XL)
    results_xl = xl.load_workbook(RESULTS_XL)
    for s in CLEAR_SHEETS:
        results_xl[s].delete_rows(2,5000)
    pals = pals_setup(wb[PALS])
    check_worksheet(pals, wb, results_xl)
    print("RAN")
    results_xl.save(RESULTS_XL)

if __name__ == '__main__':
    main()
