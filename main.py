import openpyxl as xl
import os
import copy

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

def check_worksheet(pals, ws, results):
    db = dict()
    ws.delete_rows(1, 5000)
    ws.append(["Unit Number", "Problem", "Current Name", "isActive", "Correct Name", "Correct isActive"])
    for rownum in range(2, ws.max_row+1):
        db[ws.cell(rownum, 1)] = [ws.cell(rownum, 2), ws.cell(rownum, 3)]
    for unit in db.keys():
        if unit in pals.keys:
            pals_name = pals[unit][1]
            pals_active = pals[unit][2]
            db_name = db[unit][1]
            db_active = db[unit][2]
            if pals_name == db_name and pals_active == db_active:
                results[VERIFIED].append(unit, pals_name, pals_active)
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
                ws.append(*info)
            del pals[unit]
        else:
            print(f"=====> { unit } { db[unit][1] } { db[unit][2]} ")


# Generate a dict of ret[unit_num] = (name, active)
def pals_setup(ws):
    pals = dict()
    for rownum in range(6, ws.max_row + 1):
        for c in [2, 6, 8]: #unit num columns for PALS Prod
            unit = ws.cell(rownum, c).value
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
    check_worksheet(pals, wb[DB_SHEETS[0]], results_xl)
    results_xl.save(RESULTS_XL)

    # s = DB_SHEETS[0]
    # # for s in DB_SHEETS:
    # print(f'=================\nSheet is: {s} / {s.title}\n=================')
    # res = results_xl[s]
    # dbs = wb[s]
    #
    # for record in [[k] + v for k, v in pals.items()]:
    #     res.append(record)
    #     print(*dbs.values)
    #
    # print(results_xl.sheetnames)
    #
    # print("Finished")
    # results_xl.save(RESULTS_XL)

if __name__ == '__main__':
    main()
