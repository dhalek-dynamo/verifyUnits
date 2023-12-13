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

class UnitDetail:
    def __init__(self):
        self.Unit = 0
        self.Name = ""
        self.isActive = True
        self.unitCorrect = False
        self.nameCorrect = False
        self.isActiveCorrect = False
        self.verified = False
        self.missingThisUnit = False
        self.dbName = ""

class PalsUnit:
    def __init__(self):
        self.palsUnit = -1
        self.palsName = ""
        self.palsActive = False
        self.verified = False
        self.missingFromPals = False # is the pals version missing from any of the dbs?

    dbUnits = dict() # { dbName : UnitDetail }

def check_units(pals, dbs, ws):
    for rownum in range(2, ws.max_row+1):
        unit = ws.cell(rownum, 1)
        name = ws.cell(rownum, 2)
        active = ws.cell(rownum, 3)
        isActive = ws.cell(rownum, 3)
        if unit in pals.keys():
            self.Unit = unit
            self.Name = name
            self.isActive = active
            self.nameCorrect = pals[unit][palsUnit] == name
            self.isActiveCorrect = pals[unit][palsActive] == isActive
            self.verified = self.nameCorrect and self.isActiveCorrect



# Generate a dict of ret[unit_num] = (name, active)
# redundant but oh well
def pals_setup(ws, results):
    pals = dict()
    for rownum in range(6, ws.max_row + 1):
        for c in [2, 6, 8]: #unit num columns for PALS Prod
            unit = ws.cell(rownum, c).value
            name = ws.cell(rownum, c+1).value
            active = ws.cell(rownum, 10).value is None
            if unit:
                if unit in INACTIVE_FORESTS:
                    active = False
                pals[unit] = [name, active]

    return pals

if __name__ == '__main__':
    wb = xl.load_workbook(INPUT_XL)
    results_xl = xl.load_workbook(RESULTS_XL)
    for s in CLEAR_SHEETS:
        results_xl[s].delete_rows(2,5000)

    allPals = pals_setup(wb[PALS], results_xl)
    #
    # for sheet in DB_SHEETS:
    #   call whatever function we're using
    #
    check_units(allPals, wb[DB_SHEETS[0]])

    different = [ ]
    not_in_pals = [ ]
    total_checked = nope = verified = 0

    s = DB_SHEETS[0]
    # for s in DB_SHEETS:
    print(f'=================\nSheet is: {s} / {s.title}\n=================')
    res = results_xl[s]
    res.append(["Unit", "Name", "Active?", "Verified?", "Name Correct?", "Active Correct?", "Name", "Active"])
    dbs = wb[s]

    # loop through each db - for testing doing cara test?
    # for each member of allPals add to ws, check db for that unit, etc
    # current = copy.deepcopy(allPals) # don't need this I think

    # at bottom add the ones that are in input but not pals
    n = 0

    # this won't work as well'
    # so with pals in memory, and the results assigned to ws
    # go through each row in dbs and check if it's in pals
    # OR do it the non retarded way and just add to PALS as it goes through, if it's not in there then add it

    for record in [[k] + v for k, v in allPals.items()]:
        res.append(record)
        print(*dbs.values)
        if n < 100:
            print(record)
        n += 1

    print(results_xl.sheetnames)

    # for row in range(2, ws.max_row + 1):
    #     total_checked += 1


    print("Finished")
    count = 0

    results_xl.save(RESULTS_XL)
    print(f'PALS count:        {count}')
    print(f'PALS rows:         {len(allPals)}')
    print(f'Not in PALS count: {len(not_in_pals)}')
    print( '=======================')
    print(f'Verified count:    {verified}')



#================================================================================================================



    # xyz = 0
    # for key, val in allPals.items():
    #     xyz += 1
    #     if xyz > 100:
    #         break
    #     print(key, "-", val, "-", *val)

    #         t = ws.cell(row, 1).value
    #         t_out = [ ws.cell(row, 1).value, ws.cell(row, 2).value, ws.cell(row, 3).value, ws.cell(row, 4).value]
    #         if t in allPals['four']:
    #             allPals['four'][t][2] = True
    #             if ws.cell(row, 2).value == allPals['four'][t][0] and ws.cell(row, 3).value == allPals['four'][t][1]:
    #                 # verified.append(f'{s} - {row} - {t} - {pals[t][0]}\n')
    #                 wb['VERIFIED'].append(t_out)
    #                 verified += 1
    #             else:
    #                 # add this to NOPE worksheet for now
    #                 wb['NOPE'].append(t_out)
    #                 # print(f"==> NOPE: {ws.cell(row, 1).value} ({ws.cell(row, 2).value}) {ws.cell(row, 3).value}")
    #                 nope += 1
    #         else:
    #             not_in_pals.append(f"==> NOT IN PALS ({s}): {ws.cell(row, 1).value} ({ws.cell(row, 2).value}) {ws.cell(row, 3).value} \n")
    #             wb[WORKING].append(t_out)
    #             # print("=====[ Not in PALS ]=====")

    # for unit, val in allPals.items():
    #     # for db in bunch-of-dbs ....
    #     #     do the comparisons here
    #     if val[2] is False: # val[2] is flag for it being used, true is good, false means it wasn't in the db
    #         # print(unit, *val)
    #         pass
    #     count += 1 if not val[2] else 0
