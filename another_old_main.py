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


    # loop through each db - for testing doing cara test?
    # for each member of allPals add to ws, check db for that unit, etc
    # current = copy.deepcopy(allPals) # don't need this I think

    # at bottom add the ones that are in input but not pals
    n = 0

    # this won't work as well'
    # so with pals in memory, and the results assigned to ws
    # go through each row in dbs and check if it's in pals
    # OR do it the non retarded way and just add to PALS as it goes through, if it's not in there then add it




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
