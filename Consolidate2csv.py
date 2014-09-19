import xlrd
import sys
import os
import csv
from datetime import datetime


allfiles = {}

def parse_file(datafile, type):
    workbook = xlrd.open_workbook(datafile)
    sheet = workbook.sheet_by_index(1)
    newdict = {}
    rptdate = ()
    dod_ord_tot = 0
    va_ord_tot = 0
    dod_ord_passed = 0
    va_ord_passed = 0
    dod_res_tot = 0
    va_res_tot = 0
    dod_res_passed = 0
    va_res_passed = 0
    isSunday = False
    if type == "Con":
        rptdate = xlrd.xldate_as_tuple(sheet.cell_value(0,2),0)
        dayofweek = datetime(*rptdate).weekday()
        isSunday = (dayofweek == 6)
        va_ord_tot = sheet.cell_value(6,2)
        va_ord_passed = sheet.cell_value(7,2)
        dod_ord_tot = sheet.cell_value(10,2)
        dod_ord_passed = sheet.cell_value(11,2)
        newdict = {type: [round(va_ord_tot), round(va_ord_passed),round(dod_ord_tot), round(dod_ord_passed)]}
    elif type == "Rad":
        rptdate = xlrd.xldate_as_tuple(sheet.cell_value(1,2),0)
        dayofweek = datetime(*rptdate).weekday()
        isSunday = (dayofweek == 6)
        va_ord_tot = sheet.cell_value(6,2)
        va_ord_passed = sheet.cell_value(7,2)
        dod_ord_tot = sheet.cell_value(10,2)
        dod_ord_passed = sheet.cell_value(11,2)
        newdict = {type: [round(va_ord_tot), round(va_ord_passed),round(dod_ord_tot), round(dod_ord_passed)]}
    elif type == "Lab":
        rptdate = xlrd.xldate_as_tuple(sheet.cell_value(0,2),0)
        dayofweek = datetime(*rptdate).weekday()
        isSunday = (dayofweek == 6)
        va_ord_tot = sheet.cell_value(8,2)
        va_ord_passed = sheet.cell_value(9,2)
        dod_ord_tot = sheet.cell_value(12,2)
        dod_ord_passed = sheet.cell_value(13,2)
        va_res_tot = sheet.cell_value(20,2)
        va_res_passed = sheet.cell_value(21,2)
        dod_res_tot = sheet.cell_value(24,2)
        dod_res_passed = sheet.cell_value(25,2)
        newdict = {type: [round(va_ord_tot), round(va_ord_passed),round(dod_ord_tot), round(dod_ord_passed),
                             round(va_res_tot), round(va_res_passed),round(dod_res_tot), round(dod_res_passed)]}

    if rptdate in allfiles:
        if type not in allfiles[rptdate]:
            allfiles[rptdate].update(newdict)
    else:
        allfiles[rptdate] = newdict

    if isSunday:
        constr_additional(sheet, type)

def constr_additional(sheet, type):
    for i in range(2):
        newdict = {}
        rptdate = ()
        dod_ord_tot = 0
        va_ord_tot = 0
        dod_ord_passed = 0
        va_ord_passed = 0
        dod_res_tot = 0
        va_res_tot = 0
        dod_res_passed = 0
        va_res_passed = 0
        if type == "Con":
            rptdate = xlrd.xldate_as_tuple(sheet.cell_value(0,i+3),0)
            va_ord_tot = sheet.cell_value(6,i+3)
            va_ord_passed = sheet.cell_value(7,i+3)
            dod_ord_tot = sheet.cell_value(10,i+3)
            dod_ord_passed = sheet.cell_value(11,i+3)
            newdict = {type: [round(va_ord_tot), round(va_ord_passed),round(dod_ord_tot), round(dod_ord_passed)]}
        elif type == "Rad":
            rptdate = xlrd.xldate_as_tuple(sheet.cell_value(1,i+3),0)
            va_ord_tot = sheet.cell_value(6,i+3)
            va_ord_passed = sheet.cell_value(7,i+3)
            dod_ord_tot = sheet.cell_value(10,i+3)
            dod_ord_passed = sheet.cell_value(11,i+3)
            newdict = {type: [round(va_ord_tot), round(va_ord_passed),round(dod_ord_tot), round(dod_ord_passed)]}
        elif type == "Lab":
            rptdate = xlrd.xldate_as_tuple(sheet.cell_value(0,i+3),0)
            va_ord_tot = sheet.cell_value(8,i+3)
            va_ord_passed = sheet.cell_value(9,i+3)
            dod_ord_tot = sheet.cell_value(12,i+3)
            dod_ord_passed = sheet.cell_value(13,i+3)
            va_res_tot = sheet.cell_value(20,i+3)
            va_res_passed = sheet.cell_value(21,i+3)
            dod_res_tot = sheet.cell_value(24,i+3)
            dod_res_passed = sheet.cell_value(25,i+3)
            newdict = {type: [round(va_ord_tot), round(va_ord_passed),round(dod_ord_tot), round(dod_ord_passed),
                             round(va_res_tot), round(va_res_passed),round(dod_res_tot), round(dod_res_passed)]}

        if rptdate in allfiles:
            if type not in allfiles[rptdate]:
                allfiles[rptdate].update(newdict)
        else:
            allfiles[rptdate] = newdict

def main():
    for file in os.listdir(sys.argv[1]):
        if file.endswith(".xlsx") and "Con" in file:
            parse_file(sys.argv[1]+ "/" + file, "Con")
        elif file.endswith(".xlsx") and "Rad" in file:
            parse_file(sys.argv[1]+ "/" + file, "Rad")
        elif file.endswith(".xlsx") and "Lab" in file:
            parse_file(sys.argv[1]+ "/" + file, "Lab")
    f = open('eodtest.csv', 'w')
    header = 'rptdate,type,va_ord_tot,va_ord_passed,dod_ord_tot,dod_ord_passed,va_res_tot,va_res_passed,dod_res_tot,dod_res_passed\n'
    f.writelines(header)
    outfile = ''
    for k1, v1 in allfiles.items():
        rdate = str(k1[0]) + '-' + str(k1[1]) + '-' + str(k1[2])
        for k2, v2 in v1.items():
            outfile = rdate+ ',' + k2
            for s in v2:
                outfile += ',' + str(s)
            if k2 == 'Con' or k2 == 'Rad':
                outfile += ',,,,\n'
            else:
                outfile += '\n'
            f.writelines(outfile)
    f.close()



if __name__ == '__main__':
    main()
