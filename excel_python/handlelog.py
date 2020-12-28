import re
import xlsxwriter

def ReadAlllog():
    excel_name = "core.xlsx"
    logname_1 = "x86_test_core.log"
    logname_2 = "arm_test_core.log"

    api_workbook = xlsxwriter.Workbook(excel_name, {'nan_inf_to_errors': True})
    bold = api_workbook.add_format({'align':'center'})
    sheetname_1 = GetNameList(logname_1, "sheet")
    sheetname_2 = GetNameList(logname_2, "sheet")
    testname_1 = GetNameList(logname_1, "test")
    testname_2 = GetNameList(logname_2, "test")
    api_workbook.close()

def Write_to_sheet(logname,api_workbook,sheetname,testname,type):
    for sheet in sheetname:
        sametype_test = [test for test in testname if sheet in test]
        namelist = ["height","wight","Intel","C","SpeedUp"]
        if type is "new":
            wsheet = api_workbook.add_worksheet(sheet)
            typename = "x86"
        else:
            wsheet = api_workbook.get_worksheet_by_name(sheet)
            typename = "ARM"
            namelist[2] = "ARM_NENO"
            del namelist[0:2]
        row = 0
        for test in sametype_test:
            col = 0
            if type is "new"
                wsheet.write_string(row, col, test)
            else:
                col = col + 5
            row = row + 1
            wsheet.write_string(row, col, typename)
            row = row + 1
            wsheet.write_string(row, col, namelist)
            testdata = Gettestdata(logname, test)
            datalines = re.split("[\n\r]", testdata)
            del datalines[0:2]
            for line in datalines:
                row = row + 1
                data = re.split("\t", line)
                while '' in data:
                    data.remove('')
                data = [element.strip() for element in data]
                float_data = [float(data) for element in data[2:5]]
                row_index = row + 1
                if type is "new":
                    int_data = [int(data) for element in data[0:2]]
                    speedup_formula = '{=' + chr(ord('A') + col + 3) + str(row_index) + '/' + chr(ord('A') + col + 2) + str(row_index) + '}'
                    wsheet.write_row(row, col, int_data)
                    wsheet.write_row(row, col + 2, float_data)
                    wsheet.write_row(row, col + 4, speedup_formula)
                else:
                    speedup_formula = '{=' + chr(ord('A') + col + 1) + str(row_index) + '/' + chr(ord('A') + col) + str(row_index) + '}'
                    wsheet.write_row(row, col, float_data)
                    wsheet.write_row(row, col + 2, speedup_formula)
                    
def GetNameList(log, type):
    sheetname = []
    testname = []
    with open(log, 'rb') as f:
        lines = f.readlines()
        for line in lines
            line = line.decode('utf-8')
            if line.find("[ RUN    ] PERFIPP_PERFORMANCE") > -1:
                   list1 = re.split(r'[_.]',line) 
                   if len(list1) > 5:
                       sheetname.append(list1[2] + "_" + list1[3] + "_" + list1[4])
                       testname.append(list1[2] + "_" + list1[3] + "_" + list1[4] + "." + list1[5] + "_" + list1[6])
                    else:
                        sheetname.append(list1[2])
                        testname.append(list1[2] + "." + list1[3] + "_" + list1[4])
        sheetname = sorted(set(sheetname), key = sheetname.index)
        testname = sorted(set(testname), key = testname.index)
    f.close()
    if type is "sheet"
        return sheetname
    else:
        return testname

def Gettestdata(file,test)
    pattern = "(\W\sRUN\s\W+ [A-Z]+_PERFORMANCE_test\s*)(\d+\s+\d+\s+\d+.\d+\s+\d+.\d+\s+)+"
    pattern = re.sub("test", test, pattern)
    with open(file, 'rb') as f:
        wholefile = f.read(-1)
        wholefile = wholefile.decode('utf-8')
        testdata = re.search(pattern, wholefile, re.M)
    f.close()
    if test is "MIRROR_45_135"
        return testdata.group()

ReadAlllog()
                