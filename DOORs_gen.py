from _overlapped import NULL
__author__ = 'dyu2hc'

from xlrd import open_workbook
import re, sys, os
#import DoorSys_gui
##import message


##SWT DEFINE
###1. column sequences
Col_Attributes = {
"col_TC_ID"     				:      0,
"col_Req_ID"     				:      1,
"col_Object_Text"     			:      2,
"col_TS_Test_Description"     	:      3,
"col_TS_Test_Goal"     			:      4,
"col_TS_Test_Priority"     		:      5,
"col_TS_Object_Type"     		:      6,
"col_TS_Expected_Result"     	:      7,
"col_TS_SwArchitectureDesign"   :      8,
"col_TS_Test_Enviroment"     	:      9,
"col_TS_TestLocation_TC"     	:      10,
"col_TS_Test_Case_Status"     	:      11
}


##2. Attributes name -> DOOR
AttributeName = ['"Object Text"',	'"TS_Test Description"',	'"TS_Test Goal"',	'"TS_Test Priority"',	'"TS_Object Type"',	'"TS_Expected Result"',	\
               '"TS_SwArchitectureDesign"',	'"TS_Test Environment"',	'"TS_TestLocation_TC"',	'"TS_Test Case Status"']

############



absolutenum_col = 0             #Absolute number column
comment_col = 1                 #Comment test center column
testresult = 2                  #Test Result status column name
modlink=3                       #module link
start_row_number = 1            #first line to read in excel file
sheet_number = 0                #sheet to read in excel file
number_of_record = 0        #number of requirement to update
data = []                       #data to write in DOORS
                     #link of module
skip_blank_row = True
loc = os.path.realpath(__file__)
# print (loc)
# dxl_file_name = loc.replace('DoorsUtilities_Version1.0\\DoorsFill.py','Output')
# dxl_file_name = dxl_file_name + "\\DXL.txt"
excel_file_name = loc.replace('DOOR_Generate.py','TestSpec_out.xls')
resultName = loc.replace('DOOR_Generate.py','Result.txt')
# excel_file_name = excel_file_name + "\\data.xlsx"
# print (excel_file_name)
number_of_column_in_data = len(AttributeName)
have_error = False
wb = None
s = None
# f = open(dxl_file_name, 'w')
#delete all old content of DXL file


#Define ari1yh
TestResultData= [] ##Testcase, result and Obtain
DoorData= [] ##Heading and Page
g_all_data = []

NumberOfValidColumn = 7

def Main():
    read_excel_file()
##        Sample()
##        fillData()
#         CheckXMLResult()


def Sample():

    excel = open_workbook(excel_file_name)
    sheet_0 = excel.sheet_by_index(0) # Open the first tab

    prev_row = [None for i in range(sheet_0.ncols)]
    for row_index in range(sheet_0.nrows):
        row= []
        for col_index in range(sheet_0.ncols):
            if col_index <= NumberOfValidColumn :
                CheckValid = sheet_0.cell(rowx=row_index,colx= 7).value #Read Colunm Generate or not
                if "Generate" in CheckValid:
                    value = sheet_0.cell(rowx=row_index,colx=col_index).value
                    if col_index == 0:
                        if len(value) == 0:
                            value = prev_row[col_index]


                    row.append(value)
            else:
                pass
        if len(row) >1:
            prev_row = row
            g_all_data.append(row)

    print(g_all_data)


def fillData():

    ReturnList.writelines("ListTestCases = {\t\t\t\t\t\t#Test Cases ID\t\t\t#vEgo\t\t#vObj\t#dInitObjEgo\t#testpattern\t#variant\t\t#Number of RQ\n\n" )
    counter = 1
    for index in g_all_data[1:]:
        p = re.compile('(?!=[\\d])\d+$')
        RetNum = p.search(index[0])
        if RetNum is None:
            print("<E> Cannot find absolute number in requirement %s" %index[0] )
        else:
            ReturnList.writelines("\t\t\t\t'%d'\t\t:\t\t\t('%s',\t\t%d,\t\t%d,\t\t\t%d,\t\t\t'%s',\t\t'%s',\t\t%s),\n" % (counter, index[1], index[2], index[3], index[4], index[5], index[6], RetNum.group(0).strip()))

#         print("Data : {} \n",  index)

        counter = counter + 1

    ReturnList.writelines("\t\t}")
    ReturnList.close()


    print("Complete")
#     print('Data: ', all_data[2:9])
    #for row in data:
    #    print(row[0].stripaa(), "   ", row[1].strip())
def read_excel_file():
    excel = open_workbook(excel_file_name)
    sheet_0 = excel.sheet_by_index(sheet_number)

    if excel is None:
        print("Cannot find %s" % sys.argv[1])
    print('- Reading Sheet:', sheet_0.name)
    number_of_record = sheet_0.nrows
    print("- Number of test cases to update:",  number_of_record - 1)
    for row in range(sheet_0.nrows):
        if row == 0:
            pass ##Heading row. DO not thing
        else:
            col_TC_ID = str(int(sheet_0.cell(row, Col_Attributes["col_TC_ID"]).value))
            col_TC_ID = col_TC_ID.replace('"',"")
            col_Req_ID = str(int(sheet_0.cell(row, Col_Attributes["col_Req_ID"]).value))
            col_Req_ID = col_Req_ID.replace('"',"")

            col_Object_Text =str( sheet_0.cell(row, Col_Attributes["col_Object_Text"]).value)
            col_Object_Text =col_Object_Text.replace("≤","<=")
            col_Object_Text =col_Object_Text.replace("≥",">=")
            col_Object_Text = col_Object_Text.replace('\r\n','')
            col_Object_Text = col_Object_Text.replace('"',"'")

            col_Object_Text = '"' + col_Object_Text + '"'
            print (col_Object_Text)
            print ("###############")

            col_TS_Test_Description =str( sheet_0.cell(row, Col_Attributes["col_TS_Test_Description"]).value)
            col_TS_Test_Description =col_TS_Test_Description.replace("≤","<=")
            col_TS_Test_Description =col_TS_Test_Description.replace("≥",">=")
            col_TS_Test_Description = col_TS_Test_Description.replace('"',"'")
##            col_TS_Test_Description = '"' + col_TS_Test_Description + '"'

            col_TS_Test_Goal = sheet_0.cell(row, Col_Attributes["col_TS_Test_Goal"]).value
            col_TS_Test_Priority = sheet_0.cell(row, Col_Attributes["col_TS_Test_Priority"]).value
            col_TS_Object_Type = sheet_0.cell(row, Col_Attributes["col_TS_Object_Type"]).value
            col_TS_Expected_Result = sheet_0.cell(row, Col_Attributes["col_TS_Expected_Result"]).value
            col_TS_SwArchitectureDesign = sheet_0.cell(row, Col_Attributes["col_TS_SwArchitectureDesign"]).value
            col_TS_Test_Enviroment = sheet_0.cell(row, Col_Attributes["col_TS_Test_Enviroment"]).value
            col_TS_TestLocation_TC = sheet_0.cell(row, Col_Attributes["col_TS_TestLocation_TC"]).value
            col_TS_Test_Case_Status = sheet_0.cell(row, Col_Attributes["col_TS_Test_Case_Status"]).value

            TestResultData.append([col_TC_ID, col_Req_ID, col_Object_Text, col_TS_Test_Description, \
                                  col_TS_Test_Goal, col_TS_Test_Priority, col_TS_Object_Type, col_TS_Expected_Result,\
                                  col_TS_SwArchitectureDesign,col_TS_Test_Enviroment, col_TS_TestLocation_TC,\
                                  col_TS_Test_Case_Status])


##    for item in TestResultData:
##        for eachItem in item:
##            print(eachItem)
##
##        print("#################")
            #DoorData.append([HeadingTestResultCol, Page])
def CheckingData(data = []):
    ## Checking if any Testcase, Test Result contain null
    ## Return True if one of them do not contain data
    listValidResult = ('passed', 'failed', 'n/a', 'Passed', 'Failed')
    errorList = []
    ReturnValue = False
    for row in data:
        if row[0] == '' or row[1] == '' or row[1] not in listValidResult :
            if row[0] == '':
                errorList.append("<Error> Cannot find the Test Case in row %d" %(data.index(row) + 2))
            elif row[1] == '':
                errorList.append("<Error> Cannot find the Test Result in row %d" %(data.index(row) + 2))
            else:
                errorList.append("<Error> Test result status is INVALID in row %d" %(data.index(row) + 2))
            ReturnValue = True
    return ReturnValue, errorList
def checkDoubleTC(data = []):
    isDouble = False
    errorList= []
    for index in range(len(data)):
        if index == len(data) -1:
            break
        for index2 in range(index +1, len(data)):
            if data[index][0]== data[index2][0]:
                isDouble = True
                errorList.append("<Error> Double test case number between Row: %d and Row: %d"%(index +2, index2 +2))

    return isDouble, errorList


def GetOnlyNumber():
    ## Get only the number of Test case
    ## Return True if one of them cannot convert
    ReturnVal = False

    for index in TestResultData:
        p = re.compile('(?!=[\\d])\d+$')
        RetNum = p.search(index[0])
        if RetNum is None:
            print("<E> Cannot find absolute number in requirement %s. Row : %d" %(index[0], TestResultData.index(index) + 2) )
            ReturnVal = True
        else:
            index[0] = RetNum.group(0).strip()

    return ReturnVal

def Process(ModuleTestCaseName = '', ModuleRequirementName = '',):
    if (('Ex:' in ModuleTestCaseName) or ModuleTestCaseName == ''):   #or 'Ex:' in TestResult) or \
        # 'Test Result Status' not in TestResult \
        # or pageName == ''): #or TestResult ==''):

        print('Please input correct data!')
        #message.message_info('Please input correct data!')
        print('Please input correct data!')
    else:
        DoorData.append([ModuleRequirementName,ModuleTestCaseName])

        CheckingTotal = True

        print("**** ==============================================**** ")
        print("**** ==============================================**** ")
        print("**** Reading excel file... **** ")
        read_excel_file()

##        Ret = GetOnlyNumber()
##        if Ret == False:
##            CheckingTotal = True
##    #         TestResultData = data
##        else:
##            CheckingTotal = False

        CheckingTotal = True
        if CheckingTotal == False:
            print("**** Data error ****")
            print("**** Stop generating ****")
            mss ='Generating Failed!'
            #message.message_info(mss)
            print(mss)
        else:
            print("**** Checking data...**** ")
##            Return, Listerror = CheckingData(TestResultData)
##            Return1, Listerror1 = checkDoubleTC(TestResultData)
            Return = False
            Return1 = False
            if Return == True or Return1 == True :
                print("- Some error in data:")
                for error in Listerror:
                    print(error)
                for error in Listerror1:
                    print(error)

                mss ='Generating Failed!'
                print(mss)
                #message.message_info(mss)
            else:
                print("Data is no issue")
                print("*** START GENERATING SCRIPTS ***")
                write_dxl_file()
                print("*** COMPLETE GENERATING SCRIPTS ***")


                mss ='Generating completely!'
                print(mss)
                #message.message_info(mss)
                os.system(resultName)

    TestResultData.clear() ##Reset
    DoorData.clear()

def write_dxl_file():
    ReturnList= open("Result.txt", "w+")
    ReturnList.truncate()
    g_index_row = 0

    #DXL: create array to store data

    ReturnList.writelines("Object_Text=%s\n"%AttributeName[0] )
    ReturnList.writelines("TS_Test_Description=%s\n"%AttributeName[1] )
    ReturnList.writelines("TS_Test_Goal=%s\n"%AttributeName[2] )
    ReturnList.writelines("TS_Test_Priority=%s\n"%AttributeName[3] )
    ReturnList.writelines("TS_Object_Type=%s\n"%AttributeName[4] )
    ReturnList.writelines("TS_Expected_Result=%s\n"%AttributeName[5] )
    ReturnList.writelines("TS_SwArchitectureDesign=%s\n"%AttributeName[6] )
    ReturnList.writelines("TS_Test_Enviroment=%s\n"%AttributeName[7] )
    ReturnList.writelines("TS_TestLocation_TC=%s\n"%AttributeName[8] )
    ReturnList.writelines("TS_Test_Case_Status=%s\n"%AttributeName[9] )



##    res = '"' + DoorData[0][0] + '"'
##    obtainCol = '"' + DoorData[0][0].replace("Test Result Status", "Obtained Test Results") + '"'
##    ReturnList.writelines("TestResultHeading=%s\n"%res )
##    ReturnList.writelines("Obtain=%s\n"%obtainCol )
##    ReturnList.writelines("NumberTCRequested=%d\n"%len(TestResultData) )
##    ReturnList.writelines("NumberTCUpdated=0\n")
##    ReturnList.writelines("Obtain=%s\n"%obtainCol )

    mod = '"' + DoorData[0][1] + '"'
    RqName ='"' + DoorData[0][0] + '"'
    ReturnList.writelines("RequirementModuleName=%s\n"%RqName)
    ReturnList.writelines("currMod = %s\n"%mod )
    ReturnList.writelines("Module m%d = null\n"%g_index_row)
    ReturnList.writelines("if (!null currMod){m%d = edit(currMod, true)}\n"%g_index_row)
    ReturnList.writelines("Array ArrObjectInfor%d = create(%d,%d)\n" % (g_index_row,len(TestResultData), number_of_column_in_data + 2)) ##Adding TC_ID and RQ_ID

    RequirementPathCalculation = """
//Calculate Path
Project project = getParentProject(m0)
pathToRequirement = "/" name(project) "/20_SYS/" RequirementModuleName
pathToLinkModule = "/" name(project) "/Linkmodule/Verifies_DAS_SUZ_02"

//Array data
"""
    ReturnList.writelines(RequirementPathCalculation)
    #DXL: write data of column 1 (requirement absolute number)
#     old_row_no = g_cur_row
    for row_no in TestResultData:
#         row_no = g_cur_row
        ReturnList.writelines('put(ArrObjectInfor%d, %s, %d, %d)\n' % (g_index_row, row_no[Col_Attributes['col_Object_Text']] , TestResultData.index(row_no), 0))
        ReturnList.writelines('put(ArrObjectInfor%d, "%s", %d, %d)\n' % (g_index_row, row_no[Col_Attributes['col_TS_Test_Description']], TestResultData.index(row_no), 1))
        ReturnList.writelines('put(ArrObjectInfor%d, "%s", %d, %d)\n' % (g_index_row, row_no[Col_Attributes['col_TS_Test_Goal']], TestResultData.index(row_no), 2))
        ReturnList.writelines('put(ArrObjectInfor%d, "%s", %d, %d)\n' % (g_index_row, row_no[Col_Attributes['col_TS_Test_Priority']], TestResultData.index(row_no), 3))
        ReturnList.writelines('put(ArrObjectInfor%d, "%s", %d, %d)\n' % (g_index_row, row_no[Col_Attributes['col_TS_Object_Type']], TestResultData.index(row_no), 4))
        ReturnList.writelines('put(ArrObjectInfor%d, "%s", %d, %d)\n' % (g_index_row, row_no[Col_Attributes['col_TS_Expected_Result']], TestResultData.index(row_no), 5))
        ReturnList.writelines('put(ArrObjectInfor%d, "%s", %d, %d)\n' % (g_index_row, row_no[Col_Attributes['col_TS_SwArchitectureDesign']], TestResultData.index(row_no), 6))
        ReturnList.writelines('put(ArrObjectInfor%d, "%s", %d, %d)\n' % (g_index_row, row_no[Col_Attributes['col_TS_Test_Enviroment']], TestResultData.index(row_no), 7))
        ReturnList.writelines('put(ArrObjectInfor%d, "%s", %d, %d)\n' % (g_index_row, row_no[Col_Attributes['col_TS_TestLocation_TC']], TestResultData.index(row_no), 8))
        ReturnList.writelines('put(ArrObjectInfor%d, "%s", %d, %d)\n' % (g_index_row, row_no[Col_Attributes['col_TS_Test_Case_Status']], TestResultData.index(row_no), 9))

        TCID ='put(ArrObjectInfor%d, "%s", %d, %d)\n' % (g_index_row, row_no[Col_Attributes['col_TC_ID']], TestResultData.index(row_no), 10)
        TCID = TCID.replace('"',"")
        ReturnList.writelines(TCID)

        RQID = 'put(ArrObjectInfor%d, "%s", %d, %d)\n' % (g_index_row, row_no[Col_Attributes['col_Req_ID']], TestResultData.index(row_no), 11)
        RQID = RQID.replace('"',"")
        ReturnList.writelines(RQID)
        #ReturnList.writelines('put(ArrObjectInfor%d, "%s", %d, %d)\n' % (g_index_row, row_no[Col_Attributes['col_TC_ID']], TestResultData.index(row_no), 10))
        #ReturnList.writelines('put(ArrObjectInfor%d, "%s", %d, %d)\n' % (g_index_row, row_no[Col_Attributes['col_Req_ID']], TestResultData.index(row_no), 11))
#         print(data[row_no][0]),
#         print("--> " + data[row_no][1])
#         g_cur_row += 1
        print("Test case %s : Generate completely" %row_no[0])
        #DXL: Write for loop
    forloop = """for (loop = 0; loop<%d; loop++)
{
    bool IsErrorOccured = false

	CurrTestCaseAbsID = (int get(ArrObjectInfor%d,loop,10))
	Object currOject = object(CurrTestCaseAbsID)
    if ((!null currOject) and (loop == 0))
    {
        //do nothing --> this object is the first one created by tester
    }
    else if ((null currOject) and (loop != 0))
    {

        //create new object here
        PreviousTestCaseAbsID = (int get(ArrObjectInfor0,loop-1,10))
        Object PreviousObject = object(PreviousTestCaseAbsID)
        if (!null PreviousObject)
        {
            currOject = create PreviousObject //Create new object below with same level
        }
        else
        {
            print "ERROR: Test case "  PreviousTestCaseAbsID " is not exist"
            IsErrorOccured = true
            break
        }

    }
    else
    {
        IsErrorOccured = true
        print "ERROR: Test case "  CurrTestCaseAbsID " is exist"
        break
        //Error--> Stop updating
    }

	if((!null currOject) and (IsErrorOccured == false))
	{

            //Update Test case attributes

            currOject.Object_Text		  = (string get(ArrObjectInfor0,loop,0))
            currOject.TS_Test_Description = (string get(ArrObjectInfor0,loop,1))
            currOject.TS_Test_Goal		  = (string get(ArrObjectInfor0,loop,2))
            currOject.TS_Test_Priority	  = (string get(ArrObjectInfor0,loop,3))
            currOject.TS_Object_Type	  = (string get(ArrObjectInfor0,loop,4))
            currOject.TS_Expected_Result  = (string get(ArrObjectInfor0,loop,5))
            currOject.TS_SwArchitectureDesign = (string get(ArrObjectInfor0,loop,6))
            currOject.TS_Test_Enviroment  = (string get(ArrObjectInfor0,loop,7))
            currOject.TS_TestLocation_TC  = (string get(ArrObjectInfor0,loop,8))
            currOject.TS_Test_Case_Status = (string get(ArrObjectInfor0,loop,9))


            //NumberTCUpdated = NumberTCUpdated +1
            print "Completed Test Case " CurrTestCaseAbsID "\\n"

	}
	else
	{
        IsErrorOccured = true
		print CurrTestCaseAbsID " is null or Object Type is not Test Case\\n"
        break
	}

}

//Update LinkModule
Module targetRequirementModule = null
if (IsErrorOccured == false)
{


    for (loop = 0; loop<%d; loop++)
    {
        m0 = edit(currMod, true)
        CurrTestCaseAbsID = (int get(ArrObjectInfor0,loop,10))
    	Object currOject = object(CurrTestCaseAbsID)

        //Link link
        //for link in currOject -> "*" do {
        //    delete(link)
        //}
        targetRequirementModule = read(pathToRequirement)
        filtering off
        ReqID = (int get(ArrObjectInfor0,loop,11))
        Object RequirementObject = object(ReqID)


        if ((!null RequirementObject) and (!null currOject))
         {
            currOject -> pathToLinkModule -> RequirementObject
         }
         else
         {
            print "LINK ERROR: Test case " CurrTestCaseAbsID " to " ReqID
            IsErrorOccured = true
            break
         }

    }
}
close(targetRequirementModule)
//print "Requested: " NumberTCRequested " test cases -- Updated: " NumberTCUpdated " test cases \n
""" % (len(TestResultData) ,g_index_row,len(TestResultData))
    ReturnList.writelines(forloop)
#     ReturnList.writelines( "\nsave(Module m%d)\n"%g_index_row)
#     ReturnList.writelines("close(Module m%d)\n"%g_index_row)
    g_index_row=g_index_row+1
    ReturnList.close()

if __name__== "__main__":
    Process('DAS_SUZ_02_YEA_SW_TST_ENG8_Test_Specification','System and Software Requirements and Structure_DAS_SUZ_02_YEA - ENG2 ENG3 ENG4 ENG5')
##    read_excel_file()

##    pass

# if __name__ == '__main__':
#     excel = open_workbook(excel_file_name)
#     sheet_0 = excel.sheet_by_index(sheet_number)
#     if excel is None:
#         print("Cannot find %s" % sys.argv[1])
#     else:
#         Sample()
#          read_excel_file()
#          g_oldreq = data[1][2]
#          for row_no in range(start_row_number, len(data) - 1):
#              row_no = 0
#              write_dxl_file()
#
#         print("Finished!")

#     ReturnList.close()
