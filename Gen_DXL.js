const { writeFile, readFile, utils } = require("xlsx");


const for_loop_template = `for (loop = 0; loop<%TESTCASES_CNT%; loop++)
{
    bool IsErrorOccured = false

	CurrTestCaseAbsID = (int get(ArrObjectInfor0,loop,10))
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
%CURROBJ_INFOS%
            //currOject.Object_Text		  = (string get(ArrObjectInfor0,loop,0))
            //currOject.TS_Test_Description = (string get(ArrObjectInfor0,loop,1))
            //currOject.TS_Test_Goal		  = (string get(ArrObjectInfor0,loop,2))
            //currOject.TS_Test_Priority	  = (string get(ArrObjectInfor0,loop,3))
            //currOject.TS_Object_Type	  = (string get(ArrObjectInfor0,loop,4))
            //currOject.TS_Expected_Result  = (string get(ArrObjectInfor0,loop,5))
            //currOject.TS_SwArchitectureDesign = (string get(ArrObjectInfor0,loop,6))
            //currOject.TS_Test_Enviroment  = (string get(ArrObjectInfor0,loop,7))
            //currOject.TS_TestLocation_TC  = (string get(ArrObjectInfor0,loop,8))
            //currOject.TS_Test_Case_Status = (string get(ArrObjectInfor0,loop,9))


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
    for (loop = 0; loop<%TESTCASES_CNT%; loop++)
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
`

let linkModule = "/Linkmodule/Verifies_DAS_SUZ_02";
let Req_Module = "System and Software Requirements and Structure_DAS_SUZ_02_YEA - ENG2 ENG3 ENG4 ENG5";
let TestSpec_Module = "DAS_SUZ_02_YEA_SW_TST_ENG8_Test_Specification"; 
