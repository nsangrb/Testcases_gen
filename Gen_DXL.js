const { Get_OverviewInfo } = require("./Func");
const { writeFile } = require("fs");
const { exec } = require("child_process");
const { readFile, utils } = require("xlsx");

module.exports = { Generate_Dxl };
//#region Template
const for_loop_template = `for (loop = 0; loop<%TESTCASES_CNT%; loop++)
{
    bool IsErrorOccured = false

	CurrTestCaseAbsID = (int get(ArrObjectInfor0,loop,0))
	Object currOject = object(CurrTestCaseAbsID)
    if ((!null currOject) and (loop == 0))
    {
        //do nothing --> this object is the first one created by tester
    }
    else if ((null currOject) and (loop != 0))
    {

        //create new object here
        PreviousTestCaseAbsID = (int get(ArrObjectInfor0,loop-1,0))
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
        CurrTestCaseAbsID = (int get(ArrObjectInfor0,loop,0))
    	Object currOject = object(CurrTestCaseAbsID)

        //Link link
        //for link in currOject -> "*" do {
        //    delete(link)
        //}
        targetRequirementModule = read(pathToRequirement)
        filtering off
        ReqID = (int get(ArrObjectInfor0,loop,1))
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
`;
//#endregion

function Read_Config(excel_config_path) {
  let wb = readFile(excel_config_path);
  return Get_OverviewInfo(
    {
      "TestSpec path": "",
      "Link module": "",
      "Requirement Module": "",
      "TestSpec Module": "",
    },
    wb.Sheets["Overview"],
    "B7:C11"
  );
}

function Generate_Dxl(excel_config_path) {
  let [err, Overview_info] = Read_Config(excel_config_path);
  if (err !== "") return err;
  let linkModule = Overview_info["Link module"];
  let Req_Module = Overview_info["Requirement Module"];
  let TestSpec_Module = Overview_info["TestSpec Module"];
  let wb = readFile(Overview_info["TestSpec path"]);
  var data = utils.sheet_to_json(wb.Sheets["Test spec"], {
    blankrows: false,
  });
  let Attr_keys = Object.keys(data[0]);

  let define_attrs = "";
  let currobj_infos = "";
  let result = "";
  for (var column = 2; column < Attr_keys.length; column++) {
    define_attrs += `${Attr_keys[column].replace(/ /g, "_")} = "${
      Attr_keys[column]
    }"\n`;
    currobj_infos += `\t\t\tcurrOject.${Attr_keys[column].replace(
      / /g,
      "_"
    )}\t=(string get(ArrObjectInfor0,loop,${column}))\n`;
  }

  result += define_attrs;
  result += `RequirementModuleName="${Req_Module}"\n`;
  result += `currMod="${TestSpec_Module}"\n`;
  result += `Module m0 = null\nif (!null currMod){m0 = edit(currMod, true)}\nArray ArrObjectInfor0 = create(${
    data.length - 1
  },${Attr_keys.length})\n`;
  result += `
  //Calculate Path
  Project project = getParentProject(m0)
  pathToRequirement = "/" name(project) "/20_SYS/" RequirementModuleName
  pathToLinkModule = "/" name(project) "${linkModule}"\n
  `;
  result += "//Array data\n";
  for (var line = 1; line < data.length; line++) {
    for (var column = 0; column < Attr_keys.length; column++) {
      //console.log(data[line][Attr_keys[column]]);
      let corr_str = data[line][Attr_keys[column]].toString();
      corr_str = corr_str.replace(/???/g, "<=");
      corr_str = corr_str.replace(/???/g, ">=");
      //corr_str = corr_str.replace(/\r\n/g, "");
      corr_str = corr_str.replace(/"/g, "'");
      if (column > 1)
        result += `put(ArrObjectInfor0, "${corr_str}", ${
          line - 1
        }, ${column})\n`;
      else {
        result += `put(ArrObjectInfor0, ${corr_str}, ${line - 1}, ${column})\n`;
      }
    }
  }
  result += for_loop_template
    .replace(/%TESTCASES_CNT%/g, data.length - 1)
    .replace(/%CURROBJ_INFOS%/, currobj_infos);
  writeFile("Result.dxl", result, function (error) {
    if (error) err += error.stack + "\r\n";
  });
  exec(`notepad "Result.dxl"`);
  return err;
}
