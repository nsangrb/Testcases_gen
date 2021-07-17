const { IsDefined, Get_OverviewInfo } = require("./Func");
const { exec } = require("child_process");
const { writeFile, readFile, utils } = require("xlsx");

module.exports = {
  Generate_TestSpec,
};

function Update_content(testspec_fieldname, data, req_info, mapping) {
  let err = "";
  let _regex = /\${\w+\}/g;
  let patterns = data.match(_regex);
  let temp_data = data;
  if ((patterns || []).length > 0) {
    patterns.forEach((element) => {
      if (IsDefined(mapping[element]))
        temp_data = temp_data.replace(element, req_info[mapping[element]]);
      else {
        err += `Arg ${element} in field "${testspec_fieldname}" does not exist in Mapping sheet!!\r\n`;
      }
    });
  }
  return [err, temp_data];
}

function Init_content(TestSpec_keys) {
  let args = [];
  for (var i in TestSpec_keys) {
    let element = TestSpec_keys[i];
    args[element] = "x";
  }
  return args;
}

function Read_Config(excel_config_path) {
  let wb = readFile(excel_config_path);
  let [err, Path_info] = Get_OverviewInfo(
    {
      "Requirement path": "",
      Output: "",
      "TestSpec path": "",
    },
    wb.Sheets["Overview"],
    "B3:C7"
  );
  let TestSpec = utils.sheet_to_json(wb.Sheets["TestSpec_Template"]);
  let mapping = utils.sheet_to_json(wb.Sheets["Mapping"]);
  if (TestSpec.length === 0)
    err += "Please define some testing attrs for TestSpec sheet!!\r\n";
  if (mapping.length === 0)
    err += "Please define some mapping attrs for Mapping sheet!!\r\n";
  return [err, Path_info, TestSpec, mapping];
}

function Generate_TestSpec(excel_config_path) {
  let output = [];
  let [err, Path_info, TestSpec, mapping] = Read_Config(excel_config_path);
  if (err !== "") return err;

  let wb = readFile(Path_info["Requirement path"]);
  let data = utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
    blankrows: false,
  });
  let TestSpec_keys = Object.keys(TestSpec[0]);
  output.push(Init_content(TestSpec_keys));
  data.every((el) => {
    let args = [];
    let tmp_err = "";
    for (var i in TestSpec_keys) {
      let element = TestSpec_keys[i];
      [tmp_err, args[element]] = Update_content(
        element,
        TestSpec[0][element].replace(/\r\n/g, "\n"),
        el,
        mapping[0]
      );
      err += tmp_err;
    }
    output.push(args);
    if (err !== "") return false;
    else return true;
  });
  if (err !== "") return err;
  let output_sheet = utils.json_to_sheet(output);
  let wb_out = utils.book_new();
  wb_out.SheetNames.push("Test spec");
  wb_out.Sheets["Test spec"] = output_sheet;
  writeFile(wb_out, Path_info["TestSpec path"]);
  console.log(
    `Testcases file is generated successfully!!\r\nPlease check path: ${Path_info["TestSpec path"]}`
  );
  exec(`start excel "${Path_info["TestSpec path"]}"`);
  return err;
}
