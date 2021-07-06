const { writeFile, readFile, utils } = require("xlsx");
const mapping = require("./mapping.json");
const TestSpec = require("./TestSpec.json");
var wb = readFile(
  "M:\\my_apps\\JS_Workspace\\Testcases_gen\\System and Software Requirements and Structure_DAS_SUZ_02_YEA - ENG2 ENG3 ENG4 ENG5.xlsx"
);

let output = [];

function Update_content(data, req_info) {
  let _regex = /\{\w+\}/g;
  let patterns = data.match(_regex);
  let temp_data = data;
  if ((patterns || []).length > 0) {
    patterns.forEach((element) => {
      temp_data = temp_data.replace(element, req_info[mapping[element]]);
    });
  }
  return temp_data;
}

function Init_content(TestSpec_keys) {
  let args = [];
  for (var i in TestSpec_keys) {
    let element = TestSpec_keys[i];
    args[element] = "x";
  }
  return args;
}
var data = utils.sheet_to_json(wb.Sheets["System and Software Requirement"], {
  blankrows: false,
});

let TestSpec_keys = Object.keys(TestSpec);

output.push(Init_content(TestSpec_keys));

data.forEach((el) => {
  let args = [];
  for (var i in TestSpec_keys) {
    let element = TestSpec_keys[i];
    args[element] = Update_content(TestSpec[element], el);
  }
  output.push(args);
});
let output_sheet = utils.json_to_sheet(output);
let wb_out = utils.book_new();
wb_out.SheetNames.push("Test spec");
wb_out.Sheets["Test spec"] = output_sheet;
writeFile(wb_out, "TestSpec_out.xls");
