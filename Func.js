const { existsSync } = require("fs");
const { utils } = require("xlsx");

module.exports = { Get_OverviewInfo, IsDefined, GetAbsPath, getSysArgs };

function Get_OverviewInfo(List_Info, OverviewInfo, ref) {
  let err = "";
  OverviewInfo["!ref"] = ref;
  let data = utils.sheet_to_csv(OverviewInfo, { blankrows: false });
  if (data[data.length - 1] === "\n") data = data.substring(0, data.length - 1);
  let Split_rows = data.split("\n");
  Split_rows.forEach((item) => {
    let Split_colums = item.split(",");
    if (Split_colums[1] == "")
      err += `Please choose '${Split_colums[0]}'!!\r\n`;
    List_Info[Split_colums[0]] = Split_colums[1];
  });
  return [err, List_Info];
}

function IsDefined(obj) {
  return typeof obj != "undefined";
}

function GetAbsPath(path) {
  let err = "";
  if (!existsSync(path)) {
    err = `Path ${path} NOT exist!!\r\n`;
  }
  return [err, path.replace(/[\/\\]$/, "")];
}

function getSysArgs() {
  const args = {};
  process.argv.slice(2, process.argv.length).forEach((arg) => {
    // long arg
    if (arg.slice(0, 2) === "--") {
      const longArg = arg.split("=");
      const longArgFlag = longArg[0].slice(2, longArg[0].length);
      const longArgValue = longArg.length > 1 ? longArg[1] : true;
      args[longArgFlag] = longArgValue;
    }
    // flags
    else if (arg[0] === "-") {
      const flags = arg.slice(1, arg.length).split("");
      flags.forEach((flag) => {
        args[flag] = true;
      });
    }
  });
  return args;
}
