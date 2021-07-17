const { Generate_TestSpec } = require("./Gen_TestSpec");
const { Generate_Dxl } = require("./Gen_Dxl");
const { IsDefined, GetAbsPath, getSysArgs } = require("./Func");

function Generate_func(args) {
  let err = "";
  let excel_config_path = "";

  if (IsDefined(args["excel_config"])) {
    [err, excel_config_path] = GetAbsPath(args["excel_config"]);
  } else {
    err += "Please define arg '--excel_config=...'\r\n";
  }

  if (err === "") {
    try {
      err += eval(args["func"])(excel_config_path);
    } catch (e) {
      err += e.stack;
    }
  }

  return err;
}

const sys_args = getSysArgs();
let _Is_gen_TestSpec = true;
let err = "";

if (Object.keys(sys_args).length === 1 && IsDefined(sys_args["help"])) {
  console.log("\t--func={Function_to_call}");
  console.log("\t--excel_config={Excel_config_path}");
}

if (IsDefined(sys_args["func"])) {
  err += Generate_func(sys_args);
} else {
  err += "Please choose the func to use!!! (--func=...)\r\n";
}

if (err !== "") console.log(err);
else console.log("Done!");
