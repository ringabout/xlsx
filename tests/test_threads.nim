when compileOption("threads"):
  import std/threadpool 

  import unittest

  import xlsx

  proc checkData(sheetName: string): bool =
    let data = parseExcel("tests/test.xlsx")
    return data[sheetName].data == @["name", "grade", "age",
      "simon", "", "14", "tom", "87", "34"]

    suite "Test parse Excel with threads":
      let sheetName = "Sheet1"
  
    test "Parse Excel with threads":
      when defined(gcArc) or defined(gcOrc):
        let fv = spawn(parseExcel("tests/test.xlsx"))
        let data = ^fv
        check(data[sheetName].data == @["name", "grade", "age",
          "simon", "", "14", "tom", "87", "34"])
      else:
        let fv = spawn(checkData(sheetName))
        check(^fv)

      test "Get all sheet names with threads":
        let fv = spawn(parseAllSheetName("tests/test.xlsx"))
        let data = ^fv
        check(data == @["Sheet2", "Sheet1"])
