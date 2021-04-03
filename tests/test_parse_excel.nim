import unittest

import xlsx


suite "Test parse Excel":
  let sheetName = "Sheet1"

  test "Parse Excel":
    let data = parseExcel("tests/test.xlsx")
    check(data[sheetName].data == @["name", "grade", "age",
        "simon", "", "14", "tom", "87", "34"])

  test "Parse Excel and skip headers":
    let data = parseExcel("tests/test.xlsx", skipHeaders = true)
    check(data[sheetName].data == @["simon", "", "14", "tom",
        "87", "34"])

  test "Get all sheet names":
    let data = parseAllSheetName("tests/test.xlsx")
    check(data == @["Sheet1", "Sheet2"])

  test "Read Excel by lines":
    for line in lines("tests/test.xlsx", "Sheet1"):
      discard line
