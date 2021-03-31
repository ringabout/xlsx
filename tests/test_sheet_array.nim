import unittest

import xlsx

suite "Test SheetArray":
  let sheetName = "Sheet1"

  test "Get one element from SheetArray":
    let data = parseExcel("tests/test.xlsx")
    check(data[sheetName][1, 0] == "simon")

  test "Set one element to SheetArray":
    var data = parseExcel("tests/test.xlsx")
    data[sheetName][1, 0] = "mary"
    check(data[sheetName][1, 0] == "mary")

  test "SheetArray to csv":
    let data = parseExcel("tests/test.xlsx")
    data[sheetName].toCsv("tests/test.csv")

  test "SheetArray to seq":
    let data = parseExcel("tests/test.xlsx")
    check(data[sheetName].toSeq(false) == @[@["name", "grade", "age"], @[
        "simon", "", "14"], @["tom", "87", "34"]])
    check(data[sheetName].toSeq(true) == @[@["simon", "", "14"], @["tom", "87", "34"]])
