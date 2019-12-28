import unittest

import xlsx


suite "test parse Excel":
  let sheetName = "Sheet2"

  test "parse Excel":
    let data = parseExcel("tests/test.xlsx")
    check(data[sheetName].data == @["name", "grade", "age",
        "simon", "", "14", "tom", "87", "34"])

  test "skip header":
    let data = parseExcel("tests/test.xlsx", skipHeaders = true)
    check(data[sheetName].data == @["simon", "", "14", "tom",
        "87", "34"])

  test "get one element from SheetArray":
    let data = parseExcel("tests/test.xlsx")
    check(data[sheetName][1, 0] == "simon")

  test "set one element in SheetArray":
    var data = parseExcel("tests/test.xlsx")
    data[sheetName][1, 0] = "mary" 
    check(data[sheetName][1, 0] == "mary")

  test "toCsv":
    let data = parseExcel("tests/test.xlsx")
    data[sheetName].toCsv("tests/test.csv")

  test "toSeq":
    let data = parseExcel("tests/test.xlsx")
    check(data[sheetName].toSeq(false) == @[@["name", "grade", "age"], @["simon", "", "14"], @["tom", "87", "34"]])
    check(data[sheetName].toSeq(true) == @[@["simon", "", "14"], @["tom", "87", "34"]])