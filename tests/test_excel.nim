import unittest

import xlsx


suite "test parse Excel":
  let sheetName = "Sheet2"

  test "parse Excel":
    let data = parseExcel("tests/test.xlsx")
    check(data[sheetName].data == @["name", "grade", "age",
        "simon", "", "14", "tom", "87", "34"])

  test "skip header":
    let data = parseExcel("tests/test.xlsx", skipHeader = true)
    check(data[sheetName].data == @["simon", "", "14", "tom",
        "87", "34"])

  test "toCsv":
    let data = parseExcel("tests/test.xlsx")
    data[sheetName].toCsv("tests/test.csv")

