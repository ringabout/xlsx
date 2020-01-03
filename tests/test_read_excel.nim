import unittest

import xlsx

suite "Read excel with types":
  let sheetName = "Sheet1"

  test "Read excel with type int":
    let data = readExcel[int]("tests/test_read_excel.xlsx", sheetName,
        skipHeaders = false)
    check(data.data == @[1, 4, 7, 9, 4, 7, 0, 3, 12, 54, 24, 887])

  test "Read excel with type float":
    let data = readExcel[float]("tests/test_read_excel.xlsx", sheetName,
        skipHeaders = false)
    check(data.data == @[1.0, 4.0, 7.0, 9.0, 4.0, 7.0, 0.0, 3.0,
        12.0, 54.0, 24.0, 887.0])

  test "Read excel with type string":
    let data = readExcel[string]("tests/test_read_excel.xlsx", sheetName,
        skipHeaders = false)
    check(data.data == @["1", "4", "7", "9", "4", "7", "", "3",
        "12", "54", "24", "887"])
