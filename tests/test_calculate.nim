import unittest

import xlsx

suite "read and write excel":
  let sheetName = "Sheet1"

  test "read excel with type int":
    let data = readExcel[int]("tests/test_int.xlsx", sheetName,
        skipHeaders = false)
    check(data.data == @[1, 4, 7, 9, 4, 7, 1, 3, 12, 54, 24, 887])

  test "read excel with type float":
    let data = readExcel[float]("tests/test_int.xlsx", sheetName,
        skipHeaders = false)
    check(data.data == @[1.0, 4.0, 7.0, 9.0, 4.0, 7.0, 1.0, 3.0,
        12.0, 54.0, 24.0, 887.0])

  test "read excel with type string":
    let data = readExcel[string]("tests/test_int.xlsx", sheetName,
        skipHeaders = false)
    check(data.data == @["1", "4", "7", "9", "4", "7", "1", "3",
        "12", "54", "24", "887"])
