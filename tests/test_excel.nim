import unittest

import xlsx


suite "test parse Excel":
  test "parse Excel":
    let data = parseExcel("tests/test.xlsx")
    check(data == SheetArray(shape: (3, 3), data: @["name", "grade", "age",
        "simon", "", "14", "tom", "87", "34"]))

  test "skip header":
    let data = parseExcel("tests/test.xlsx", skipHeader = true)
    check(data == SheetArray(shape: (2, 3), data: @["simon", "", "14", "tom",
        "87", "34"]))
