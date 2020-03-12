import unittest

import xlsx


suite "Test parse Excel with date":
  let sheetName = "Sheet1"

  test "Parse Excel with date":
    let data = parseExcel("tests/test_date.xlsx")
    discard data[sheetName]
