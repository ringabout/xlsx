import unittest

import xlsx


suite "Test parse Excel with dateTime":
  let sheetName = "Sheet1"

  test "Parse Excel with dateTime":
    let data = parseExcel("tests/test_date.xlsx")
    discard data[sheetName]
