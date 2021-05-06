import unittest

import xlsx

suite "Test parse Excel with formula":
  let sheetName = "Sheet1"

  test "Parse transposed rows":
    let data = parseExcel("tests/test_formula.xlsx", skipHeaders=false)
    check(data[sheetName].toSeq(false) == @[@["foo", "bar"], @["foo", ""], @["bar", ""], @["foo", "bar"]])
