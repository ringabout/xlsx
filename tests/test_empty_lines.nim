import unittest

import xlsx

suite "skip empty lines":
  let sheetName = "Cx"
  test "parse Excel and skip empty lines":
    let data = parseExcel("tests/test_empty_lines.xlsx", header = false, trailingRows = true)
    let rows = data[sheetName].toSeq(false)
    check rows.len == 10
    check rows[0].len == 4