[![Build Status](https://dev.azure.com/xzsflywind/xlsx/_apis/build/status/xflywind.xlsx?branchName=master)](https://dev.azure.com/xzsflywind/xlsx/_build/latest?definitionId=2&branchName=master)

# xlsx [![nimble](https://raw.githubusercontent.com/yglukhov/nimble-tag/master/nimble.png)](https://github.com/yglukhov/nimble-tag)
Parse xlsx written in Nim.[WIP]

### Docs

Docs in https://xflywind.github.io/xlsx/utils.html

### Usage

#### Parse Excel without header.

```nim
import xlsx


let
  data = parseExcel("tests/test.xlsx")
  sheetName = "Sheet2"
echo data[sheetName]
```

output:

```text
+----------+----------+----------+
|name      |grade     |age       |
|simon     |          |14        |
|tom       |87        |34        |
+----------+----------+----------+
```

#### Parse Excel with header.

```nim
import xlsx


let
  data = parseExcel("tests/test.xlsx", header = true)
  sheetName = "Sheet2"
echo data[sheetName]
```

output:

```text
+----------+----------+----------+
|name      |grade     |age       |
+----------+----------+----------+
|simon     |          |14        |
|tom       |87        |34        |
+----------+----------+----------+
```

#### Parse Excel and skip header for data processing.

```nim
import xlsx


let
  data = parseExcel("tests/test.xlsx", skipHeaders = true)
  sheetName = "Sheet2"
echo data[sheetName]
```

output:

```text
+----------+----------+----------+
|simon     |          |14        |
|tom       |87        |34        |
+----------+----------+----------+
```

#### Convert to Csv

```nim
import xlsx


let sheetName = "Sheet2"
let data = parseExcel("tests/test.xlsx")
data[sheetName].toCsv("tests/test.csv", sep = ",")
```

output:

```text
name,grade,age
simon,,14
tom,87,34
```

#### Loop through rows:
```nim
import xlsx

let sheetName = "Sheet2"
let data = parseExcel("tests/test.xlsx")
let rows = data[sheetName].toSeq(false)
for row in rows:
  echo row
```

output:

```text
@["name", "grade", "age"]
@["simon", "", "14"]
@["tom", "87", "34"]
```

#### Loop through rows and skip headers:
```nim
import xlsx

let sheetName = "Sheet2"
let data = parseExcel("tests/test.xlsx")
let rows = data[sheetName].toSeq(true)
for row in rows:
  echo "Name is: " & row[0]
```

output:

```text
Name is: simon
Name is: tom
```
