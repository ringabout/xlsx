# xlsx
Parse xlsx written by Nim.[WIP]

### Usage

Parse Excel without header.

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

Parse Excel with header.

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

Parse Excel and skip header for data processing.

```nim
import xlsx


let 
  data = parseExcel("tests/test.xlsx", skipHeader = true)
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

convert to csv

```nim
import xlsx


let sheetName = "Sheet2"
let data = parseExcel("tests/test.xlsx")
data[sheetName].toCsv("test.csv", sep = ",")
```

output:

```text
name,grade,age
simon,,14
tom,87,34
```
