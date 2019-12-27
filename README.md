# xlsx
Parse xlsx written by Nim.[WIP]

### Usage

Parse Excel without header.

```nim
import xlsx


let data = parseExcel("tests/test.xlsx")
echo data
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


let data = parseExcel("tests/test.xlsx", header = true)
echo data
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


let data = parseExcel("tests/test.xlsx", skipHeader = true)
echo data
```

output:

```text
+----------+----------+----------+
|simon     |          |14        |
|tom       |87        |34        |
+----------+----------+----------+
```
