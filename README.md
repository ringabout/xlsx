# xlsx
Parse xlsx written by Nim.[WIP]

### Usage

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
