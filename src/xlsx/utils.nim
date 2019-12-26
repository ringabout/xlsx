import os, streams, parsexml, parseutils, tables, times, unicode
import strutils except alignLeft

import zip / zipfiles


const
  UpperLetters = {'A' .. 'Z'}
  CharDataOption = {xmlCharData, xmlWhitespace}
let TempDir* = getTempDir() / "docx_windx_tmp"


type
  # newException(SheetDataKindError, "unKnown sheet data kind")
  XlsxError* = object of Exception
  SheetDataKindError* = object of XlsxError
  SheetDataKind* {.pure.} = enum
    Initial, Boolean, Date, Error, InlineStr, Num, SharedString, Formula
  sdk = SheetDataKind
  WorkBook* = Table[string, string]
  ContentTypes* = seq[string]
  SharedStrings* = seq[string]
  SheetData* = object
    case kind: SheetDataKind
    of sdk.Boolean:
      bvalue: string
    of sdk.Date:
      dvalue: string
    of sdk.InlineStr:
      isvalue: string
    of sdk.Num:
      nvalue: string
    of sdk.SharedString:
      svalue: string
    of sdk.Formula:
      fvalue: string
      fnvalue: string
    of sdk.Error:
      error: string
    else:
      discard
  SheetInfo* = tuple
    rows, cols: int
    start: string
  Sheet* = object
    info: SheetInfo
    data: seq[SheetData]
  SheetArray* = object
    shape: tuple[rows: int, cols: int]
    data: seq[string]

proc extractXml*(src: string, dest: string = TempDir) =
  if not existsFile(src):
    raise newException(IOError, "No such file: " & src)
  var z: ZipArchive
  if not z.open(src):
    raise newException(IOError, "[ZIP] Can't open file: " & src)
  z.extractAll(dest)
  z.close()

template `=?=`(a, b: string): bool =
  cmpIgnoreCase(a, b) == 0

proc matchKindName(x: XmlParser, kind: XmlEventKind, name: string): bool {.inline.} =
  x.kind == kind and x.elementName =?= name

proc parseContentTypes*(fileName: string): ContentTypes =
  # open xml file
  var s = newFileStream(fileName, fmRead)
  if s == nil: quit("cannot open the file" & fileName)
  var x: XmlParser
  defer: x.close()
  open(x, s, fileName)

  while true:
    x.next()
    case x.kind
    of xmlElementOpen:
      # catch <Override
      if x.elementName =?= "Override":
        # ignore xmlElementOpen with name "Override"
        x.next()
        # maybe many attrs
        while true:
          case x.kind
          of xmlAttribute:
            # match attr PartName
            if x.attrKey =?= "PartName":
              result.add x.attrValue
          of xmlElementEnd:
            break
          else: discard
          x.next()
    of xmlElementEnd:
      discard
    of xmlEof:
      break # end the world
    else:
      discard

proc parseStringTable*(x: var XmlParser, res: var seq[string]) =
  var count = 0
  while true:
    # match <si>
    if x.matchKindName(xmlElementStart, "si"):
      # ignore <si>
      x.next()
      # match attrs in <si>
      # maybe <t> , <phoneticPr and so on.
      while true:
        # macth <t>
        if x.matchKindName(xmlElementStart, "t"):
          # ignore <t>
          x.next()
          # match charData in <t>
          while x.kind in CharDataOption:
            res[count] &= x.charData
            x.next()
          # seq index
          inc(count)
          # if match chardata, end loop
          break
        else:
          discard
        # switch to the next element
        x.next()
    elif x.kind == xmlEof: # end the world
      break
    else:
      discard
    x.next()

proc parseSharedString*(fileName: string): SharedStrings =
  # open xml file
  var s = newFileStream(fileName, fmRead)
  if s == nil: quit("cannot open the file" & fileName)
  var x: XmlParser
  defer: x.close()
  open(x, s, fileName, {reportWhitespace})

  while true:
    x.next()
    case x.kind
    of xmlElementOpen:
      # match <sst>
      if x.elementName =?= "sst":
        # match attrs in <sst>
        while true:
          # ignore <sst
          x.next()
          case x.kind
          of xmlAttribute:
            # match attr count
            if x.attrKey =?= "count":
              # initial seq that stores strings
              result = newSeq[string](parseInt(x.attrValue))
          of xmlElementStart:
            # match <si>, then parse StringTable
            x.parseStringTable(result)
            break
          else:
            discard
    of xmlEof:
      break # end the world
    else:
      discard

proc praseWorkBook*(fileName: string): WorkBook =
  # open xml file
  var s = newFileStream(fileName, fmRead)
  if s == nil: quit("cannot open the file" & fileName)
  var x: XmlParser
  defer: x.close()
  open(x, s, fileName)

  var name: string
  while true:
    x.next()
    if x.matchKindName(xmlElementStart, "sheets"):
    # catch <sheets>
      # ignore sheets
      x.next()
      # parse name: sheetId
      while x.matchKindName(xmlElementOpen, "sheet"):
        # ignore xmlElementOpen with name "sheet"
        x.next()
        # maybe many sheets
        while true:
          case x.kind
          of xmlAttribute:
            # parse name -> "Sheet1"
            if x.attrKey =?= "name":
              name = x.attrValue
            # parse sheetId -> "s1"
            if x.attrKey =?= "sheetId":
              result[name] = x.attrValue
          of xmlElementEnd:
            break
          else: discard
          # ignore element
          x.next()
        # ignore xmlElementEnd />
        x.next()
      # over
      break

proc parseSheetDataBoolean(x: var XmlParser): SheetData {.inline.} =
  result = SheetData(kind: sdk.Boolean)
  # ignore <v>
  x.next()
  while x.kind == xmlCharData:
    result.bvalue &= x.charData
    x.next()
  # ignore </v>
  x.next()
  # point to </c>

proc parseSheetDataNum(x: var XmlParser): SheetData {.inline.} =
  result = SheetData(kind: sdk.Num)
  # ignore <v>
  x.next()
  while x.kind == xmlCharData:
    result.nvalue &= x.charData
    x.next()
  # ignore </v>
  x.next()
  # point to </c>

proc parseSheetDataSharedString(x: var XmlParser): SheetData {.inline.} =
  result = SheetData(kind: sdk.SharedString)
  # ignore <v>
  x.next()
  while x.kind in CharDataOption:
    result.svalue &= x.charData
    x.next()
  # ignore </v>
  x.next()
  # point to </c>

proc parseSheetDataFormula(x: var XmlParser): SheetData {.inline.} =
  result = SheetData(kind: sdk.Formula)
  # ignore <f>
  x.next()
  while x.kind in CharDataOption:
    result.fvalue &= x.charData
    x.next()
  # ignore </f>
  x.next()
  # ignore <v>
  x.next()
  while x.kind == xmlCharData:
    result.fnvalue &= x.charData
    x.next()
  # ignore </v>
  x.next()
  # point to </c>

proc parseSheetDataInlineStr(x: var XmlParser): SheetData {.inline.} =
  result = SheetData(kind: sdk.InlineStr)
  # ignore <is>
  x.next()
  # ignore <t>
  while x.kind in CharDataOption:
    result.isvalue &= x.charData
    x.next()
  # ignore </t>
  x.next()
  # ignore </is>
  x.next()
  # point to </c>s

proc parseSheetDate(x: var XmlParser): SheetData {.inline.} =
  result = SheetData(kind: sdk.Date)
  # ignore <v>
  x.next()
  while x.kind in CharDataOption:
    result.nvalue &= x.charData
    x.next()
  # ignore </v>
  x.next()
  # point to </c>

proc calculatePolynomial(a: string): int =
  for i in 0 .. a.high:
    # !Maybe raise alpha
    result = result * 27 + (ord(a[i]) - ord('A') + 1)

proc parseDimension*(x: string): SheetInfo =
  # A1:B3
  var
    rowLeft, rowRight: int
    colLeft, colRight: string
    row, col: int
    pos = 0
  pos += parseWhile(x, colLeft, UpperLetters, pos)
  pos += parseInt(x, rowLeft, pos)
  pos += skip(x, ":", pos)
  pos += parseWhile(x, colRight, UpperLetters, pos)
  pos += parseInt(x, rowRight, pos)
  row = rowRight - rowLeft + 1
  col = calculatePolynomial(colRight) - calculatePolynomial(colLeft) + 1
  result = (row, col, colLeft & $rowLeft)

proc parsePos*(x: string, s: SheetInfo): int =
  var
    rowRight, rowLeft: int
    colRight, colLeft: string
    row, col: int
    pos = 0
  pos += parseWhile(x, colRight, UpperLetters, pos)
  pos += parseInt(x, rowRight, pos)
  pos = 0
  pos += parseWhile(s.start, colLeft, UpperLetters, pos)
  pos += parseInt(s.start, rowLeft, pos)
  row = rowRight - rowLeft
  col = calculatePolynomial(colRight) - calculatePolynomial(colLeft)
  result = row * s.cols + col

proc parseDataKind(s: string): SheetDataKind {.inline.} =
  # <c r="A2" t="s">
  ## convert string to SheetDataKind
  result = case s
  of "b": sdk.Boolean
  of "d": sdk.Date
  of "e": sdk.Error
  of "inlineStr": sdk.InlineStr
  of "n": sdk.Num
  of "s": sdk.SharedString
  of "str": sdk.Formula
  else: raise

# proc parseValue(x: var XmlParser): string =
#   while true:
#     if x.matchKindName(xmlElementOpen, "v"):

proc parseRowMetaData(x: var XmlParser, s: SheetInfo): (int, SheetData) =
  # <c r="A2" t="s">
  var
    pos: int
    kind: SheetDataKind
    value: SheetData
  while true:
    x.next()
    case x.kind
    of xmlAttribute:
      # catch key "r"
      if x.attrKey =?= "r":
        pos = parsePos(x.attrValue, s)
      # catch key "t"
      elif x.attrKey =?= "t":
        kind = parseDataKind(x.attrValue)
    of xmlElementClose, xmlEof:
      break
    else:
      discard
  # if omit key "t", it should be sdk.Num kind.
  if kind == sdk.Initial:
    kind = sdk.Num
  case kind
  of sdk.Boolean:
    while true:
      x.next()
      case x.kind
      of xmlElementStart:
        if x.elementName =?= "v":
          value = x.parseSheetDataBoolean
      of xmlElementEnd:
        if x.elementName =?= "c":
          break
      of xmlEof:
        break
      else:
        discard
  of sdk.Date:
    while true:
      x.next()
      case x.kind
      of xmlElementStart:
        if x.elementName =?= "v":
          value = x.parseSheetDate
      of xmlElementEnd:
        if x.elementName =?= "c":
          break
      of xmlEof:
        break
      else:
        discard
  of sdk.InlineStr:
    while true:
      x.next()
      case x.kind
      of xmlElementStart:
        if x.elementName =?= "is":
          value = x.parseSheetDataInlineStr
      of xmlElementEnd:
        if x.elementName =?= "c":
          break
      of xmlEof:
        break
      else:
        discard
  of sdk.Num:
    while true:
      x.next()
      case x.kind
      of xmlElementStart:
        if x.elementName =?= "v":
          value = x.parseSheetDataNum
      of xmlElementEnd:
        if x.elementName =?= "c":
          break
      of xmlEof:
        break
      else:
        discard
  of sdk.SharedString:
    while true:
      x.next()
      case x.kind
      of xmlElementStart:
        if x.elementName =?= "v":
          value = x.parseSheetDataSharedString
      of xmlElementEnd:
        if x.elementName =?= "c":
          break
      of xmlEof:
        break
      else:
        discard
  of sdk.Formula:
    while true:
      x.next()
      case x.kind
      of xmlElementStart:
        if x.elementName =?= "f":
          value = x.parseSheetDataFormula
      of xmlElementEnd:
        if x.elementName =?= "c":
          break
      of xmlEof:
        break
      else:
        discard
  else:
    raise newException(XlsxError, "not support" & $kind)
  result = (pos, value)

proc parseRowData*(x: var XmlParser, s: var Sheet) =
  while true:
    x.next()
    case x.kind
    of xmlElementOpen:
      if x.elementName =?= "c":
        let (pos, value) = parseRowMetaData(x, s.info)
        s.data[pos] = value
    of xmlEof:
      break
    else:
      discard
  # ignore />
  x.next()

proc parseSheet*(fileName: string): Sheet =
  # open xml file
  var s = newFileStream(fileName, fmRead)
  if s == nil: quit("cannot open the file" & fileName)
  var x: XmlParser
  defer: x.close()
  open(x, s, fileName, {reportWhitespace})

  x.next()
  # parse Dimension
  while true:
    x.next()
    if x.matchKindName(xmlElementOpen, "dimension"):
      x.next()
      while true:
        case x.kind
        of xmlAttribute:
          if x.attrKey =?= "ref":
            result.info = parseDimension(x.attrValue)
            result.data = newSeq[SheetData](result.info.rows * result.info.cols)
        of xmlElementEnd:
          break
        else:
          discard
        x.next()
      # discard />
      x.next()
      break
  # parse data
  while true:
    x.next()
    case x.kind
    of xmlElementStart:
      if x.elementName =?= "sheetData":
        # ignore <sheetData>
        x.next()
        while true:
          case x.kind
          of xmlElementOpen:
            if x.elementName =?= "row":
              parseRowData(x, result)
          of xmlEof:
            break
          else:
            discard
          x.next()
    of xmlEof:
      break
    else:
      discard

proc getKindString(item: SheetData, str: SharedStrings): string =
  case item.kind
  of sdk.SharedString:
    result = str[parseInt(item.svalue)]
  of sdk.Num:
    result = item.nvalue
  else:
    result = ""

proc xlsxToCsv(s: Sheet, str: SharedStrings, fileName = "test.csv", sep = ",") =
  let f = open(fileName, fmWrite)
  defer: f.close()
  let (rows, cols, _) = s.info
  for i in 0 ..< rows:
    var res = ""
    for j in 0 ..< cols:
      let item = s.data[i * cols + j]
      res.add getKindString(item, str)
      if j < cols - 1:
        res.add sep
    f.writeLine res

iterator getAlign*(s: Sheet, str: SharedStrings, sep = ","): string =
  let (rows, cols, _) = s.info
  for i in 0 ..< rows:
    var res = "|"
    for j in 0 ..< cols:
      let item = s.data[i * cols + j]
      res.add getKindString(item, str)
    yield res

proc plotSym(cols: int, width = 10): string =
  #+------------+------------+
  for i in 0 ..< cols:
    result.add "+"
    result.add repeat("-", width)
  result.add "+"

iterator get*(s: Sheet, str: SharedStrings, sep = "|", width = 10): string =
  let (rows, cols, _) = s.info
  for i in 0 ..< rows:
    var res = sep
    for j in 0 ..< cols:
      let item = s.data[i * cols + j]
      res.add alignLeft(getKindString(item, str), width)
      res.add sep
    yield res

proc getSheetArray(s: Sheet, str: SharedStrings): SheetArray =
  let (rows, cols, _) = s.info
  result.shape = (rows, cols)
  result.data = newseq[string](rows * cols)
  for idx, item in s.data:
    result.data[idx] = getKindString(item, str)

proc parseExcel(fileName: string): SheetArray =
  extractXml(fileName)
  defer: removeDir(TempDir)
  let 
    contentTypes = parseContentTypes(TempDir / "[Content_Types].xml")
    workbook = praseWorkBook(TempDir / "xl/workbook.xml")
    sharedstring = parseSharedString("files/td/xl/sharedStrings.xml")
    sheet = parseSheet("files/td/xl/worksheets/sheet2.xml")

  result = getSheetArray(sheet, sharedstring)
  # echo plotSym(sheet.info.cols)
  # for item in sheet.get(sharedstring, sep = "|"):
  #   echo item
  # echo plotSym(sheet.info.cols)
  # xlsxToCsv(sheet, sharedstring, sep = ", ")


when isMainModule:
  import timeit
  timeOnce("test"):
    echo parseExcel("./test.xlsx")
  # echo repeat("-", 40)
  # echo parsePos("A2", (3, 2, "A1"))
  # echo repeat("-", 40)
