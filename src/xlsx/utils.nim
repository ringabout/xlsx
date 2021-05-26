import os, streams, parsexml, parseutils, tables, times, strutils
import format
import sugar

import zip / zipfiles



const
  UpperLetters = {'A' .. 'Z'}
  CharDataOption = {xmlCharData, xmlWhitespace}
  DEFAULT_WORKBOOK_PATH = "xl/workbook.xml"
let TempDir* = getTempDir() / "xlsx_windx_tmp" ## temp dir for all extracted xml files from Excel


type
  XlsxError* = object of Exception
  NotExistsXlsxFileError* = object of XlsxError
  InvalidXlsxFileError* = object of XlsxError
  NotFoundSheetError* = object of XlsxError
  UnKnownSheetDataKindError* = object of XlsxError

  SheetDataKind* {.pure.} = enum
    Initial, Boolean, Date, Time, Error, InlineStr, Num, SharedString, Formula
  sdk = SheetDataKind

  WorkBook = object
    data: Table[string, string]
    date1904: bool
  ContentTypes = Table[string, string]
  Relationships = Table[string, string]
  SharedStrings = seq[string]
  Styles = seq[string]

  SheetData = object
    case kind: SheetDataKind
    of sdk.Boolean:
      bvalue: string
    of sdk.Date:
      dvalue: string
    of sdk.Time:
      tValue: string
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

  SheetInfo = tuple
    rows, cols: int
    start: string
    date1904: bool

  Sheet = object
    info: SheetInfo
    data: seq[SheetData]

  SheetArray* = object ## SheetArray
    shape*: tuple[rows: int, cols: int]
    data*: seq[string]
    header*: bool
    colType*: seq[SheetDataKind]

  SheetTensor*[T] = object
    shape*: tuple[rows: int, cols: int]
    data*: seq[T]

  SheetTable* = object
    data*: Table[string, SheetArray]

when defined(windows):
  {.passl: "-lz".}

proc extractXml*(src: string, dest: string = TempDir) {.inline.} =
  ## extract xml file from excel using zip,
  ## default path is TempDir.
  if not existsFile(src):
    raise newException(NotExistsXlsxFileError, "No such xlsx file: " & src)
  var z: ZipArchive
  if not z.open(src):
    raise newException(InvalidXlsxFileError, "[ZIP] Can't open xlsx file: " & src)
  z.extractAll(dest)
  z.close()

template checkIndex(cond: untyped, msg = "") =
  when compileOption("boundChecks"):
    {.line.}:
      if not cond:
        raise newException(IndexError, msg)

template `=?=`(a, b: string): bool =
  cmpIgnoreCase(a, b) == 0

proc matchKindName(x: XmlParser, kind: XmlEventKind, name: string): bool {.inline.} =
  x.kind == kind and x.elementName =?= name

proc parseNameSpace(x: var XmlParser, nameSpaceDelimiter = ":"): string {.inline.} =
  x.next()
  if x.kind == xmlPI:
    x.next()
  if x.kind == xmlElementOpen:
    let
      name = x.elementName
      idx = rfind(name, nameSpaceDelimiter)
    if idx > 0:
      result = name[0 .. idx]
    elif idx == -1:
      result = ""

proc parseOverride(x: var XmlParser, res: var ContentTypes) {.inline.} =
  # ignore xmlElementOpen with name "Override"
  # maybe many attrs
  while true:
    x.next()
    case x.kind
    of xmlAttribute:
      # match attr PartName
      if x.attrKey =?= "PartName":
        let
          path = x.attrValue
          name = path.splitFile.name
        res[name] = path
    of xmlElementEnd:
      break
    else:
      discard

proc parseContentTypes(fileName: string): ContentTypes {.inline.} =
  # TODO maybe add namespaceUri
  # open xml file
  var s = newFileStream(fileName, fmRead)
  if s == nil: quit("Unable to read file: " & fileName)
  var x: XmlParser
  open(x, s, fileName)
  defer: x.close()

  let name = parseNameSpace(x)

  while true:
    case x.kind
    of xmlElementOpen:
      # catch <namespace:Override
      if x.elementName =?= (name & "Override"):
        x.parseOverride(result)
    of xmlElementEnd:
      discard
    of xmlEof:
      break # end the world
    else:
      discard
    x.next()

  let workBookKey = "workbook"
  if workBookKey notin result or result[workBookKey] == "":
    result[workBookKey] = DEFAULT_WORKBOOK_PATH

proc parseRelationship(x: var XmlParser, res: var Relationships) {.inline.} =
  var id, target: string

  while true:
    x.next()
    case x.kind
    of xmlAttribute:
      # match attr Id
      if x.attrKey =?= "Id":
        id = x.attrValue
      # match attr Target
      if x.attrKey =?= "Target":
        target = x.attrValue
    of xmlElementEnd:
      break
    else:
      discard

  res[id] = target

proc parseRelationships(fileName: string): Relationships {.inline.} =
  var s = newFileStream(fileName, fmRead)
  if s == nil: quit("Unable to read file: " & fileName)
  var x: XmlParser
  open(x, s, fileName)
  defer: x.close()

  let name = parseNameSpace(x)

  while true:
    case x.kind
    of xmlElementOpen:
      # catch <namespace:Relationship
      if x.elementName =?= (name & "Relationship"):
        x.parseRelationship(result)
    of xmlElementEnd:
      discard
    of xmlEof:
      break
    else:
      discard
    x.next()

proc parseTData(x: var XmlParser, res: var seq[string], escapeStrings: bool,
    count: var int) {.inline.} =
  # ignore <t>
  x.next()
  # match charData in <t>
  while x.kind in CharDataOption:
    res[count] &= x.charData
    x.next()
  # escape strings
  if escapeStrings:
    res[count] = escape(res[count])
  # seq index
  inc(count)

proc parseStringTable(x: var XmlParser, res: var seq[string],
    escapeStrings: bool) {.inline.} =
  var count = 0
  while true:
    # match <si>
    if x.matchKindName(xmlElementStart, "si"):
      # ignore <si>
      # match attrs in <si>
      # maybe <t> , <phoneticPr and so on.
      while true:
        x.next()
        # macth <t>
        case x.kind
        of xmlElementStart:
          if x.elementName =?= "t":
            x.parseTData(res, escapeStrings, count)
            # if match chardata, end loop
            break
        of xmlElementOpen:
          if x.elementName =?= "t":
            while x.kind != xmlElementClose:
              x.next()
            x.parseTData(res, escapeStrings, count)
            # if match chardata, end loop
            break
        else:
          discard
        # switch to the next element
    elif x.kind == xmlEof: # end the world
      break
    else:
      discard
    x.next()

proc parseSharedString(fileName: string,
    escapeStrings = false): SharedStrings {.inline.} =
  # open xml file
  var s = newFileStream(fileName, fmRead)
  if s == nil: quit("Unable to read file: " & fileName)
  var x: XmlParser
  open(x, s, fileName, {reportWhitespace})
  defer: x.close()

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
            if x.attrKey =?= "uniqueCount":
              # initial seq that stores strings
              result = newSeq[string](parseInt(x.attrValue))
          of xmlElementStart:
            # match <si>, then parse StringTable
            x.parseStringTable(result, escapeStrings)
            break
          else:
            discard
    of xmlEof:
      break # end the world
    else:
      discard

proc parseNumFmts(x: var XmlParser): TableRef[string, string] {.inline.} =
  # -> parse numFmts
  # <numFmts count="2">
  # <numFmt formatCode="dd"-"mm"-"yyyy" "hh:mm:ss" numFmtId="176"/>
  # <numFmt formatCode="hh:mm:ss" numFmtId="177"/>
  # </numFmts>
  new result
  while true:
    x.next()
    case x.kind
    of xmlElementOpen:
      if x.elementName =?= "numFmt":
        var key, value: string
        while true:
          x.next()
          case x.kind
          of xmlAttribute:
            if x.attrKey =?= "numFmtId":
              key = x.attrValue
            elif x.attrKey =?= "formatCode":
              value = x.attrValue.toLower
          of xmlElementEnd:
            result[key] = value
            break
          of xmlEof:
            # Maybe raise error
            break
          else:
            discard
    of xmlElementEnd:
      if x.elementName =?= "numFmts":
        break
    of xmlEof:
      break
    else:
      discard

proc parseCellXfs(x: var XmlParser, count: int, t: TableRef[string,
    string]): seq[string] {.inline.} =
  # -> parse cellXfs
  # <cellXfs count="6">
  # <xf numFmtId="0" borderId="0" fillId="0" fontId="0" xfId="0"/>
  # <xf numFmtId="14" borderId="0" fillId="0" fontId="0" xfId="0" applyNumberFormat="1"/>
  # <xf numFmtId="177" borderId="1" fillId="2" fontId="0" xfId="0" applyNumberFormat="1" applyBorder="1" applyFill="1"/>
  # </cellXfs>
  # open xml file
  result = newSeqOfCap[string](count)
  while true:
    x.next()
    case x.kind
    of xmlElementOpen:
      if x.elementName =?= "xf":
        while true:
          x.next()
          case x.kind
          of xmlAttribute:
            if x.attrKey =?= "numFmtId":
              if x.attrValue in t:
                result.add move(t[x.attrValue])
              elif x.attrValue in STANDARD_FORMATS:
                result.add STANDARD_FORMATS[x.attrValue]
              else:
                result.add "none"
          of xmlElementEnd:
            break
          of xmlEof:
            # Maybe raise error
            break
          else:
            discard
    of xmlElementEnd:
      if x.elementName =?= "cellXfs":
        break
    of xmlEof:
      break
    else:
      discard

proc parseStyles(fileName: string): Styles {.inline.} =
  # parse numFmts and cellXfs
  var s = newFileStream(fileName, fmRead)
  if s == nil: quit("Unable to read file: " & fileName)
  var x: XmlParser
  open(x, s, fileName, {reportWhitespace})
  defer: x.close()
  var table = newTable[string, string]()

  # open xml file
  while true:
    x.next()
    case x.kind
    of xmlElementOpen:
      # <numFmts  first
      if x.elementName =?= "numFmts":
        table = x.parseNumFmts
      # <cellXfs  second
      elif x.elementName =?= "cellXfs":
        var count: string
        while true:
          x.next()
          case x.kind
          of xmlAttribute:
            if x.attrKey =?= "count":
              count = x.attrValue
          of xmlElementClose, xmlEof:
            break
          else:
            discard
        result = x.parseCellXfs(count.parseInt, table)
    of xmlEof:
      break # end the world
    else:
      discard

proc parseSheetNameInWorkBook(x: var XmlParser, nameSpace: string): Table[
    string, string] {.inline.} =
  var name: string

  while x.matchKindName(xmlElementOpen, nameSpace & "sheet"):
    # ignore xmlElementOpen with name "sheet"
    x.next()
    # maybe many sheets
    while true:
      case x.kind
      of xmlAttribute:
        # parse name -> "Sheet1"
        if x.attrKey =?= "name":
          name = x.attrValue
        # parse r:id -> "rId1"
        if x.attrKey =?= "r:id":
          result[name] = x.attrValue
      of xmlElementEnd:
        break
      else: discard
      # ignore element
      x.next()
    # ignore xmlElementEnd />
    x.next()

proc parseDateInWorkBook(x: var XmlParser, nameSpace: string): bool {.inline.} =
  # catch <workbookPr/> or <workbookPr date1904="1"/>
  while true:
    x.next()
    case x.kind
    of xmlAttribute:
      if x.attrKey =?= nameSpace & "date1904" and parseBool(x.attrValue):
        result = true
    of xmlElementEnd, xmlEof:
      break
    else:
      discard

proc parseWorkBook(fileName: string): WorkBook {.inline.} =
  # open xml file
  var s = newFileStream(fileName, fmRead)
  if s == nil: quit("Unable to read file: " & fileName)
  var x: XmlParser
  open(x, s, fileName)
  defer: x.close()

  let name = x.parseNameSpace

  while true:
    case x.kind
    of xmlElementStart:
      if x.elementName =?= name & "sheets":
        x.next()
        result.data = x.parseSheetNameInWorkBook(name)
        break
    of xmlElementOpen:
      if x.elementName =?= name & "workbookPr":
        # ignore <workbookPr
        result.date1904 = x.parseDateInWorkBook(name)
    of xmlEof:
      break
    else:
      discard
    x.next()

proc parseSheetDataBoolean(x: var XmlParser): SheetData {.inline.} =
  result = SheetData(kind: sdk.Boolean)
  # ignore <v>
  x.next()
  while x.kind == xmlCharData:
    result.bvalue &= x.charData
    x.next()
  # ignore </v>
  # x.next()
  # # point to </c>

proc parseSheetDataNum(x: var XmlParser): SheetData {.inline.} =
  result = SheetData(kind: sdk.Num)
  # ignore <v>
  x.next()
  while x.kind == xmlCharData:
    result.nvalue &= x.charData
    x.next()
  # ignore </v>
  # x.next()
  # # point to </c>

proc parseSheetDataSharedString(x: var XmlParser): SheetData {.inline.} =
  result = SheetData(kind: sdk.SharedString)
  # ignore <v>
  x.next()
  while x.kind in CharDataOption:
    result.svalue &= x.charData
    x.next()
  # ignore </v>
  # x.next()
  # # point to </c>

proc parseSheetDataFormula(x: var XmlParser): SheetData {.inline.} =
  result = SheetData(kind: sdk.Formula)
  # ignore <f>, <t>, <ref>, </ref>
  while not (x.kind in CharDataOption):
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
  # x.next()
  # # point to </c>

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
  # x.next()
  # # point to </c>

proc parseSheetDate(x: var XmlParser): SheetData {.inline.} =
  result = SheetData(kind: sdk.Date)
  # ignore <v>
  x.next()
  while x.kind in CharDataOption:
    result.dvalue &= x.charData
    x.next()
  # ignore </v>
  # x.next()
  # # point to </c>

proc parseSheetTime(x: var XmlParser): SheetData {.inline.} =
  result = SheetData(kind: sdk.Time)
  # ignore <v>
  x.next()
  while x.kind in CharDataOption:
    result.tvalue &= x.charData
    x.next()
  # ignore </v>
  # x.next()
  # # point to </c>

proc calculatePolynomial(a: string): int {.inline.} =
  for i in 0 .. a.high:
    result = result * 27 + (ord(a[i]) - ord('A') + 1)

proc parseDimension(x: string, date1904: bool): SheetInfo {.inline.} =
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
  result = (row, col, colLeft & $rowLeft, date1904)

proc parsePos(x: string, s: SheetInfo): int {.inline.} =
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

proc parseDataKind(s: string): (SheetDataKind, string) {.inline.} =
  # <c r="A2" t="s">
  ## convert string to SheetDataKind
  result = case s
  of "b": (sdk.Boolean, "bool")
  of "d": (sdk.Date, "date")
  of "e": (sdk.Error, "error")
  of "inlineStr": (sdk.InlineStr, "string")
  of "n": (sdk.Num, "num")
  of "s": (sdk.SharedString, "string")
  of "str": (sdk.Formula, "formula")
  else:
    raise newException(UnKnownSheetDataKindError, "unsupport sheet data kind")

proc parseAttr(s: string, styles: Styles): (SheetDataKind, string) {.inline.} =
  let
    style = styles[s.parseInt]

  var kind: string
  if style in FORMATS:
    kind = FORMATS[style]
  else:
    kind = "date"

  case kind
  of "float":
    result = (sdk.Num, style)
  of "date":
    result = (sdk.Date, style)
  of "time":
    result = (sdk.Time, style)
  of "percentage":
    # TODO may add new kind
    result = (sdk.Initial, style)
  else:
    result = (sdk.Initial, style)

proc toDuration(x: string): Duration {.inline.} =
  var
    pos: int
    tok: string
  pos += x.parseUntil(tok, {'.'}, pos)
  let
    intPart = parseInt(tok)

  var floatPart: float

  if pos > x.len:
    # in seconds
    floatPart = parseFloat(x[pos ..< ^1]) * 24 * 3600

  result = initDuration(days = intPart, seconds = int(floatPart))

proc parseRowMetaData(x: var XmlParser, s: SheetInfo, styles: Styles): (int, SheetData) =
  # <c r="A2" t="s">
  var
    pos: int
    kind: SheetDataKind
    value: SheetData
    style: string
  while true:
    x.next()
    case x.kind
    of xmlAttribute:
      case x.attrKey
      # catch key "r"
      of "r":
        pos = parsePos(x.attrValue, s)
      # catch key "t"
      of "t":
        (kind, style) = parseDataKind(x.attrValue)
      of "s":
        (kind, style) = parseAttr(x.attrValue, styles)
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
          let dur = value.dvalue
          if s.date1904:
            value.dvalue = $(initDateTime(1, mJan, 1904, 0, 0, 0, local()) +
                toDuration(dur))
          else:
            value.dvalue = $(initDateTime(30, mDec, 1899, 0, 0, 0, local()) +
                toDuration(dur))
      of xmlElementEnd:
        if x.elementName =?= "c":
          break
      of xmlEof:
        break
      else:
        discard
  of sdk.Time:
    while true:
      x.next()
      case x.kind
      of xmlElementStart:
        if x.elementName =?= "v":
          value = x.parseSheetTime
          let
            originTime = parseFloat(value.tvalue) * 24 * 3600
            durTime = initDuration(seconds = int(originTime)).toParts
            durDateTime = initDateTime(1, mJan, 2020,
                durTime[Hours], durTime[Minutes], durTime[Seconds], local())
          value.tvalue = $(durDateTime.format("HH:mm:ss"))
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
      of xmlElementOpen:
        if x.elementName =?= "f":
          value = x.parseSheetDataFormula
      of xmlElementStart:
        if x.elementName =?= "f":
          value = x.parseSheetDataFormula
        elif x.elementName =?= "v":
          value = x.parseSheetDataNum
      of xmlElementEnd:
        if x.elementName =?= "c":
          break
      of xmlEof:
        break
      else:
        discard
  else:
    value = SheetData(kind: sdk.Error, error: "error")
    # raise newException(XlsxError, "not support " & $kind)
  result = (pos, value)

proc parseRowData(x: var XmlParser, s: var Sheet, styles: Styles, position: var int) {.inline.} =
  
  while true:
    x.next()
    case x.kind
    of xmlElementOpen:
      if x.elementName =?= "c":
        let (pos, value) = parseRowMetaData(x, s.info, styles)
        position = pos
        s.data[pos] = value
    of xmlEof:
      break
    else:
      discard
  # ignore />
  x.next()

proc parseSheet(fileName: string, styles: Styles, date1904: bool,
    trailingRows = false): Sheet {.inline.} =
  # open xml file
  var s = newFileStream(fileName, fmRead)
  if s == nil: quit("Unable to read file: " & fileName)
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
            result.info = parseDimension(x.attrValue, date1904)
            result.data = newSeq[SheetData](result.info.rows *
                result.info.cols)

        of xmlElementEnd:
          break
        else:
          discard
        x.next()
      # discard />
      x.next()
      break

  var position: int
  # parse data
  while true:
    x.next()
    case x.kind
    of xmlElementStart:
      if x.elementName =?= "sheetData":
        # ignore <sheetData>
        while true:
          x.next()
          case x.kind
          of xmlElementOpen:
            if x.elementName =?= "row":
              x.parseRowData(result, styles, position)
          of xmlEof:
            break
          else:
            discard
    of xmlEof:
      break
    else:
      discard
  
  if trailingRows:
    result.data.setLen(position + 1)
    let 
      rows = (position + 1) div result.info.cols
    result.info.rows = rows

proc getKindString(item: SheetData, str: SharedStrings): string {.inline.} =
  case item.kind
  of sdk.Boolean:
    result = $parseBool(item.bvalue)
  of sdk.SharedString:
    result = str[parseInt(item.svalue)]
  of sdk.Num:
    result = item.nvalue
  of sdk.InlineStr:
    result = item.isvalue
  of sdk.Date:
    result = item.dvalue
  of sdk.Time:
    result = item.tvalue
  of sdk.Formula:
    result = item.fnvalue
  else:
    result = ""

iterator get(s: Sheet, str: SharedStrings, sep = ","): string {.inline.} =
  let (rows, cols, _, _) = s.info
  for i in 0 ..< rows:
    var res: string
    for j in 0 ..< cols:
      let item = s.data[i * cols + j]
      res.addSep(sep)
      res.add getKindString(item, str)
    yield res

proc plotSym(cols: int, width = 10): string {.inline.} =
  #+------------+------------+
  for i in 0 ..< cols:
    result.add "+"
    result.add repeat("-", width)
  result.add "+\n"

proc getSheetArray(s: Sheet, str: SharedStrings, header: bool,
    skipHeaders: bool): SheetArray =
  let (rows, cols, _, _) = s.info
  result.shape = (rows, cols)
  result.header = header
  # ignore header
  if skipHeaders:
    dec(result.shape.rows)

  result.data = newSeq[string](result.shape.rows * cols)
  result.colType = newSeq[SheetDataKind](cols)
  if likely(not skipHeaders):
    var
      start: int
      over: int
      skipCount: int
    if result.header:
      start = cols
      skipCount = cols
      over = min(2 * cols, result.data.len)
    else:
      start = 0
      over = min(cols, result.data.len)
    for idx, item in s.data:
      if header:
        if skipCount == 0 and start < over:
          result.colType[start - cols] = item.kind
          inc(start)
        if skipCount > 0:
          dec(skipCount)
      else:
        if start < over:
          result.colType[start] = item.kind
          inc(start)
      result.data[idx] = getKindString(item, str)
  else:
    var
      skipCount = cols
      pos, skipType = 0
    for item in s.data:
      if skipCount > 0:
        dec(skipCount)
        continue
      if skipType < cols:
        # may lose some types because of missings of first row
        result.colType[skipType] = item.kind
        inc(skipType)
      result.data[pos] = getKindString(item, str)
      inc(pos)

    # if skip header, header should be false
    result.header = false

proc parseAllSheetName*(fileName: string): seq[string] {.inline.} =
  ## get all sheet name
  extractXml(fileName)
  defer: removeDir(TempDir)

  let
    contentTypes = parseContentTypes(TempDir / "[Content_Types].xml")
    workbook = parseWorkBook(TempDir / contentTypes["workbook"])
  result = newSeqOfCap[string](workbook.data.len)
  for key in workbook.data.keys:
    result.add(key)

proc getRelsFileName(fileName: string): string =
  let pathParts = fileName.splitFile
  # Valid even when name and ext are empty (i.e. /_rels/.rels)
  result = pathParts.dir / "_rels" / pathParts.name & pathParts.ext & ".rels"

proc parseSheetFileNames(contentTypes: ContentTypes, workbook: WorkBook): Table[string, string] =
  let
    workbookFileName = TempDir / contentTypes["workbook"]
    workbookRels = parseRelationships(workbookFileName.getRelsFileName)

  result = collect(initTable(workbook.data.len)):
    for sheetName, rId in workbook.data.pairs:
      let
        internalSheetName = workbookRels[rId].splitFile.name
        sheetFileName = contentTypes[internalSheetName]

      {sheetName: sheetFileName}

proc parseExcel*(fileName: string, sheetName = "", header = false,
    skipHeaders = false, escapeStrings = false, trailingRows = false): SheetTable =
  ## Parse excel and return SheetTable which contains
  ## all sheetArray.
  ## 
  ## `trailingRows` is used to skip empty lines after last row with elements.
  ## But If there are some empty lines before last row with elements, these lines will be kept.
  ## 
  ## If you want to skip all empty lines, you should use iterator `lines` and set `skipEmptyLines` = `true`.
  runnableExamples:
    let
      data = parseExcel("tests/test.xlsx")
      sheetName = "Sheet2"
    echo data[sheetName]

  extractXml(fileName)
  defer: removeDir(TempDir)
  let
    contentTypes = parseContentTypes(TempDir / "[Content_Types].xml")
    workbook = parseWorkBook(TempDir / contentTypes["workbook"])
    styles = parseStyles(TempDir / contentTypes["styles"])
    sheetFileNames = parseSheetFileNames(contentTypes, workbook)

  var sharedString: SharedStrings
  if "sharedStrings" in contentTypes:
    sharedString = parseSharedString(TempDir / contentTypes["sharedStrings"],
        escapeStrings = escapeStrings)

  if sheetName == "":
    for name, fileName in sheetFileNames.pairs:
      var sheet = parseSheet(TempDir / fileName, styles,
          workbook.date1904, trailingRows)
      result.data[name] = getSheetArray(sheet, sharedString, header, skipHeaders)
    return

  if sheetName notin sheetFileNames:
    raise newException(NotFoundSheetError, "no such sheet name: " & sheetName)

  var
    sheet = parseSheet(TempDir / sheetFileNames[sheetName], styles,
        workbook.date1904, trailingRows)
  result.data[sheetName] = getSheetArray(sheet, sharedString, header, skipHeaders)

iterator lines*(fileName: string, sheetName: string,
    escapeStrings = false, skipEmptyLines = false): string =
  ## return lines of xlsx
  runnableExamples:
    for i in lines("tests/test.xlsx", "Sheet2"):
      echo i
  extractXml(fileName)
  defer: removeDir(TempDir)
  let
    contentTypes = parseContentTypes(TempDir / "[Content_Types].xml")
    workbook = parseWorkBook(TempDir / contentTypes["workbook"])
    styles = parseStyles(TempDir / contentTypes["styles"])
    sheetFileNames = parseSheetFileNames(contentTypes, workbook)

  var sharedString: SharedStrings
  if "sharedStrings" in contentTypes:
    sharedString = parseSharedString(TempDir / contentTypes["sharedStrings"],
        escapeStrings = escapeStrings)

  if sheetName notin sheetFileNames:
    raise newException(NotFoundSheetError, "no such sheet name: " & sheetName)

  var
    sheet = parseSheet(TempDir / sheetFileNames[sheetName], styles,
        workbook.date1904)
  for item in get(sheet, sharedString):
    if skipEmptyLines and len(item) == 0:
      continue
    yield item

proc `[]`*(s: SheetArray, i, j: Natural): string {.inline.} =
  # get element from SheetArray
  checkIndex(i < s.shape.rows)
  checkIndex(j < s.shape.cols)
  s.data[i * s.shape.cols + j]

proc `[]=`*(s: var SheetArray, i, j: Natural, value: string) {.inline.} =
  # set element from SheetArray
  checkIndex(i < s.shape.rows)
  checkIndex(j < s.shape.cols)
  s.data[i * s.shape.cols + j] = value

template `[]`*(s: SheetTable, key: string): SheetArray =
  s.data[key]

proc `$`*(s: SheetArray): string =
  ## display SheetArray
  let
    (rows, cols) = s.shape
    width = 10
  var header = s.header
  result.add plotSym(cols, width)
  for i in 0 ..< rows:
    var res = "|"
    for j in 0 ..< cols:
      let item = s.data[i * cols + j]
      res.add alignLeft(item[0 ..< min(width, item.len)], width)
      res.add "|"
    result.add res
    result.add "\n"
    if header:
      result.add plotSym(cols)
      header = false
  result.add plotSym(cols)

proc show*(s: SheetArray, rmax = 20, cmax = 5, width = 10) =
  ## display SheetArray with more control
  runnableExamples:
    let
      sheetName = "sheet2"
      excel = "tests/nim.xlsx"
      data = parseExcel(excel, sheetName = sheetName, header = true,
          skipHeaders = false)

    data[sheetName].show(width = 20)

  let
    (rows, cols) = s.shape
    width = width
    rowFlag = rows <= rmax
    colFlag = cols <= cmax

  if rowFlag and colFlag:
    echo $s
    return

  var
    header = s.header
    result = ""

  if rowFlag:
    # succ: the next value
    result.add plotSym(succ(cmax), width)
    for i in 0 ..< rows:
      var res = "|"
      for j in 0 .. cmax:
        var item: string
        if j == cmax:
          # if last row
          item = repeat(".", min(width, 3))
          res.add center(item[0 ..< min(width, item.len)], width)
        else:
          item = s.data[i * cols + j]
          res.add alignLeft(item[0 ..< min(width, item.len)], width)
        res.add "|"
      result.add res
      result.add "\n"
      if header:
        result.add plotSym(succ(cmax), width)
        header = false
    result.add plotSym(succ(cmax), width)
    echo result
    return

  if colFlag:
    result.add plotSym(cols, width)
    for i in 0 .. rmax:
      var res = "|"
      for j in 0 ..< cols:
        var item: string
        if i == rmax:
          # if last row
          item = repeat(".", min(width, 3))
          res.add center(item[0 ..< min(width, item.len)], width)
        else:
          item = s.data[i * cols + j]
          res.add alignLeft(item[0 ..< min(width, item.len)], width)
        res.add "|"
      result.add res
      result.add "\n"
      if header:
        result.add plotSym(cols, width)
        header = false
    result.add plotSym(cols, width)
    echo result
    return

  result.add plotSym(succ(cmax), width)
  for i in 0 .. rmax:
    var res = "|"
    for j in 0 .. cmax:
      var item: string
      if i == rmax:
        # if last col,
        item = repeat(".", min(width, 3))
        res.add center(item[0 ..< min(width, item.len)], width)
      elif j == cmax:
        # if last row
        item = repeat(".", min(width, 3))
        res.add center(item[0 ..< min(width, item.len)], width)
      else:
        item = s.data[i * cols + j]
        res.add alignLeft(item[0 ..< min(width, item.len)], width)
      res.add "|"
    result.add res
    result.add "\n"
    if header:
      result.add plotSym(succ(cmax), width)
      header = false
  result.add plotSym(succ(cmax), width)
  echo result

proc toCsv*(s: SheetArray, dest: string, sep = ",") {.inline.} =
  ## Parse SheetArray and write a Csv file
  runnableExamples:
    let sheetName = "Sheet2"
    let data = parseExcel("tests/test.xlsx")
    data[sheetName].toCsv("tests/test.csv", sep = ",")

  let f = open(dest, fmWrite)
  defer: f.close()
  let (rows, cols) = s.shape
  for i in 0 ..< rows:
    var res = ""
    for j in 0 ..< cols:
      res.add s.data[i * cols + j]
      if j < pred(cols):
        res.add sep
    f.writeLine res

proc toSeq*(s: SheetArray, skipHeaders = false): seq[seq[string]] {.inline.} = # <-- HERE
  ## Parse SheetArray and return a seq[seq[string]]
  runnableExamples:
    let sheetName = "Sheet2"
    let data = parseExcel("tests/test.xlsx")
    let rows = data[sheetName].toSeq(false)
    for row in rows:
      echo row

  let
    sheet = s.data
    (_, cols) = s.shape
  var
    cCount = 1
    firstRow = true
    row: seq[string]
  # Loop through sheet
  for r in sheet:
    # Check is headers should be included
    if skipHeaders and firstRow: # <-- HERE
      # If last column end header loop
      if cCount == cols:
        firstRow = false
        cCount = 0
      cCount += 1
      continue
    # Add data to internal seq[string]
    row.add(r)
    # Add internal seq[string] to result var
    if cCount == cols:
      cCount = 1
      result.add(row)
      row = @[]
    else:
      cCount += 1

proc parseData[T](x: sink string): T {.inline.} =
  when T is SomeSignedInt:
    try:
      result = T(x.parseInt)
    except ValueError:
      discard
  elif T is SomeUnsignedInt:
    try:
      result = T(x.parseUInt)
    except ValueError:
      discard
  elif T is SomeFloat:
    try:
      result = T(x.parseFloat)
    except ValueError:
      discard
  elif T is bool:
    try:
      result = T(x.parseBool)
    except ValueError:
      discard
  elif T is string:
    result = move(x)

proc getSheetTensor[T](s: Sheet, str: SharedStrings,
    skipHeaders: bool): SheetTensor[T] {.inline.} =
  let (rows, cols, _, _) = s.info
  result.shape = (rows, cols)
  # ignore header
  if skipHeaders:
    dec(result.shape.rows)
  result.data = newseq[T](result.shape.rows * cols)
  if not skipHeaders:
    for idx, item in s.data:
      result.data[idx] = parseData[T](getKindString(item, str))
  else:
    var
      skipCount = cols
      pos = 0
    for item in s.data:
      if skipCount > 0:
        dec(skipCount)
        continue
      result.data[pos] = parseData[T](getKindString(item, str))
      inc(pos)
    result.data = result.data

proc readExcel*[T: SomeNumber|bool|string](fileName: string,
    sheetName: string, skipHeaders = false, escapeStrings = false): SheetTensor[T] =
  ## read excel for scientific calculation
  # for arraymancy https://github.com/mratsim/Arraymancer/blob/master/src/io
  runnableExamples:
    let sheetName = "Sheet1"
    let data = readExcel[int]("tests/test_read_excel.xlsx", sheetName,
        skipHeaders = false)
    # if missing value, will fill default value of T
    assert(data.data == @[1, 4, 7, 9, 4, 7, 0, 3, 12, 54, 24, 887])

  extractXml(fileName)
  defer: removeDir(TempDir)
  let
    contentTypes = parseContentTypes(TempDir / "[Content_Types].xml")
    workbook = parseWorkBook(TempDir / contentTypes["workbook"])
    styles = parseStyles(TempDir / contentTypes["styles"])
    sheetFileNames = parseSheetFileNames(contentTypes, workbook)

  var sharedString: SharedStrings
  if "sharedStrings" in contentTypes:
    sharedString = parseSharedString(TempDir / contentTypes["sharedStrings"],
        escapeStrings = escapeStrings)

  if sheetName notin sheetFileNames:
    raise newException(NotFoundSheetError, "no such sheet name: " & sheetName)

  let
    sheet = parseSheet(TempDir / sheetFileNames[sheetName], styles,
        workbook.date1904)

  result = getSheetTensor[T](sheet, sharedString, skipHeaders)


when isMainModule:
  let excel = "../../tests/test_date_time.xlsx"
  let sheetName = "Sheet1"
  echo parseAllSheetName(excel)
  let data = parseExcel(excel, sheetName = sheetName)
  echo data[sheetName]
  # let
  #   sheetName = "Sheet1"
  #   excel = "../../tests/test_dateTime.xlsx"
  #   data = parseExcel(excel, sheetName = sheetName, header = false,
  #       skipHeaders = false, escapeStrings = true)

  # data[sheetName].show(width = 20)
  # data[sheetName].show(width = 20)
  # data[sheetName].show(width = 20)
  # for i in lines("../../tests/test.xlsx", "Sheet2", skipEmptyLines = true):
  #   echo i

  # echo parseAllSheetName("../../tests/test_int.xlsx")

  # echo readExcel[float]("../../tests/test_int.xlsx", "Sheet1")
