import os, streams, parsexml, strutils, tables, times

import zip / zipfiles


const fileName = "./test.xlsx"
assert existsFile(fileName)


type
  # newException(SheetDataKindError, "unKnown sheet data kind")
  SheetDataKindError* = ref Exception
  SheetDataKind* {.pure.} = enum
    Boolean, Date, Error, InLineStr, Num, SharedString, Formula
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

  Sheet* = object
    rows, cols: int
    data: seq[SheetData]


proc extractXml*(fileName: string) =
  var z: ZipArchive
  if not z.open(fileName):
    echo "Opening zip failed"
    quit(1)
  z.extractAll("files/td")
  z.close()
  assert existsDir("files/td/xl/worksheets")
  assert existsFile("files/td/xl/worksheets/sheet1.xml")

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
              break
          of xmlElementClose:
            break
          else: discard
          x.next()
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
          while x.kind == xmlCharData:
            res[count] &= x.charData
            x.next()
          # seq index
          count += 1
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
  open(x, s, fileName)

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
          of xmlElementClose:
            break
          else: discard
          # ignore element
          x.next()
        # ignore xmlElementClose
        x.next()
      # over
      break


# b for boolean
# d for date
# e for error
# inlineStr for an inline string (i.e., not stored in the shared strings part, but directly in the cell)
# n for number
# s for shared string (so stored in the shared strings part and not in the cell)
# str for a formula (a string representing the formula)


#[
<c r="B1" t="s">
<v>1</v>
</c>
]#

# Boolean, Date, Error, InLineStr, Num, SharedString, Formula

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
  while x.kind == xmlCharData:
    result.svalue &= x.charData
    x.next()
  # ignore </v>
  x.next()
  # point to </c>

proc parseSheetDataFormula(x: var XmlParser): SheetData {.inline.} =
  result = SheetData(kind: sdk.Formula)
  # ignore <f>
  x.next()
  while x.kind == xmlCharData:
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


proc parseSheet*(fileName: string) = discard

when isMainModule:
  echo parseContentTypes("files/td/[Content_Types].xml")
  echo praseWorkBook("files/td/xl/workbook.xml")
  echo parseSharedString("files/td/xl/sharedStrings.xml")
