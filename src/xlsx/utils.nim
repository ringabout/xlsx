import os, streams, parsexml, strutils, tables

import zip / zipfiles


const fileName = "./test.xlsx"
assert existsFile(fileName)

type
  WorkBook* = Table[string, string]
  ContentTypes* = seq[string]


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
      break
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
    case x.kind
    of xmlElementStart:
      # catch <sheets>
      if x.elementName =?= "sheets":
        # ignore sheets
        x.next()
        # parse name: sheetId
        while x.kind == xmlElementOpen and x.elementName =?= "sheet":
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
    else:
      discard


when isMainModule:
  echo parseContentTypes("files/td/[Content_Types].xml")
  echo praseWorkBook("files/td/xl/workbook.xml")
