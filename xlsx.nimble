# Package

version       = "0.4.8"
author        = "flywind"
description   = "Read and parse Excel files"
license       = "MIT"
srcDir        = "src"



# Dependencies

requires "nim >= 1.0.0"
requires "zippy >= 0.10.4"

# tests
task test, "Run all tests":
  exec "nim c -r tests/alltests.nim"

task test_arc, "Run all tests with arc":
  exec "nim c -r --gc:arc tests/alltests.nim"

# docs
task docs, "Generate docs":

  exec "nim doc2 " &
    "--git.commit:master " &
    "--index:on " &
    "--git.devel:master " &
    "--git.url:https://github.com/ringabout/xlsx " &
    "-o:docs/utils.html " &
    "src/xlsx/utils.nim"
