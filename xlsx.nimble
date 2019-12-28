# Package

version       = "0.1.2"
author        = "flywind"
description   = "Read and parse Excel files"
license       = "MIT"
srcDir        = "src"



# Dependencies

requires "nim >= 1.0.0"
requires "zip >= 0.2.1"

# tests
task test, "Run all tests":
  exec "nim c -r tests/alltests.nim"
