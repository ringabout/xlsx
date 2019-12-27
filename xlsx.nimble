# Package

version       = "0.1.0"
author        = "flywind"
description   = "A new awesome nimble package"
license       = "MIT"
srcDir        = "src"



# Dependencies

requires "nim >= 1.0.0"
requires "zip >= 0.2.1"

# tests
task test, "Run all tests":
  exec "nim c -r tests/alltests.nim"