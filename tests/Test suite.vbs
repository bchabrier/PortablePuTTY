const ForReading = 1
const ForWriting = 2

' setup, called before every test in the file
sub setup()
  clean

  dim fso: set fso = CreateObject("Scripting.FileSystemObject")
  origvbs = fso.GetFile("..\Portable PuTTY.vbs").OpenAsTextStream(ForReading, -2).ReadAll()
  targetvbs = origvbs

  Set regEx = new RegExp

  ' stub call to putty.exe
  regEx.Pattern = """putty.exe"""
  regEx.IgnoreCase = False
  regEx.MultiLine = True
  regEx.Global = False
  targetvbs = regEx.Replace(targetvbs, """stubputty.bat""")

  ' write file
  fso.CreateTextFile("Portable PuTTY.vbs")
  set f = fso.GetFile("Portable PuTTY.vbs")
  set ts = f.OpenAsTextStream(ForWriting, -2)
  ts.Write targetvbs
  ts.Close

  set fso = Nothing
    
end sub

' teardown, called after every test in the file
sub tearDown()
  Assert.trace "setup"
  clean
end sub

' delete a file if it exists
sub deleteFile(filespec)
  dim fso: set fso = CreateObject("Scripting.FileSystemObject")
  if fso.FileExists(filespec) then
    fso.DeleteFile filespec
  end if
  set fso = Nothing
end sub

' check if a file if it exists
function FileExists(filespec)
  dim fso: set fso = CreateObject("Scripting.FileSystemObject")
  FileExists = fso.FileExists(filespec)
  set fso = Nothing
end function

' prepare the test by
' copying test data into the tests directory
sub prepareTest(dir)
  dim fso: set fso = CreateObject("Scripting.FileSystemObject")
  fso.CopyFolder "test data\" & dir, "."
  set fso = Nothing
end sub

' cleans the directory, removing all generated files
sub clean()
  deleteFile("Portable PuTTY.lnk")
end sub

' execute Portable PuTTY.vbs
sub runPortablePuTTY
  Set WshShell = CreateObject("WScript.Shell" )
  Assert.trace WshShell.CurrentDirectory
  WshShell.Run """Portable PuTTY.vbs""", 1, true
  Set WshShell = Nothing
end sub

' verifies that a shortcut is created
' when Portable PuTTY is run
sub TestShortcutCreated
  runPortablePuTTY
  Assert.IsTrue FileExists("Portable PuTTY.lnk"), "Shortcut was not created"
end sub

' verifies that a shortcut is created
' when Portable PuTTY is run
sub TestShortcutCreated
  prepareTest("Empty")
  runPortablePuTTY
  Assert.IsTrue FileExists("Portable PuTTY.lnk"), "Shortcut was not created"
end sub



