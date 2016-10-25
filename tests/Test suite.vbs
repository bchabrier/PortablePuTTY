' setup, called before every test in the file
sub setup()
  Assert.trace "setup"
  clean
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

' cleans the directory, removing all generated files
sub clean()
  deleteFile("Portable PuTTY.lnk")
end sub

' execute Portable PuTTY.vbs
sub runPortablePuTTY
  Set WshShell = CreateObject("WScript.Shell" )
  Assert.trace WshShell.CurrentDirectory
  WshShell.Run """..\Portable PuTTY.vbs""", 1, true
  Set WshShell = Nothing
end sub

' verifies that a shortcut is created
' when Portable PuTTY is run
sub TestShortcutCreated
  runPortablePuTTY
  Assert.IsTrue FileExists("Portable PuTTY.lnk"), "Shortcut was not created"
end sub



