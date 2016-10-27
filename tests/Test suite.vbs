const ForReading = 1
const ForWriting = 2
const ForAppending = 8

function readFile(filespec)
  dim fso: set fso = CreateObject("Scripting.FileSystemObject")
  readFile = fso.GetFile(filespec).OpenAsTextStream(ForReading, -2).ReadAll()
  set fso = Nothing
end function

function dumpReg()
  dim fso: set fso = CreateObject("Scripting.FileSystemObject")

  const TemporaryFolder = 2
  Set tfolder = fso.GetSpecialFolder(TemporaryFolder)
  tempregfilename = tfolder.Path & "\" & fso.GetTempName    

  Set WshShell = CreateObject("WScript.Shell" ) 
  WshShell.Run "reg export HKEY_CURRENT_USER\Software\TestSimonTatham\PuTTY " & tempregfilename & " /y", 0, true
  Set WshShell = Nothing

  if fso.FileExists(tempregfilename) then
     dumpReg = fso.GetFile(tempregfilename).OpenAsTextStream(ForReading, -2).ReadAll()
     fso.deleteFile(tempregfilename)
  else
     dumpReg = ""
  end if

  set fso = Nothing
end function

' setup, called before every test in the file
sub setup()
  cleanDir

  origvbs = readFile("..\Portable PuTTY.vbs")
  targetvbs = origvbs

  Set regEx = new RegExp

  ' stub call to putty.exe
  regEx.Pattern = """putty.exe"""
  regEx.IgnoreCase = False
  regEx.MultiLine = True
  regEx.Global = False
  targetvbs = regEx.Replace(targetvbs, """""""test data\stubputty.bat""""""")

  ' stub call to msgbox
  regEx.Pattern = "msgbox"
  regEx.IgnoreCase = true
  regEx.MultiLine = True
  regEx.Global = true
  targetvbs = regEx.Replace(targetvbs, "stubmsgbox")

  ' change registers from SimonTatham to TestSimonTatham
  regEx.Pattern = "SimonTatham"
  regEx.IgnoreCase = false
  regEx.MultiLine = True
  regEx.Global = true
  targetvbs = regEx.Replace(targetvbs, "TestSimonTatham")

  ' write file
  dim fso: set fso = CreateObject("Scripting.FileSystemObject")
  fso.CreateTextFile "Portable PuTTY.vbs"
  set f = fso.GetFile("Portable PuTTY.vbs")
  set ts = f.OpenAsTextStream(ForWriting, -2)
  ts.Write targetvbs
  stubmsgbox = fso.GetFile("test data\stubmsgbox.vbs").OpenAsTextStream(ForReading, -2).ReadAll()
  ts.Write stubmsgbox
  ts.Close

  set fso = Nothing
    
end sub

' teardown, called after every test in the file
sub tearDown()
  cleanDir
  cleanRegistry
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
' and loading init.reg if any
sub prepareTest(dir)
  dim fso: set fso = CreateObject("Scripting.FileSystemObject")
  fso.CopyFolder "test data\" & dir, "."

  if fso.FileExists("init.reg") then
    Set WshShell = CreateObject("WScript.Shell" ) 
    WshShell.Run "reg import init.reg", 0, true
    Set WshShell = Nothing
  end if

  set fso = Nothing
end sub

' cleans the directory, removing all generated files
sub cleanDir()
  deleteFile("Portable PuTTY.lnk")
  deleteFile("stubputty.bat")
  deleteFile("stubmsgbox.bat")
  deleteFile("init.reg")
  deleteFile("run.reg")
  deleteFile("putty.reg")
end sub

' cleans the registry
sub cleanRegistry()
  Set WshShell = CreateObject("WScript.Shell" ) 
  WshShell.Run "reg delete HKEY_CURRENT_USER\Software\TestSimonTatham /f", 0, true
  Set WshShell = Nothing
end sub

' execute Portable PuTTY.vbs
sub runPortablePuTTY
  Set WshShell = CreateObject("WScript.Shell" )
  Assert.trace WshShell.CurrentDirectory
  WshShell.Run """Portable PuTTY.vbs""", 1, true
  Set WshShell = Nothing
end sub

' ignores the RandSeedFile part
function filterRandSeedFile(regstr)
    Set regEx = new RegExp

    dim result: result = regstr

    ' ignore RandSeedFile
    regEx.Pattern = """RandSeedFile""=.*"
    regEx.IgnoreCase = False
    regEx.MultiLine = True
    regEx.Global = False
    result = regEx.Replace(result, "<ignore>")

    filterRandSeedFile = result
end function

' verifies that a shortcut is created
' when Portable PuTTY is run
sub iTestShortcutCreated
  runPortablePuTTY
  Assert.IsTrue FileExists("Portable PuTTY.lnk"), "Shortcut was not created"
end sub

' check that putty.exe is downloaded
sub ignoreTestPuttyDownloaded
  Assert.IsTrue false
end sub

sub TestFirstLaunchNoLocalWithModif
  prepareTest("TestFirstLaunchNoLocalWithModif")
  runPortablePuTTY
  Assert.IsTrue FileExists("putty.reg"), "putty.reg was not created"
  Assert.Equal readFile("putty.reg"), readFile("run.reg"), "putty.reg and run.reg differ"
  Assert.Equal dumpReg(), readFile("run.reg"), "registry not as modified by putty.exe"
end sub

sub TestFirstLaunchWithLocalWithModif
  prepareTest("TestFirstLaunchWithLocalWithModif")
  runPortablePuTTY
  Assert.IsTrue FileExists("putty.reg"), "putty.reg was not created"
  Assert.Equal readFile("putty.reg"), readFile("run.reg"), "putty.reg and run.reg differ"
  Assert.Equal dumpReg(), readFile("run.reg"), "registry not as modified by putty.exe"
end sub

sub TestFirstLaunchWithLocalNoModif
  prepareTest("TestFirstLaunchWithLocalNoModif")
  runPortablePuTTY
  Assert.IsTrue FileExists("putty.reg"), "putty.reg was not created"
  Assert.Equal readFile("putty.reg"), readFile("init.reg"), "putty.reg and init.reg differ"
  Assert.Equal dumpReg(), readFile("init.reg"), "registry not as initial"
end sub

sub TestFreshLaunchNoModif
  prepareTest("TestFreshLaunchNoModif")
  runPortablePuTTY
  Assert.IsTrue FileExists("putty.reg"), "putty.reg was not created"
  Assert.Equal dumpReg(), readFile("putty.reg"), "registry not as saved"
end sub

sub TestNormalLaunchNoModif
  prepareTest("TestNormalLaunchNoModif")
  Assert.Equal readFile("putty.reg"), readFile("init.reg"), "putty.reg and init.reg should be identical at the start of the test"
  runPortablePuTTY
  Assert.IsTrue FileExists("putty.reg"), "putty.reg was not created"
  Assert.Equal filterRandSeedFile(readFile("putty.reg")), filterRandSeedFile(readFile("init.reg")), "putty.reg and init.reg differ"
  Assert.Equal dumpReg(), readFile("putty.reg"), "registry not as saved"
end sub

sub TestUseLocalNoModif
  prepareTest("TestUseLocalNoModif")
  runPortablePuTTY
  Assert.Equal dumpReg(), readFile("init.reg"), "register should not be modified"
end sub

sub TestUseSavedNoModif
  prepareTest("TestUseSavedNoModif")
  runPortablePuTTY
  Assert.IsTrue FileExists("putty.reg"), "putty.reg was not created"
  Assert.Equal dumpReg(), readFile("putty.reg"), "putty.reg should be dumped"
end sub

sub TestCancelUseLocal
  prepareTest("TestCancelUseLocal")
  runPortablePuTTY
  Assert.NotEqual dumpReg(), readFile("run.reg"), "putty.exe should not have run"
end sub



' preexisting       confirm    used by    modified by     after
'local    saved     use local  putty.exe  putty.exe   local    saved    Scenario
'------------------------------------------------------------------------------------------------
'empty    empty     N/A        empty      yes         modified modified FirstLaunchNoLocalWithModif 
'existing empty     N/A        existing   yes         modified modified FirstLaunchWithLocalWithModif
'existing empty     N/A        existing   no          existing existing FirstLaunchWithLocalNoModif
'empty    saved     N/A        saved      no          saved    saved	FreshLaunchNoModif
'empty    saved     N/A        saved      yes         empty    modified not tested
'existing =existing N/A        saved      no          saved    saved  	NormalLaunchNoModif
'existing =existing N/A        saved      yes         modified modified not tested
'existing saved     yes        existing   no          existing N/A     	UseLocalNoModif
'existing saved     yes        existing   yes         modified N/A    	not tested
'existing saved     no         saved      no          saved    saved    UseSavedNoModif
'existing saved     no         saved      yes         modified modified not tested

'existing saved     cancel     N/A        N/A         N/A      N/A      CancelUseLocal
