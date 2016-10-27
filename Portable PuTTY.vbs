Set WshShell = CreateObject("WScript.Shell" ) 
dim fso: set fso = CreateObject("Scripting.FileSystemObject")
dim CurrentDirectory: CurrentDirectory = WshShell.CurrentDirectory

const ForReading = 1
const TemporaryFolder = 2
Set tfolder = fso.GetSpecialFolder(TemporaryFolder)

' create a shortcut with the proper icon
' this will be useful to put a shortcut in another directory
set oShellLink = WshShell.CreateShortcut(CurrentDirectory & "\Portable PuTTY.lnk")
oShellLink.TargetPath = WScript.ScriptFullName
oShellLink.WindowStyle = 1
oShellLink.IconLocation = CurrentDirectory & "\putty.exe, 0"
oShellLink.Description = "Shortcut to Portable PuTTY"
oShellLink.WorkingDirectory = CurrentDirectory
oShellLink.Save

tempregfilename = tfolder.Path & "\" & fso.GetTempName    

function dumpReg()
	WshShell.Run "reg export HKEY_CURRENT_USER\Software\SimonTatham\PuTTY " & tempregfilename & " /y", 0, true
	if fso.FileExists(tempregfilename) then
	   dumpReg = fso.GetFile(tempregfilename).OpenAsTextStream(ForReading, -2).ReadAll()
	else
	   dumpReg = ""
	end if
end function

function filterString(s)
    Set regEx = new RegExp

    dim result: result = s

    ' ignore first line
    regEx.Pattern = ".*"
    regEx.IgnoreCase = False
    regEx.MultiLine = True
    regEx.Global = False
    result = regEx.Replace(result, "<ignore>")

    ' ignore RandSeedFile
    regEx.Pattern = """RandSeedFile""=.*"
    regEx.IgnoreCase = False
    regEx.MultiLine = True
    regEx.Global = False
    result = regEx.Replace(result, "<ignore>")

    ' ignore Recent sessions
    regEx.Pattern = """Recent sessions""=.*"
    regEx.IgnoreCase = False
    regEx.MultiLine = True
    regEx.Global = False
    result = regEx.Replace(result, "<ignore>")

    filterString = result
end function

function debug(msg)
   if not fso.fileexists("output.log") then
      fso.createTextFile("ouput.log")
   end if
   fso.GetFile("output.log").OpenAsTextStream(8, -2).Write(msg & Chr(10))
end function

' get current list of sessions
if fso.FileExists("putty.reg") then
   WshShell.Run "cmd /d /c echo " & "putty.reg exists" & " >> output.log", 0, true
   WshShell.Run "cmd /d /c dir >> output.log", 0, true
   savedreg = fso.GetFile("putty.reg").OpenAsTextStream(ForReading, -2).ReadAll()
   savedsessions = filterString(savedreg)
else
   savedsessions = "<ignore>"
end if

localsessions = filterString(dumpReg())

' compare to local sessions
useLocalSessions=false
if localsessions <> savedsessions and savedsessions <> "<ignore>" and localsessions <> "<ignore>" then

    ' if different, propose to clear previous one
    ret = msgbox("Local sessions already exist:" & chr(10) & "coucou" & chr(10) & chr(10) & "You can choose to use them or to overwrite them with the saved session." & chr(10) & chr(10) & "Do you want to use the local sessions?", vbYesNoCancel, "Use local sessions?")
    Select case ret
    case vbCancel
       Wscript.Quit
    case vbNo
       WshShell.Run "reg delete HKEY_CURRENT_USER\Software\SimonTatham\PuTTY\Sessions /va /f", 0, true
    case vbYes
       useLocalSessions=true
    End Select
end if

if not useLocalSessions then
  ' import saved sessions
  WshShell.Run "reg import putty.reg", 0, true

  ' update link to rnd file
  WshShell.Run "reg add HKEY_CURRENT_USER\Software\SimonTatham\PuTTY /v ""RandSeedFile"" /d """ & fso.BuildPath(CurrentDirectory, "putty.rnd") & """ /f", 0, true
end if

' run executable, wait for its end
WshShell.Run "putty.exe", 1, true

if not useLocalSessions then

  ' dump current sessions
  currentreg = dumpReg()

  ' check if we need to save the current sessions
  if currentreg <> savedreg then
    ' if needed, save them
    if fso.FileExists("putty.bak") then
        if fso.FileExists("putty.bak.bak") then
	        fso.DeleteFile("putty.bak.bak")
        end if
        fso.MoveFile "putty.bak", "putty.bak.bak"
    end if
    if fso.FileExists("putty.reg") then
	    fso.MoveFile "putty.reg", "putty.bak"
    end if
    fso.CopyFile tempregfilename, "putty.reg"
  end if
end if

' do some cleaning
if fso.FileExists(tempregfilename) then
   fso.DeleteFile(tempregfilename)
end if
Set WshShell = Nothing
set fso = Nothing



rem regedit /s puttydel.reg
rem pause


