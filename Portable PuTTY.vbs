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
    ' note: The pattern [\s\S] matches any whitespace or non-whitespace character. 
    ' You can't use . in this case, because that special character does not match newlines, 
    ' and you want the expression to span multiple lines. The qualifiers * and ? mean 
    ' "match any number of characters" and "use the shortest match" respectively. 
    regEx.Pattern = """Recent sessions""=[\s\S]*?" & vbCrLf & vbCrLf 
    regEx.IgnoreCase = False
    regEx.MultiLine = True
    regEx.Global = False
    result = regEx.Replace(result, "<ignore>" & vbCrLf & vbCrLf)

    filterString = result
end function

Dim objProgressMsg
Function ProgressMsg( strMessage, strWindowTitle )
    ' inspired from Denis St-Pierre (http://www.robvanderwoude.com/vbstech_ui_progress.php)
    Dim wshShell: Set wshShell = WScript.CreateObject( "WScript.Shell" )

    If strMessage = "" Then
        On Error Resume Next
        objProgressMsg.Terminate( )
        On Error Goto 0
        Exit Function
    End If

    strTempVBS = tfolder.Path + "\" & fso.GetTempName & ".vbs"

    Set objTempMessage = fso.CreateTextFile( strTempVBS, True )
    objTempMessage.WriteLine( "Msg" & "Box """ & strMessage & """, 4096, """ & strWindowTitle & """" )
    objTempMessage.Close

    On Error Resume Next
    objProgressMsg.Terminate( )
    On Error Goto 0

    Set objProgressMsg = WshShell.Exec( "%windir%\system32\wscript.exe " & strTempVBS )

    Set wshShell = Nothing

End Function

' debugging facility
' can also use things like:
'   WshShell.Run "cmd /d /c dir >> output.log", 0, true
function debug(msg)
   if not fso.fileexists("output.log") then
      fso.createTextFile("ouput.log")
   end if
   fso.GetFile("output.log").OpenAsTextStream(8, -2).Write(msg & Chr(10))
end function

' get current list of sessions
if fso.FileExists("putty.reg") then
   savedreg = fso.GetFile("putty.reg").OpenAsTextStream(ForReading, -2).ReadAll()
   savedsessions = filterString(savedreg)
else
   savedsessions = "<ignore>"
end if

' check if putty.exe exists, if not, download it
if not fso.FileExists("putty.exe") then
   puttyUrl = "https://the.earth.li/~sgtatham/putty/latest/x86/putty.exe"
   ret = msgbox("Download putty.exe from """ & puttyUrl & """?", vbOKCancel, "Download putty.exe?")
   if ret = vbCancel then
     Wscript.Quit
   end if

   wTitle = "Portable PuTTY"
   
   ' download putty.exe
   ProgressMsg "Preparing download", wTitle
   dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
   dim bStrm: Set bStrm = createobject("Adodb.Stream")
   xHttp.Open "GET", puttyUrl, False
   xHttp.Send


   ProgressMsg "Downloading """"putty.exe""""", wTitle
   with bStrm
     .type = 1 '//binary
     .open
     .write xHttp.responseBody
     .savetofile "putty.exe", 2 '//overwrite
   end with

   set xHttp = Nothing
   set bStrm = Nothing

   ProgressMsg "", wTitle
end if

function getLocalSessionsList(str)
    Set regEx = new RegExp

    regEx.Pattern = "\[.*\\SimonTatham\\PuTTY\\Sessions\\([^]]*)\]"
    regEx.IgnoreCase = False
    regEx.MultiLine = True
    regEx.Global = True
    set Matches = regEx.Execute(str)

    result = ""
    For Each Match in Matches
       result = result & Match.SubMatches(0) & Chr(10)
    Next

    getLocalSessionsList = result
    
end function

localsessions = filterString(dumpReg())

' compare to local sessions
useLocalSessions=false
if localsessions <> savedsessions and savedsessions <> "<ignore>" and localsessions <> "<ignore>" then

    ' if different, propose to clear previous one
    ret = msgbox("Local sessions already exist:" & chr(10) & getLocalSessionsList(localsessions) & chr(10) & "You can choose to use them or to overwrite them with the saved session." & chr(10) & chr(10) & "Do you want to use the local sessions?", vbYesNoCancel, "Use local sessions?")
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


