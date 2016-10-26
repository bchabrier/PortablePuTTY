function stubmsgbox(msg, buttons, title)
   if fso.FileExists("stubmsgbox.bat") then
      stubmsgbox = WshShell.Run ("stubmsgbox.bat """ & Replace(msg, Chr(10), " ") & """ """ & buttons & """ """ & title & """", 1, true)
      if stubmsgbox = 9999 then
      	 stubmsgbox = MsgBox(msg, buttons, title)
      end if
   else
      stubmsgbox = MsgBox(msg, buttons, title)
   end if
end function