sub TestShortcutCreated
  WshShell.Run "..\Portable PuTTY.vbs", 1, true
  Assert.IsTrue "a" = "a"
  Assert.IsTrue "a" = "a", "a and a don't match"
  Assert.IsTrue 123 = 123
  Assert.IsTrue 123 = 123, "123 should match"
  Assert.IsTrue 123 > 1, "123 is not more than 1"
end sub


