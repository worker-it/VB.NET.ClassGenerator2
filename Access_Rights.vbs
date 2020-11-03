Dim WshShell, oExec, lclFolder

lclFolder = "C:\Program Files\itworker\Class Generator\Database"
Set WshShell = CreateObject("WScript.Shell")

Set oExec = WshShell.Exec("icacls ""C:\Program Files\itworker\Class Generator\Database"" /grant:r Everyone:F /T")

'Do While oExec.Status = 0
'    WScript.Sleep 100
'Loop

'WScript.Echo oExec.Status