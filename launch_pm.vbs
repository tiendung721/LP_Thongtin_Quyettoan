Option Explicit

Dim fileSystem, shell, projectRoot, batchPath, exitCode

Set fileSystem = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
projectRoot = fileSystem.GetParentFolderName(WScript.ScriptFullName)
batchPath = fileSystem.BuildPath(projectRoot, "Chay_Phan_Mem.bat")
shell.CurrentDirectory = projectRoot

exitCode = shell.Run("""" & batchPath & """", 0, True)
If exitCode <> 0 Then
    MsgBox "Khong the khoi dong KIKAI Quyet Toan PM (ma loi " & exitCode & ")." & vbCrLf & _
        "Hay chay Chay_Phan_Mem.bat de xem thong tin loi.", vbExclamation, "KIKAI Quyet Toan PM"
End If
