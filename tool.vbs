Dim fso,wsh,temphtml,temppath,fhta,objWMIService,objProcess,strComputer,colProcesses
Set fso = CreateObject("Scripting.FileSystemObject")
Set wsh = wscript.CreateObject("wscript.Shell")

'***********************************
Sub Clock
    temppath = fso.GetSpecialFolder(2).ShortPath & "\"
    temphtml = fso.GetTempName & ".hta"
    Set fhta = fso.OpenTextFile(temppath & temphtml,2,True)

    Call CreateHTA
    wsh.run (temppath & temphtml),0,false
End Sub
'***********************************
Sub KillClock(FileName)
    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process")
    For Each objProcess in colProcesses
        If InStr(objProcess.CommandLine,FileName) > 0 Then
                objProcess.Terminate(0) 
        End If
    Next

    wsh.run ("cmd /c del " & temppath & temphtml),0,false

End Sub
'***********************************
Sub CreateHTA
    fhta.WriteLine "<html>"
    fhta.WriteLine "<script language=""VBScript"">"
    fhta.WriteLine "window.resizeTo 0, 0"
    fhta.WriteLine "Sub Window_OnLoad"
    fhta.WriteLine "width = 400 : height = 200"
    fhta.WriteLine "window.resizeTo width, height"
    fhta.WriteLine "window.moveTo screen.availWidth\2 - width\2, screen.availHeight\2 - height\2"
    fhta.WriteLine "End Sub"
    fhta.WriteLine "</script>"
    fhta.WriteLine "<hta:application id=""oHTA""" 
    fhta.WriteLine "border=""none""" 
    fhta.WriteLine "caption=""no""" 
    fhta.WriteLine "contextmenu=""no""" 
    fhta.WriteLine "innerborder=""yes""" 
    fhta.WriteLine "scroll=""no""" 
    fhta.WriteLine "showintaskbar=""no""" 
    fhta.WriteLine "/>"
    fhta.WriteLine "<img src=""data:image/gif;base64,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""  style=""margin:0 0 0 150px"">"
    fhta.WriteLine "<p align=""center"">文件处理中...</p>"
    fhta.WriteLine "</html>"

End Sub
'***********************************


Set oFs = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("Shell.Application")
Set oWsh = CreateObject("WScript.Shell")
Set oSelFolder = oShell.BrowseForFolder(0, "请选择要处理的 Excel 文件所在的文件夹；处理时间较长，请耐心等待", 0, 0)

If (oSelFolder is Nothing) Then
  WScript.Quit
End If

Call Clock

Set oFolder = oFs.GetFolder(oSelFolder.Self.Path)
Set oExcel = CreateObject("Excel.Application")
oExcel.Visible = False
oExcel.Application.DisplayAlerts =  False
Set oResults = CreateObject("Scripting.Dictionary")

Function ReadFile(filePath, fileName)
  Set oWorkbook = oExcel.Workbooks.Open(filePath)
  For Each oSheet In oWorkbook.Sheets
    If NOT(oResults.Exists(oSheet.Name)) Then
      oResults.Add oSheet.Name,oExcel.Workbooks.Add
    End If
    Set oResSheets = oResults(oSheet.Name).Sheets
    oSheet.Copy null,oResSheets(oResSheets.Count)
    oResSheets(oSheet.Name).Name = fileName
  Next
  oWorkbook.Close
End Function

For Each oFile In oFolder.Files
  filePath = oFile.Path
  fileExt = oFs.GetExtensionName(filePath)
  fileName = Replace(oFile.Name, "."&fileExt, "")
  If InStr(1, fileName, "~$") = 1 Then
  ElseIf (fileExt = "xls" OR fileExt = "xlsx") Then
    ReadFile filePath,fileName
  End If
Next

For Each sKey in oResults.keys
  Set oWb = oResults(sKey)
  On Error Resume Next
  oWb.Sheets("Sheet1").Delete
  On Error Goto 0
  oWb.Sheets(1).Activate
  sFullPath = oFolder.Path & "\处理结果\"
  If NOT oFs.FolderExists(sFullPath) Then
    oFs.CreateFolder sFullPath
  End If
  oWb.SaveAs sFullPath & sKey
  oWb.Close
Next

Call KillClock(temphtml)
WScript.Echo "处理完毕."
oShell.Explore oFolder.Path

oExcel.Application.DisplayAlerts =  True
oExcel.Quit
Set oResults = Nothing
Set oExcel = Nothing
Set oFolder = Nothing
Set oSelFolder = Nothing
Set oShell = Nothing
Set oFs = Nothing
