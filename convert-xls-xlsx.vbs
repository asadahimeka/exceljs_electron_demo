Set args = WScript.Arguments
strPath = args(0)
strPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(strPath)
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False
Set objFso = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFso.GetFolder(strPath)
For Each objFile In objFolder.Files
    fileName = objFile.Path
    If (objFso.GetExtensionName(objFile.Path) = "xls") Then
        Set objWorkbook = objExcel.Workbooks.Open(fileName)
        saveFileName = Replace(fileName,".xls",".xlsx")
        objWorkbook.SaveAs saveFileName,51
        objWorkbook.Close()
        objExcel.Application.DisplayAlerts =  True
    End If
Next
MsgBox "Finished conversion"
