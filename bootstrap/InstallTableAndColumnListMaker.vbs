Option Explicit

Const RAW_MODULE_URL = "https://raw.githubusercontent.com/SSsan-00/TableAndColumnListMaker/main/src/TableAndColumnListMaker.bas"
Const OUTPUT_WORKBOOK_NAME = "TableAndColumnListMaker.xlsm"
Const XLSM_FILE_FORMAT = 52

Dim shell
Dim fso
Dim tempModulePath
Dim outputWorkbookPath
Dim importError
Dim excelApp
Dim workbook

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

tempModulePath = shell.ExpandEnvironmentStrings("%TEMP%") & "\TableAndColumnListMaker.bas"
outputWorkbookPath = shell.SpecialFolders("Desktop") & "\" & OUTPUT_WORKBOOK_NAME

On Error Resume Next
DownloadModule RAW_MODULE_URL, tempModulePath
If Err.Number <> 0 Then
    MsgBox "Failed to download the VBA module." & vbCrLf & Err.Description, vbCritical, "TableAndColumnListMaker Bootstrap"
    WScript.Quit 1
End If
On Error GoTo 0

Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = True
excelApp.DisplayAlerts = False

On Error Resume Next
Set workbook = excelApp.Workbooks.Add()
workbook.VBProject.VBComponents.Import tempModulePath
If Err.Number <> 0 Then
    importError = Err.Description
    excelApp.DisplayAlerts = True
    MsgBox "Failed to import the VBA module." & vbCrLf & vbCrLf & _
        "Enable Excel's setting: Trust access to the VBA project object model, then run this bootstrap again." & vbCrLf & vbCrLf & _
        importError, vbCritical, "TableAndColumnListMaker Bootstrap"
    WScript.Quit 1
End If
On Error GoTo 0

If fso.FileExists(outputWorkbookPath) Then
    fso.DeleteFile outputWorkbookPath, True
End If

workbook.SaveAs outputWorkbookPath, XLSM_FILE_FORMAT
excelApp.DisplayAlerts = True

MsgBox "Bootstrap completed." & vbCrLf & vbCrLf & _
    "Created workbook: " & outputWorkbookPath & vbCrLf & _
    "Open the workbook in Excel and run RunTableAndColumnListMaker.", _
    vbInformation, "TableAndColumnListMaker Bootstrap"

Sub DownloadModule(ByVal sourceUrl, ByVal destinationPath)
    Dim http
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", sourceUrl, False
    http.Send

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 1000, "DownloadModule", "HTTP status: " & CStr(http.Status)
    End If

    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "Shift_JIS"
    stream.Open
    stream.WriteText http.ResponseText
    stream.SaveToFile destinationPath, 2
    stream.Close
End Sub
