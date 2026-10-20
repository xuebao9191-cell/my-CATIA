Attribute VB_Name = "批量导出STP"
Option Explicit

' 功能：导出指定文件夹中的 CATPart / CATProduct 为 STEP。
' 注意：只扫描当前文件夹，不递归子文件夹；已有同名 .stp 文件会被覆盖。
' 说明：本宏打开的是磁盘上干净的文件副本（未修改），Close 不会触发保存询问。
Sub 批量导出STP()
    Const INPUT_FOLDER As String = "D:\CATIA\Parts"

    Dim fso As Object, folder As Object, file As Object
    Dim catApp As Object, doc As Object
    Dim ext As String, outPath As String
    Dim okCount As Long, failCount As Long

    On Error GoTo FatalError
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(INPUT_FOLDER) Then Err.Raise vbObjectError + 200, , "输入目录不存在：" & INPUT_FOLDER
    Set folder = fso.GetFolder(INPUT_FOLDER)
    Set catApp = GetObject(, "CATIA.Application")

    For Each file In folder.Files
        ext = LCase$(fso.GetExtensionName(file.Name))
        If ext = "catpart" Or ext = "catproduct" Then
            On Error GoTo OneFileError
            outPath = fso.BuildPath(INPUT_FOLDER, fso.GetBaseName(file.Name) & ".stp")
            Set doc = catApp.Documents.Open(file.Path)
            doc.ExportData outPath, "stp"
            ' 关闭刚打开的干净文档；如个别版本仍询问保存，可改用如下安全关闭：
            ' On Error Resume Next: doc.Close: On Error GoTo OneFileError
            doc.Close
            Debug.Print "OK," & outPath
            okCount = okCount + 1
            On Error GoTo FatalError
        End If
ContinueFile:
    Next file

    MsgBox "STEP 导出结束。" & vbCrLf & "成功：" & okCount & vbCrLf & "失败：" & failCount, vbInformation
    Exit Sub

OneFileError:
    Debug.Print "ERROR," & file.Name & "," & Err.Number & "," & Err.Description
    failCount = failCount + 1
    Err.Clear
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close
    Set doc = Nothing
    On Error GoTo FatalError
    Resume ContinueFile

FatalError:
    MsgBox "导出宏中止：" & Err.Number & " - " & Err.Description, vbExclamation
End Sub
