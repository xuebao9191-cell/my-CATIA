Attribute VB_Name = "批量另存重命名"
Option Explicit

' 功能：将当前打开的 CATPart / CATProduct 另存到输出目录，并添加前缀。
' 运行前：修改 OUTPUT_FOLDER；建议先用测试文件验证。
Sub 批量另存重命名()
    Const OUTPUT_FOLDER As String = "D:\CATIA\Output"
    Const NEW_PREFIX As String = "CA-"

    Dim catApp As Object, doc As Object, fso As Object
    Dim targetPath As String, targetName As String
    Dim okCount As Long, failCount As Long, msg As String

    On Error GoTo FatalError
    Set catApp = GetObject(, "CATIA.Application")
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(OUTPUT_FOLDER) Then
        fso.CreateFolder OUTPUT_FOLDER
    End If

    For Each doc In catApp.Documents
        If doc.GetType = "CATPart" Or doc.GetType = "CATProduct" Then
            On Error GoTo OneDocError
            targetName = NEW_PREFIX & fso.GetFileName(doc.Name)
            targetPath = fso.BuildPath(OUTPUT_FOLDER, targetName)

            ' 不覆盖已有文件：自动追加编号。
            If fso.FileExists(targetPath) Then
                targetPath = UniquePath(fso, OUTPUT_FOLDER, NEW_PREFIX, doc.Name)
            End If

            doc.SaveAs targetPath
            Debug.Print "OK," & targetPath
            okCount = okCount + 1
            On Error GoTo FatalError
        End If
ContinueDocument:
    Next doc

    MsgBox "批量另存结束。" & vbCrLf & "成功：" & okCount & vbCrLf & "失败：" & failCount, vbInformation
    Exit Sub

OneDocError:
    Debug.Print "ERROR," & doc.Name & "," & Err.Number & "," & Err.Description
    failCount = failCount + 1
    Err.Clear
    On Error GoTo FatalError
    Resume ContinueDocument

FatalError:
    MsgBox "宏中止：" & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Function UniquePath(ByVal fso As Object, ByVal folder As String, _
                            ByVal prefix As String, ByVal originalName As String) As String
    Dim baseName As String, ext As String, candidate As String, n As Long
    baseName = fso.GetBaseName(originalName)
    ext = fso.GetExtensionName(originalName)
    n = 1
    Do
        candidate = fso.BuildPath(folder, prefix & baseName & "_" & Format$(n, "000") & "." & ext)
        If Not fso.FileExists(candidate) Then Exit Do
        n = n + 1
    Loop
    UniquePath = candidate
End Function
