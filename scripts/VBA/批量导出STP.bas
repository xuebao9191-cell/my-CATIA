Attribute VB_Name = "批量导出STP_IGES"
' ============================================================
' 功能：批量将当前文件夹下所有 .CATPart / .CATProduct
'       导出为 STEP (.stp) 格式
' 用法：修改 strFolder 为你的零件所在文件夹，然后运行宏
' 环境：CATIA V5 + VBA (Tools -> Macros)
' ============================================================

Sub 批量导出STP()

    Dim strFolder As String
    Dim strFile As String
    Dim objFilesystem As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim oCATIA As Object
    Dim oDocuments As Object
    Dim oDocument As Object
    Dim oProduct As Object

    strFolder = "D:\CATIA\Parts"   ' <-- 修改为你自己的文件夹

    Set objFilesystem = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFilesystem.GetFolder(strFolder)

    Set oCATIA = GetObject(, "CATIA.Application")
    Set oDocuments = oCATIA.Documents

    For Each objFile In objFolder.Files

        If LCase(objFile.Name) Like "*.catpart" Or LCase(objFile.Name) Like "*.catproduct" Then

            Set oDocument = oDocuments.Open(strFolder & "\" & objFile.Name)
            Set oProduct = oDocument.GetItem("Product1")

            Call oDocument.ExportData(strFolder & "\" & Replace(objFile.Name, ".CATPart", ".stp"), "stp")

            oDocument.Close
            MsgBox "已导出: " & objFile.Name

        End If

    Next objFile

    MsgBox "批量导出完成！"

End Sub
