Attribute VB_Name = "批量另存重命名"
' ============================================================
' 功能：将当前打开的所有文档批量另存为指定前缀的新名称
' 用法：修改 strNewPrefix 为你要加的前缀后运行
' 环境：CATIA V5 + VBA
' ============================================================

Sub 批量另存重命名()

    Dim oCATIA As Object
    Dim oDoc As Object
    Dim strNewPrefix As String
    Dim strNewName As String

    strNewPrefix = "CA-"   ' <-- 修改前缀

    Set oCATIA = GetObject(, "CATIA.Application")

    For Each oDoc In oCATIA.Documents

        If oDoc.GetType = "CATProduct" Or oDoc.GetType = "CATPart" Then

            strNewName = strNewPrefix & oDoc.Name
            Call oDoc.SaveAs("D:\CATIA\Output\" & strNewName)
            Debug.Print "已保存: " & strNewName

        End If

    Next oDoc

    MsgBox "批量另存完成！"

End Sub
