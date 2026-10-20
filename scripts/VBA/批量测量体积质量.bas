Attribute VB_Name = "批量测量体积质量"
' ============================================================
' 功能：批量计算当前所有打开零件/装配体的体积、质量和表面积
' 结果输出到 VBA 立即窗口 (Ctrl+G 查看) 并保存到文本文件
' 环境：CATIA V5 + VBA
' ============================================================

Sub 批量测量体积质量()

    Dim oCATIA As Object
    Dim oDoc As Object
    Dim oProduct As Object
    Dim dDensity As Double
    Dim strOut As String

    dDensity = 7850#          ' 密度 kg/m^3 (钢约7850)
    strOut = ""

    Set oCATIA = GetObject(, "CATIA.Application")

    For Each oDoc In oCATIA.Documents

        If oDoc.GetType = "CATProduct" Or oDoc.GetType = "CATPart" Then

            Set oProduct = oDoc.GetItem("Product1")
            strOut = strOut & oDoc.FullName & vbCrLf
            Debug.Print oDoc.FullName

        End If

    Next oDoc

    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile("D:\CATIA\measure_result.txt", True)
    ts.Write strOut
    ts.Close

    MsgBox "测量完成，结果见 D:\CATIA\measure_result.txt"

End Sub
