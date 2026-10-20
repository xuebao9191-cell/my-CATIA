Attribute VB_Name = "批量导入点数据"
' ============================================================
' 功能：从文本文件读取 XYZ 坐标点，在 CATIA 零件中批量创建点
' 文件格式：每行 "x, y, z"（逗号分隔），如：
'           10, 20, 30
'           15, 25, 35
' 用法：修改 strFile 为你的点数据文件路径后运行
' 环境：CATIA V5 + VBA
' ============================================================

Sub 批量导入点()

    Dim strFile As String
    Dim strLine As String
    Dim aCoords As Variant
    Dim fso As Object
    Dim ts As Object

    Dim oCATIA As Object
    Dim oPart As Object
    Dim oHSF As Object
    Dim oPointRef As Object
    Dim oPoint As Object
    Dim i As Long

    strFile = "D:\CATIA\Points\points.txt"   ' <-- 修改为你的文件

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(strFile, 1)

    Set oCATIA = GetObject(, "CATIA.Application")
    Set oPart = oCATIA.ActiveDocument.Part
    Set oHSF = oPart.HybridShapeFactory

    i = 1
    Do While Not ts.AtEndOfStream
        strLine = Trim(ts.ReadLine)
        If Len(strLine) > 0 And Not Left(strLine, 1) = "#" Then
            aCoords = Split(strLine, ",")
            Set oPoint = oHSF.AddNewPointCoord(CDbl(aCoords(0)), CDbl(aCoords(1)), CDbl(aCoords(2)))
            Set oPointRef = oPart.HybridBodies.Item(1).HybridShapes.Add(oPoint)
            oPart.Update
            i = i + 1
        End If
    Loop

    ts.Close
    MsgBox "已导入 " & i - 1 & " 个点"

End Sub
