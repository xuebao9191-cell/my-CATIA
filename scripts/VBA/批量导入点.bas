Attribute VB_Name = "批量导入点数据"
Option Explicit

' 输入格式：每行 x,y,z；空行和以 # 开头的行为忽略。
' 运行前：修改 POINTS_FILE。宏会把点放入现有第一个几何图形集，或新建几何图形集。
Sub 批量导入点()
    Const POINTS_FILE As String = "D:\CATIA\Points\points.txt"

    Dim catApp As Object, doc As Object, part As Object, hsf As Object
    Dim hb As Object, pointObj As Object
    Dim fso As Object, ts As Object
    Dim lineText As String, parts As Variant
    Dim x As Double, y As Double, z As Double
    Dim countOK As Long, countBad As Long, lineNo As Long

    On Error GoTo FatalError
    Set catApp = GetObject(, "CATIA.Application")
    If catApp.Documents.Count = 0 Then Err.Raise vbObjectError + 100, , "请先打开一个 CATPart。"
    Set doc = catApp.ActiveDocument
    If doc.GetType <> "CATPart" Then Err.Raise vbObjectError + 101, , "当前活动文档不是 CATPart。"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(POINTS_FILE) Then Err.Raise vbObjectError + 102, , "找不到点文件：" & POINTS_FILE
    Set ts = fso.OpenTextFile(POINTS_FILE, 1, False)

    Set part = doc.Part
    Set hsf = part.HybridShapeFactory

    If part.HybridBodies.Count = 0 Then
        Set hb = part.HybridBodies.Add
        hb.Name = "Imported_Points"
    Else
        Set hb = part.HybridBodies.Item(1)
    End If

    Do While Not ts.AtEndOfStream
        lineNo = lineNo + 1
        lineText = Trim$(ts.ReadLine)
        If Len(lineText) > 0 And Left$(lineText, 1) <> "#" Then
            parts = Split(lineText, ",")
            If UBound(parts) <> 2 Then
                countBad = countBad + 1
                Debug.Print "跳过第 " & lineNo & " 行：需要三个逗号分隔的坐标。"
            ElseIf Not IsNumeric(Trim$(parts(0))) Or Not IsNumeric(Trim$(parts(1))) Or Not IsNumeric(Trim$(parts(2))) Then
                countBad = countBad + 1
                Debug.Print "跳过第 " & lineNo & " 行：坐标不是有效数字。"
            Else
                x = CDbl(Trim$(parts(0)))
                y = CDbl(Trim$(parts(1)))
                z = CDbl(Trim$(parts(2)))
                Set pointObj = hsf.AddNewPointCoord(x, y, z)
                hb.HybridShapes.Add pointObj
                countOK = countOK + 1
            End If
        End If
    Loop

    ts.Close
    part.Update
    MsgBox "点导入完成。" & vbCrLf & "成功：" & countOK & vbCrLf & "跳过/错误行：" & countBad, vbInformation
    Exit Sub

FatalError:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    MsgBox "导入失败：" & Err.Number & " - " & Err.Description, vbExclamation
End Sub
