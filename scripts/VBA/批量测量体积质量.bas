Attribute VB_Name = "批量测量体积质量"
Option Explicit

' 第一阶段仅测量已打开的 CATPart；CATProduct 装配体暂不自动递归统计。
' 密度单位假定 kg/m^3，体积单位假定 m^3；请在本机 CATIA V5 中用已知模型校核单位。
' 结果写入 D:\CATIA\measure_result.csv。
Sub 批量测量体积质量()
    Const OUTPUT_FILE As String = "D:\CATIA\measure_result.csv"
    Const DENSITY_KG_M3 As Double = 7850#

    Dim catApp As Object, doc As Object, part As Object, body As Object
    Dim spa As Object, ref As Object, measurable As Object
    Dim fso As Object, ts As Object
    Dim volume As Double, area As Double, mass As Double
    Dim okCount As Long, failCount As Long

    On Error GoTo FatalError
    Set catApp = GetObject(, "CATIA.Application")
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists("D:\CATIA") Then fso.CreateFolder "D:\CATIA"
    Set ts = fso.CreateTextFile(OUTPUT_FILE, True, True)
    ts.WriteLine "Document,Type,Volume_m3,Area_m2,Density_kg_m3,EstimatedMass_kg,Status"

    For Each doc In catApp.Documents
        If doc.GetType = "CATPart" Then
            On Error GoTo OneDocError
            Set part = doc.Part
            If part.Bodies.Count = 0 Then Err.Raise vbObjectError + 300, , "零件中没有实体 Body。"
            Set body = part.Bodies.Item(1)
            Set spa = doc.GetWorkbench("SPAWorkbench")
            Set ref = part.CreateReferenceFromObject(body)
            Set measurable = spa.GetMeasurable(ref)

            volume = measurable.Volume
            area = measurable.Area
            mass = volume * DENSITY_KG_M3

            ts.WriteLine CsvEscape(doc.Name) & ",CATPart," & Replace(CStr(volume), ",", ".") & "," & _
                         Replace(CStr(area), ",", ".") & "," & CStr(DENSITY_KG_M3) & "," & _
                         Replace(CStr(mass), ",", ".") & ",OK"
            Debug.Print doc.Name & " | V=" & volume & " | A=" & area & " | estimated mass=" & mass
            okCount = okCount + 1
            On Error GoTo FatalError
        ElseIf doc.GetType = "CATProduct" Then
            ts.WriteLine CsvEscape(doc.Name) & ",CATProduct,,,,,SKIPPED: assembly recursion not implemented"
        End If
ContinueDoc:
    Next doc

    ts.Close
    MsgBox "测量结束。" & vbCrLf & "成功零件：" & okCount & vbCrLf & _
           "失败零件：" & failCount & vbCrLf & "结果：" & OUTPUT_FILE & vbCrLf & _
           "注意：质量按统一密度估算，不等于材料属性的真实质量。", vbInformation
    Exit Sub

OneDocError:
    ts.WriteLine CsvEscape(doc.Name) & ",CATPart,,,,,ERROR " & Err.Number & ": " & CsvEscape(Err.Description)
    Debug.Print "ERROR," & doc.Name & "," & Err.Number & "," & Err.Description
    failCount = failCount + 1
    Err.Clear
    Set measurable = Nothing
    On Error GoTo FatalError
    Resume ContinueDoc

FatalError:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    MsgBox "测量宏中止：" & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Function CsvEscape(ByVal value As String) As String
    CsvEscape = """" & Replace(value, """", """"""") & """"
End Function
