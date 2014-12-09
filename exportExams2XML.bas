Attribute VB_Name = "exportXML"
Sub exportXML()
Dim lRows As Long
Dim fsT As Object
Dim oStr As String
Dim checkStr As String

lRows = 24
FullPath = ThisWorkbook.Path & "\output.xml"

Set fsT = CreateObject("ADODB.Stream")
fsT.Type = 2 'Specify stream type - we want To save text/string data.
fsT.Charset = "utf-8" 'Specify charset For the source text data.
fsT.Open 'Open the stream And write binary data To the object

Set mySheet = ThisWorkbook.Worksheets(1)
MsgBox (Cells(1, 1))

For k = 0 To 20
    fsT.WriteText "/*" & Cells(1, (2 + k * 6)) & "*/" & vbNewLine
    For i = 4 To lRows
        checkStr = ""
        oStr = "<tr>" & vbNewLine
        oStr = oStr & "   <td>" & Format(Cells(i, 1), "dd.mm") & "</td>" & vbNewLine
        For j = (2 + k * 6) To (2 + k * 6) + 5
            If j > 1 Then
                checkStr = checkStr & Cells(i, j)
            End If
            If Format(Cells(i, j), "General number") <= 1 Then
                oStr = oStr & "   <td>" & Format(Cells(i, j), "hh:mm") & "</td>" & vbNewLine
            ElseIf Cells(i, j) = "" Then
                oStr = oStr & "   <td>-</td>" & vbNewLine
            Else
                oStr = oStr & "   <td>" & Cells(i, j) & "</td>" & vbNewLine
            End If
        Next j
        oStr = oStr & "</tr>" & vbNewLine
        If Len(checkStr) > 3 Then
            fsT.WriteText oStr
        End If
    Next i
Next k

fsT.SaveToFile FullPath, 2 'Save binary data To disk

End Sub
