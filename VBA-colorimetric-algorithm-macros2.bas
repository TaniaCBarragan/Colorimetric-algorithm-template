Attribute VB_Name = "Módulo2"
Sub Validation()
    Dim wsOut As Worksheet, wsSpec As Worksheet, wsRL As Worksheet
    Set wsOut = ActiveSheet
    Set wsSpec = Sheets("Specimens")
    Set wsRL = Sheets("Reference Library")

    Dim idSearch As String
    idSearch = Trim(wsOut.Range("C4").Value)
    If idSearch = "" Then
        MsgBox "Write an ID in C4 cell.", vbExclamation
        Exit Sub
    End If

    ' Find ID in Specimens sheet
    Dim found As Range
    Set found = wsSpec.Range("B:B").Find(idSearch, LookIn:=xlValues, LookAt:=xlWhole)

    If Not found Is Nothing Then
        Dim R1 As Double, G1 As Double, B1 As Double
        R1 = Val(found.Offset(0, 1).Value)
        G1 = Val(found.Offset(0, 2).Value)
        B1 = Val(found.Offset(0, 3).Value)

        ' Clear previous results
        Dim lastOutRow As Long
        lastOutRow = wsOut.Cells(wsOut.Rows.Count, "E").End(xlUp).Row
        If lastOutRow >= 4 Then
            wsOut.Range("E4:K" & lastOutRow).ClearContents
        End If

        Dim lastRowRL As Long, i As Long, outRow As Long
        Dim R2 As Double, G2 As Double, B2 As Double
        Dim currentDistance As Double, similarityPct As Double
        
        lastRowRL = wsRL.Cells(wsRL.Rows.Count, "D").End(xlUp).Row
        outRow = 4

        For i = 4 To lastRowRL
            R2 = Val(wsRL.Cells(i, 4).Value)
            G2 = Val(wsRL.Cells(i, 5).Value)
            B2 = Val(wsRL.Cells(i, 6).Value)

            currentDistance = Sqr((R1 - R2) ^ 2 + (G1 - G2) ^ 2 + (B1 - B2) ^ 2)

            ' Chromatic Similarity % (umbral 441.6729)
            similarityPct = (1 - (currentDistance / 441.6729)) * 100
            If similarityPct < 0 Then similarityPct = 0

            wsOut.Cells(outRow, "E").Value = wsRL.Cells(i, 2).Value ' ID of Reference Library
            wsOut.Cells(outRow, "F").Value = wsRL.Cells(i, 3).Value ' Defined sex
            wsOut.Cells(outRow, "G").Value = R2                     ' R
            wsOut.Cells(outRow, "H").Value = G2                     ' G
            wsOut.Cells(outRow, "I").Value = B2                     ' B
            wsOut.Cells(outRow, "J").Value = Round(currentDistance, 2)
            wsOut.Cells(outRow, "K").Value = Round(similarityPct, 2)

            outRow = outRow + 1
        Next i

        ' Sort from lowest chromatic distance to highest
        If outRow > 5 Then
            wsOut.Sort.SortFields.Clear
            wsOut.Sort.SortFields.Add Key:=wsOut.Range("J4:J" & outRow - 1), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With wsOut.Sort
                .SetRange wsOut.Range("E4:K" & outRow - 1)
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .Apply
            End With
        End If
    Else
        MsgBox "ID not found in Specimens list.", vbCritical
    End If
End Sub
