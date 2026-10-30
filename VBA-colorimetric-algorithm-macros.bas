Attribute VB_Name = "Módulo1"
Sub Registration()

    Dim wsReg As Worksheet, wsSpec As Worksheet
    Set wsReg = Sheets("Registration and results")
    Set wsSpec = Sheets("Specimens")
    
    Dim newID As String
    newID = Trim(wsReg.Range("C6").Value)

    If newID = "" Then
        MsgBox "Please write an ID in C6 cell.", vbExclamation, "ID Empty"
        Exit Sub
    End If

    Dim existID As Range
    Set existID = wsSpec.Range("B:B").Find(newID, LookIn:=xlValues, LookAt:=xlWhole)

    If Not existID Is Nothing Then
        MsgBox "ID '" & newID & "' is already used. Try again.", vbCritical, "ID Duplicate"
        Exit Sub
    End If

    Dim newRow As Long
    newRow = wsSpec.Cells(wsSpec.Rows.Count, "B").End(xlUp).Row + 1
    If newRow < 4 Then newRow = 4
    
    wsSpec.Cells(newRow, "B").Value = wsReg.Range("C6").Value ' ID
    wsSpec.Cells(newRow, "C").Value = wsReg.Range("C7").Value ' R
    wsSpec.Cells(newRow, "D").Value = wsReg.Range("C8").Value ' G
    wsSpec.Cells(newRow, "E").Value = wsReg.Range("C9").Value ' B

    Dim lastRow As Long
    lastRow = wsSpec.Cells(wsSpec.Rows.Count, "B").End(xlUp).Row

    With wsSpec.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsSpec.Range("B4:B" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange wsSpec.Range("B4:E" & lastRow)
        .Header = xlGuess
        .Apply
    End With

    wsReg.Range("C6:C9").ClearContents
    Application.Goto wsSpec.Range("A1")
    Application.Goto wsReg.Range("C6")
End Sub
Sub AddRL()
Attribute AddRL.VB_Description = "Añadir a la base de datos cositas asies\r\n"
Attribute AddRL.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim wsRLr As Worksheet, wsBD As Worksheet
    Set wsRLr = Sheets("Add to Reference Library")
    Set wsRL = Sheets("Reference Library")

    Dim newID As String
    newID = Trim(wsRLr.Range("C6").Value)

    If newID = "" Then
        MsgBox "Please write an ID in C6 cell.", vbExclamation, "ID Empty"
        Exit Sub
    End If

    Dim existID As Range
    Set existID = wsRL.Range("B:B").Find(newID, LookIn:=xlValues, LookAt:=xlWhole)

    If Not existID Is Nothing Then
        MsgBox "ID '" & newID & "' is already used. Try again.", vbCritical, "ID Duplicate"
        Exit Sub
    End If

    Dim newRow As Long
    newRow = wsRL.Cells(wsRL.Rows.Count, "B").End(xlUp).Row + 1
    If newRow < 4 Then newRow = 4
    
    wsRL.Cells(newRow, "B").Value = wsRLr.Range("C6").Value ' ID
    wsRL.Cells(newRow, "C").Value = wsRLr.Range("C7").Value ' Defined Sex
    wsRL.Cells(newRow, "D").Value = wsRLr.Range("C8").Value ' R
    wsRL.Cells(newRow, "E").Value = wsRLr.Range("C9").Value ' G
    wsRL.Cells(newRow, "F").Value = wsRLr.Range("C10").Value ' B
    wsRL.Cells(newRow, "G").Value = wsRLr.Range("C11").Value ' Observations

    Dim lastRow As Long
    lastRow = wsRL.Cells(wsRL.Rows.Count, "C").End(xlUp).Row
    
    With wsRL.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsRL.Range("B4:B" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange wsRL.Range("B4:G" & lastRow)
        .Header = xlNo
        .Apply
    End With

    wsRLr.Range("C6:C11").ClearContents
    Application.Goto wsRL.Range("A1")
    Application.Goto wsRLr.Range("C6")
End Sub
Sub Search()

    Dim wsReg As Worksheet, wsSpec As Worksheet, wsRL As Worksheet
    Set wsReg = Sheets("Registration and results")
    Set wsSpec = Sheets("Specimens")
    Set wsRL = Sheets("Reference Library")

    Dim idSearch As String
    idSearch = wsReg.Range("H5").Value
    If idSearch = "" Then MsgBox "Write an ID in H5 cell": Exit Sub

    Dim found As Range
    Set found = wsSpec.Range("B:B").Find(idSearch, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not found Is Nothing Then
        Dim R1 As Double, G1 As Double, B1 As Double
        R1 = Val(found.Offset(0, 1).Value)
        G1 = Val(found.Offset(0, 2).Value)
        B1 = Val(found.Offset(0, 3).Value)
        
        wsReg.Range("H10").Value = R1
        wsReg.Range("H11").Value = G1
        wsReg.Range("H12").Value = B1
        
        ' Euclidian Algoritm
        Dim minimumDistance As Double, currentDistance As Double
        Dim lastRowRL As Long, i As Long
        Dim R2 As Double, G2 As Double, B2 As Double
        Dim predictedSex As String
        
        minimumDistance = 999999
        lastRowRL = wsRL.Cells(wsRL.Rows.Count, "D").End(xlUp).Row
        
        For i = 2 To lastRowRL
            R2 = Val(wsRL.Cells(i, 4).Value)
            G2 = Val(wsRL.Cells(i, 5).Value)
            B2 = Val(wsRL.Cells(i, 6).Value)
            
            currentDistance = Sqr((R1 - R2) ^ 2 + (G1 - G2) ^ 2 + (B1 - B2) ^ 2)
            
            If currentDistance < minimumDistance Then
                minimumDistance = currentDistance
                predictedSex = wsRL.Cells(i, 3).Value
            End If
        Next i
        
        wsReg.Range("H13").Value = predictedSex
        Dim foundRecord As Range
        Dim saveRow As Long
        
        Set foundRecord = wsSpec.Range("H:H").Find(idSearch, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundRecord Is Nothing Then
            saveRow = foundRecord.Row
        Else
            saveRow = wsSpec.Cells(wsSpec.Rows.Count, "H").End(xlUp).Row + 1
        End If
        wsSpec.Cells(saveRow, "H").Value = idSearch
        wsSpec.Cells(saveRow, "I").Value = R1
        wsSpec.Cells(saveRow, "J").Value = G1
        wsSpec.Cells(saveRow, "K").Value = B1
        wsSpec.Cells(saveRow, "L").Value = predictedSex
        wsSpec.Cells(saveRow, "M").Value = minimumDistance
        
        Application.OnTime Now + TimeValue("00:00:05"), "ClearResults"
    Else
        MsgBox "ID not found in Specimens list."
    End If
End Sub
Sub ClearResults()
    Dim wsReg As Worksheet
    Set wsReg = Sheets("Registration and results")
    
    wsReg.Range("H5").ClearContents
    wsReg.Range("H10:H13").ClearContents
End Sub


