Attribute VB_Name = "CostsAndHrs"
Option Explicit

Sub RefreshData()
    Dim objConn As Variant, bBG As Boolean
    Application.ScreenUpdating = False
    For Each objConn In ThisWorkbook.Connections 'https://stackoverflow.com/questions/22083668/wait-until-activeworkbook-refreshall-finishes-vba
        Debug.Print objConn.Name
        bBG = objConn.OLEDBConnection.BackgroundQuery 'Get current background-refresh value
        objConn.OLEDBConnection.BackgroundQuery = False 'Temporarily disable background-refresh
        objConn.Refresh 'Refresh this connection
        objConn.OLEDBConnection.BackgroundQuery = bBG 'Set background-refresh value back to original value
    Next
    Call UpdateCosts
    Sheet3.Range("A1").Value = "Updated: " & Format(Now, "yyyy mmm dd, HH:MM:SS")
    Call ShortParts
    Sheet9.Range("A1").Value = "Updated: " & Format(Now, "yyyy mmm dd, HH:MM:SS")
    Call LastPartsDue
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(1).Activate
    ThisWorkbook.Worksheets(1).Range("A1").Select
End Sub

Sub UpdateCosts()
    Dim wSheet As Worksheet
    Dim iRow As Long, i As Integer
    Dim sMach As String, sJob As String
    Dim oDictCodes As Object, varRng As Variant
    
    Set wSheet = ThisWorkbook.Worksheets(3)
    iRow = wSheet.Range("A3").End(xlDown).Row
    If iRow < 10000 Then
        wSheet.Rows("3:" & iRow).Delete
    Else
        wSheet.Rows(3).Delete
    End If
    
    wSheet.Range("A3:A501").Value = Sheet2.Range("A2:A" & 500).Value
    iRow = 3
    sJob = UCase(wSheet.Range("A" & iRow).Value)
    Do While sJob <> ""
        With wSheet
            .Range("B" & iRow).Formula = "=MAX(SUMIFS('Material (Line Items)'!I:I,'Material (Line Items)'!A:A,A" & iRow & ")," & _
                                                "SUMIFS('Material (Planned)'!I:I,'Material (Planned)'!A:A,A" & iRow & "))"
            sMach = GetMach(sJob) 'find smach (prod line not impt)
            If sMach <> "" Then
                Set oDictCodes = GetCodes(sMach)
                varRng = GetRng(sJob)
                On Error Resume Next
                    i = UBound(varRng)
                    Err.Clear
                On Error GoTo 0 'errhandler
                If i > 0 Then 'sJob (the CO) was found in the labor table
                    .Range("C" & iRow).Value = GetHrs(varRng, oDictCodes, "ME") 'sum up ME
                    .Range("D" & iRow).Value = GetHrs(varRng, oDictCodes, "EE") 'sum up EE
                    .Range("E" & iRow).Value = GetHrs(varRng, oDictCodes, "SW") 'sum up SW
                    .Range("F" & iRow).Value = GetHrs(varRng, oDictCodes, "MA") 'sum up MA
                    .Range("G" & iRow).Value = GetHrs(varRng, oDictCodes, "EA") 'sum up EA
                    .Range("H" & iRow).Value = GetHrs(varRng, oDictCodes, "TS") 'sum up TS
                End If
                i = 0
            End If
            iRow = iRow + 1
            sJob = .Range("A" & iRow).Value
        End With
    Loop
    Exit Sub
errhandler:
    MsgBox "Error"
End Sub
Function GetMach(sJob As String) As String
'''Returns machine as string
    Dim wSheet As Worksheet
    Dim i As Integer, iCnt As Integer, iRow As Integer, iEmpty As Integer
    Dim bKeepLooking As Boolean
    Dim sProj As String, sOut As String, sCarrSufx(7) As String
    Dim vRng As Variant, vVar As Variant
    
    sCarrSufx(0) = "MIN"
    sCarrSufx(1) = "PLT"
    sCarrSufx(2) = "P12"
    sCarrSufx(3) = "P18"
    sCarrSufx(4) = "P06"
    sCarrSufx(5) = "UNI"
    sCarrSufx(6) = "U2K"
    sCarrSufx(7) = "V12"
    
    sJob = UCase(sJob)
    Set wSheet = Sheet11 'try hrs
    iCnt = WorksheetFunction.CountIf(wSheet.Range("A:A"), sJob & "*")
    If iCnt > 0 Then 'job is in list
        vRng = wSheet.Range("A1:C5000").Value2
        i = 1
        Do While iEmpty < 3
            If Trim(UCase(vRng(i, 1))) = sJob Then
                sProj = Trim(UCase(vRng(i, 3)))
                If sProj Like "C*" Then 'carr
                    If sProj Like "C*CELL8" Then
                        sOut = "CELL8"
                        Exit Do
                    ElseIf sProj Like "C*LAB3" Then
                        sOut = "LAB3"
                        Exit Do
                    ElseIf sProj Like "C0*" Then
                        sProj = UCase(Right(sProj, 3))
                        For Each vVar In sCarrSufx
                            If vVar = sProj Then
                                sOut = vVar
                                Exit For
                            End If
                        Next
                        If sOut <> "" Then Exit Do
                    End If
                ElseIf sProj Like "W####-*" Then 'mateer
                    If Mid(sProj, 2, 1) < 3 Then 'semiautomatic
                        sOut = "SEMI"
                        Exit Do
                    ElseIf Mid(sProj, 2, 1) < 6 Or Mid(sProj, 2, 1) = 9 Then 'automatic
                        sOut = "AUTO"
                        Exit Do
                    ElseIf Mid(sProj, 2, 1) = 6 Then 'rotary
                        sOut = "ROTARY"
                        Exit Do
                    End If
                ElseIf sProj Like "W4*" Or sProj Like "W7*" Then 'burt?
                    sOut = "408/704"
                    Exit Do
                End If
            ElseIf vRng(i, 1) = 0 Then
                iEmpty = iEmpty + 1
            Else
                iEmpty = 0
            End If
            i = i + 1
        Loop
    End If
    
    If sOut = "" Then
        Set wSheet = Sheet10 'try mat'l (LI)
        
        
        If sOut = "" Then
            Set wSheet = Sheet8 'try mat'l (Plan)
            
            
        End If
    End If
    
    If sOut = "PLT" Then sOut = "POWERFUGE"
    
    GetMach = sOut
    
End Function

Function GetCodes(sMach As String) As Object
'''Returns dictionary of labor codes Key=labor type, Item=array of codes
    Dim iRow As Integer, iCol As Integer, i As Integer
    Dim oDictCodes As Object, oDictCols As Object 'color values
    Dim bFound As Boolean, lCol As Long
    Dim sKey As String, arrVals As Variant
    Dim wSheet As Worksheet
    
    On Error GoTo errhandler
    
    Set wSheet = Sheet4
    Set oDictCols = CreateObject("Scripting.Dictionary") 'color values for each category
    Set oDictCodes = CreateObject("Scripting.Dictionary") 'indices for output arrays
    
    With wSheet
        iCol = WorksheetFunction.Match(sMach, .Range("2:2"), 0) 'find machine column
        iRow = 1 'find color codes
        Do While UCase(.Cells(iRow, 1).Value) <> "ENGINEERING" 'skip bast burt machines
            iRow = iRow + 1
        Loop
        Do While .Cells(iRow, 1).Value > 0 'get colors
            If Not oDictCols.exists("ME") And UCase(.Cells(iRow, 1).Value) = "ME" Then
                oDictCols.Add "ME", .Cells(iRow, 1).Interior.Color
            ElseIf Not oDictCols.exists("EE") And UCase(.Cells(iRow, 1).Value) = "EE" Then
                oDictCols.Add "EE", .Cells(iRow, 1).Interior.Color
            ElseIf Not oDictCols.exists("SW") And UCase(.Cells(iRow, 1).Value) = "SW" Then
                oDictCols.Add "SW", .Cells(iRow, 1).Interior.Color
            ElseIf Not oDictCols.exists("TS") And UCase(.Cells(iRow, 1).Value) = "TS" Then
                oDictCols.Add "TS", .Cells(iRow, 1).Interior.Color
            ElseIf Not oDictCols.exists("MA") And UCase(.Cells(iRow, 1).Value) = "MA" Then
                oDictCols.Add "MA", .Cells(iRow, 1).Interior.Color
            ElseIf Not oDictCols.exists("EA") And UCase(.Cells(iRow, 1).Value) = "EA" Then
                oDictCols.Add "EA", .Cells(iRow, 1).Interior.Color
'            ElseIf Not oDictCols.exists("TS") And UCase(.Cells(iRow, 1).Value) = "TS" Then 'only used if Assy has diff test codes
'                oDictCols.Add "TS", .Cells(iRow, 1).Interior.Color
            End If
            iRow = iRow + 1
        Loop

        iRow = 3 ' go through codes for machine
        Do While .Cells(iRow, iCol).Value > 0
            lCol = .Cells(iRow, iCol).Interior.Color
            bFound = False
            For i = 0 To oDictCols.Count - 1 'find matching category
                If lCol = oDictCols.items()(i) Then 'color found in color dict
                    bFound = True
                    If .Cells(iRow, iCol).Value > 0 Then
                        If oDictCodes.exists(oDictCols.keys()(i)) Then 'reassign with new code added
                            sKey = oDictCols.keys()(i)
                            arrVals = oDictCodes(sKey)
                            ReDim Preserve arrVals(UBound(arrVals) + 1)
                            arrVals(UBound(arrVals)) = .Cells(iRow, iCol).Value
                            oDictCodes(sKey) = arrVals
                        Else 'add first code
                            oDictCodes.Add oDictCols.keys()(i), Array(.Cells(iRow, iCol).Value)
                        End If
                    End If
                    Exit For
                End If
            Next
            If Not bFound And lCol <> 255 Then
                Debug.Print "Unable to find color " & lCol
            End If
            iRow = iRow + 1
        Loop
    End With
    
    Set GetCodes = oDictCodes
    
    Exit Function
errhandler:
    MsgBox "Error in GetCodes function... can't find the labor codes"
End Function

Function GetRng(sJob As String) As Variant
'''Returns matrix of relevant rows from PQ table (faster than iterating through table)
    Dim varOut As Variant, wSheet As Worksheet
    Dim i As Integer, iStart As Integer, iEnd As Integer
    
    Set wSheet = Sheet11 'labor hours
    i = 1
    With wSheet
        Do While Trim(.Range("A" & i).Value) <> sJob
            i = i + 1
            If i > 5000 Then Exit Do
        Loop
        If i < 5000 Then iStart = i
        Do While Trim(.Range("A" & i).Value) = sJob
            i = i + 1
            If i > 5000 Then Exit Do
        Loop
        If i < 5000 Then iEnd = i - 1
        If iStart = 0 Or iEnd = 0 Then Exit Function
        varOut = .Range("A" & iStart & ":M" & iEnd).Value2
    End With
    GetRng = varOut
End Function
Function GetHrs(varRng As Variant, oDict As Object, sTyp As String) As Single
'''Returns hours for labor type
'styp can be ME, EE, SW, MA, EA, TS
    Dim arrCodes As Variant, varVar As Variant
    Dim i As Integer, iTst As Integer 'to test arrCodes
    Dim sOutTot As Single

    arrCodes = oDict(sTyp)
    On Error Resume Next
        iTst = arrCodes(0)
        Err.Clear
    On Error GoTo 0 'errhandler
    If iTst > 0 Then 'arrCodes has at least 1 code
        For Each varVar In arrCodes 'if it's in the job, add the hours to total
            If IsNumeric(varVar) Then 'must be a labor code
                For i = 1 To UBound(varRng)
                    If CInt(varRng(i, 4)) = CInt(varVar) Then
                        sOutTot = sOutTot + varRng(i, 13)
                    End If
                Next
            End If
        Next
    End If
    GetHrs = sOutTot
    
End Function

