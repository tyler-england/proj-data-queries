Attribute VB_Name = "LastParts"
Option Explicit

Sub LastPartsDue()
'''Lists last parts on Summary tab (along with description & date due)
    Dim i As Integer, sJob As String
    Dim sParts As String, oDictParts As Object
    
    i = 3
    sJob = UCase(Sheet3.Range("A" & i).Value)
    
    Do While sJob <> ""
        sParts = ""
        sParts = GetAllParts(sJob) 'find parts & dates --> assign to dict
        Sheet3.Range("J" & i).Value = sParts
        i = i + 1
        sJob = UCase(Sheet3.Range("A" & i).Value)
    Loop
    
End Sub

Function GetAllParts(sJob As String) As String
'''Returns string of parts & dates
    Dim sOut As String, wsShort As Worksheet
    Dim sArrOld As Variant, sArrNew(2, 19) As String
    Dim i As Integer, j As Integer, jInd As Integer
    Dim iStart As Integer, iEnd As Long
    Dim dDateOld As Date, dDateNew As Date
    Dim vRng As Variant
    
    Set wsShort = Sheet9
    On Error Resume Next
        iStart = WorksheetFunction.Match("*" & sJob & "**", wsShort.Range("A:A"), 0)
    On Error GoTo errhandler
    
    If iStart > 0 Then
        iEnd = wsShort.Range("A" & iStart).End(xlDown).Row - 1
        If iEnd > 10000 Then iEnd = wsShort.Range("C" & iEnd).End(xlUp).Row
        vRng = wsShort.Range("C" & iStart & ":I" & iEnd).Value2
        For i = 1 To iEnd - iStart + 1
            If vRng(i, 6) > 0 And InStr(vRng(i, 7), sJob) = 0 Then 'part with due date
                dDateNew = vRng(i, 7)
                If sArrNew(2, 19) <> "" Then 'old date is oldest in list
                    dDateOld = CDate(sArrNew(2, 19))
                Else 'list of parts/dates isn't full yet
                    dDateOld = 0
                End If
                If dDateNew > dDateOld Then 'new due date later than latest existing due date
                    sArrOld = sArrNew
                    jInd = 0
                    For j = 0 To 19 'find new date's place in the descending order
                        If sArrOld(0, j) = "" Then
                            jInd = j
                        ElseIf dDateNew > CDate(sArrOld(2, j)) Then
                            jInd = j
                        End If
                        If jInd > 0 Then Exit For
                    Next
                    If jInd > 0 Then 'a date needs to be replaced
                        For j = 0 To jInd - 1
                            sArrNew(0, j) = sArrOld(0, j)
                            sArrNew(1, j) = sArrOld(1, j)
                            sArrNew(2, j) = Format(sArrOld(2, j), "YYYY MMM DD")
                        Next
                        sArrNew(0, jInd) = vRng(i, 1)
                        sArrNew(1, jInd) = vRng(i, 2)
                        sArrNew(2, jInd) = Format(vRng(i, 7), "YYYY MMM DD")
                        For j = jInd + 1 To 19
                            sArrNew(0, j) = sArrOld(0, j - 1)
                            sArrNew(1, j) = sArrOld(1, j - 1)
                            sArrNew(2, j) = Format(sArrOld(2, j - 1), "YYYY MMM DD")
                        Next
                    End If
                End If
            End If
        Next
        For j = 0 To 19
            If sArrNew(0, j) <> "" Then sOut = sOut & sArrNew(0, j) & "//" & sArrNew(1, j) & "//" & sArrNew(2, j) & ";;"
        Next
    End If
    
    GetAllParts = sOut
    Exit Function
    
errhandler:
    Debug.Print "Error in GetAllParts function (" & sJob & ")"
    GetAllParts = ""
End Function
