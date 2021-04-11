Attribute VB_Name = "PartShorts"
Type zQryShort
'MOMAST (mm, T01)
OSTAT As String
ORDNO As String
FITEM As String
FDESC As String
FSKLC As String
STARTDATE As String     'from SSTDT
DUEDATE As String       'from ODUDT
OPENQTY As String
JOBNO As String
REFNO As String
MOWH As String
'MODATA (md, T02)
CITEM As String
CITEMnoWH As String
CDESC As String
QTREQ As String
ISQTY As String
LIDTFORMAT As String    'from LISDT
USRSQ As String
'SLQNTY (sq, T05)
LLOCN As String
LQNTY As String
LBHNO As String 'batch/lot
'ITEMBL (ib, T06)
ALLOCATED As String
LTCOD As String
PLANIB As String
'MOMAST (mm2, T07)
MCORDNO As String
MCOPENQTY As String
MCDUEDATE As String
MCMNTDATE As String
MCDPTNO As String
MCOSTAT As String
MCREMOPS As String
MCFSKLC As String
'POITEM (pi, T08)
POITNBR As String
PONUM As String
POACTQTY As String
PODUEDATE As String
POMNTDATE As String
poBUYNO As String
POSTATUS As String
poBLCOD As String
poPMPSTTS As String 'PO Mast - PO Sts
'REQHDF (rq, T09)
rqREQNO As String
rqQTYOR As String
rqDUEDT As String
rqMDATE As String
rqDPTNO As String
rqNAMER As String
rqORDNO As String
'ITEMASA (imf, for filtering by Item Type on the Finished Item)
imfITTYPE As String
'ITEMASA (imc, for UCDEF on components)
imcUCDEF As String
'VENNAM (vn)
vnVNAMA As String 'vn joined to component PO Item
vnibVNAMA As String 'vnib joined to component item balance
End Type

'User-defined variables and arrays:
Public MOs() As zQryShort           'data for open MOs
Public MOsHiNdx As Long             'highest index in MOs() array
Public MOsUB As Long                'upper bound of MOs() array

Function Load_MOs(Optional COnum = 0, Optional StartRw As Integer = 4)
Dim strSQL As String, COsql As String, MOsql As String, strTracer As String, strStatus As String, Env As String, EnvLtr As String, SLQTYJoin As String, MCJoin As String, PIJoin As String, whTags As Boolean
Dim TitleStyle, Server, ITsql As String, WHsql
'On Error GoTo errhandler
whTags = False
'If Range("showWHtags") = "Always" Then: whTags = True
COsql = ""

If COnum > 0 Then
    COsql = "mm.JOBNO like '%" & COnum & "%'"
Else
'    For c = 2 To 151
'        If UCase(Left(Cells(2, c), 1)) = "M" And Len(Cells(2, c)) = 7 Then 'MO
'            COsql = COsql & "mm.ORDNO = '" & UCase(Cells(2, c)) & "' or "
'        ElseIf Cells(2, c) <> "" Then 'CO
'            COsql = COsql & "mm.JOBNO like '%" & UCase(Cells(2, c)) & "%' or md.CITEM = '" & UCase(Cells(2, c)) & "' or "
'        End If
'    Next c
    c = 2
    Do While Sheet2.Range("A" & c).Value > 0
        COsql = COsql & "mm.JOBNO like '%" & UCase(Sheet2.Range("A" & c).Value) & "%' or md.CITEM = '" & UCase(Sheet2.Range("A" & c).Value) & "' or "
        c = c + 1
    Loop
    If COsql = "" Then
        MsgBox "No CO numbers entered!", vbExclamation
        Exit Function
    End If
    COsql = Left(COsql, Len(COsql) - 4)
End If

'ITsql = ""
'For Each Cl In Worksheets("Options").Range("ITselect")
'    If Cl = "Yes" Then
'        ITsql = ITsql & "imf.ITTYP = '" & Left(Cl.Offset(0, -1), 1) & "' or "
'    End If
'Next Cl
'If ITsql <> "" Then: ITsql = " AND (" & Left(ITsql, Len(ITsql) - 4) & ")"

ITsql = "imf.ITTYP = '0' or imf.ITTYP = '1' or imf.ITTYP = '2' or imf.ITTYP = '9'"
'Item type codes 0=Phantom, 1=Assembly, 2=Fabricated, 9=User Option
ITsql = " AND (" & ITsql & ")"

'WHsql = ""
'If Range("MOhouse") <> "" Then: WHsql = " AND mm.FITWH='" & Range("MOhouse") & "'"

Env = "AKR"
EnvLtr = "W"
'TitleStyle = Worksheets("Options").Range("TitleStyle")

  MOsHiNdx = 0
'  If blnAbort Then Exit Function
'  If blnAbortGlobal Then Exit Function
  
  'strTracer = "AA100"

  'NOTE:  ADODB objects require a reference to
  '       Microsoft ActiveX Data Objects 2.8 Library,
  '       which also provides enumerations (adChar, adparamInput, etc.)
  Dim conX As New ADODB.Connection 'cnXAPSAAKR
  Dim cmdX As New ADODB.Command 'cm_cmalib" & envltr & "_pdb100
  Dim rstX As New ADODB.Recordset

  'strTracer = "AA150"
  conX.Provider = "IBMDA400"
  conX.Properties("Force Translate") = 0
  conX.Open "Provider=IBMDA400;Data Source=XAPSA" & Env, "", ""
  
  Set cmdX.ActiveConnection = conX
  cmdX.CommandType = adCmdText

  'Application.StatusBar = "Waiting for XA to return data..."

'Set Join Criteria for Stock Location Qty based on warehouse pref:
SLQTYJoin = " on md.CITEM = sq.ITNBR"
MCJoin = " on md.CITEM = mc.FITEM and mc.ORQTY + mc.QTDEV - mc.QTSPL - mc.QTSCP - mc.QTYRC > 0 and mc.OSTAT < '45' and md.ORDNO<>mc.ORDNO" 'added last "and" to prevent MO from referencing itself with the new Clearwater process
PIJoin = " on (md.CITEM = pi.ITNBR or mc.ORDNO = pi.JOBNO or mm.ORDNO = pi.JOBNO) and pi.STKQT < pi.QTYOR + pi.QTDEV and pi.STAIC < '50'"
'If Range("MOCompWH") = "MO Comp WH Only" Then
    SLQTYJoin = SLQTYJoin & " and md.CITWH=sq.HOUSE"
    MCJoin = MCJoin & " and md.CITWH=mc.FITWH"
    PIJoin = PIJoin & " and md.CITWH=pi.HOUSE"
'End If

strSQL = "select"
strSQL = strSQL & " mm.OSTAT, mm.ORDNO as MOORDNO, mm.FITEM, mm.FDESC, mm.FSKLC, mm.SSTDT, mm.ODUDT, mm.ORQTY, mm.JOBNO, mm.QTYRC, mm.QTDEV as MOQTDEV, mm.QTSCP, mm.QTSPL, mm.REFNO, mm.FITWH as MOWH,"
strSQL = strSQL & " md.ORDNO, md.CITEM, md.CDESC, md.QTREQ, md.ISQTY, md.LISDT, md.SEQNM, md.CITWH as CompWH," 'SEQNM was user sequence (USRSQ), changed 5/16/18
strSQL = strSQL & " ib.MALQT, ib.PLREQ, ib.ITNBR, ib.HOUSE, ib.MOHTQ, ib.MPRPQ, ib.MPUPQ, ib.LTCOD, ib.PLANIB,"
strSQL = strSQL & " sq.LLOCN, sq.LQNTY, sq.LBHNO, sq.ITNBR, sq.HOUSE as LocHouse,"
strSQL = strSQL & " pi.ORDNO as POORDNO, pi.QTYOR, pi.QTDEV as POQTDEV, pi.DUEDT, pi.MDATE, pi.BUYNO, pi.STAIC, pi.ITNBR as POITNBR, pi.STKQT as POSTKQT, pi.LINSQ, pi.POISQ, pi.BLCOD, pi.JOBNO as PIJOBNO,"
strSQL = strSQL & " pb.ORDNO as pbORDNO, pb.ITNBR, pb.LINSQ, pb.POISQ, pb.STAIC as pbSTAIC, pb.STKQT as pbSTKQT, pb.RELQT as pbRELQT, pb.RELDT as pbRELDT, pb.BLKSQ as pbBLKSQ,"
strSQL = strSQL & " pm.ORDNO as pmORDNO, pm.PSTTS as pmPSTTS,pm.HOUSE as pmWH,"
strSQL = strSQL & " mc.ORDNO as MCORDNO, mc.ORQTY as MCORQTY, mc.ODUDT as MCODUDT, mc.MDATE as MCMDATE, mc.DPTNO as MCDPTNO, mc.OSTAT as MCOSTAT, mc.OPSNS as MCOPSNS, mc.WCCUR as MCWCCUR, "
strSQL = strSQL & " mc.FSKLC as MCFSKLC, mc.FITEM as MCFITEM, mc.QTYRC as MCQTYRC, mc.QTDEV as MCQTDEV, mc.QTSCP as MCQTSCP, mc.QTSPL as MCQTSPL, mc.JOBNO as MCJOBNO, mc.FITWH as mcHDRWH, "
strSQL = strSQL & " imf.ITNBR, imf.ITTYP,"
strSQL = strSQL & " imc.ITNBR, imc.UCDEF,"
strSQL = strSQL & " vn.VNAMA, vn.VNDNR,"
strSQL = strSQL & " vnib.VNAMA as vnibVNAMA,"
strSQL = strSQL & " rq.ITNBR, rq.REQNO, rq.QTYOR as RQQTYOR, rq.DUEDT as RQDUEDT, rq.MDATE as RQMDATE, rq.DPTNO as RQDPTNO, rq.NAMER, rq.ORDNO as RQORDNO"

strSQL = strSQL & " from amflib" & EnvLtr & ".MOMAST mm"
strSQL = strSQL & " inner join amflib" & EnvLtr & ".MODATA md"
strSQL = strSQL & " on mm.ORDNO = md.ORDNO"

strSQL = strSQL & " inner join amflib" & EnvLtr & ".ITEMBL ib"
strSQL = strSQL & " on md.CITEM = ib.ITNBR and md.CITWH=ib.HOUSE"

strSQL = strSQL & " left outer join amflib" & EnvLtr & ".SLQNTY sq"
strSQL = strSQL & SLQTYJoin 'see above

strSQL = strSQL & " left outer join amflib" & EnvLtr & ".MOMAST mc"
strSQL = strSQL & MCJoin

strSQL = strSQL & " left outer join amflib" & EnvLtr & ".POITEM pi"
strSQL = strSQL & PIJoin
'STRsql = STRsql & " on (md.CITEM = pi.ITNBR or mc.ORDNO = pi.JOBNO) and pi.STKQT < pi.QTYOR + pi.QTDEV and pi.STAIC < '50'"
'STRsql = STRsql & " left outer join amflib" & EnvLtr & ".POITEM pi"
'STRsql = STRsql & " on mc.ORDNO = pi.JOBNO"

strSQL = strSQL & " left outer join amflib" & EnvLtr & ".POMAST pm"
strSQL = strSQL & " on pi.ORDNO = pm.ORDNO"

strSQL = strSQL & " left outer join amflib" & EnvLtr & ".POBLKT pb"
strSQL = strSQL & " on pi.ITNBR=pb.ITNBR and pi.ORDNO=pb.ORDNO and pi.LINSQ=pb.LINSQ and pi.POISQ=pb.POISQ and pb.STKQT < pb.RELQT and pb.STAIC < '50'"

strSQL = strSQL & " left outer join amflib" & EnvLtr & ".ITMRVA imf"
strSQL = strSQL & " on mm.FITEM = imf.ITNBR"

strSQL = strSQL & " left outer join amflib" & EnvLtr & ".ITMRVA imc"
strSQL = strSQL & " on md.CITEM = imc.ITNBR"

strSQL = strSQL & " left outer join amflib" & EnvLtr & ".VENNAM vn"
strSQL = strSQL & " on pi.VNDNR = vn.VNDNR"

strSQL = strSQL & " left outer join amflib" & EnvLtr & ".VENNAM vnib"
strSQL = strSQL & " on ib.VNDNR = vnib.VNDNR"

strSQL = strSQL & " left outer join amflib" & EnvLtr & ".REQHDF rq"
strSQL = strSQL & " on md.CITEM = rq.ITNBR and rq.ORDNO = ''"

strSQL = strSQL & " where (mm.OSTAT = '10' or mm.OSTAT = '40') and (" & COsql & ")" & ITsql & WHsql
'STRsql = STRsql & " and md.ISQTY > md.QTREQ"
'STRsql = STRsql & " where (mm.OSTAT = '10' or mm.OSTAT = '40') and mm.JOBNO = '361496'"
'STRsql = STRsql & " where (mm.OSTAT = '10' or mm.OSTAT = '40') and mm.ORDNO = 'M138660'"
'If TitleStyle = "By Stock Loc" Then
'    strSQL = strSQL & " order by mm.JOBNO, mm.FSKLC, mm.SSTDT, mm.ORDNO, mm.FITEM" 'CO->StkLoc->StartDt->MO->FnItem
'Else
    strSQL = strSQL & " order by mm.JOBNO, mm.SSTDT, mm.ORDNO, mm.FITEM" 'CO->StartDt->MO->FnItem
'End If
'STRsql = STRsql & " order by mm.JOBNO, mm.FSKLC, mm.SSTDT, mm.ORDNO, mm.FITEM" ', md.USRSQ, md.CITEM" 'Remove sort by FSKLC at Michele's request, if Akron wants it back then make configurable.

Sheet9.Range("J1").Value = strSQL

  'strTracer = "AA200"
  cmdX.CommandText = strSQL
  rstX.CursorLocation = adUseClient
  rstX.CacheSize = 100
  rstX.Open cmdX
  
  'strTracer = "AA250"
  MOsUB = 50000
  ReDim MOs(0 To MOsUB)

  'strStatus = "Retrieving data from XA..."
  'Application.StatusBar = strStatus

While Not rstX.EOF
  MOsHiNdx = MOsHiNdx + 1
  If MOsHiNdx > MOsUB Then 'increase upper bound
    MOsUB = MOsUB + 50000
    ReDim Preserve MOs(0 To MOsUB)
  End If
  With MOs(MOsHiNdx)
    .OSTAT = RTrim$(CStr(rstX.Fields("OSTAT").Value))
    .ORDNO = RTrim$(CStr(rstX.Fields("MOORDNO").Value))
    .FITEM = RTrim$(CStr(rstX.Fields("FITEM").Value))
    .FDESC = RTrim$(CStr(rstX.Fields("FDESC").Value))
    .FSKLC = RTrim$(CStr(rstX.Fields("FSKLC").Value))
    .OPENQTY = rstX.Fields("ORQTY").Value - rstX.Fields("QTYRC").Value + rstX.Fields("MOQTDEV").Value - rstX.Fields("QTSCP").Value - rstX.Fields("QTSPL").Value
    .JOBNO = RTrim$(CStr(rstX.Fields("JOBNO").Value))
    '.imcUCDEF = RTrim$(CStr(rstX.Fields("UCDEF").Value))
    .CITEM = RTrim$(CStr(rstX.Fields("CITEM").Value))
    .CITEMnoWH = RTrim$(CStr(rstX.Fields("CITEM").Value))
    .CDESC = RTrim$(CStr(rstX.Fields("CDESC").Value))
    .QTREQ = RTrim$(CStr(rstX.Fields("QTREQ").Value))
    .ISQTY = RTrim$(CStr(rstX.Fields("ISQTY").Value))
    .USRSQ = RTrim$(CStr(rstX.Fields("SEQNM").Value)) 'changed from USRSQ to SEQNM 5/16/18, keeping the .USRSQ variable for convenience
    .ALLOCATED = rstX.Fields("MALQT").Value + rstX.Fields("PLREQ").Value
    .LTCOD = RTrim$(CStr(rstX.Fields("LTCOD").Value))
    .PLANIB = RTrim$(CStr(rstX.Fields("PLANIB").Value))
    .MOWH = "[" & RTrim$(CStr(rstX.Fields("MOWH").Value)) & "]"
    If RTrim$(CStr(rstX.Fields("CompWH").Value)) <> RTrim$(CStr(rstX.Fields("MOWH").Value)) Or whTags = True Then: .CITEM = .CITEM & " [" & RTrim$(CStr(rstX.Fields("CompWH").Value)) & "]"
    If Not IsNull(rstX.Fields("vnibVNAMA")) Then: .vnibVNAMA = "(Def) " & RTrim$(CStr(rstX.Fields("vnibVNAMA").Value))
    .imfITTYPE = RTrim$(CStr(rstX.Fields("ITTYP").Value)) 'Item Type from Item Master (using Item Master to avoid issues with Clearwater)
    .REFNO = RTrim$(CStr(rstX.Fields("REFNO").Value))
    
    'strTracer = "AA420"
    If Not IsNull(rstX.Fields("POORDNO")) Then
        'strTracer = "AA422"
        If Not IsNull(rstX.Fields("pbORDNO")) Then
            .PONUM = RTrim$(CStr(rstX.Fields("POORDNO").Value)) & "(" & RTrim$(CStr(rstX.Fields("pbBLKSQ").Value)) & ")"
            .POACTQTY = rstX.Fields("pbRELQT").Value - rstX.Fields("pbSTKQT").Value
            .POSTATUS = RTrim$(CStr(rstX.Fields("pbSTAIC").Value))
            .PODUEDATE = ""
            strX = RTrim$(CStr(rstX.Fields("pbRELDT").Value))
            If Len(strX) = 7 Then
                .PODUEDATE = Mid(strX, 4, 2) & "/" & Mid(strX, 6, 2) & "/" & Mid(strX, 2, 2)
            End If
        Else
            .PONUM = RTrim$(CStr(rstX.Fields("POORDNO").Value))
            .POACTQTY = rstX.Fields("QTYOR").Value + rstX.Fields("POQTDEV").Value - rstX.Fields("POSTKQT").Value
            If RTrim$(CStr(rstX.Fields("BLCOD").Value)) = 1 Then 'Blanket PO with no open releases
                .PONUM = .PONUM & "(B!)"
                .POACTQTY = 0
            End If
            .POSTATUS = RTrim$(CStr(rstX.Fields("STAIC").Value))
            .PODUEDATE = ""
            strX = RTrim$(CStr(rstX.Fields("DUEDT").Value))
            If Len(strX) = 7 Then
                .PODUEDATE = Mid(strX, 4, 2) & "/" & Mid(strX, 6, 2) & "/" & Mid(strX, 2, 2)
            End If
        End If
        .POITNBR = RTrim$(CStr(rstX.Fields("POITNBR").Value))
        .poPMPSTTS = rstX.Fields("pmPSTTS") 'PO Mast - PO Status
        If RTrim$(CStr(rstX.Fields("CompWH").Value)) <> RTrim$(CStr(rstX.Fields("pmWH").Value)) Or whTags = True Then: .PONUM = .PONUM & " [" & RTrim$(CStr(rstX.Fields("pmWH").Value)) & "]"
        .POMNTDATE = ""
        strX = RTrim$(CStr(rstX.Fields("MDATE").Value))
        If Len(strX) = 7 Then
          .POMNTDATE = Mid(strX, 4, 2) & "/" & Mid(strX, 6, 2) & "/" & Mid(strX, 2, 2)
        End If
        .poBUYNO = RTrim$(CStr(rstX.Fields("BUYNO").Value))
        If Not IsNull(rstX.Fields("VNAMA")) Then: .vnVNAMA = RTrim$(CStr(rstX.Fields("VNAMA").Value))
        .poBLCOD = RTrim$(CStr(rstX.Fields("BLCOD").Value)) 'needed? Yes! (for Jim Guth's issue-Blanket PO with no open releases)
    End If

    'strTracer = "AA403"
    
    If Not IsNull(rstX.Fields("REQNO")) Then
        .rqREQNO = RTrim$(CStr(rstX.Fields("REQNO").Value))
        .rqQTYOR = RTrim$(CStr(rstX.Fields("RQQTYOR").Value))
        .rqDPTNO = RTrim$(CStr(rstX.Fields("RQDPTNO").Value))
        .rqNAMER = RTrim$(CStr(rstX.Fields("NAMER").Value))
        .rqORDNO = RTrim$(CStr(rstX.Fields("RQORDNO").Value))
        .rqDUEDT = ""
        strX = RTrim$(CStr(rstX.Fields("RQDUEDT").Value))
        If Len(strX) = 7 Then
          .rqDUEDT = Mid(strX, 4, 2) & "/" & Mid(strX, 6, 2) & "/" & Mid(strX, 2, 2)
        End If
        .rqMDATE = ""
        strX = RTrim$(CStr(rstX.Fields("RQMDATE").Value))
        If Len(strX) = 7 Then
          .rqMDATE = Mid(strX, 4, 2) & "/" & Mid(strX, 6, 2) & "/" & Mid(strX, 2, 2)
        End If
    End If

    If Not IsNull(rstX.Fields("MCORDNO")) Then
        .MCORDNO = RTrim$(CStr(rstX.Fields("MCORDNO").Value))
        If RTrim$(CStr(rstX.Fields("CompWH").Value)) <> RTrim$(CStr(rstX.Fields("mcHDRWH").Value)) Or whTags = True Then: .MCORDNO = .MCORDNO & " [" & RTrim$(CStr(rstX.Fields("mcHDRWH").Value)) & "]"
        .MCOPENQTY = rstX.Fields("MCORQTY").Value - rstX.Fields("MCQTYRC").Value + rstX.Fields("MCQTDEV").Value - rstX.Fields("MCQTSCP").Value - rstX.Fields("MCQTSPL").Value
        '.MCORQTY = RTrim$(CStr(rstX.Fields("MCORQTY").Value))
        .MCDPTNO = RTrim$(CStr(rstX.Fields("MCDPTNO").Value))
        .MCOSTAT = RTrim$(CStr(rstX.Fields("MCOSTAT").Value))
        .MCFSKLC = RTrim$(CStr(rstX.Fields("MCFSKLC").Value))
        .MCREMOPS = RTrim$(CStr(rstX.Fields("MCWCCUR").Value)) & "(" & RTrim$(CStr(rstX.Fields("MCOPSNS").Value)) & ")"
        .MCDUEDATE = ""
        strX = RTrim$(CStr(rstX.Fields("MCODUDT").Value))
        If Len(strX) = 7 Then
          .MCDUEDATE = Mid(strX, 4, 2) & "/" & Mid(strX, 6, 2) & "/" & Mid(strX, 2, 2)
        End If
        .MCMNTDATE = ""
        strX = RTrim$(CStr(rstX.Fields("MCMDATE").Value))
        If Len(strX) = 7 Then
          .MCMNTDATE = Mid(strX, 4, 2) & "/" & Mid(strX, 6, 2) & "/" & Mid(strX, 2, 2)
        End If
    End If
    
    If Not IsNull(rstX.Fields("LLOCN")) Then
        .LLOCN = RTrim$(CStr(rstX.Fields("LLOCN").Value))
        If rstX.Fields("LBHNO") <> "" Then: .LLOCN = .LLOCN & "(" & RTrim$(CStr(rstX.Fields("LBHNO").Value)) & ")"
        If RTrim$(CStr(rstX.Fields("CompWH").Value)) <> RTrim$(CStr(rstX.Fields("LocHouse").Value)) Or whTags = True Then: .LLOCN = .LLOCN & " [" & RTrim$(CStr(rstX.Fields("LocHouse").Value)) & "]"
    End If
    If Not IsNull(rstX.Fields("LQNTY")) Then: .LQNTY = RTrim$(CStr(rstX.Fields("LQNTY").Value))
    
    .STARTDATE = ""
    strX = RTrim$(CStr(rstX.Fields("SSTDT").Value))
    If Len(strX) = 7 Then
      .STARTDATE = Mid(strX, 4, 2) & "/" & Mid(strX, 6, 2) & "/" & Mid(strX, 2, 2)
    End If
    
    .DUEDATE = ""
    strX = RTrim$(CStr(rstX.Fields("ODUDT").Value))
    If Len(strX) = 7 Then
      .DUEDATE = Mid(strX, 4, 2) & "/" & Mid(strX, 6, 2) & "/" & Mid(strX, 2, 2)
    End If
    
    .LIDTFORMAT = ""
    strX = RTrim$(CStr(rstX.Fields("LISDT").Value))
    If Len(strX) = 7 Then
      .LIDTFORMAT = Mid(strX, 4, 2) & "/" & Mid(strX, 6, 2) & "/" & Mid(strX, 2, 2)
    End If
    strTracer = "AA404"

'    .ORQTY = decZero
'    strX = RTrim$(CStr(rstX.Fields("QTYOR").Value))
'    If IsNumeric(strX) Then
'      .ORQTY = CDec(strX)
'    End If
'    .actDt = ""
  End With
  lngY = (MOsHiNdx + 1) Mod 1000
  If lngY = 0 Then
    Application.StatusBar = strStatus & CStr(MOsHiNdx + 1)
  End If
  rstX.MoveNext
  DoEvents
Wend
rstX.Close
Set rstX = Nothing
conX.Close
Set conX = Nothing
Application.StatusBar = ""

subexit:
  Exit Function

errhandler:
  lngErr = Err.Number
  strErr = Err.Description
  blnAbort = True
  If rstX.State = adStateOpen Then
    rstX.Close
  End If
  If conX.State = adStateOpen Then
    conX.Close
  End If
  Application.Cursor = xlDefault  'no more hourglass
  strMsgTitle = "Program Error"
  strMsgPrompt = "The following error occurred:" & Chr(10) & Chr(10) & _
                 "Error No. " & CStr(lngErr) & Chr(10) & _
                 "Error Description: " & strErr & Chr(10) & _
                 "Function: z1_GetShipmentData" & Chr(10) & _
                 "Tracer: " & strTracer & Chr(10) & _
                 "SQL: " & strSQL
  MsgBox strMsgPrompt, , strMsgTitle
  Application.StatusBar = ""
  Resume subexit

End Function

Sub ShortParts(Optional COnum = 0, Optional BorderClr = "", Optional BorderInrClr = "")
    Dim Cl, Rw, LstRw, LstRwMORtg, NewMO As Boolean, NewComp As Boolean, NewCompRow, NewLoc As Boolean, NewPO As Boolean, NewMC As Boolean, NewRQ As Boolean
    Dim LocNextRow, PONextRow, NewCompN, PrevCO, LastCOrow, PartsToReturn, COname, DlrToReturn, Adj, DateStyle, OSPitem As Boolean, OSPrec As Integer
    Dim TitleMsg, TitleStyle, PrevStkLoc, AddTitle As Boolean, bShort As Boolean, wSheet As Worksheet
    'LstRw = Cells.Find(What:="<END>", After:=[B1], SearchOrder:=xlByColumns, SearchDirection:=xlNext, LookIn:=xlFormulas).Row
    'LstRwMORtg = Cells.Find(What:="*", After:=[U1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookIn:=xlFormulas).Row
    'If LstRwMORtg > LstRw Then: LstRw = LstRwMORtg
    'If LstRw > 3 Then: Rows("4:" & LstRw).Delete
    Set wSheet = Sheet9
    LstRw = wSheet.Range("C3").End(xlDown).Row
    If LstRw < 10000 Then wSheet.Rows("3:" & LstRw).Delete
    Call Load_MOs(COnum)
    'Worksheets("MOs").CommandButton1.Caption = "Filling..."
    DoEvents
    Cl = 1
    Rw = 3
    Adj = -1 'used to compensate for inserting or removing columns after "Part Description"
    PrevCO = ""
    PrevStkLoc = "XXXXXXXXXX"
    LastCOrow = 4
    PartsToReturn = 0
    DlrToReturn = 0
    AddTitle = False
    DateStyle = "yyyy mmm dd"
    
    For N = 1 To MOsHiNdx
        With MOs(N)
            If CStr(.JOBNO) <> CStr(PrevCO) Then 'Add CO Title Row
                PrevCO = .JOBNO
                PartsToReturn = 0
                DlrToReturn = 0
                If IsNumeric(PrevCO) Then
                    TitleMsg = "CO" & PrevCO & " - " & .REFNO
                Else
                    TitleMsg = PrevCO & " - " & .REFNO
                End If
                AddTitle = True
            End If
            If AddTitle = True Then
                wSheet.Cells(Rw, 1) = TitleMsg
                wSheet.Range("A" & Rw & ":I" & Rw).Interior.Color = 0
                wSheet.Range("A" & Rw & ":I" & Rw).Font.ColorIndex = 2
                AddTitle = False
            End If
            NewMO = False
            If .ORDNO <> MOs(N - 1).ORDNO Then NewMO = True 'new MO Row
            If NewMO Then  'Populate MO Header Row and component column names
                If OSPitem = True Then 'At least one PO Item didn't match the CITEM in the last MO, so populate that data
                    wSheet.Cells(Rw, 3) = MOs(OSPrec).POITNBR
                    'wsheet.Cells(Rw, 4) = .CDESC 'think about pulling in the PO Item Description extension for the Ext Desc
                    wSheet.Cells(Rw, 11) = MOs(OSPrec).PONUM
                    wSheet.Cells(Rw, 12) = MOs(OSPrec).POACTQTY
                    wSheet.Cells(Rw, 13) = MOs(OSPrec).PODUEDATE
                    wSheet.Cells(Rw, 13).NumberFormat = DateStyle
                    wSheet.Cells(Rw, 14) = MOs(OSPrec).POMNTDATE
                    wSheet.Cells(Rw, 14).NumberFormat = DateStyle
                    wSheet.Cells(Rw, 15) = MOs(OSPrec).poBUYNO
                    wSheet.Cells(Rw, 16) = MOs(OSPrec).POSTATUS
                    wSheet.Cells(Rw, 17) = MOs(OSPrec).vnVNAMA
                    'Right now this will only pull the first OSP Item. I'll need to loop for other POs to find more.
                    OSPitem = False
                    Rw = Rw + 1
                End If
                If wSheet.Cells(Rw - 1, 2).Value > 0 And wSheet.Cells(Rw - 1, 1).Value = 0 Then 'no parts on the last MO --> delete last MO
                    'Debug.Print "deleted: " & Rw - 1
                    wSheet.Rows(Rw - 1).Delete
                    Rw = Rw - 1
                End If
                wSheet.Cells(Rw, 2) = .ORDNO
                wSheet.Cells(Rw, 3) = .FITEM
                wSheet.Cells(Rw, 4) = .FDESC
                wSheet.Cells(Rw, 8) = .REFNO
                wSheet.Cells(Rw, 9) = .JOBNO
                wSheet.Range("A" & Rw & ":I" & Rw).Font.Bold = True
                wSheet.Range("H" & Rw & ":I" & Rw).Font.Italic = True
                Rw = Rw + 1
            End If
            
            If Not IsNumeric(.LQNTY) Then .LQNTY = 0
            If Not IsNumeric(.ISQTY) Then .ISQTY = 0
            If .CITEM = MOs(N - 1).CITEM Then 'repeat item -> new PO#
                If IsItShort(MOs(N)) And .PONUM <> "" Then
                    wSheet.Cells(Rw - 1, 8).Value = wSheet.Cells(Rw - 1, 8).Value & "/" & .PONUM
                End If
            Else 'not a repeat item
                bShort = IsItShort(MOs(N))
                If bShort Then
                    If NewMO = True Or .CITEM <> MOs(N - 1).CITEM Or .USRSQ <> MOs(N - 1).USRSQ Then 'New MO or New Component
                        'Populate Component Item Number, Description, Req Qty, Issued Qty, Issue Date, Total allocated, LT Code, Planner
                        wSheet.Cells(Rw, 3) = .CITEM
                        wSheet.Cells(Rw, 4) = .CDESC
                        'wsheet.Cells(Rw, 5) = .ISQTY - .QTREQ
                        wSheet.Cells(Rw, 5) = .QTREQ
                        wSheet.Cells(Rw, 6) = .ISQTY
                        wSheet.Cells(Rw, 7) = .LQNTY
                        NewComp = True
                        NewCompRow = Rw
                        NewCompN = N
                        NewLoc = True
                        LocNextRow = Rw
                        NewPO = True
                        PONextRow = Rw
                        NewMC = True
                        NewRQ = True
                        For r = Rw To NewCompRow Step -1
                            If .LLOCN = Cells(r, 9 + Adj) Then: NewLoc = False 'Check if this location has already been displayed for this component
                            If wSheet.Cells(r, 9 + Adj) = "" Then: LocNextRow = r 'Find the next blank cell in the location column to use if needed
                            If .PONUM = Cells(r, 12 + Adj) Then: NewPO = False 'Check if this PO has been displayed for this component
                            If wSheet.Cells(r, 12 + Adj) = "" Then: PONextRow = r 'Find the next blank cell in the MO/PO/Req column to use if needed
                            If .MCORDNO = Cells(r, 12 + Adj) Then: NewMC = False 'Check if this MO has been displayed for this component
                            If .rqREQNO = Cells(r, 12 + Adj) Then: NewRQ = False 'Check if this Req has been displayed for this component
                        Next r
                        If NewLoc = True Then 'Populate the location name and qty using row identified above
                            wSheet.Cells(LocNextRow, 9 + Adj) = .LLOCN
                            wSheet.Cells(LocNextRow, 10 + Adj) = .LQNTY
                        End If
                        If NewPO = True And .poPMPSTTS < "50" Then
                            If .POITNBR = .CITEMnoWH Then
                                'only display PO info if new PO or new component or new MO using row identified above
                                wSheet.Cells(Rw, 8) = .PONUM
                                wSheet.Cells(Rw, 9) = .PODUEDATE
                            Else
                                If OSPitem = False Then: OSPrec = N
                                OSPitem = True
                          End If
                        End If
                        If NewRQ = True Then
                            'If wSheet.Cells(PONextRow, 11) <> "" Then: PONextRow = PONextRow + 1 'If a PO just took the empty row, increment up 1
                            wSheet.Cells(Rw, 8) = .rqREQNO
                            wSheet.Cells(Rw, 9) = .rqDUEDT
                        End If
                        If NewMC = True Then
                            'If wsheet.Cells(PONextRow, 11) <> "" Then: PONextRow = PONextRow + 1 'If a PO or Req just took the empty row, increment up 1
                            wSheet.Cells(Rw, 8) = .MCORDNO
                            wSheet.Cells(Rw, 9) = .MCDUEDATE
                        End If
                        wSheet.Cells(Rw, 9).NumberFormat = DateStyle
                    Else 'not a new component
                        NewComp = False
                    End If
                End If
            End If
    '        Do Until wSheet.Cells(Rw, 3) = "" And wSheet.Cells(Rw, 9 + Adj) = "" And wSheet.Cells(Rw, 12 + Adj) = "" 'Find the next row that doesn't have a component, location, or MO/PO/Req
    '            Rw = Rw + 1
    '        Loop
            Do While wSheet.Range("C" & Rw).Value > 0
                Rw = Rw + 1
            Loop
        End With
    Next N
    wSheet.Columns("A:I").AutoFit
    If wSheet.Cells(Rw - 1, 2).Value > 0 Then wSheet.Rows(Rw - 1).Delete
    
    If OSPitem = True Then 'At least one PO Item didn't match the CITEM in the last MO, so populate that data
        wSheet.Cells(Rw, 3) = MOs(OSPrec).POITNBR
        'wsheet.Cells(Rw, 4) = .CDESC 'think about pulling in the PO Item Description extension for the Ext Desc
        wSheet.Cells(Rw, 12 + Adj) = MOs(OSPrec).PONUM
        wSheet.Cells(Rw, 13 + Adj) = MOs(OSPrec).POACTQTY
        wSheet.Cells(Rw, 14 + Adj) = MOs(OSPrec).PODUEDATE
        wSheet.Cells(Rw, 14 + Adj).NumberFormat = DateStyle
        wSheet.Cells(Rw, 15 + Adj) = MOs(OSPrec).POMNTDATE
        wSheet.Cells(Rw, 15 + Adj).NumberFormat = DateStyle
        wSheet.Cells(Rw, 16 + Adj) = MOs(OSPrec).poBUYNO
        wSheet.Cells(Rw, 17 + Adj) = MOs(OSPrec).POSTATUS
        wSheet.Cells(Rw, 18 + Adj) = MOs(OSPrec).vnVNAMA
        'Right now this will only pull the first OSP Item. I'll need to loop for other POs to find more.
        OSPitem = False
        Rw = Rw + 1
    End If

End Sub

Function IsItShort(zQItem As zQryShort) As Boolean
    IsItShort = True
    With zQItem
        If .ISQTY >= .QTREQ Then
            IsItShort = False
        ElseIf .LQNTY >= .QTREQ - .ISQTY Then
            IsItShort = False
        ElseIf .MCORDNO <> "" Then
            IsItShort = False
        End If
    End With
End Function

