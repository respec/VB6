VERSION 5.00
Begin VB.Form frmAddMet 
   Caption         =   "WinHSPF - Edit Met Segment"
   ClientHeight    =   3408
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6492
   HelpContextID   =   36
   Icon            =   "frmAddMet.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3408
   ScaleWidth      =   6492
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboName 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   5175
   End
   Begin ATCoCtl.ATCoGrid agdMet 
      Height          =   2295
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   6015
      _ExtentX        =   10605
      _ExtentY        =   4043
      SelectionToggle =   0   'False
      AllowBigSelection=   -1  'True
      AllowEditHeader =   0   'False
      AllowLoad       =   0   'False
      AllowSorting    =   0   'False
      Rows            =   2
      Cols            =   6
      ColWidthMinimum =   300
      gridFontBold    =   0   'False
      gridFontItalic  =   0   'False
      gridFontName    =   "MS Sans Serif"
      gridFontSize    =   8
      gridFontUnderline=   0   'False
      gridFontWeight  =   400
      gridFontWidth   =   0
      Header          =   ""
      FixedRows       =   1
      FixedCols       =   0
      ScrollBars      =   3
      SelectionMode   =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorBkg    =   -2147483637
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      InsideLimitsBackground=   -2147483643
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      ComboCheckValidValues=   0   'False
   End
   Begin VB.CommandButton cmdMet 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdMet 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label lblName 
      Caption         =   "<none>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmAddMet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SegName$
Dim curtext$
Dim lMetSeg As HspfMetSeg
Dim AddEdit&, aMetDetails$()

Private Sub agdMet_CommitChange(ChangeFromRow As Long, ChangeToRow As Long, ChangeFromCol As Long, ChangeToCol As Long)
  DoLimits agdMet
  If curtext <> agdMet.TextMatrix(agdMet.row, agdMet.col) Then
    If agdMet.row = 1 And (agdMet.col = 1 Or agdMet.col = 2 Or agdMet.col = 3) And _
       AddEdit = 1 Then
      'cant change these, set it back
      MsgBox "The Precip Data Set defines the met segment and thus cannot be changed." & vbCrLf & _
             "Use 'Add' to add a new met segment.", vbOKOnly, "Edit Met Segment Problem"
      agdMet.TextMatrix(agdMet.row, agdMet.col) = curtext
    Else
      ClearRestofRow
    End If
  End If
  If AddEdit = 0 Then
    'adding, check for precip data set
    If agdMet.row = 1 And (agdMet.col = 1 Or agdMet.col = 2 Or agdMet.col = 3) Then
      If Len(Trim(agdMet.TextMatrix(1, 1))) > 0 And _
         Len(Trim(agdMet.TextMatrix(1, 2))) > 0 And _
         Len(Trim(agdMet.TextMatrix(1, 3))) > 0 Then
        SegName = myUci.GetWDMAttr(agdMet.TextMatrix(1, 1), agdMet.TextMatrix(1, 3), "LOC")
        lblName.Caption = SegName
      End If
    End If
  End If
End Sub

Private Sub ClearRestofRow()
  With agdMet
    If .col = 1 Then  'wdm ids, clear tstypes and dsns
      .TextMatrix(.row, 2) = ""
      .TextMatrix(.row, 3) = ""
    ElseIf .col = 2 Then  'tstypes, clear dsns
      .TextMatrix(.row, 3) = ""
    End If
  End With
End Sub

Private Sub agdMet_RowColChange()
  curtext = agdMet.TextMatrix(agdMet.row, agdMet.col)
  DoLimits agdMet
End Sub

Private Sub agdMet_TextChange(ChangeFromRow As Long, ChangeToRow As Long, ChangeFromCol As Long, ChangeToCol As Long)
  DoLimits agdMet
End Sub

Private Sub cboName_Click()
    Dim delim$, quote$, lMetDetails$, basedsn&, SDate&(6), EDate&(6)
    Dim metwdmid$, i&, Id&, tstypecol As CollString
    Dim tsfile As ATCclsTserFile, j&
    
    'fill in grid from the met data details
    delim = ","
    quote = """"
    lMetDetails = aMetDetails(cboName.ListIndex + 1)
    basedsn = StrSplit(lMetDetails, delim, quote)
    For i = 0 To 5
      SDate(i) = StrSplit(lMetDetails, delim, quote)
    Next i
    For i = 0 To 5
      EDate(i) = StrSplit(lMetDetails, delim, quote)
    Next i
    metwdmid = lMetDetails
    With agdMet
      For i = 1 To 8
        .TextMatrix(i, 1) = metwdmid
        If AddEdit <> 2 Then
          Select Case i
            Case 1:
              .TextMatrix(i, 3) = basedsn
              .TextMatrix(i, 2) = "PREC"
              .TextMatrix(i, 4) = 1
              .TextMatrix(i, 5) = 1
            Case 2:
              .TextMatrix(i, 3) = basedsn + 2
              .TextMatrix(i, 2) = "ATEM"
              .TextMatrix(i, 4) = 1
              .TextMatrix(i, 5) = 1
            Case 3:
              .TextMatrix(i, 3) = basedsn + 6
              .TextMatrix(i, 2) = "DEWP"
              .TextMatrix(i, 4) = 1
              .TextMatrix(i, 5) = 1
            Case 4:
              .TextMatrix(i, 3) = basedsn + 3
              .TextMatrix(i, 2) = "WIND"
              .TextMatrix(i, 4) = 1
              .TextMatrix(i, 5) = 1
            Case 5:
              .TextMatrix(i, 3) = basedsn + 4
              .TextMatrix(i, 2) = "SOLR"
              .TextMatrix(i, 4) = 1
              .TextMatrix(i, 5) = 1
            Case 6:
              .TextMatrix(i, 3) = basedsn + 7
              .TextMatrix(i, 2) = "CLOU"
              .TextMatrix(i, 4) = 0
              .TextMatrix(i, 5) = 1
            Case 7:
              .TextMatrix(i, 3) = basedsn + 5
              .TextMatrix(i, 2) = "PEVT"
              .TextMatrix(i, 4) = 1
              .TextMatrix(i, 5) = 0
            Case 8:
              .TextMatrix(i, 3) = basedsn + 1
              .TextMatrix(i, 2) = "EVAP"
              .TextMatrix(i, 4) = 0
              .TextMatrix(i, 5) = 1
          End Select
        Else
          'on create
          Id = CInt(Mid(metwdmid, 4, 1))
          If IsNumeric(Id) Then
            j = CInt(Id)
            Set tsfile = myUci.GetWDMObj(j)
            If Not tsfile Is Nothing Then
              Set tstypecol = uniqueAttributeValues("TSTYPE", tsfile.DataCollection)
            End If
          End If
          Select Case i
            Case 1:
              .TextMatrix(i, 3) = basedsn
              Id = CInt(Mid(metwdmid, 4, 1))
              .TextMatrix(i, 2) = myUci.GetDataSetFromDsn(Id, basedsn).attrib("TSTYPE")
              .TextMatrix(i, 4) = 1
              .TextMatrix(i, 5) = 1
            Case 2:
              .TextMatrix(i, 3) = 0
              .TextMatrix(i, 2) = ""
              If tstypecol("ATEM") = "ATEM" Then
                .TextMatrix(i, 2) = "ATEM"
              ElseIf tstypecol("ATMP") = "ATMP" Then
                .TextMatrix(i, 2) = "ATMP"
              End If
              .TextMatrix(i, 4) = 1
              .TextMatrix(i, 5) = 1
            Case 3:
              .TextMatrix(i, 3) = 0
              .TextMatrix(i, 2) = ""
              If tstypecol("DEWP") = "DEWP" Then
                .TextMatrix(i, 2) = "DEWP"
              End If
              .TextMatrix(i, 4) = 1
              .TextMatrix(i, 5) = 1
            Case 4:
              .TextMatrix(i, 3) = 0
              .TextMatrix(i, 2) = ""
              If tstypecol("WIND") = "WIND" Then
                .TextMatrix(i, 2) = "WIND"
              ElseIf tstypecol("WNDH") = "WNDH" Then
                .TextMatrix(i, 2) = "WNDH"
              End If
              .TextMatrix(i, 4) = 1
              .TextMatrix(i, 5) = 1
            Case 5:
              .TextMatrix(i, 3) = 0
              .TextMatrix(i, 2) = ""
              If tstypecol("SOLR") = "SOLR" Then
                .TextMatrix(i, 2) = "SOLR"
              End If
              .TextMatrix(i, 4) = 1
              .TextMatrix(i, 5) = 1
            Case 6:
              .TextMatrix(i, 3) = 0
              .TextMatrix(i, 2) = ""
              If tstypecol("CLOU") = "CLOU" Then
                .TextMatrix(i, 2) = "CLOU"
              End If
              .TextMatrix(i, 4) = 0
              .TextMatrix(i, 5) = 1
            Case 7:
              .TextMatrix(i, 3) = 0
              .TextMatrix(i, 2) = ""
              If tstypecol("PEVT") = "PEVT" Then
                .TextMatrix(i, 2) = "PEVT"
              ElseIf tstypecol("EVAP") = "EVAP" Then
                .TextMatrix(i, 2) = "EVAP"
              End If
              .TextMatrix(i, 4) = 1
              .TextMatrix(i, 5) = 0
            Case 8:
              .TextMatrix(i, 3) = 0
              .TextMatrix(i, 2) = ""
              If tstypecol("EVAP") = "EVAP" Then
                .TextMatrix(i, 2) = "EVAP"
              End If
              .TextMatrix(i, 4) = 0
              .TextMatrix(i, 5) = 1
          End Select
          'default dsns if possible
          If Not tsfile Is Nothing Then
            If Len(.TextMatrix(i, 2)) > 0 And .TextMatrix(i, 3) = 0 Then
              For j = 1 To tsfile.DataCount
                If tsfile.Data(j).attrib("TSTYPE") = .TextMatrix(i, 2) Then
                  .TextMatrix(i, 3) = tsfile.Data(j).Header.Id
                  Exit For
                End If
              Next j
            End If
          End If
        End If
      Next i
    End With
    SegName = myUci.GetWDMAttr(agdMet.TextMatrix(1, 1), agdMet.TextMatrix(1, 3), "LOC")
End Sub

Private Sub cmdMet_Click(Index As Integer)
  Dim r&, vMetSeg As Variant, ifound As Boolean
  If Index = 0 Then 'okay
    If AddEdit = 1 Then 'editing
      If Not lMetSeg Is Nothing Then
        For r = 1 To 8
          lMetSeg.MetSegRec(r).Source.volname = agdMet.TextMatrix(r, 1)
          lMetSeg.MetSegRec(r).Source.member = agdMet.TextMatrix(r, 2)
          lMetSeg.MetSegRec(r).Source.volid = agdMet.TextMatrix(r, 3)
          lMetSeg.MetSegRec(r).MFactP = agdMet.TextMatrix(r, 4)
          lMetSeg.MetSegRec(r).MFactR = agdMet.TextMatrix(r, 5)
        Next r
        lMetSeg.ExpandMetSegName agdMet.TextMatrix(1, 1), agdMet.TextMatrix(1, 3)
      End If
    ElseIf AddEdit = 0 And Len(SegName) > 0 Then 'add new met seg
      Set lMetSeg = New HspfMetSeg
      Set lMetSeg.Uci = myUci
      For r = 1 To 8
        If Len(agdMet.TextMatrix(r, 1)) > 0 And _
           Len(agdMet.TextMatrix(r, 2)) > 0 And _
           Len(agdMet.TextMatrix(r, 3)) > 0 Then
          lMetSeg.MetSegRec(r).Source.volname = agdMet.TextMatrix(r, 1)
          lMetSeg.MetSegRec(r).Source.member = agdMet.TextMatrix(r, 2)
          lMetSeg.MetSegRec(r).Source.volid = agdMet.TextMatrix(r, 3)
          If Len(Trim(agdMet.TextMatrix(r, 4))) = 0 Then
            lMetSeg.MetSegRec(r).MFactP = 0#
          Else
            lMetSeg.MetSegRec(r).MFactP = agdMet.TextMatrix(r, 4)
          End If
          If Len(Trim(agdMet.TextMatrix(r, 5))) = 0 Then
            lMetSeg.MetSegRec(r).MFactR = 0#
          Else
            lMetSeg.MetSegRec(r).MFactR = agdMet.TextMatrix(r, 5)
          End If
          lMetSeg.MetSegRec(r).Sgapstrg = ""
          lMetSeg.MetSegRec(r).Ssystem = "ENGL"
          lMetSeg.MetSegRec(r).Tran = "SAME"
          lMetSeg.MetSegRec(r).typ = r
        End If
      Next r
      
      lMetSeg.ExpandMetSegName agdMet.TextMatrix(1, 1), agdMet.TextMatrix(1, 3)
      ifound = False
      For Each vMetSeg In myUci.MetSegs
        If vMetSeg.Compare(lMetSeg, "PERLND") And vMetSeg.Compare(lMetSeg, "RCHRES") Then
          'already exists
          ifound = True
        End If
      Next vMetSeg
      If Not ifound Then
        lMetSeg.Id = myUci.MetSegs.Count + 1
        myUci.MetSegs.Add lMetSeg
      End If
    ElseIf AddEdit = 2 Then 'from create
      Set lMetSeg = New HspfMetSeg
      Set lMetSeg.Uci = myUci
      For r = 1 To 8
        If Len(agdMet.TextMatrix(r, 1)) > 0 And _
           Len(agdMet.TextMatrix(r, 2)) > 0 And _
           Len(agdMet.TextMatrix(r, 3)) > 0 Then
          lMetSeg.MetSegRec(r).Source.volname = agdMet.TextMatrix(r, 1)
          lMetSeg.MetSegRec(r).Source.member = agdMet.TextMatrix(r, 2)
          lMetSeg.MetSegRec(r).Source.volid = agdMet.TextMatrix(r, 3)
          If Len(Trim(agdMet.TextMatrix(r, 4))) = 0 Then
            lMetSeg.MetSegRec(r).MFactP = 0#
          Else
            lMetSeg.MetSegRec(r).MFactP = agdMet.TextMatrix(r, 4)
          End If
          If Len(Trim(agdMet.TextMatrix(r, 5))) = 0 Then
            lMetSeg.MetSegRec(r).MFactR = 0#
          Else
            lMetSeg.MetSegRec(r).MFactR = agdMet.TextMatrix(r, 5)
          End If
          lMetSeg.MetSegRec(r).Sgapstrg = ""
          lMetSeg.MetSegRec(r).Ssystem = "ENGL"
          lMetSeg.MetSegRec(r).Tran = "SAME"
          lMetSeg.MetSegRec(r).typ = r
        End If
      Next r
      If Len(agdMet.TextMatrix(1, 1)) > 0 And Len(agdMet.TextMatrix(1, 3)) > 0 Then
        lMetSeg.ExpandMetSegName agdMet.TextMatrix(1, 1), agdMet.TextMatrix(1, 3)
      End If
      lMetSeg.Id = myUci.MetSegs.Count + 1
      myUci.MetSegs.Add lMetSeg
      If lMetSeg.MetSegRec(1).Source.volid = 0 Or _
         lMetSeg.MetSegRec(7).Source.volid = 0 Then
        'warn user that precip and evap are required
        myMsgBox.Show "Precipitation and Evapotraspiration data are required to run HSPF Hydrology." & _
          vbCrLf & vbCrLf & "At least one of these has not been specified.", "WinHSPF Create Warning", "+OK"
      End If
    End If
  Else 'cancel
  End If
  Unload Me
End Sub

Private Sub Form_Load()
  Dim i&, r&, nwdm&, aunits&(), j&, WDMId$
  Dim numMetSeg&, arrayMetSegs$(), cntMetSeg&
  Dim lMetDetails$(), lMetDescs$()
  
  If AddEdit = 1 Then
    Me.Caption = "WinHSPF - Edit Met Segment"
    cboName.Visible = False
  Else
    If AddEdit = 2 Then
      Me.Caption = "WinHSPF - Initial Met Segment"
    Else
      Me.Caption = "WinHSPF - Add Met Segment"
    End If
    lblName.Visible = False
    'add candidate met seg names to list
    myUci.GetWDMUnits nwdm, aunits
    cntMetSeg = 0
    cboName.Clear
    For i = 1 To nwdm
      myUci.GetMetSegNames aunits(i), numMetSeg, arrayMetSegs, lMetDetails, lMetDescs
      myUci.GetWDMIDFromUnit aunits(i), WDMId
      If numMetSeg > 0 Then
        cntMetSeg = cntMetSeg + numMetSeg
        ReDim Preserve aMetDetails(cntMetSeg)
        For j = 1 To numMetSeg
          cboName.AddItem arrayMetSegs(j - 1) & ":" & lMetDescs(j - 1)
          aMetDetails(cntMetSeg - numMetSeg + j) = lMetDetails(j - 1) & ", " & WDMId
        Next j
      End If
    Next i
    If numMetSeg > 0 Then
      cboName.ListIndex = 0
    End If
  End If
  
  lblName.Caption = SegName
  With agdMet
    .TextMatrix(0, 0) = "Constituent"
    .TextMatrix(0, 1) = "WDM ID"
    .TextMatrix(0, 2) = "TSTYPE"
    .TextMatrix(0, 3) = "DSN"
    .TextMatrix(0, 4) = "Mfact P/I"
    .TextMatrix(0, 5) = "Mfact R"
    .TextMatrix(1, 0) = "Precip"
    .TextMatrix(2, 0) = "Air Temp"
    .TextMatrix(3, 0) = "Dew Point"
    .TextMatrix(4, 0) = "Wind"
    .TextMatrix(5, 0) = "Solar Rad"
    .TextMatrix(6, 0) = "Cloud"
    .TextMatrix(7, 0) = "Evapotrans"
    .TextMatrix(8, 0) = "Pot Evap"
    .ColEditable(1) = True
    .ColEditable(2) = True
    .ColEditable(3) = True
    .ColEditable(4) = True
    .ColEditable(5) = True
    .ColType(3) = ATCoInt
    .ColType(4) = ATCoSng
    .ColType(5) = ATCoSng
    If AddEdit = 1 Then
      Set lMetSeg = Nothing
      For i = 1 To myUci.MetSegs.Count  'find which met seg this is
        If myUci.MetSegs(i).Name = SegName Then
          Set lMetSeg = myUci.MetSegs(i)
        End If
      Next i
      If Not lMetSeg Is Nothing Then
        For r = 1 To 8
          .TextMatrix(r, 1) = lMetSeg.MetSegRec(r).Source.volname
          .TextMatrix(r, 2) = lMetSeg.MetSegRec(r).Source.member
          .TextMatrix(r, 3) = lMetSeg.MetSegRec(r).Source.volid
          .TextMatrix(r, 4) = lMetSeg.MetSegRec(r).MFactP
          .TextMatrix(r, 5) = lMetSeg.MetSegRec(r).MFactR
        Next r
      End If
    End If
  End With
End Sub

Public Sub Init(n$, iopt&)
  SegName = n
  AddEdit = iopt
End Sub

Private Sub DoLimits(g As Object)
  Dim i&, s$
  Dim tstypecol As CollString
  Dim tsfile As ATCclsTserFile
  Dim vname As Variant
    
  g.ClearValues
  If g.col = 1 Then  'valid wdm ids
    If myUci.wdmcount = 1 Then
      g.addvalue "WDM1"
    Else
      For i = 2 To myUci.wdmcount
        g.addvalue "WDM" & CStr(i)
      Next i
    End If
    'For i = 1 To myUci.filesblock.Count
    '  If Mid(myUci.filesblock.Value(i).typ, 1, 3) = "WDM" Then
    '    g.AddValue myUci.filesblock.Value(i).typ
    '  End If
    'Next i
  ElseIf g.col = 2 Then  'valid tstypes
    s = Mid(g.TextMatrix(g.row, 1), 4, 1)
    If IsNumeric(s) Then
      i = CInt(s)
      Set tsfile = myUci.GetWDMObj(i)
      If Not tsfile Is Nothing Then
        Set tstypecol = uniqueAttributeValues("TSTYPE", tsfile.DataCollection)
        For Each vname In tstypecol
          g.addvalue (vname)
        Next
      End If
    End If
  ElseIf g.col = 3 Then 'valid dsns
    s = Mid(g.TextMatrix(g.row, 1), 4, 1)
    If IsNumeric(s) Then
      i = CInt(s)
      Set tsfile = myUci.GetWDMObj(i)
      If Not tsfile Is Nothing Then
        For i = 1 To tsfile.DataCount
          If tsfile.Data(i).attrib("TSTYPE") = g.TextMatrix(g.row, 2) Then
            g.addvalue tsfile.Data(i).Header.Id
          End If
        Next i
      End If
    End If
  End If
End Sub

Private Sub Form_Resize()
  If Not (Me.WindowState = vbMinimized) Then
    If width < 1500 Then width = 1500
    If height < 2000 Then height = 2000
    agdMet.width = width - 500
    cmdMet(0).Left = (width / 2) - cmdMet(0).width - 200
    cmdMet(1).Left = (width / 2) + 200
    cmdMet(0).Top = height - cmdMet(0).height - 600
    cmdMet(1).Top = height - cmdMet(1).height - 600
    agdMet.height = height - 1600
  End If
End Sub
