VERSION 5.00
Object = "{8B6FC3CA-0323-4FCA-8A85-2668B1192E25}#1.0#0"; "ATCoCtl.ocx"
Begin VB.Form frmStaData 
   Caption         =   "Station Data"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   13875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraFilter 
      Caption         =   "Filter by Statistic Type"
      Height          =   800
      HelpContextID   =   21
      Left            =   40
      TabIndex        =   11
      Top             =   3620
      Width           =   2592
      Begin VB.ComboBox cboFilter 
         Height          =   288
         HelpContextID   =   21
         Left            =   70
         TabIndex        =   1
         Top             =   460
         Width           =   2460
      End
      Begin VB.Label lblFIlter 
         Caption         =   "Statistic Type:"
         Height          =   252
         Left            =   80
         TabIndex        =   12
         Top             =   220
         Width           =   1212
      End
   End
   Begin VB.Frame fraGridButtons 
      Caption         =   "Grid Commands"
      Height          =   800
      HelpContextID   =   23
      Left            =   7850
      TabIndex        =   9
      Top             =   3620
      Width           =   1812
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         HelpContextID   =   23
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   732
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         HelpContextID   =   23
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame fraStatButtons 
      Caption         =   "Statistic Commands"
      Height          =   800
      HelpContextID   =   21
      Left            =   2700
      TabIndex        =   8
      Top             =   3620
      Width           =   5052
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         HelpContextID   =   22
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   732
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         HelpContextID   =   24
         Left            =   4200
         TabIndex        =   4
         Top             =   240
         Width           =   732
      End
      Begin VB.Label lblStatSel 
         BackColor       =   &H00E0E0E0&
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   460
         Width           =   3132
      End
      Begin VB.Label lblStatSel 
         Caption         =   "Active Statistic: "
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   220
         Width           =   1452
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   13080
      TabIndex        =   7
      Top             =   3840
      Width           =   732
   End
   Begin ATCoCtl.ATCoGrid grdStaData 
      Height          =   3615
      HelpContextID   =   23
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   6376
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
      BackColorBkg    =   -2147483632
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      InsideLimitsBackground=   -2147483643
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      ComboCheckValidValues=   0   'False
   End
End
Attribute VB_Name = "frmStaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim newStat() As Boolean
Dim Changes() As String
Dim SelStats() As ssStatistic
Dim station As ssStation
Dim StatType As String
Dim SaidYes As Boolean

Private Sub cboFilter_Click()
  Dim i&, ListItem&
  
  ListItem = cboFilter.ListIndex
  Set station.SelStats = Nothing
  Set station.Statistics = Nothing
  For i = 1 To station.Statistics.Count
    With station.Statistics(i)
        If cboFilter.List(ListItem) = "All" Then
'          station.SelStats.Add _
'              station.Statistics(i), CStr(.code) & "_" & CStr(.value)
          station.SelStats.Add _
              station.Statistics(i), CStr(.code) & "_" & CStr(.value) & "_" & .SourceID
        Else
          If .StatTypeID = ListItem + 1 Then
'            station.SelStats.Add _
'                station.Statistics(i), CStr(.code) & "_" & CStr(.value)
            station.SelStats.Add _
                station.Statistics(i), CStr(.code) & "_" & CStr(.value) & "_" & .SourceID
          End If
        End If
    End With
  Next i
  StatType = cboFilter.List(ListItem)
  SetGrid
End Sub

Private Sub cmdAdd_Click()
  Dim i&
  With grdStaData
    For i = 0 To .Cols - 1
      If i = 5 Or i = 16 Then
        .ColEditable(i) = False
      Else
        .ColEditable(i) = True
      End If
    Next i
    .Rows = .Rows + 1
    .row = .Rows
    .col = 0
    ReDim Preserve newStat(1 To .Rows)
    newStat(.Rows) = True
    ReDim Preserve SelStats(1 To .Rows)
    Set SelStats(.Rows) = New ssStatistic
    Set SelStats(.Rows).DB = SSDB
    Set SelStats(.Rows).station = station
    .row = .Rows
    .col = 3
  End With
End Sub

Private Sub cmdCancel_Click()
  SetGrid
End Sub

Private Sub cmdDelete_Click()
  Dim row&, col&, response&
  Dim myStatistic As ssStatistic
  
  If grdStaData.Rows = 0 Then Exit Sub
  If newStat(grdStaData.row) Then
    response = 2
  Else
    response = myMsgBox.Show("Are you certain you want to delete the " & _
        Trim(grdStaData.TextMatrix(grdStaData.row, 1)) & vbCrLf & _
        "statistic from the list of available statistics?", _
        "User Action Verification", "+&Cancel", "-&Yes")
  End If
  If response = 2 Then
    If Not newStat(grdStaData.row) Then
      Set myStatistic = station.Statistics(grdStaData.row)
      myStatistic.Delete
    End If
    With grdStaData
      For row = .row To .Rows - 1
        Set SelStats(row) = SelStats(row + 1)
        newStat(row) = newStat(row + 1)
        For col = 0 To .Cols - 1
          .TextMatrix(row, col) = .TextMatrix(row + 1, col)
        Next col
      Next row
      .Rows = .Rows - 1
      If UBound(SelStats) > 1 Then
        ReDim Preserve SelStats(1 To UBound(SelStats) - 1)
        ReDim Preserve newStat(1 To UBound(newStat) - 1)
      Else
        ReDim SelStats(1)
        ReDim newStat(1)
      End If
    End With
  End If
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim row&, col&, response&, fldID&
  Dim madeChanges As Boolean
  Dim lOldVal As String
  Dim lNewVal As String
  Dim lRecIDStr As String
  Dim lRowNeedsUpdate&
  
  'Perform QA check on values selected/entered in grid
  Dim lmsg As String
  lmsg = ""
  If Not QACheck(lmsg) Then
    If Len(lmsg) > 0 Then MsgBox (lmsg)
    GoTo NoChanges
  End If
  
  'Record the changes made to statistic values
  ChangesMade madeChanges
  If Not madeChanges Then GoTo NoChanges
  
  If SaidYes Then
    response = 1
  Else
    response = myMsgBox.Show("Are you certain you want to write all " & _
        "new values in the Station Data grid to the database?", _
        "User Action Verification", "+&Yes", "-&Cancel")
  End If
  If response = 1 Then  'Make changes to database
    frmUserInfo.Show vbModal, Me
    If Not UserInfoOK Then GoTo NoChanges
    Me.MousePointer = vbHourglass
    For row = 1 To grdStaData.Rows
      lRowNeedsUpdate = 0
      If newStat(row) Then
        Set SelStats(row) = New ssStatistic
        Set SelStats(row).DB = SSDB
        Set SelStats(row).station = station
        SelStats(row).code = Changes(2, row, 2)
        lRowNeedsUpdate = 1
        SelStats(row).Add Changes(), row
      Else
        For col = 1 To UBound(Changes, 3)
          If Changes(1, row, col) = "" Then
          Else
            lRowNeedsUpdate = 1
            Exit For
          End If
        Next
        If lRowNeedsUpdate = 1 Then
          SelStats(row).Edit Changes(), row
        End If
      End If
      
      If lRowNeedsUpdate = 1 Then
        'Write to DetailedLog table if this row has been changed
        For col = 1 To UBound(Changes, 3)
          If col = 2 Or col = 5 Or col = 7 Or col = 8 Then
            If Changes(0, row, col) = "1" Or Changes(0, row, col) = "2" Then
              Select Case col
                Case 2: fldID = 1 '3
                Case 5: fldID = 3 '4
                Case 7: fldID = 4 '5
                Case 8: fldID = 2
              End Select
              If col = 7 Then 'record change in source ID (source can be too long for field)
                lOldVal = SelStats(row).GetSourceID(Changes(1, row, col))
                lNewVal = SelStats(row).GetSourceID(Changes(2, row, col))
              Else
                lOldVal = Changes(1, row, col)
                lNewVal = Changes(2, row, col)
              End If
              lRecIDStr = station.id & " - " & SelStats(row).code
              'SSDB.RecordChanges TransID, "STATISTIC", fldID, CStr(station.id), lOldVal, lNewVal
              SSDB.RecordChanges TransID, "STATISTIC", fldID, lRecIDStr, lOldVal, lNewVal
            End If
          End If
        Next col
      End If
    Next row
    cboFilter_Click
    lblStatSel(1).Caption = grdStaData.TextMatrix(grdStaData.row, 1)
    Me.MousePointer = vbDefault
  End If
NoChanges:
  SaidYes = False
End Sub

Private Function QACheck(Optional aMsg As String = "") As Boolean
  Dim row&, col&, i&, response&
  Dim Val$
  Dim mySource As ssSource
  Dim rowListing As String
  
  QACheck = True
  With grdStaData
    'Perform QA check on fields in grid
    For row = 1 To .Rows
      For col = 0 To .Cols - 1
        If col < 3 Then 'Check to ensure selections made for first 3 fields in grid
          If Trim(.TextMatrix(row, col)) = "" Then
            MsgBox "The " & .TextMatrix(0, col) & " field in row " & row & _
                " of the grid is blank." & vbCrLf & "All Type, Code, and " & _
                "Name fields in the grid all require a selection." & vbCrLf & _
                "Make the necessary selections then click the 'Save' button again."
            QACheck = False
            Exit Function
          End If
'        ElseIf col = 3 Then  'Check that value column is numeric
'          val = .TextMatrix(row, col)
'          If Left(val, 1) = "<" Then val = Mid(val, 2)
'          If Not IsNumeric(val) And Len(val) > 0 Then
'            MsgBox "The Value field for " & .TextMatrix(row, 1) & _
'                   " must contain a numeric entry."
'            QACheck = False
'            Exit Function
'          End If
        ElseIf col = 7 Then 'Check whether Citation exists. Add to DB if not?
          With grdStaData
            If Trim(.TextMatrix(row, col)) = "" Then
              .TextMatrix(row, col) = "none"
              'Exit Function
            End If
            For i = 1 To SSDB.Sources.Count
              If .TextMatrix(row, col) = SSDB.Sources(i).Name Then Exit For
            Next i
            If i > SSDB.Sources.Count Then  'Citation does not exist in DB
              If newStat(row) Then
                response = myMsgBox.Show("The citation you entered for " & _
                    .TextMatrix(row, 2) & " does not match any " & _
                    "currently existing in the database." & vbCrLf & vbCrLf & _
                    "Would you like to save this new citation to the database or cancel " & _
                    "this save operation and select one of the existing citations?", _
                    "User Action Verification", "+&Save", "-&Cancel")
              Else
                response = myMsgBox.Show("You have edited the citation for " & _
                    .TextMatrix(row, 1) & "." & vbCrLf & vbCrLf & _
                    "Would you like to Add this citation as a new entry to the database," & _
                    vbCrLf & "Overwrite the previously existing citation (this option " & _
                    "will change the citation for all data citing this source)," & vbCrLf & _
                    "or Cancel this save operation and re-select one of the existing citations?", _
                    "User Action Verification", "+&Add", "-&Overwrite", "-&Cancel")
              End If
              If response = 1 Then
                Set mySource = New ssSource
                Set mySource.DB = SSDB
                mySource.Add .TextMatrix(row, col), .TextMatrix(row, col + 1)
                Set SSDB.Sources = Nothing
                SaidYes = True
              ElseIf response = 2 And Not newStat(row) Then
                Set mySource = _
                    SSDB.Sources(CStr(station.Statistics(grdStaData.row).SourceID))
                mySource.Edit .TextMatrix(row, col), .TextMatrix(row, col + 1)
                For i = row + 1 To .Rows
                  'Change Citation of other data in grid with same source
                  If .TextMatrix(i, col) = station.Statistics(row).Source Then
                    .TextMatrix(i, col) = .TextMatrix(row, col)
                    'also update source url
                    .TextMatrix(i, col + 1) = .TextMatrix(row, col + 1)
                  End If
                Next i
                Set SSDB.Sources = Nothing
                SaidYes = True
              Else
                QACheck = False
                SaidYes = False
                Exit Function
              End If
            End If
          End With
        End If
      Next col
    Next row
    
    'A new check on duplicate stats
    Dim lRowFurther As Integer
    For row = 1 To .Rows
      For lRowFurther = 1 To .Rows
        If lRowFurther <> row Then
          If .TextMatrix(row, 1) = .TextMatrix(lRowFurther, 1) And .TextMatrix(row, 7) = .TextMatrix(lRowFurther, 7) Then
            'found a duplicate stat with same name and same source, Not allowed
            aMsg = aMsg & .TextMatrix(row, 1) & " has duplicate version from same source. On rows " & row & ", " & lRowFurther
            QACheck = False
          End If
        End If
      Next lRowFurther
    Next row
    
    'A new check on duplicate IsPreferred
    For row = 1 To .Rows
      For lRowFurther = 1 To .Rows
        If lRowFurther <> row Then
          If .TextMatrix(row, 1) = .TextMatrix(lRowFurther, 1) And _
            .TextMatrix(row, 4) = .TextMatrix(lRowFurther, 4) And _
            LCase(.TextMatrix(row, 4)) = "yes" Then
            'found a stat with two IsPreferred values, Not Allowed
            aMsg = aMsg & .TextMatrix(row, 1) & " has more than one IsPreferred values. On rows " & row & ", " & lRowFurther
            QACheck = False
          End If
        End If
      Next lRowFurther
    Next row
  End With
End Function

Private Sub Form_Load()
  Dim i&
  
  Set station = SSDB.state.SelStation
  Me.Caption = "Station Data - " & station.label
  'Populate statistic type combo box
  For i = 1 To SSDB.StatisticTypes.Count
    cboFilter.AddItem SSDB.StatisticTypes(i).Name
  Next i
  cboFilter.AddItem "All"
  cboFilter.ListIndex = cboFilter.ListCount - 1
  'Set SelStats collection = entire Statistics collection of station
  Set station.SelStats = Nothing
  For i = 1 To station.Statistics.Count
'    station.SelStats.Add _
'        station.Statistics(i), CStr(station.Statistics(i).code) & "_" & CStr(station.Statistics(i).value)
    With station.Statistics(i)
      station.SelStats.Add _
          station.Statistics(i), CStr(.code) & "_" & CStr(.value) & "_" & .SourceID
    End With

  Next i
  SetGrid
End Sub

Private Sub Form_Resize()
  If Me.Width > 8000 Then
    grdStaData.Width = Me.Width - 230
    cmdExit.Left = Me.Width - cmdExit.Width - 200
  End If
  If Me.height > 3500 Then
    grdStaData.height = Me.height - fraStatButtons.height - 700
    fraStatButtons.Top = grdStaData.height + 20
    fraFilter.Top = fraStatButtons.Top
    fraGridButtons.Top = fraStatButtons.Top
    cmdExit.Top = fraStatButtons.Top + 250
  End If
End Sub

Private Sub grdStaData_RowColChange()
  Dim i&
  Dim statTypeCode$
  
  'Fill in combo box entries
  With grdStaData
    If .row = 0 Then Exit Sub
    lblStatSel(1).Caption = .TextMatrix(.row, 2)
    .ClearValues
    SizeGrid
    Select Case .col
      'Fill list of Statistic Types in first column
      Case 0:
        For i = 1 To SSDB.StatisticTypes.Count
          .addValue SSDB.StatisticTypes(i).Name
        Next i
        .ComboCheckValidValues = True
      'Fill list of Stat Abbreviations in second column
      Case 1:
        If Len(Trim(.TextMatrix(.row, 0))) > 0 Then
          statTypeCode = GetStatTypeCode(.TextMatrix(.row, 0))
          For i = 1 To SSDB.StatisticTypes(statTypeCode).StatLabels.Count
            If SSDB.StatisticTypes(statTypeCode).StatLabels(i).id > 14 Then
              .addValue SSDB.StatisticTypes(statTypeCode).StatLabels(i).Name
            End If
          Next i
          .ComboCheckValidValues = True
        End If
      'Fill list of Statistic Names in third column
      Case 2:
        If Len(Trim(.TextMatrix(.row, 0))) > 0 Then
          statTypeCode = GetStatTypeCode(.TextMatrix(.row, 0))
          For i = 1 To SSDB.StatisticTypes(statTypeCode).StatLabels.Count
            .addValue SSDB.StatisticTypes(statTypeCode).StatLabels(i).code
          Next i
          .ComboCheckValidValues = True
        End If
'      'Fill list of Units in sixth column
'      Case 4: For i = 1 To SSDB.Units.Count
'                .addValue SSDB.Units(i).EnglishLabel
'              Next i
'              .ComboCheckValidValues = True
      'Fill list of Citations in seventh column
      Case 4: 'IsPreferred
        If Len(Trim(.TextMatrix(.row, 0))) > 0 Then
          .addValue "No"
          .addValue "Yes"
          .ComboCheckValidValues = True
        End If
      Case 7:
        If Len(Trim(.TextMatrix(.row, 0))) > 0 Then
          For i = 1 To SSDB.Sources.Count
            .addValue SSDB.Sources(i).Name
          Next i
          .ComboCheckValidValues = False
        End If
      Case 8:
        If Len(Trim(.TextMatrix(.row, 0))) > 0 Then
          For i = 1 To SSDB.Sources.Count
            If Len(SSDB.Sources(i).URL) > 0 Then .addValue SSDB.Sources(i).URL
          Next i
          .ComboCheckValidValues = False
        End If
    End Select
  End With
End Sub

Private Sub grdStaData_CommitChange(ChangeFromRow As Long, ChangeToRow As Long, _
                                    ChangeFromCol As Long, ChangeToCol As Long)
  Dim i&, response&
  Dim statTypeCode$, str$
  Dim CitationCode$
  Dim lmsg$, lResponse$
  Dim lOriginalIsPreferred As Integer
  
  'Adjust appropriate columns in row when a field is edited
  Select Case ChangeFromCol
    Case 0:
      'Clear the Code, Name, and Units fields when new Type selected
      For i = 1 To SSDB.StatisticTypes.Count
        With grdStaData
          If .TextMatrix(ChangeFromRow, 0) = SSDB.StatisticTypes(i).Name Then
            If .TextMatrix(ChangeFromRow, 0) <> SelStats(ChangeFromRow).StatType Then
              .TextMatrix(ChangeFromRow, 1) = ""
              .TextMatrix(ChangeFromRow, 2) = ""
              .TextMatrix(ChangeFromRow, 5) = ""
              Exit Sub
            End If
          End If
        End With
      Next i
    Case 1:
      'Make sure this Statistic does not already exist for this station
      'Make sure this Statistic has only one that is labelled as IsPreferred = Y
      With grdStaData
        For i = 1 To .Rows
          If i <> ChangeFromRow Then
            If .TextMatrix(ChangeFromRow, 1) = .TextMatrix(i, 1) Then 'Compare Stat name
              If .TextMatrix(ChangeFromRow, 7) = .TextMatrix(i, 7) Then 'Compare data source, if yes, then problem
                lmsg = "An existing version of " & .TextMatrix(ChangeFromRow, 1) & " came from a same source." & vbCrLf & _
                       "Pick/enter another source."
                MsgBox (lmsg)
                .TextMatrix(ChangeFromRow, 7) = "Needs a different source"
              End If
              If .TextMatrix(ChangeFromRow, 4) = .TextMatrix(i, 4) And .TextMatrix(i, 4) = "Yes" Then
                lmsg = "For each statistic, there can be only one preferred for a station." & vbCrLf
                lmsg = lmsg & "Please specify which is preferred?"
                
                lResponse = myMsgBox.Show(lmsg, "User Action Verification", "+&Original", "-&Current", "-&None", "-&Cancel")
                lOriginalIsPreferred = OriginalIsPreferredRow(.TextMatrix(i, 1))
                If lResponse = 1 Then
                    If lOriginalIsPreferred <> 0 Then .TextMatrix(lOriginalIsPreferred, 4) = "Yes"
                    If ChangeFromRow <> lOriginalIsPreferred Then .TextMatrix(ChangeFromRow, 4) = "No"
                    If i <> lOriginalIsPreferred Then .TextMatrix(i, 4) = "No"
                ElseIf lResponse = 2 Then
                    .TextMatrix(ChangeFromRow, 4) = "Yes"
                    .TextMatrix(i, 4) = "No"
                    If lOriginalIsPreferred <> ChangeFromRow Then .TextMatrix(lOriginalIsPreferred, 4) = "No"
                ElseIf lResponse = 3 Then
                    .TextMatrix(lOriginalIsPreferred, 4) = "No"
                    .TextMatrix(i, 4) = "No"
                    .TextMatrix(ChangeFromRow, 4) = "No"
                Else
                  '.TextMatrix(ChangeFromRow, 1) = SelStats(ChangeFromRow).Name
                  .TextMatrix(ChangeFromRow, 1) = ""
                  .TextMatrix(lOriginalIsPreferred, 4) = "Yes"
                  If ChangeFromRow <> lOriginalIsPreferred Then .TextMatrix(ChangeFromRow, 4) = "No"
                  If i <> lOriginalIsPreferred Then .TextMatrix(i, 4) = "No"
                  Exit Sub
                End If
              End If
            End If
          End If
        Next i
        'Change the Code and Unit fields to match the selected Name field
        statTypeCode = GetStatTypeCode(.TextMatrix(ChangeFromRow, 0))
        If statTypeCode <> "" Then
          For i = 1 To SSDB.StatisticTypes(statTypeCode).StatLabels.Count
            If .TextMatrix(ChangeFromRow, 1) = SSDB.StatisticTypes(statTypeCode).StatLabels(i).Name Then
              .TextMatrix(ChangeFromRow, 2) = SSDB.StatisticTypes(statTypeCode).StatLabels(i).code
              .TextMatrix(ChangeFromRow, 5) = SSDB.StatisticTypes(statTypeCode).StatLabels(i).Units
              Exit Sub
            End If
          Next i
        End If
      End With
    Case 2:
      'Make sure this Statistic does not already exist for this station
      With grdStaData
        For i = 1 To .Rows
          If i <> ChangeFromRow Then
            If .TextMatrix(ChangeFromRow, 2) = .TextMatrix(i, 2) Then
              If .TextMatrix(ChangeFromRow, 4) = .TextMatrix(i, 4) And .TextMatrix(i, 4) = "Yes" Then
                lmsg = "For each statistic, there can be only one preferred for a station." & vbCrLf
                lmsg = lmsg & "Please specify which is preferred?"
                
                lResponse = myMsgBox.Show(lmsg, "User Action Verification", "+&Original", "-&Current", "-&None", "-&Cancel")
                lOriginalIsPreferred = OriginalIsPreferredRow(.TextMatrix(i, 1))
                If lResponse = 1 Then
                    If lOriginalIsPreferred <> 0 Then .TextMatrix(lOriginalIsPreferred, 4) = "Yes"
                    If ChangeFromRow <> lOriginalIsPreferred Then .TextMatrix(ChangeFromRow, 4) = "No"
                    If i <> lOriginalIsPreferred Then .TextMatrix(i, 4) = "No"
                ElseIf lResponse = 2 Then
                    .TextMatrix(ChangeFromRow, 4) = "Yes"
                    .TextMatrix(i, 4) = "No"
                    If lOriginalIsPreferred <> ChangeFromRow Then .TextMatrix(lOriginalIsPreferred, 4) = "No"
                ElseIf lResponse = 3 Then
                    .TextMatrix(lOriginalIsPreferred, 4) = "No"
                    .TextMatrix(i, 4) = "No"
                    .TextMatrix(ChangeFromRow, 4) = "No"
                Else
                  .TextMatrix(ChangeFromRow, 2) = SelStats(ChangeFromRow).Abbrev
                  .TextMatrix(lOriginalIsPreferred, 4) = "Yes"
                  If ChangeFromRow <> lOriginalIsPreferred Then .TextMatrix(ChangeFromRow, 4) = "No"
                  If i <> lOriginalIsPreferred Then .TextMatrix(i, 4) = "No"
                  Exit Sub
                End If
              End If
            End If
          End If
        Next i
      End With
      'Change the Name and Unit fields to match the selected Code field
      With grdStaData
        statTypeCode = GetStatTypeCode(.TextMatrix(ChangeFromRow, 0))
        If statTypeCode = "" Then Exit Sub
        For i = 1 To SSDB.StatisticTypes(statTypeCode).StatLabels.Count
          If .TextMatrix(ChangeFromRow, 2) = SSDB.StatisticTypes(statTypeCode).StatLabels(i).code Then
            .TextMatrix(ChangeFromRow, 1) = SSDB.StatisticTypes(statTypeCode).StatLabels(i).Name
            .TextMatrix(ChangeFromRow, 5) = SSDB.StatisticTypes(statTypeCode).StatLabels(i).Units
            Exit Sub
          End If
        Next i
      End With
    Case 4: 'IsPreferred
      With grdStaData
        If .TextMatrix(ChangeFromRow, ChangeFromCol) = "Yes" Then
          For i = 1 To .Rows
            If i <> ChangeFromRow Then
              If .TextMatrix(ChangeFromRow, 1) = .TextMatrix(i, 1) Or .TextMatrix(ChangeFromRow, 2) = .TextMatrix(i, 2) Then
                If .TextMatrix(i, ChangeFromCol) = "Yes" Then
                lmsg = "For each statistic, there can be only one preferred for a station." & vbCrLf
                lmsg = lmsg & "Please specify which is preferred?"
                
                lResponse = myMsgBox.Show(lmsg, "User Action Verification", "+&Original", "-&Current", "-&None", "-&Cancel")
                lOriginalIsPreferred = OriginalIsPreferredRow(.TextMatrix(i, 1))
                  If lResponse = 1 Then
                    If lOriginalIsPreferred <> 0 Then .TextMatrix(lOriginalIsPreferred, 4) = "Yes"
                    If ChangeFromRow <> lOriginalIsPreferred Then .TextMatrix(ChangeFromRow, 4) = "No"
                    If i <> lOriginalIsPreferred Then .TextMatrix(i, 4) = "No"
                  ElseIf lResponse = 2 Then
                    .TextMatrix(ChangeFromRow, 4) = "Yes"
                    .TextMatrix(i, 4) = "No"
                    If lOriginalIsPreferred <> ChangeFromRow Then .TextMatrix(lOriginalIsPreferred, 4) = "No"
                  ElseIf lResponse = 3 Then
                    .TextMatrix(lOriginalIsPreferred, 4) = "No"
                    .TextMatrix(i, 4) = "No"
                    .TextMatrix(ChangeFromRow, 4) = "No"
                  Else
                    .TextMatrix(lOriginalIsPreferred, 4) = "Yes"
                    If ChangeFromRow <> lOriginalIsPreferred Then .TextMatrix(ChangeFromRow, 4) = "No"
                    If i <> lOriginalIsPreferred Then .TextMatrix(i, 4) = "No"
                    Exit Sub
                  End If
                End If
              End If
            End If
          Next i
        End If
      End With

'    Case 5:
'      With grdStaData
'        If Len(.TextMatrix(ChangeFromRow, ChangeFromCol)) > 20 Then
'          .TextMatrix(ChangeFromRow, ChangeFromCol) = Left(.TextMatrix(ChangeFromRow, ChangeFromCol), 20)
'          MsgBox "The Date field for the " & .TextMatrix(ChangeFromRow, 2) & " statistic" & vbCrLf & _
'              "has been truncated to 20 characters, its maximum allowable length."
'        End If
'      End With
    Case 7: 'Citation
    
      With grdStaData
        For i = 1 To .Rows
          If i <> ChangeFromRow Then
            If .TextMatrix(ChangeFromRow, ChangeFromCol) = .TextMatrix(i, ChangeFromCol) Or _
               (.TextMatrix(ChangeFromRow, ChangeFromCol) = "" And LCase(.TextMatrix(i, ChangeFromCol)) = "none") Or _
               (LCase(.TextMatrix(ChangeFromRow, ChangeFromCol)) = "none" And .TextMatrix(i, ChangeFromCol) = "") Then 'same source
              If .TextMatrix(ChangeFromRow, 1) = .TextMatrix(i, 1) Then 'same name
                
                MsgBox ("An existing version of " & .TextMatrix(ChangeFromRow, 1) & " came from a same source." & vbCrLf & "Pick/enter another source.")
                '.TextMatrix(ChangeFromRow, ChangeFromCol) = SelStats(ChangeFromRow).Source
                .TextMatrix(ChangeFromRow, ChangeFromCol) = ""
                If .TextMatrix(ChangeFromRow, ChangeFromCol) = "" Then .TextMatrix(ChangeFromRow, ChangeFromCol + 1) = ""
                Exit Sub
              End If
            End If
          End If
        Next i
      End With
      
      'updated Citation, update Citation URL also
      For i = 1 To SSDB.Sources.Count
        If grdStaData.TextMatrix(ChangeFromRow, ChangeFromCol) = SSDB.Sources(i).Name Then
          grdStaData.TextMatrix(ChangeFromRow, ChangeFromCol + 1) = SSDB.Sources(i).URL
          Exit For
        End If
      Next i
    Case 8:
      With grdStaData
        For i = 1 To .Rows
          If i <> ChangeFromRow Then
            If .TextMatrix(ChangeFromRow, ChangeFromCol) = .TextMatrix(i, ChangeFromCol) And .TextMatrix(ChangeFromRow, ChangeFromCol) <> "" Then 'same URL
              If .TextMatrix(ChangeFromRow, 1) = .TextMatrix(i, 1) Then
                MsgBox ("An existing version of " & .TextMatrix(ChangeFromRow, 1) & " came from a same source." & vbCrLf & "Pick/enter another source.")
                '.TextMatrix(ChangeFromRow, ChangeFromCol) = SelStats(ChangeFromRow).SourceURL
                .TextMatrix(ChangeFromRow, ChangeFromCol) = ""
                'If .TextMatrix(ChangeFromRow, ChangeFromCol) = "" Then .TextMatrix(ChangeFromRow, ChangeFromCol - 1) = ""
                Exit Sub
              End If
            End If
          End If
        Next i
      End With
      'updated Citation URL, update Citation also
      For i = 1 To SSDB.Sources.Count
        If grdStaData.TextMatrix(ChangeFromRow, ChangeFromCol) = SSDB.Sources(i).URL Then
          grdStaData.TextMatrix(ChangeFromRow, ChangeFromCol - 1) = SSDB.Sources(i).Name
          Exit For
        End If
      Next i
  End Select
End Sub

Private Sub ChangesMade(madeChanges As Boolean)
  Dim row&
  Dim OldVals() As String
  
  ReDim OldVals(grdStaData.Rows, 1 To UBound(DataFields))
  For row = 1 To grdStaData.Rows
    If Not newStat(row) Then
      OldVals(row, 1) = SelStats(row).StatType
      OldVals(row, 2) = SelStats(row).Name
      OldVals(row, 3) = SelStats(row).Abbrev
      OldVals(row, 4) = SelStats(row).value
      If SelStats(row).IsPreferred Then
        OldVals(row, 5) = "Yes"
      Else
        OldVals(row, 5) = "No"
      End If
      OldVals(row, 6) = SelStats(row).Units.id
      OldVals(row, 7) = SelStats(row).YearsRec
      OldVals(row, 8) = SelStats(row).Source
      OldVals(row, 9) = SelStats(row).SourceURL
      
      OldVals(row, 10) = SelStats(row).StdError
      OldVals(row, 11) = SelStats(row).Variance
      OldVals(row, 12) = SelStats(row).LowerCI
      OldVals(row, 13) = SelStats(row).UpperCI
      OldVals(row, 14) = SelStats(row).StatStartDate
      OldVals(row, 15) = SelStats(row).StatEndDate
      OldVals(row, 16) = SelStats(row).StatisticRemarks
      OldVals(row, 17) = SelStats(row).Statistic_md
      
    End If
  Next row
  RecordChanges OldVals(), madeChanges
End Sub

Private Sub RecordChanges(OldVals() As String, madeChanges As Boolean)
  Dim row&, col&, statCnt&
  Dim myStat As ssStatistic
  
  Set myStat = New ssStatistic
  Set myStat.DB = SSDB
  statCnt = grdStaData.Rows
  ReDim Changes(0 To 2, 1 To statCnt, 1 To UBound(DataFields))
  For row = 1 To statCnt
    For col = 1 To UBound(DataFields)
      If grdStaData.TextMatrix(row, col - 1) <> OldVals(row, col) Then
        If newStat(row) Then Changes(0, row, col) = "2" Else Changes(0, row, col) = "1"
        Changes(1, row, col) = OldVals(row, col)
        madeChanges = True
      End If
      If col = 2 Or col = 3 Then 'convert label string to index
        Changes(2, row, col) = GetLabelID(grdStaData.TextMatrix(row, col - 1), SSDB)
      Else
        Changes(2, row, col) = grdStaData.TextMatrix(row, col - 1)
      End If
    Next col
  Next row
End Sub

Private Sub SetGrid()
  Dim statCount&, statNumber&, col&
  Dim myStat As ssStatistic

  statCount = station.SelStats.Count
  With grdStaData
    .ClearData
    .Rows = statCount
    SizeGrid
    If statCount = 0 Then Exit Sub
    ReDim newStat(1 To statCount)
    ReDim SelStats(1 To statCount)
    For statNumber = 1 To statCount
      Set SelStats(statNumber) = station.SelStats(statNumber)
      .TextMatrix(statNumber, 0) = SelStats(statNumber).StatType
      .TextMatrix(statNumber, 1) = SelStats(statNumber).Name
      .TextMatrix(statNumber, 2) = SelStats(statNumber).Abbrev
      .TextMatrix(statNumber, 3) = SelStats(statNumber).value
      If SelStats(statNumber).IsPreferred Then
        .TextMatrix(statNumber, 4) = "Yes"
      ElseIf Not SelStats(statNumber).IsPreferred Then
        .TextMatrix(statNumber, 4) = "No"
      Else
        .TextMatrix(statNumber, 4) = "No"
      End If
      .TextMatrix(statNumber, 5) = SelStats(statNumber).Units.id
      .TextMatrix(statNumber, 6) = SelStats(statNumber).YearsRec
      .TextMatrix(statNumber, 7) = SelStats(statNumber).Source
      .TextMatrix(statNumber, 8) = SelStats(statNumber).SourceURL
      
      .TextMatrix(statNumber, 9) = SelStats(statNumber).StdError
      .TextMatrix(statNumber, 10) = SelStats(statNumber).Variance
      .TextMatrix(statNumber, 11) = SelStats(statNumber).LowerCI
      .TextMatrix(statNumber, 12) = SelStats(statNumber).UpperCI
      .TextMatrix(statNumber, 13) = SelStats(statNumber).StatStartDate
      .TextMatrix(statNumber, 14) = SelStats(statNumber).StatEndDate
      .TextMatrix(statNumber, 15) = SelStats(statNumber).StatisticRemarks
      .TextMatrix(statNumber, 16) = SelStats(statNumber).Statistic_md
      
      newStat(statNumber) = False
    Next statNumber
    For col = 0 To .Cols - 1
      .ColEditable(col) = True
    Next col
    .ColEditable(5) = False
    '.ColEditable(8) = False
    .ColEditable(16) = False
    If .Rows > 0 Then
      lblStatSel(1).Caption = .TextMatrix(1, 1)
    End If
  End With
End Sub

Private Function OriginalIsPreferredRow(aStatName As String) As Integer
    Dim lInd As Integer
    For lInd = 1 To UBound(SelStats, 1)
      If LCase(aStatName) = LCase(SelStats(lInd).Name) Then
        If SelStats(lInd).IsPreferred Then
          OriginalIsPreferredRow = lInd
          Exit For
        End If
      End If
    Next
End Function

Private Sub SizeGrid()
  With grdStaData
    .TextMatrix(0, 0) = "Statistic Type"
    .ColWidth(0) = 2500
    .TextMatrix(0, 1) = "Name"
    .ColWidth(1) = 2700
    .TextMatrix(0, 2) = "Code"
    .ColWidth(2) = 1000
    .TextMatrix(0, 3) = "Value"
    .ColWidth(3) = 600
    .TextMatrix(0, 4) = "IsPreferred"
    .ColWidth(4) = 600
    .TextMatrix(0, 5) = "Conv Flg"
    .ColWidth(5) = 740
    .TextMatrix(0, 6) = "YearsRec"
    .ColWidth(6) = 750
    .TextMatrix(0, 7) = "Citation"
    .ColWidth(7) = 3000
    .TextMatrix(0, 8) = "Citation URl"
    .ColWidth(8) = 3000
    
    .TextMatrix(0, 9) = "StdErr"
    .ColWidth(9) = 600
    .TextMatrix(0, 10) = "Variance"
    .ColWidth(10) = 600
    .TextMatrix(0, 11) = "LowerCI"
    .ColWidth(11) = 600
    .TextMatrix(0, 12) = "UpperCI"
    .ColWidth(12) = 600
    .TextMatrix(0, 13) = "StatStartDate"
    .ColWidth(13) = 800
    .TextMatrix(0, 14) = "StatEndDate"
    .ColWidth(14) = 800
    .TextMatrix(0, 15) = "Remarks"
    .ColWidth(15) = 1500
    .TextMatrix(0, 16) = "LastModified"
    .ColWidth(16) = 1750
  End With
  
End Sub
