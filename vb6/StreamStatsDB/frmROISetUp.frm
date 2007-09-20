VERSION 5.00
Object = "*\A..\..\VBEXPE~2\ATCoCtl\ATCoCtl.vbp"
Begin VB.Form frmROISetUp 
   Caption         =   "Add New ROI"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraImportFiles 
      Caption         =   "Import Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   40
      TabIndex        =   12
      Top             =   40
      Width           =   8295
      Begin VB.CommandButton cmdRHOFile 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7410
         TabIndex        =   18
         ToolTipText     =   "This file will overwrite the RHO matrix currently on file for this state"
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox txtRHOFile 
         Height          =   288
         Left            =   3990
         TabIndex        =   17
         ToolTipText     =   "This file will overwrited the RHO matrix currently on file for this state"
         Top             =   240
         Width           =   3300
      End
      Begin VB.TextBox txtMConFile 
         Height          =   288
         Left            =   3990
         TabIndex        =   16
         ToolTipText     =   "This file will overwrite the concurrent matrix currently on file for this state"
         Top             =   696
         Width           =   3300
      End
      Begin VB.CommandButton cmdMConFile 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7410
         TabIndex        =   15
         ToolTipText     =   "This file will overwrite the concurrent matrix currently on file for this state"
         Top             =   696
         Width           =   800
      End
      Begin VB.CommandButton cmdStaData 
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7410
         TabIndex        =   14
         ToolTipText     =   "Station data not already on file will be imported to the database from this file"
         Top             =   1150
         Width           =   800
      End
      Begin VB.TextBox txtStaDataFile 
         Height          =   288
         Left            =   3990
         TabIndex        =   13
         ToolTipText     =   "Station data not already on file will be imported to the database from this file"
         Top             =   1164
         Width           =   3300
      End
      Begin VB.Label lblRHOFile 
         Alignment       =   1  'Right Justify
         Caption         =   "Smoothed Correlation (RHO) File:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "This file will overwrite the RHO matrix currently on file for this state"
         Top             =   240
         Width           =   3795
      End
      Begin VB.Label lblMConFile 
         Alignment       =   1  'Right Justify
         Caption         =   "Years of Concurrent Record (MCon) File:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "This file will overwrite the concurrent matrix currently on file for this state"
         Top             =   720
         Width           =   3795
      End
      Begin VB.Label lblStaDataFile 
         Alignment       =   1  'Right Justify
         Caption         =   "Station Data File:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Station data not already on file will be imported to the database from this file"
         Top             =   1200
         Width           =   3795
      End
   End
   Begin VB.Frame fraVars 
      Caption         =   "User Input Variables by Region"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   40
      TabIndex        =   7
      Top             =   1600
      Width           =   8295
      Begin VB.ListBox lstRegions 
         Height          =   1815
         HelpContextID   =   16
         Left            =   60
         TabIndex        =   23
         Top             =   240
         Width           =   1755
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
         Height          =   340
         Left            =   5520
         TabIndex        =   11
         Top             =   1720
         Width           =   780
      End
      Begin VB.CommandButton cmdDelVar 
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
         Height          =   340
         Left            =   4560
         TabIndex        =   10
         Top             =   1720
         Width           =   780
      End
      Begin VB.CommandButton cmdAddVar 
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
         Height          =   340
         Left            =   3600
         TabIndex        =   9
         Top             =   1720
         Width           =   780
      End
      Begin ATCoCtl.ATCoGrid grdROIParms 
         Height          =   1425
         Left            =   1860
         TabIndex        =   8
         ToolTipText     =   "Only rows with every field entered will be saved to the database"
         Top             =   240
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   2514
         SelectionToggle =   0   'False
         AllowBigSelection=   -1  'True
         AllowEditHeader =   0   'False
         AllowLoad       =   0   'False
         AllowSorting    =   0   'False
         Rows            =   2
         Cols            =   4
         ColWidthMinimum =   300
         gridFontBold    =   0   'False
         gridFontItalic  =   0   'False
         gridFontName    =   "MS Sans Serif"
         gridFontSize    =   8
         gridFontUnderline=   0   'False
         gridFontWeight  =   400
         gridFontWidth   =   0
         Header          =   ""
         FixedRows       =   2
         FixedCols       =   0
         ScrollBars      =   2
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
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7605
      TabIndex        =   5
      Top             =   5880
      Width           =   735
   End
   Begin VB.Frame fraAnalysisOptions 
      Caption         =   "Analysis Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   4680
      TabIndex        =   1
      Top             =   3840
      Width           =   3660
      Begin ATCoCtl.ATCoText atxSimStations 
         Height          =   255
         Left            =   2640
         TabIndex        =   26
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         InsideLimitsBackground=   16777215
         OutsideHardLimitBackground=   8421631
         OutsideSoftLimitBackground=   8454143
         HardMax         =   100
         HardMin         =   20
         SoftMax         =   100
         SoftMin         =   20
         MaxWidth        =   -999
         Alignment       =   1
         DataType        =   1
         DefaultValue    =   "30"
         Value           =   "30"
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkUseRegions 
         Caption         =   "distinct regions within state"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   3495
      End
      Begin VB.CheckBox chkRegress 
         Caption         =   "back-step least-squares regression"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   3495
      End
      Begin VB.CheckBox chkCF 
         Caption         =   "use climate factor in similarity calc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   3495
      End
      Begin VB.CheckBox chkDistance 
         Caption         =   "use distance in similarity calc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblSimStations 
         Caption         =   "Number of similar stations:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Frame fraReturnPds 
      Caption         =   "Return Periods"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2440
      Left            =   40
      TabIndex        =   0
      Top             =   3840
      Width           =   4620
      Begin StreamStatsDB.ATCoSelectListSortByProp lstReturnPeriods 
         Height          =   2175
         Left            =   120
         TabIndex        =   24
         Top             =   165
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3836
         RightLabel      =   "Selected:"
         LeftLabel       =   "Available:"
      End
   End
End
Attribute VB_Name = "frmROISetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CurRegion As nssRegion
Private SelectStatsOnFile() As ssStatistic
Private ExistingRegions As New FastCollection 'of nssRegion

Private Sub cmdAddVar_Click()
  Dim col As Long
  Dim newParm As New nssParameter

  Set newParm.Region = CurRegion
  CurRegion.ROIParameters.Add newParm
  With grdROIParms
    For col = 0 To .Cols - 1
      .ColEditable(col) = True
    Next col
    .Rows = .Rows + 1
    .row = .Rows
    .col = 5
  End With
End Sub

Private Sub cmdCancel_Click()
  CurRegion.ROIParameters.Clear
  Set CurRegion.ROIParameters = Nothing
  PopulateGrid
End Sub

Private Sub cmdDelVar_Click()
  With grdROIParms
    If .Rows = 0 Then Exit Sub
    If .row <= CurRegion.ROIParameters.Count Then CurRegion.ROIParameters.RemoveByIndex .row
    .DeleteRow .row
  End With
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdRHOFile_Click()
  Dim Filename$, BinFile$
  
  On Error GoTo x

  With frmCDLG.CDLG
    .DialogTitle = "Select the RHO flat file for import"
    .Filename = ExePath & "*.rho"
    .Filter = "(*.rho)|*.rho|(*.txt)|*.txt|(All Files)|*.*"
    .filterIndex = 1
    .CancelError = True
    .ShowOpen
    If Len(Dir(.Filename, vbDirectory)) > 1 Then
      Me.MousePointer = vbHourglass
      BinFile = MatrixToBinary(.Filename, 1, SSDB.state.Abbrev)
      Me.MousePointer = vbDefault
      If Len(BinFile) > 0 Then
        MsgBox "RHO file stored in binary format as " & BinFile, vbInformation, "StreamStatsDB"
        SaveSetting "StreamStatsDB", "Defaults", SSDB.state.Abbrev & "_RHOImportFile", BinFile
        .Filename = BinFile
      Else
        MsgBox "Problem reading RHO flat file into binary format.", vbExclamation, "StreamStatsDB"
        .Filename = ""
      End If
    End If
    txtRHOFile.Text = .Filename
  End With
x:
  Unload frmCDLG
End Sub

Private Sub cmdMConFile_Click()
  Dim Filename$, BinFile$
  
  On Error GoTo x
  
  With frmCDLG.CDLG
    .DialogTitle = "Select the MCon flat file for import"
    .Filename = ExePath & "*.rec"
    .Filter = "(*.rec)|*.rec|(*.txt)|*.txt|(All Files)|*.*"
    .filterIndex = 1
    .CancelError = True
    .ShowOpen
    If Len(Dir(.Filename, vbDirectory)) > 1 Then
      Me.MousePointer = vbHourglass
      BinFile = MatrixToBinary(.Filename, 2, SSDB.state.Abbrev)
      Me.MousePointer = vbDefault
      If Len(BinFile) > 0 Then
        MsgBox "MCon file stored in binary format as " & BinFile, vbInformation, "StreamStatsDB"
        SaveSetting "StreamStatsDB", "Defaults", SSDB.state.Abbrev & "_MConImportFile", BinFile
        .Filename = BinFile
      Else
        MsgBox "Problem reading MCon flat file into binary format.", vbExclamation, "StreamStatsDB"
        .Filename = ""
      End If
    End If
    txtMConFile.Text = .Filename
  End With
x:
  Unload frmCDLG
End Sub

Private Sub cmdSave_Click()
  Dim row As Long, col As Long, regionCnter As Long, ParmCnter As Long
  Dim str As String
  Dim myparm As nssParameter
  Dim MyDepVar As nssDepVar
  
  'Perform QA check on values selected/entered in grid
  If Not QACheck Then Exit Sub
  
  'Record the changes made to statistic values
  frmUserInfo.Show vbModal, Me
  If Not UserInfoOK Then GoTo NoChanges
  
  Me.MousePointer = vbHourglass

  For row = 1 To lstReturnPeriods.RightCount
    If row = lstReturnPeriods.RightCount Then
      str = str & lstReturnPeriods.RightItem(row - 1)
    Else
      str = str & lstReturnPeriods.RightItem(row - 1) & ","
    End If
  Next row
  SSDB.state.Edit str, -chkCF.value, -chkDistance.value, -chkRegress.value, -chkUseRegions.value, atxSimStations.value 'useRegions
  For regionCnter = 1 To lstRegions.ListCount
    Set CurRegion = ExistingRegions(regionCnter)
    CurRegion.ClearROIUserparms
    For ParmCnter = 1 To CurRegion.ROIParameters.Count
      Set myparm = CurRegion.ROIParameters(ParmCnter)
      If myparm.Name <> "Not Assigned" Then
        Set myparm.Region = CurRegion
        myparm.Add CurRegion, myparm.Abbrev, myparm.GetMin(False), _
            myparm.GetMax(False), myparm.Units.id
        myparm.AddROIUserParm CurRegion, myparm.Abbrev, _
            myparm.CorrelationType, myparm.SimulationVar, myparm.RegressionVar
      End If
    Next ParmCnter
    For row = 1 To lstReturnPeriods.RightCount
      Set MyDepVar = New nssDepVar
      MyDepVar.Add True, CurRegion, lstReturnPeriods.RightItem(row - 1)
    Next row
  Next regionCnter
NoChanges:
  Me.MousePointer = vbDefault
End Sub

Private Sub cmdStaData_Click()
  Dim Filename$
  Dim ff As ATCoFindFile
  
  On Error GoTo x
  
  ImportedNewData = False
  Filename = GetSetting("StreamStatsDB", "Defaults", SSDB.state.Abbrev & "_StaDataImportFile")
  With frmCDLG.CDLG
    .DialogTitle = "Select the Station Data file for import"
    .Filename = Filename
    .Filter = "(All Files)|*.*"
    .filterIndex = 1
    .CancelError = True
    .ShowOpen
    
    On Error GoTo Errhand
    
    txtStaDataFile.Text = .Filename
    If Len(Dir(.Filename, vbDirectory)) > 1 Then
      SaveSetting "StreamStatsDB", "Defaults", SSDB.state.Abbrev & "_StaDataImportFile", .Filename
      frmImportStations.OpenDataFile .Filename
      frmImportStations.Show vbModal, Me
      If ImportedNewData Then
        SSDB.state.StatsOnFile.Clear
        Set SSDB.state.StatsOnFile = Nothing
        SSDB.state.Regions.Clear
        Set SSDB.state.Regions = Nothing
        Form_Load
      End If
    End If
  End With
Errhand:
  If Err.Number > 0 Then MsgBox Err.Description, vbCritical, "Error Opening Station Data File"
x:
  Unload frmCDLG
End Sub

Private Sub Form_Load()
  Dim i As Long, ParmCnt As Long, regnCnter As Long, height As Long
  Dim statAbbrev As String

  Me.Caption = "Add new ROI data for " & SSDB.state.Name
  'Retrieve name of station data import file from registry
  txtStaDataFile.Text = GetSetting("StreamStatsDB", "Defaults", SSDB.state.Abbrev & "_StaDataImportFile")
ImportedData:
  grdROIParms.Rows = 0
  ReDim SelectStatsOnFile(0)
  'Create array with superset of ssStatistics on file for regions in state
  For ParmCnt = 1 To SSDB.state.StatsOnFile.Count
    Select Case SSDB.state.StatsOnFile(ParmCnt).statTypeCode
      Case "D":   'none of these stats are mathematical quantities
      Case "PFS": 'do not count actual peak flows or their std dev
        statAbbrev = SSDB.state.StatsOnFile(ParmCnt).Abbrev
        If Not (Left(statAbbrev, 1) = "P" And _
            IsNumeric(Mid(statAbbrev, 2)) Or statAbbrev = "SDPK") Then
          ReDim Preserve SelectStatsOnFile(UBound(SelectStatsOnFile) + 1)
          Set SelectStatsOnFile(UBound(SelectStatsOnFile)) = SSDB.state.StatsOnFile(ParmCnt)
        End If
      Case Else   'count all remaining stats
        ReDim Preserve SelectStatsOnFile(UBound(SelectStatsOnFile) + 1)
        Set SelectStatsOnFile(UBound(SelectStatsOnFile)) = SSDB.state.StatsOnFile(ParmCnt)
    End Select
  Next ParmCnt

  'set number of similar stations to use
  If SSDB.state.ROISimStations > 0 Then atxSimStations.value = SSDB.state.ROISimStations
  'Select appropriate check boxes
  If SSDB.state.ROIDistance Then chkDistance.value = 1 Else chkDistance.value = 0
  If SSDB.state.ROIClimateFactor Then chkCF.value = 1 Else chkCF.value = 0
  If SSDB.state.ROIRegress Then chkRegress.value = 1 Else chkRegress.value = 0
  If SSDB.state.ROIUseRegions Then chkUseRegions.value = 1 Else chkUseRegions.value = 0

  'Fill in ROI regions, if there are any
  For regnCnter = 1 To SSDB.state.Regions.Count
    If SSDB.state.Regions(regnCnter).ROIRegnID > 0 Then
      If height = 0 Then
        height = 255
      Else
        height = height + 195
      End If
      lstRegions.AddItem SSDB.state.Regions(regnCnter).Name
      ExistingRegions.Add SSDB.state.Regions(regnCnter)
    End If
  Next regnCnter
  If height > 1815 Then height = 1815
  If lstRegions.ListCount > 0 Then
    ImportedNewData = True
    lstRegions.height = height
    lstRegions.Selected(0) = True
  Else
    lstRegions.height = 255
    MsgBox "There are no ROI data on file." & vbCrLf & vbCrLf & _
        "You must import station data for " & SSDB.state.Name & vbCrLf & _
        "before specifying the ROI parameters.", , "need station data"
    cmdStaData_Click
    If ImportedNewData Then  'reset all region data in state
      SSDB.state.Regions.Clear
      Set SSDB.state.Regions = Nothing
      GoTo ImportedData
    Else
      Exit Sub
    End If
  End If
  'Fill in names of matrix files
  txtRHOFile.Text = GetSetting("StreamStatsDB", "Defaults", SSDB.state.Abbrev & "_RHOImportFile")
  txtMConFile.Text = GetSetting("StreamStatsDB", "Defaults", SSDB.state.Abbrev & "_MConImportFile")
  
  PopulateReturnPeriods
  
  With grdROIParms
    .Rows = 0
    .Cols = 7
    For i = 0 To .Cols - 1
      .ColEditable(i) = True
    Next i
    .col = 5
  End With
  lstRegions_Click
  
End Sub

Private Sub PopulateReturnPeriods()
  Dim vRetPd As Variant
  Dim i As Long, j As Long
  Dim statAbbrev As String
  
  With lstReturnPeriods
    .ClearRight
    .ClearLeft
    For i = 1 To SSDB.state.StatsOnFile.Count
      statAbbrev = SSDB.state.StatsOnFile(i).Abbrev
      If (SSDB.state.StatsOnFile(i).statTypeCode = "PFS" And _
          Left(statAbbrev, 1) = "P" And _
          IsNumeric(Mid(statAbbrev, 2))) Then
        .LeftItem(j) = statAbbrev
        .LeftItemData(j) = SSDB.state.StatsOnFile(i).code
        j = j + 1
      End If
    Next i
    For Each vRetPd In SSDB.state.ROIPeakFlows
      For i = 0 To .LeftCount - 1
        If vRetPd.code = .LeftItem(i) Then
          .MoveRight (i)
        End If
      Next i
    Next vRetPd
  End With
End Sub

Private Sub grdROIParms_CommitChange(ChangeFromRow As Long, ChangeToRow As Long, _
                                     ChangeFromCol As Long, ChangeToCol As Long)
  Dim i&, response&
  Dim str$, statTypeCode$
  Dim myparm As nssParameter
  
  Set myparm = CurRegion.ROIParameters(ChangeFromRow)
  'Adjust appropriate columns in row when a field is edited
  Select Case ChangeFromCol
    Case 0:
      'Make sure this Statistic does not already exist for this station
      With grdROIParms
        For i = 1 To .Rows
          If i <> ChangeFromRow Then
            If .TextMatrix(ChangeFromRow, 0) = .TextMatrix(i, 0) Then
              MsgBox "This parameter already exists for this station."
              .TextMatrix(ChangeFromRow, 0) = myparm.Name
              .TextMatrix(ChangeFromRow, 1) = myparm.GetMin(False)
              .TextMatrix(ChangeFromRow, 2) = myparm.GetMax(False)
              .TextMatrix(ChangeFromRow, 3) = myparm.SimulationVar
              .TextMatrix(ChangeFromRow, 4) = myparm.RegressionVar
              If myparm.CorrelationType > 100 Then
                .TextMatrix(ChangeFromRow, 5) = "Positive or Negative"
              ElseIf myparm.CorrelationType > 0 Then
                .TextMatrix(ChangeFromRow, 5) = "Positive"
              ElseIf myparm.CorrelationType < 0 Then
                .TextMatrix(ChangeFromRow, 5) = "Negative"
              Else
                .TextMatrix(ChangeFromRow, 5) = "None"
              End If
              If myparm.CorrelationType > 100 Then
                .TextMatrix(ChangeFromRow, 6) = myparm.CorrelationType / 1000
              Else
                .TextMatrix(ChangeFromRow, 6) = myparm.CorrelationType
              End If
              Exit Sub
            End If
          End If
        Next i
        'Find chosen stat and assign attributes to parm
        For i = 1 To UBound(SelectStatsOnFile)
          If .TextMatrix(ChangeFromRow, ChangeFromCol) = SelectStatsOnFile(i).Name Then Exit For
        Next i
        If SSDB.Parameters.KeyExists(SelectStatsOnFile(i).Abbrev) Then
          Set CurRegion.ROIParameters(ChangeFromRow).Units = SelectStatsOnFile(i).Units
          CurRegion.ROIParameters(ChangeFromRow).LabelCode = SelectStatsOnFile(i).code
          CurRegion.ROIParameters(ChangeFromRow).Abbrev = SelectStatsOnFile(i).Abbrev
          CurRegion.ROIParameters(ChangeFromRow).Name = SelectStatsOnFile(i).Name
        End If
      End With
  End Select

  With grdROIParms
    Select Case ChangeFromCol
      Case 0:
        CurRegion.ROIParameters(ChangeFromRow).Name = .TextMatrix(ChangeFromRow, ChangeFromCol)
      Case 1:
        If IsNumeric(.TextMatrix(ChangeFromRow, ChangeFromCol)) Then
          CurRegion.ROIParameters(ChangeFromRow).SetMin CDbl(.TextMatrix(ChangeFromRow, ChangeFromCol)), False
        End If
      Case 2:
        If IsNumeric(.TextMatrix(ChangeFromRow, ChangeFromCol)) Then
          CurRegion.ROIParameters(ChangeFromRow).SetMax .TextMatrix(ChangeFromRow, ChangeFromCol), False
        End If
      Case 3:
        CurRegion.ROIParameters(ChangeFromRow).SimulationVar = .TextMatrix(ChangeToRow, ChangeFromCol)
      Case 4:
        CurRegion.ROIParameters(ChangeFromRow).RegressionVar = .TextMatrix(ChangeToRow, ChangeFromCol)
        If .TextMatrix(ChangeFromRow, ChangeFromCol) = "False" Then
          .TextMatrix(ChangeFromRow, ChangeFromCol + 1) = ""
        End If
      Case Else
        If IsNumeric(.TextMatrix(ChangeFromRow, 6)) Then
          Select Case LCase(.TextMatrix(ChangeFromRow, 5))
            Case "positive only": CurRegion.ROIParameters(ChangeFromRow).CorrelationType = .TextMatrix(ChangeFromRow, 6)
            Case "negative only": 'make sure resulting value is negative
              If .TextMatrix(ChangeFromRow, 6) > 0 Then
                CurRegion.ROIParameters(ChangeFromRow).CorrelationType = -.TextMatrix(ChangeFromRow, 6)
              Else
                CurRegion.ROIParameters(ChangeFromRow).CorrelationType = .TextMatrix(ChangeFromRow, 6)
              End If
            Case "positive or negative": CurRegion.ROIParameters(ChangeFromRow).CorrelationType = 1000 * Abs(.TextMatrix(ChangeFromRow, 6))
            Case Else
              CurRegion.ROIParameters(ChangeFromRow).CorrelationType = 0
          End Select
        Else
          CurRegion.ROIParameters(ChangeFromRow).CorrelationType = 0
        End If
    End Select
  End With

End Sub

Private Sub grdROIParms_RowColChange()
  Dim i&
  Dim statAbbrev$
  
  'Fill in combo box entries
  With grdROIParms
    'PopulateGrid
    If .row = 0 Then Exit Sub
    .ClearValues
    SizeGrid
    Select Case .col
      'Fill list of Statistic Names in second column
      Case 0:
        For i = 1 To SSDB.state.StatsOnFile.Count
          Select Case SSDB.state.StatsOnFile(i).statTypeCode
            Case "D":   'none of these parameters are mathematical quantities
            Case "PFS":
              statAbbrev = SSDB.state.StatsOnFile(i).Abbrev
              If Not (Left(statAbbrev, 1) = "P" And IsNumeric(Mid(statAbbrev, 2)) Or statAbbrev = "SDPK") Then
                .addValue SSDB.state.StatsOnFile(i).Name
              End If
            Case Else
              .addValue SSDB.state.StatsOnFile(i).Name
          End Select
        Next i
        .ComboCheckValidValues = True
      Case 3:
        .addValue "True"
        .addValue "False"
        .ComboCheckValidValues = True
      Case 4:
        .addValue "True"
        .addValue "False"
        .ComboCheckValidValues = True
      Case 5:
        If .TextMatrix(.row, 4) = "True" Then
          .addValue "Positive Only"
          .addValue "Negative Only"
          .addValue "Positive or Negative"
          .addValue "None"
        Else
          .addValue ""
        End If
        .ComboCheckValidValues = True
    End Select
  End With
End Sub

Private Sub PopulateGrid()
  Dim i As Long, corrIndex As Long, curRow As Long
  Dim myparm As nssParameter
  
  SizeGrid
  With grdROIParms
    .ClearData
    .Rows = 0
    For i = 1 To CurRegion.ROIParameters.Count
      Set myparm = CurRegion.ROIParameters(i)
      .TextMatrix(i, 0) = myparm.Name
      .TextMatrix(i, 1) = myparm.GetMin(False)
      .TextMatrix(i, 2) = myparm.GetMax(False)
      If myparm.SimulationVar Then
        .TextMatrix(i, 3) = "True"  'used in similarity calcs
      Else
        .TextMatrix(i, 3) = "False" 'not used in similarity calcs
      End If
      If myparm.RegressionVar Then
        .TextMatrix(i, 4) = "True"  'used in regression analysis
        If myparm.CorrelationType > 100 Then 'positive or negative correlation
          .TextMatrix(i, 5) = "Positive or Negative"
        ElseIf myparm.CorrelationType > 0 Then
          .TextMatrix(i, 5) = "Positive Only"
        ElseIf myparm.CorrelationType < 0 Then
          .TextMatrix(i, 5) = "Negative Only"
        Else 'no correlation
          .TextMatrix(i, 5) = "None"
        End If
        If myparm.CorrelationType > 100 Then
          .TextMatrix(i, 6) = myparm.CorrelationType / 1000
        Else
          .TextMatrix(i, 6) = myparm.CorrelationType
        End If
      Else
        .TextMatrix(i, 4) = "False" 'not used in regression analysis
        .TextMatrix(i, 5) = ""      'no regression correlation
        .TextMatrix(i, 6) = ""
      End If
    Next i
  End With
End Sub

Private Sub SizeGrid()
  With grdROIParms
    .Cols = 7
    .TextMatrix(-1, 0) = "Parameter"
    .TextMatrix(0, 0) = "Name"
    .TextMatrix(-1, 1) = "Minimum"
    .TextMatrix(0, 1) = "Value"
    .TextMatrix(-1, 2) = "Maximum"
    .TextMatrix(0, 2) = "Value"
    .TextMatrix(-1, 3) = "Use for"
    .TextMatrix(0, 3) = "Similarity"
    .TextMatrix(-1, 4) = "Use in"
    .TextMatrix(0, 4) = "Regression"
    .TextMatrix(-1, 5) = "Regression"
    .TextMatrix(0, 5) = "Correlation"
    .TextMatrix(-1, 6) = "T-beta"
    .TextMatrix(0, 6) = "Limit"
    .ColWidth(0) = 1600
    .ColWidth(1) = 660
    .ColWidth(2) = 700
    .ColWidth(3) = 710
    .ColWidth(4) = 890
    .ColWidth(5) = 940
    .ColWidth(6) = 600
  End With
End Sub

Private Sub lblCF_Click()
  With chkCF
    If .value = 1 Then
      .value = 0
    Else
      .value = 1
    End If
  End With
End Sub

Private Sub lblDistance_Click()
  With chkDistance
    If .value = 1 Then
      .value = 0
    Else
      .value = 1
    End If
  End With
End Sub

Private Sub lblRegress_Click()
  With chkRegress
    If .value = 1 Then
      .value = 0
    Else
      .value = 1
    End If
  End With
End Sub

Private Sub lblUseRegions_Click()
  With chkUseRegions
    If .value = 1 Then
      .value = 0
    Else
      .value = 1
    End If
  End With
End Sub

Private Function QACheck() As Boolean
  Dim row As Long, col As Long
  Dim str As String
  
  With grdROIParms
    'Perform QA check on fields in grid
    For row = 1 To .Rows
      For col = 0 To .Cols - 2
        If Len(.TextMatrix(row, col)) = 0 Then
          Select Case col
            Case 0:
              str = "No parameter 'Name' has been selected for row " & .row & " of the grid." & _
                  vbCrLf & "You must make this selection or delete that row from the grid."
            Case 1:
              .TextMatrix(row, col) = "0.0"
            Case 2:
              str = "No 'Minimum Value' has been entered for the parameter " & _
                  .TextMatrix(row, 0) & "." & vbCrLf & _
                  "Click this field to enter a selection then try again to save."
            Case 3:
              str = "No selection has been made in the 'Use for Similarity' field for " & _
                  " the parameter " & .TextMatrix(row, 0) & "." & vbCrLf & _
                  "Double-click this field to make a selection then try again to save."
            Case 4:
              str = "No selection has been made in the 'Use in Regression' field for " & _
                  " the parameter " & .TextMatrix(row, 0) & "." & vbCrLf & _
                  "Double-click this field to make a selection then try again to save."
          End Select
        End If
        If Len(str) > 0 Then
          MsgBox str, vbCritical, "missing fields"
          Exit Function
        End If
      Next col
    Next row
  End With
  QACheck = True
End Function

Private Function MatrixToBinary(FName As String, MatrixType As Integer, stName As String) As String
  'FName - flat file containing matrix data
  'MatrixType - 1 - RHO, 2 - REC/MCon (coincident years)
  'StName - 2 character state abbreviation
  Dim FirstVal As Long, FldLen As Long, ipos As Long, Funit As Long
  Dim Maxnv As Long, nv As Long
  Dim RVals() As Single, TotVal As Single
  Dim IVals() As Integer
  Dim Fstr As String, istr As String, Rstr As String, FType(2) As String

  Fstr = WholeFileString(FName)
  FirstVal = 0
  While FirstVal = 0 And Len(Fstr) > 0
    istr = StrSplit(Fstr, vbCrLf, "")
    If IsNumeric(istr) Then
      If MatrixType = 1 Then 'RHO file should have first value of 1.0
        If CSng(istr) = 1# Then
          FirstVal = 1
          FldLen = Len(istr)
        End If
      Else 'assume first valid in REC file
        FirstVal = 1
      End If
    End If
  Wend
  If FirstVal = 1 Then
    Maxnv = 500000
    If MatrixType = 1 Then
      ReDim Preserve RVals(Maxnv)
      RVals(1) = 1# '1st RHO value is always 1.0
      TotVal = RVals(1)
    Else
      ReDim Preserve IVals(Maxnv)
      IVals(1) = CInt(istr)
      TotVal = IVals(1)
    End If
    nv = 1
    While Len(Fstr) > 0 'process rest of file
      Rstr = StrSplit(Fstr, vbCrLf, "") 'next record
      If MatrixType = 1 Then
        ipos = 1
        While ipos < Len(Rstr)
          nv = nv + 1
          RVals(nv) = Mid(Rstr, ipos, FldLen)
          TotVal = TotVal + RVals(nv)
          ipos = ipos + FldLen
        Wend
      Else
        While Len(Rstr) > 0
          istr = StrSplit(Rstr, " ", "")
          nv = nv + 1
          IVals(nv) = CInt(istr)
          TotVal = TotVal + IVals(nv)
        Wend
      End If
    Wend
    'write out the binary version
    FType(1) = ".rho"
    FType(2) = ".rec"
    Fstr = PathNameOnly(FName) & "\" & stName & FType(MatrixType) & ".bin"
    Funit = FreeFile(0)
    Open Fstr For Binary As #Funit
    Put #Funit, , MatrixType
    Put #Funit, , nv
    Put #Funit, , TotVal
    For ipos = 1 To nv
      If MatrixType = 1 Then
        Put #Funit, , RVals(ipos)
      Else
        Put #Funit, , IVals(ipos)
      End If
    Next ipos
    Close #Funit
  Else
    Fstr = ""
  End If
  MatrixToBinary = Fstr

End Function

Private Sub lstRegions_Click()
  Set CurRegion = SSDB.state.Regions(lstRegions.List(lstRegions.ListIndex))
  PopulateGrid
End Sub
