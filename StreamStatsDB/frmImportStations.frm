VERSION 5.00
Begin VB.Form frmImportStations 
   Caption         =   "Import Stations"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   7380
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   7095
      Begin VB.CommandButton cmdLastRow 
         Cancel          =   -1  'True
         Caption         =   "Ignore Lines After Selection"
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
         Left            =   2880
         TabIndex        =   4
         ToolTipText     =   "Ignore data from file below selected row"
         Top             =   0
         Width           =   2775
      End
      Begin VB.CommandButton cmdFirstRow 
         Caption         =   "Ignore Lines Before Selection"
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
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Ignore data from file above selected row"
         Top             =   0
         Width           =   2775
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Import Now"
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
         Left            =   5760
         TabIndex        =   2
         Top             =   0
         Width           =   1335
      End
   End
   Begin ATCoCtl.ATCoGrid agd 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4683
      SelectionToggle =   0   'False
      AllowBigSelection=   -1  'True
      AllowEditHeader =   0   'False
      AllowLoad       =   0   'False
      AllowSorting    =   0   'False
      Rows            =   78
      Cols            =   2
      ColWidthMinimum =   300
      gridFontBold    =   0   'False
      gridFontItalic  =   0   'False
      gridFontName    =   "MS Sans Serif"
      gridFontSize    =   8
      gridFontUnderline=   0   'False
      gridFontWeight  =   400
      gridFontWidth   =   0
      Header          =   "Specify Statistics for columns to import. Ignore parts of file if necessary."
      FixedRows       =   0
      FixedCols       =   1
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
Attribute VB_Name = "frmImportStations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FirstDataRow = 5

Private NameDataFile As String
Private DataFilePos As Long
Private DataFileString As String
Private LenDataFileString As String
Private DataFileLineEnd As String
Private LenDataFileLineEnd As Long

Private SkipDataLinesStart As Long 'How many lines of data file to skip before reading/populating grid
Private SkipDataLinesEnd As Long   'How many lines of data file to skip at the end of the file
Private DataFileLine() As String   'Array containing each line of data file
Private DataFileLines As Long
Private DataFileLinesAllocated As Long

Private Cols As Long

Private SamplePopulated As Boolean
Private Delimited As Boolean
Private Delimiter As String
' for fixed-width columns (non-delimited files)
Private ColStart() As Long
Private ColWidth() As Long

Private Sub DiscoverFormat()
  Dim fileline As Long
  Dim col As Long
  Dim notfirstline As Boolean
  Dim InColumn As Boolean
  Dim buf As String
  Dim LenBuf As Long
  Dim CharPos As Long
  Dim countTabs As Long
  Dim countCommas As Long
  Dim IsWhitespace() As Boolean  ' for fixed-width columns (non-delimited files)
  Dim LenFirstLine As Long
  
  For fileline = 1 + SkipDataLinesStart To DataFileLines - SkipDataLinesEnd - SkipDataLinesStart
    buf = DataFileLine(fileline)
    LenBuf = Len(buf)
    If notfirstline Then
      If countTabs > 0 Then
        If CountString(buf, vbTab) <> countTabs Then countTabs = 0
      End If
      If countCommas > 0 Then
        If CountString(buf, ",") <> countCommas Then countCommas = 0
      End If
      If LenBuf > LenFirstLine Then LenBuf = LenFirstLine
      For CharPos = 1 To LenBuf
        If IsWhitespace(CharPos) Then
          If Mid(buf, CharPos, 1) <> " " Then IsWhitespace(CharPos) = False
        End If
      Next
    Else
      LenFirstLine = LenBuf
      countTabs = CountString(buf, vbTab)
      countCommas = CountString(buf, ",")
      ReDim IsWhitespace(LenFirstLine)
      For CharPos = 1 To LenBuf
        If Mid(buf, CharPos, 1) = " " Then IsWhitespace(CharPos) = True
      Next
      notfirstline = True
    End If
  Next
  If countTabs > 0 Then
    Cols = countTabs
    Delimited = True
    Delimiter = vbTab
  ElseIf countCommas > 0 Then
    Cols = countCommas
    Delimited = True
    Delimiter = ","
  Else
    Delimited = False
    InColumn = False
    col = 0
    For CharPos = 1 To LenFirstLine
      If IsWhitespace(CharPos) Then
        If InColumn Then
          ColWidth(col) = CharPos - ColStart(col)
          InColumn = False
        End If
      Else 'Is not whitespace
        If Not InColumn Then
          InColumn = True
          col = col + 1
          ReDim Preserve ColStart(col)
          ReDim Preserve ColWidth(col)
          ColStart(col) = CharPos
        End If
      End If
    Next
    Cols = col
    If InColumn Then ColWidth(col) = CharPos - ColStart(col)
  End If
End Sub

Public Sub OpenDataFile(Optional filename As String = "")
  Dim msg As String
  Dim m_FileName As String
  Dim ff As ATCoFindFile
  
  If Len(filename) > 0 Then
    If Len(Dir(filename)) > 0 Then
      NameDataFile = filename
      GoTo OpenIt
    End If
  End If
  
  Set ff = New ATCoFindFile
  ff.SetDialogProperties "Open Station File", filename
  ff.SetRegistryInfo "StreamStats", "files", "Stations"
  NameDataFile = ff.GetName(True)
    
OpenIt:
  DataFileString = WholeFileString(NameDataFile)
  LenDataFileString = Len(DataFileString)
  DataFileLinesAllocated = LenDataFileString / 80
  ReDim DataFileLine(DataFileLinesAllocated)
  DataFileLines = 0
  While Not DataEOF
    DataFileLines = DataFileLines + 1
    If DataFileLinesAllocated < DataFileLines Then
      DataFileLinesAllocated = DataFileLinesAllocated + 500
      ReDim Preserve DataFileLine(DataFileLinesAllocated)
    End If
    DataFileLine(DataFileLines) = DataNextLine
  Wend
  ReDim Preserve DataFileLine(DataFileLines)
  DiscoverFormat
  PopulateSample
End Sub

' Subroutine ===============================================
' Name:      PopulateSample
' Purpose:   Populates the sample grid with data from file.
'
Private Sub PopulateSample()
  Dim row As Long
  Dim col As Long
  Dim cbuff As String
  Dim parsed() As String
  Dim pcols As Long
  Dim fileline As Long
  Me.MousePointer = vbHourglass
  With agd
    .TextMatrix(1, 0) = "Input Column"
    .TextMatrix(2, 0) = "Attribute Type"
    .TextMatrix(3, 0) = "Attribute Name"
    .TextMatrix(4, 0) = "Convert from Log"
    .Rows = FirstDataRow
    row = .Rows - 1
    For fileline = 1 + SkipDataLinesStart To DataFileLines - SkipDataLinesEnd
      row = row + 1
      cbuff = DataFileLine(fileline)
      pcols = ParseInputLine(cbuff, parsed)
      For col = 1 To pcols
        .TextMatrix(row, col) = parsed(col)
      Next col
    Next
    For col = 1 To Cols
      .TextMatrix(1, col) = col
    Next col
    .ColsSizeByContents
  End With
  SamplePopulated = True
  Me.MousePointer = vbDefault
End Sub

' Name:     ParseInputLine (3 arguments)
' Purpose:  Returns number of columns parsed from buffer into array
'           Populates parsed array from element 1 to index returned
'
Private Function ParseInputLine( _
                 ByVal inbuf As String, _
                 parsed() As String) As Long

  Dim parseCol As Long
  Dim fromCol As Long
  Dim toCol As Long
  parseCol = 0
  fromCol = 1
  ReDim parsed(Cols)
  If Delimited Then 'parse delimited text
    While fromCol <= Len(inbuf) And parseCol < UBound(parsed)
      toCol = InStr(fromCol, inbuf, Delimiter)
      'If chkCollapseDelim.Value = 1 Then 'treat multiple contiguous delimiters as one
       ' While toCol = fromCol And toCol < Len(inBuf)
        '  toCol = fromCol + 1
         ' toCol = InStr(fromCol, inBuf, delim)
       ' Wend
      'End If
      If toCol < fromCol Then toCol = Len(inbuf) + 1
      parseCol = parseCol + 1
      parsed(parseCol) = Mid(inbuf, fromCol, toCol - fromCol)
      fromCol = toCol + 1
    Wend
  Else
    For parseCol = 1 To Cols
      parsed(parseCol) = Trim(Mid(inbuf, ColStart(parseCol), ColWidth(parseCol)))
    Next
  End If
  If parseCol > UBound(parsed) Then parseCol = UBound(parsed)
  ParseInputLine = parseCol
End Function

Function DataEOF() As Boolean
  If DataFilePos > LenDataFileString Then DataEOF = True
End Function
Function DataNextLine() As String
  Dim EOLpos As Long
  
  If DataFilePos = 0 Then DataFilePos = 1
  If DataFilePos > LenDataFileString Then
    DataNextLine = ""
  Else
    While Mid(DataFileString, DataFilePos, 1) = vbCr _
       Or Mid(DataFileString, DataFilePos, 1) = vbLf
      DataFilePos = DataFilePos + 1
    Wend
    If LenDataFileLineEnd = 0 Then FindDataFileLineEnd
    If LenDataFileLineEnd > 0 Then
      EOLpos = InStr(DataFilePos, DataFileString, DataFileLineEnd)
      If EOLpos = 0 Then EOLpos = LenDataFileString + 1
      DataNextLine = Mid(DataFileString, DataFilePos, EOLpos - DataFilePos)
      DataFilePos = EOLpos + LenDataFileLineEnd
    End If
  End If
End Function

Private Sub FindDataFileLineEnd()
  Dim CRpos As Long
  Dim LFpos As Long
  CRpos = InStr(DataFileString, vbCr)
  LFpos = InStr(DataFileString, vbLf)
  If CRpos > 0 And LFpos = CRpos + 1 Then
    DataFileLineEnd = vbCrLf
  ElseIf LFpos > 0 And (LFpos < CRpos Or CRpos = 0) Then
    DataFileLineEnd = vbLf
  ElseIf CRpos > 0 Then
    DataFileLineEnd = vbCr
  Else
    MsgBox "Could not find CR or LF in file, so could not break it into lines", vbOKOnly, "DataFileLineEnd"
  End If
  LenDataFileLineEnd = Len(DataFileLineEnd)
End Sub

Private Sub agd_CommitChange(ChangeFromRow As Long, ChangeToRow As Long, ChangeFromCol As Long, ChangeToCol As Long)
  Dim i&
  With agd
    Select Case .TextMatrix(2, ChangeFromCol)
      Case "Station ID":
        .TextMatrix(4, ChangeFromCol) = ""
        .ColWidth(ChangeFromCol) = 1000
      Case "", "Descriptive Information":
        .TextMatrix(4, ChangeFromCol) = ""
      Case Else:
        If .row = 2 And .TextMatrix(4, ChangeFromCol) = "" Then
          .TextMatrix(4, ChangeFromCol) = "No"
        End If
    End Select
    If LCase(.TextMatrix(ChangeFromRow, ChangeFromCol)) = "latitude" Then
      MsgBox "Coordinates assumed to be [decimal degrees] or [MMDDSS]." & vbCrLf & vbCrLf & _
          "If format is otherwise, change import file then attempt import again.", , "Coordinate Type"
    End If
    If .row = 3 Then .ColWidth(ChangeFromCol) = 2000
    .ColWidth(0) = 1400
  End With
End Sub

Private Sub agd_RowColChange()
  Dim i As Long
  Dim statTypeCode As String
  
  If LenDataFileString = 0 Then OpenDataFile
  With agd
    .ClearValues
    Select Case .row
      Case 2:                                 'Fill list of Statistic Types
        .addValue "Station ID"
        .addValue ""
        For i = 1 To SSDB.StatisticTypes.Count
          .addValue SSDB.StatisticTypes(i).Name
        Next i
        .ComboCheckValidValues = True
        .ColEditable(.col) = True
      Case 3:                                 'Fill list of Stat Names
        Select Case LCase(Trim(.TextMatrix(2, .col)))
          Case "", "station id":
            .ColEditable(.col) = False
          Case Else
            .addValue ""
            statTypeCode = GetStatTypeCode(.TextMatrix(2, .col))
            For i = 1 To SSDB.StatisticTypes(statTypeCode).StatLabels.Count
              .addValue SSDB.StatisticTypes(statTypeCode).StatLabels(i).Name
            Next i
            .ComboCheckValidValues = True
            .ColEditable(.col) = True
        End Select
      Case 4:
        Select Case LCase(.TextMatrix(2, .col))
          Case "", "station id", "descriptive information":
            .ColEditable(.col) = False
          Case Else:
            .addValue "Yes"
            .addValue "No"
            .ComboCheckValidValues = True
            .ColEditable(.col) = True
        End Select
      Case Else
        .ColEditable(.col) = False
    End Select
  End With
End Sub

Private Sub cmdFirstRow_Click()
  If SkipDataLinesStart > 0 Then
    SkipDataLinesStart = 0
    cmdFirstRow.Caption = "Ignore Lines Before Selection"
  ElseIf agd.row > FirstDataRow Then
    SkipDataLinesStart = agd.row - FirstDataRow
    cmdFirstRow.Caption = "Un-Ignore First Lines"
  End If
  DiscoverFormat
  PopulateSample
End Sub

Private Sub cmdLastRow_Click()
  If SkipDataLinesEnd > 0 Then
    SkipDataLinesEnd = 0
    cmdLastRow.Caption = "Ignore Lines After Selection"
  ElseIf agd.row > FirstDataRow Then
    SkipDataLinesEnd = DataFileLines - (agd.row - FirstDataRow) - SkipDataLinesStart - 1
    cmdLastRow.Caption = "Un-Ignore Last Lines"
  End If
  DiscoverFormat
  PopulateSample
End Sub

Private Sub cmdOK_Click()
  Import
  Unload Me
End Sub

Private Sub Form_Resize()
  Me.Caption = "Import station data for " & SSDB.state.Name
  fraButtons.Left = (Me.ScaleWidth - fraButtons.Width) / 2
  fraButtons.Top = Me.ScaleHeight - fraButtons.height * 1.5
  If fraButtons.Top > fraButtons.height Then
    agd.height = fraButtons.Top - fraButtons.height / 2
  End If
  agd.Width = Me.ScaleWidth
End Sub

Sub Import()
  Dim staCnt&, attCnt&, i&, staIndex&, attIndex&
  Dim textLine$, str$, badFields$, badStats$
  Dim stationValues() As String, dataValues() As String
  Dim myRegion As nssRegion
  Dim myStation As ssStation
  Dim myStatistic As ssStatistic
  Dim myCitation As New ssSource
  Dim errorFilename As String
  Dim fileline As Long
  Dim parsed() As String
  Dim takeLog() As Boolean
  Dim pcols As Long, col As Long
  Dim roiRegnCount As Long
  Dim attribAbbrev As String

  On Error GoTo errTrap

  IPC.SendMonitorMessage "(OPEN StreamStatsDB)"
  IPC.SendMonitorMessage "(BUTTOFF DETAILS)"
  IPC.SendMonitorMessage "(BUTTON CANCEL)"
  IPC.SendMonitorMessage "(BUTTON PAUSE)"
  IPC.SendMonitorMessage "(MSG1 Importing Station file)"
  IPC.SendMonitorMessage "(PROGRESS 0)"
  
'  Set myCitation.Db = SSDB
'  myCitation.Add "Imported from station file " & FilenameNoPath(NameDataFile)
  roiRegnCount = 0
  Set myStatistic = New ssStatistic
  Set myStatistic.Db = SSDB

  For fileline = 1 + SkipDataLinesStart To DataFileLines - SkipDataLinesEnd - SkipDataLinesStart
    textLine = DataFileLine(fileline)
    'Progress bar
    IPC.SendMonitorMessage "(PROGRESS " & _
          (fileline - SkipDataLinesStart) * 100 / _
          (DataFileLines - SkipDataLinesEnd - SkipDataLinesStart) & ")"
    'Cancel/Pause
    Select Case NextPipeCharacter(IPC.hPipeReadFromProcess("Status"))
      Case "P"
        While NextPipeCharacter(IPC.hPipeReadFromProcess("Status")) <> "R"
          DoEvents
        Wend
      Case "C"
        MsgBox "The station import operation has been interrupted." & _
            vbCrLf & vbCrLf & "Data has been imported for stations in " & _
            SSDB.state.Name & " up to station " & stationValues(2, 1, 1) & ".", , _
            "Import interrupted"
        Err.Raise 999
    End Select
    
    'Read in data from current line
    pcols = ParseInputLine(textLine, parsed)
    If fileline = 1 + SkipDataLinesStart Then  'on first line
      ReDim takeLog(1 To pcols)
      With agd
        For col = 1 To pcols
          If UCase(.TextMatrix(4, col)) = "YES" Then
            takeLog(col) = True
          Else
            takeLog(col) = False
          End If
        Next col
      End With
    End If
    attCnt = 0
    ReDim stationValues(2 To 2, 1 To 1, 1 To UBound(StationFields))
    ReDim dataValues(2 To 2, 1 To pcols - 1, 1 To UBound(DataFields))
    For col = 1 To pcols
      If Len(parsed(col)) > 0 Then
        Select Case LCase(Trim(agd.TextMatrix(2, col)))
          Case "":             'Ignore columns with no category
          Case "station id":
            stationValues(2, 1, 1) = parsed(col)
            If Len(stationValues(2, 1, 1)) = 7 Then stationValues(2, 1, 1) = "0" & stationValues(2, 1, 1)
          Case "Descriptive Information":
            Select Case LCase(Trim(agd.TextMatrix(3, col)))
              Case "station_name": stationValues(2, 1, 2) = parsed(col)
              Case "station_type": stationValues(2, 1, 3) = parsed(col)
              Case "regulated": stationValues(2, 1, 4) = parsed(col)
              Case "period_of_record": stationValues(2, 1, 5) = parsed(col)
              Case "remarks": stationValues(2, 1, 6) = parsed(col)
              Case "latitude":
                If parsed(col) > 2000000 Then
                  MsgBox "The coordinates for Latitude and Longitude are not " & _
                      "in [decimal degrees] or [DDMMSS]." & vbCrLf & vbCrLf & _
                      "This import is being aborted." & vbCrLf & _
                      "Correct the import file and attempt the import again.", _
                      vbCritical, "Bad Coordinates!"
                  GoTo errTrap
                ElseIf parsed(col) > 360 Then 'convert DDMMSS to decimal
                  stationValues(2, 1, 7) = DMS2Decimal(parsed(col))
                Else
                  stationValues(2, 1, 7) = parsed(col)
                End If
              Case "longitude":
                If parsed(col) > 2000000 Then
                  MsgBox "The coordinates for Latitude and Longitude are not " & _
                      "in [decimal degrees] or [DDMMSS]." & vbCrLf & vbCrLf & _
                      "This import is being aborted." & vbCrLf & _
                      "Correct the import file and attempt the import again.", _
                      vbCritical, "Bad Coordinates!"
                  GoTo errTrap
                ElseIf parsed(col) > 360 Then 'convert DDMMSS to decimal
                  stationValues(2, 1, 8) = DMS2Decimal(parsed(col))
                Else
                  stationValues(2, 1, 8) = parsed(col)
                End If
              Case "hydrologic_unit_number": stationValues(2, 1, 9) = parsed(col)
              Case "drainage_basin_code": stationValues(2, 1, 10) = parsed(col)
              Case "county_code": stationValues(2, 1, 11) = parsed(col)
              Case "mcd_code": stationValues(2, 1, 12) = parsed(col)
              Case "directions": stationValues(2, 1, 13) = parsed(col)
              Case "state_code":  'already being stored
              Case Else: GoTo Stat
            End Select
          Case Else
Stat:
            attribAbbrev = agd.TextMatrix(3, col)
            If Len(attribAbbrev) > 0 Then
              attCnt = attCnt + 1
              'If multiple regions then check for new region
              If attribAbbrev = "ROI_Region_ID" Then  'multiple regions
                If roiRegnCount = 0 Then  'first region
                  roiRegnCount = 1
                  ReDim ROIImportRegnIDs(1 To 1)
                  ROIImportRegnIDs(1) = parsed(col)
                Else
                  For roiRegnCount = 1 To UBound(ROIImportRegnIDs)
                    If parsed(col) = ROIImportRegnIDs(roiRegnCount) Then Exit For
                  Next
                  If roiRegnCount > UBound(ROIImportRegnIDs) Then  'add new region
                    ReDim Preserve ROIImportRegnIDs(1 To roiRegnCount)
                    ROIImportRegnIDs(roiRegnCount) = parsed(col)
                  End If
                End If
              End If
              dataValues(2, attCnt, 2) = GetLabelID(attribAbbrev, SSDB)
              dataValues(2, attCnt, 3) = attribAbbrev
              If takeLog(col) Then
                dataValues(2, attCnt, 4) = CSng(10 ^ parsed(col))
              Else
                dataValues(2, attCnt, 4) = parsed(col)
              End If
              dataValues(2, attCnt, 7) = "Imported from ROI file"  '"Imported from station file " & FilenameNoPath(NameDataFile)
            End If
        End Select
      End If
    Next
    
    If fileline = DataFileLines - SkipDataLinesEnd - SkipDataLinesStart Then  'Add ROI region(s) to database
      If roiRegnCount = 0 Then  'no regions specified; assume single statewide region
        Set myRegion = New nssRegion
        Set myRegion.Db = SSDB
        myRegion.Add True, "ROI_Statewide", False, 0, 0, 1, True
      Else
        SortIntegerArray 1, UBound(ROIImportRegnIDs), ROIImportRegnIDs, ROIImportRegnIDs
        MsgBox "The import file had " & UBound(ROIImportRegnIDs) & " ROI regions identified only by indices." & vbCrLf & "Fill out the following form to assign names to these ROI regions.", , "ROI Region Names"
        frmROIImportRegions.Show vbModal, Me
        For roiRegnCount = 1 To UBound(ROIImportRegnIDs)
          Set myRegion = New nssRegion
          Set myRegion.Db = SSDB
          myRegion.Add True, ROIImportRegnNames(roiRegnCount), False, 0, 0, ROIImportRegnIDs(roiRegnCount), True
        Next
      End If
    End If
    If Len(stationValues(2, 1, 1)) > 0 Then 'If there is a station ID, try adding to database
      ImportedNewData = True
      staCnt = staCnt + 1
      staIndex = SSDB.state.Stations.IndexFromKey(stationValues(2, 1, 1))
      If staIndex = -1 Then 'Station does not exist - add it and its statistics
        Set myStation = New ssStation
        Set myStation.Db = SSDB
        Set myStation.state = SSDB.state
        myStation.IsROI = True
        myStation.Add stationValues(), -staCnt, 1 'use -stacnt as ROI station index
        myStation.id = stationValues(2, 1, 1)
        For i = 1 To attCnt
          Set myStatistic = New ssStatistic
          With myStatistic
            Set .Db = SSDB
            Set .station = myStation
            .Add dataValues(), i, 1
            If dataValues(2, i, 2) = "bad" Then
              badFields = badFields & vbCrLf & "Numeric identifier " & str & _
                  " for station " & myStation.id & "."
            End If
          End With
        Next i
      ElseIf staIndex > 0 Then 'Station exists - check whether its stats also exist
        Set myStation = SSDB.state.Stations(staIndex)
        'make sure station is marked for ROI
        myStation.IsROI = True
        'this Add will only update the ROI field on the Station State table
        myStation.Add stationValues(), -staCnt, 1 'use -stacnt as ROI station index
        For i = 1 To attCnt
          attIndex = myStation.Statistics.IndexFromKey(dataValues(2, i, 2))
          If attIndex = -1 Then 'Statistic does not exist
            Set myStatistic = New ssStatistic
            With myStatistic
              Set .Db = SSDB
              Set .station = myStation
              .Add dataValues(), i, 1
            End With
          ElseIf attIndex > 0 Then 'Statistic already exists - write to text file
            badStats = badStats & vbCrLf & vbTab & _
                Left(myStation.id & "       ", 15) & vbTab & _
                Left(dataValues(2, i, 2) & "      ", 9) & vbTab & _
                Left(myStation.Statistics(attIndex).value & "        ", 8) & _
                vbTab & dataValues(2, i, 4)
          End If
        Next i
      End If
    End If
  Next
errTrap:
  If Err.Number = 999 Then IPC.SendMonitorMessage _
      "(MSG1 User Canceled Import)"
  IPC.SendMonitorMessage "(CLOSE)"
  
  If Len(Trim(badStats)) > 0 Or Len(Trim(badFields)) > 0 Then
    If Len(Trim(badStats)) > 0 Then
      str = "This file contains a list of the attribute values"
      str = str & vbCrLf & "imported from the station file: "
      str = str & vbCrLf & "'" & NameDataFile & "' that already had values"
      str = str & vbCrLf & "stored in the database for " & SSDB.state.Name & "."
      str = str & vbCrLf & "These values were not imported to the database, but are included"
      str = str & vbCrLf & "below along with the pre-existing values for informational purposes."
      str = str & vbCrLf & vbCrLf
      str = str & vbCrLf & vbTab & "               " & vbTab & "       " & vbTab & "Existing" & vbTab & "BCF  "
      str = str & vbCrLf & vbTab & "Station ID     " & vbTab & "Stat ID" & vbTab & "Value   " & vbTab & "Value"
      str = str & vbCrLf & badStats
    End If
    If Len(Trim(badFields)) > 0 Then
      str = str & vbCrLf & vbCrLf & vbCrLf
      str = str & vbCrLf & "One or more attributes were not imported " & _
          "because they contained unrecognized numeric identifiers."
      str = str & vbCrLf & badFields
    End If
    errorFilename = ExePath & "ImportErrors-" & SSDB.state.Abbrev & ".txt"
    SaveFileString errorFilename, str
    OpenFile errorFilename
  End If
End Sub

'Private Sub LLUTMLL(ALAT, ALON, UE, UN, N, IZONE, IERR)
'
''SUBROUTINE CONVERTS LATITUDE-LONGITUDE TO UNIVERSAL
''TRANSVERSE MERCATOR COORDINATES AND VICE VERSA.
''ENTRY LLUTM CONVERTSLAT-LNG TO UTM.  ENTRY UTMLL
''CONVERTS UTM TO LAT-LNG.
'
'  'SET UP COEFFICIENTS FOR CONVERTING GEODETIC TO RECTIFYING LATITUDE
'  'AND CONVERSELY  PROGRAM NO. B836 BY DAVID HANDWERKER, U.S.G.S.
'  Dim A(16) As Double, B(12) As Double, FAC As Double, RN As Double, T As Double, _
'      TS As Double, ETAS As Double, SLAT As Double, SLON As Double, X As Double, _
'      Y As Double, SINP As Double, COSP As Double, SINW As Double, COSW As Double
'  Dim tmpZone As Long, ifl As Long
'
''  'ConvFlag = 0 --> LatLng to UTM
''  'ConvFlag = 1 --> UTM to LatLng
''  Select Case longitudeDegrees
''    Case 66 To 71: IZone = 19
''    Case 72 To 77: IZone = 18
''    Case 78 To 83: IZone = 17
''    Case 84 To 89: IZone = 16
''    Case 90 To 95: IZone = 15
''    Case 96 To 101: IZone = 14
''    Case 102 To 107: IZone = 13
''    Case 108 To 113: IZone = 12
''    Case 114 To 119: IZone = 11
''    Case 120 To 125: IZone = 10
''  End Select
'
'  A(1) = 0#
'  A(2) = 0#
'  A(3) = 0#
'  A(4) = 0#
'  A(5) = 500000#  'central false easting in each zone (meters)
'  A(6) = 0#
'  A(7) = 0#
'  A(8) = 0.9996
'  A(9) = 0#
'  A(10) = 0#
'  A(11) = 0#
'  A(12) = 0#
'  A(13) = 0#
'  A(14) = 0#
'  A(15) = 6378206.4
'  A(16) = 0.006768657997291
'  tmpZone = 0
'  ifl = 1
'
'100:
'  If (InZone = tmpZone) Then GoTo 102
'
'101:
'  tmpZone = InZone
'  'SET UP FOR CLARKE 1866 SPHEROID IN METERS
'  A(10) = (((A(16) * (7# / 32#) + (5# / 16#)) * A(16) + 0.5) * A(16) + 1#) * A(16) * 0.25
'  A(1) = -(((A(10) * (195# / 64#) + 3.25) * A(10) + 3.75) * A(10) + 3#) * A(10)
'  A(2) = (((1455# / 32#) * A(10) + (70# / 3#)) * A(10) + 7.5) * A(10) ^ 2
'  A(3) = -((70# / 3#) + A(10) * (945# / 8#)) * A(10) ^ 3
'  A(4) = (315# / 4#) * A(10) ^ 4
'  A(11) = (((7.75 - (657# / 64#) * A(10)) * A(10) - 5.25) * A(10) + 3#) * A(10)
'  A(12) = (((5045# / 32#) * A(10) - (151# / 3#)) * A(10) + 10.5) * A(10) ^ 2
'  A(13) = ((151# / 3#) - (3291# / 8#) * A(10)) * A(10) ^ 3
'  A(14) = (1097# / 4#) * A(10) ^ 4
'  'A(1) TO A(4) ARE FOR GEODETIC TO RECTIFYING LATITUDE CONVERSION WHILE
'  'A(11) TO A(14) ARE COEFFICIENTS FOR RECTIFYING TO GEODETIC CONVERSION
'  FAC = A(10) * A(10)
'  A(10) = (((225# / 64#) * FAC + 2.25) * FAC + 1#) * (1# - FAC) * (1# - A(10)) * A(15)
'  'A(10) IS NOW SET TO RADIUS OF SPHERE WITH GREAT CIRCLE LENGTH EQUAL
'  'TO SPHEROID MERIDIAN LENGTH
'  A(9) = (InZone * 6# + 3#) * 3600#
'
'102:
'  If ifl = 1 Then
'    GoTo 103
'  ElseIf ifl = 2 Then
'    GoTo 200
'  ElseIf ifl = 3 Then
'    GoTo 300
'  End If
'
'103:
'  Return
'  ENTRY LLUTM(ALAT, ALON, UE, UN, N, InZone, IERR)
'  'CONVERTS LATITUDE AND LONGITUDE IN SECONDS (SLAT AND SLON) TO X AND
'  'Y ON TRANSVERSE MERCATOR PROJECTION. A(1) TO A(4) ARE COEFFICIENTS
'  'USED TO CONVERT GEODETIC LATITUDE TO RECTIFYING LATITUDE, A(5) IS
'  'FALSE EASTING, A(6) IS FALSE NORTHING, A(7) IS MERIDIONAL DISTANCE OF
'  'ORIGIN FROM EQUATOR, A(8) IS SCALE FACTOR, A(9) IS CENTRAL MERIDIAN
'  'IN SECONDS, A(10) IS RADIUS OF SPHERE HAVING GREAT CIRCLE LENGTH
'  'EQUAL TO SPHEROID MERIDIAN LENGTH, A(11) TO A(14) ARE COEFFICIENTS
'  'TO CONVERT RECTIFYING LATITUDE TO GEODETIC LATITUDE, A(15) IS
'  'SEMIMAJOR AXIS OF SPHEROID, AND A(16) IS SQUARE OF ECCENTRICITY.
'  'IERR SET TO 1 IF LAT EXCEEDS 80.5 DEGREES OR LONG EXCEEDS 0.1 RADIANS
'  ifl = 2
'  GoTo 100
'200:
'  IERR = 0
'  For i = 1 To N
'    SLAT = ALAT(i) * 3600#
'    SLON = ALON(i) * 3600#
'    If (DABS(SLAT) > 289800#) Then IERR = 1
'    IERR = 1
'    'IF LATITUDE MAGNITUDE EXCEEDS 80.5 DEGREES OR ABSOLUTE VALUE OF
'    'LONGITUDE DIFFERENCE EXCEEDS 0.1 RADIAN, EXIT WITH X = Y = 0.0.
'    B(10) = (SLON - A(9)) * 4.84813681109536E-06
'    If (DABS(B(10)) > 0.1) Then IERR = 2
'    B(9) = SLAT * 4.84813681109536E-06
'    SINP = DSIN(B(9))
'    COSP = DCOS(B(9))
'    RN = A(15) / DSQRT(1# - A(16) * SINP * SINP)
'    T = SINP / COSP
'    TS = T * T
'    B(11) = COSP * COSP
'    ETAS = A(16) * B(11) / (1# - A(16))
'    B(1) = RN * COSP
'    B(3) = (1# - TS + ETAS) * B(1) * B(11) / 6#
'    B(5) = ((TS - 18#) * TS + 5# + (14# - 58# * TS) * ETAS) * B(1) * B(11) * B(11) / 120#
'    B(7) = (((179# - TS) * TS - 479#) * TS + 61#) * B(1) * B(11) ^ 3 / 5040#
'    B(12) = B(10) * B(10)
'    X = -(((B(7) * B(12) + B(5)) * B(12) + B(3)) * B(12) + B(1)) * B(10) * A(8) + A(5)
'    B(2) = RN * B(11) * T / 2#
'    B(4) = (ETAS * (9# + 4# * ETAS) + 5# - TS) * B(2) * B(11) / 12#
'    B(6) = ((TS - 58#) * TS + 61# + (270# - 330# * TS) * ETAS) * B(2) * B(11) * B(11) / 360#
'    B(8) = (((543# - TS) * TS - 3111#) * TS + 1385#) * B(2) * B(11) ^ 3 / 20160#
'    Y = (((B(8) * B(12) + B(6)) * B(12) + B(4)) * B(12) + B(2)) * B(12) + ((((A(4) * B(11) + A(3)) * B(11) + A(2)) * B(11) + A(1)) * SINP * COSP + B(9)) * A(10)
'    Y = (Y - A(7)) * A(8) + A(6)
'    UE(i) = X
'    UN(i) = Y
'  Next i
'  Return
'  ENTRY UTMLL(ALAT, ALON, UE, UN, N, InZone, IERR)
'  'COMPUTES LATITUDE AND LONGITUDE IN SECONDS (SLAT AND SLON) FROM GIVEN
'  'RECTANGULAR COORDINATES X AND Y FOR TRANSVERSE MERCATOR PROJECTION.
'  'A IS ARRAY OF PARAMETERS USED IN COMPUTATION, DESCRIBED BY COMMENTS
'  'FOR TMFWD SUBROUTINE. IERR SET TO 1 IF GRID DISTANCE FROM CENTRAL
'  'MERIDIAN EXCEEDS 0.1 OF SPHEROID SEMIMAJOR AXIS NUMERICALLY OR IF
'  'ABSOLUTE VALUE OF RECTIFYING LATITUDE EXCEEDS 1.4 RADIANS. SOUTH
'  'LATITUDES AND EAST LONGITUDE ARE NEGATIVE.
'  ifl = 3
'  GoTo 100
'300:
'  IERR = 0
'  For i = 1 To N
'    X = UE(i)
'    Y = UN(i)
'    B(9) = ((A(5) - X) * 0.000001) / A(8)
'    If (DABS(B(9)) > 0.0000001 * A(15)) Then IERR = 1
'    IERR = 1
'    B(10) = ((Y - A(6)) / A(8) + A(7)) / A(10)
'    If (DABS(B(10)) > 1.4) Then IERR = 2
'    SINW = DSIN(B(10))
'    COSW = DCOS(B(10))
'    B(12) = COSW * COSW
'    B(11) = (((A(14) * B(12) + A(13)) * B(12) + A(12)) * B(12) + A(11)) * SINW * COSW + B(10)
'    SINW = DSIN(B(11))
'    COSW = DCOS(B(11))
'    RN = DSQRT(1# - A(16) * SINW * SINW) * 1000000# / A(15)
'    T = SINW / COSW
'    TS = T * T
'    B(12) = COSW * COSW
'    ETAS = A(16) * B(12) / (1# - A(16))
'    B(1) = RN / COSW
'    B(2) = -T * (1# + ETAS) * RN * RN / 2#
'    B(3) = -(1# + 2# * TS + ETAS) * B(1) * RN * RN / 6#
'    B(4) = (((-6# - ETAS * 9#) * ETAS + 3#) * TS + (6# - ETAS * 3#) * ETAS + 5#) * T * RN ^ 4 / 24#
'    B(5) = ((TS * 24# + ETAS * 8# + 28#) * TS + ETAS * 6# + 5#) * B(1) * RN ^ 4 / 120#
'    B(6) = (((ETAS * 45# - 45#) * TS + ETAS * 162# - 90#) * TS - ETAS * 107# - 61#) * T * RN ^ 6 / 720#
'    B(7) = -(((TS * 720# + 1320#) * TS + 662#) * TS + 61#) * B(1) * RN ^ 6 / 5040#
'    B(8) = (((TS * 1575# + 4095#) * TS + 3633#) * TS + 1385#) * T * RN ^ 8 / 40320#
'    B(10) = B(9) * B(9)
'    SLAT = ((((B(8) * B(10) + B(6)) * B(10) + B(4)) * B(10) + B(2)) * B(10) + B(11)) * 206264.806247096
'    SLON = (((B(7) * B(10) + B(5)) * B(10) + B(3)) * B(10) + B(1)) * B(9) * 206264.806247096 + A(9)
'    ALAT(i) = SLAT / 3600#
'    ALON(i) = SLON / 3600#
'  Next i
'End Sub
