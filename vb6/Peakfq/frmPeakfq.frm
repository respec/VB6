VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "*\A..\ATCoCtl\ATCoCtl.vbp"
Begin VB.Form frmPeakfq 
   Caption         =   "PKFQWin"
   ClientHeight    =   5430
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13275
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPeakfq.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   13275
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   6000
      TabIndex        =   29
      Top             =   4920
      Width           =   5175
      Begin VB.CommandButton cmdRun 
         Caption         =   "&Run PEAKFQ"
         Height          =   375
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Specs"
         Height          =   375
         Left            =   1800
         TabIndex        =   31
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   4320
         TabIndex        =   30
         Top             =   0
         Width           =   855
      End
   End
   Begin TabDlg.SSTab sstPfq 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6588
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Station Specifications"
      TabPicture(0)   =   "frmPeakfq.frx":01CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grdSpecs"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Output Options"
      TabPicture(1)   =   "frmPeakfq.frx":01E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraOutFile"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraAddOut"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraOutRight"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Results"
      TabPicture(2)   =   "frmPeakfq.frx":0202
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraOutFileRes(1)"
      Tab(2).Control(1)=   "fraGraphics"
      Tab(2).Control(2)=   "fraOutFileRes(0)"
      Tab(2).Control(3)=   "grdGraphs"
      Tab(2).ControlCount=   4
      Begin ATCoCtl.ATCoGrid grdGraphs 
         Height          =   2295
         Left            =   -66960
         TabIndex        =   33
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   4048
         SelectionToggle =   -1  'True
         AllowBigSelection=   -1  'True
         AllowEditHeader =   0   'False
         AllowLoad       =   0   'False
         AllowSorting    =   0   'False
         Rows            =   2
         Cols            =   1
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
         SelectionMode   =   1
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
      Begin VB.Frame fraOutRight 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   8040
         TabIndex        =   21
         Top             =   600
         Width           =   2895
         Begin VB.OptionButton optGraphFormat 
            Caption         =   "WMF"
            Height          =   255
            Index           =   4
            Left            =   1200
            TabIndex        =   39
            Top             =   1600
            Width           =   975
         End
         Begin VB.OptionButton optGraphFormat 
            Caption         =   "PS"
            Height          =   255
            Index           =   3
            Left            =   2160
            TabIndex        =   38
            Top             =   1320
            Width           =   735
         End
         Begin VB.OptionButton optGraphFormat 
            Caption         =   "CGM"
            Height          =   255
            Index           =   2
            Left            =   1200
            TabIndex        =   37
            Top             =   1320
            Width           =   975
         End
         Begin VB.OptionButton optGraphFormat 
            Caption         =   "BMP"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   36
            Top             =   1600
            Width           =   975
         End
         Begin VB.OptionButton optGraphFormat 
            Caption         =   "None"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   35
            Top             =   1320
            Width           =   975
         End
         Begin VB.CheckBox chkPlotPos 
            Caption         =   "Print Plotting Positions"
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   360
            Width           =   2535
         End
         Begin VB.CheckBox chkLinePrinter 
            Caption         =   "Line Printer Plots"
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   720
            Width           =   2775
         End
         Begin VB.CheckBox chkIntRes 
            Caption         =   "Output Intermediate Results"
            Height          =   255
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   2895
         End
         Begin ATCoCtl.ATCoText txtCL 
            Height          =   255
            Left            =   1800
            TabIndex        =   25
            Top             =   2400
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            InsideLimitsBackground=   16777215
            OutsideHardLimitBackground=   8421631
            OutsideSoftLimitBackground=   8454143
            HardMax         =   0.995
            HardMin         =   0.5
            SoftMax         =   0.995
            SoftMin         =   0.5
            MaxWidth        =   -999
            Alignment       =   1
            DataType        =   2
            DefaultValue    =   0.95
            Value           =   "0.95"
            Enabled         =   -1  'True
         End
         Begin ATCoCtl.ATCoText txtPlotPos 
            Height          =   255
            Left            =   1800
            TabIndex        =   26
            Top             =   2040
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   450
            InsideLimitsBackground=   16777215
            OutsideHardLimitBackground=   8421631
            OutsideSoftLimitBackground=   8454143
            HardMax         =   0.5
            HardMin         =   0
            SoftMax         =   0.5
            SoftMin         =   0
            MaxWidth        =   -999
            Alignment       =   1
            DataType        =   2
            DefaultValue    =   0
            Value           =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.Label lblGraphics 
            Caption         =   "Graphic Plot Format"
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Label lblPlotPos 
            Caption         =   "Plotting Position:"
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblCL 
            Caption         =   "Confidence Limits:"
            Height          =   255
            Left            =   0
            TabIndex        =   27
            Top             =   2400
            Width           =   1815
         End
      End
      Begin VB.Frame fraOutFileRes 
         Caption         =   "Output File"
         Height          =   1215
         Index           =   0
         Left            =   -74760
         TabIndex        =   15
         Top             =   600
         Width           =   7455
         Begin VB.CommandButton cmdOutFileView 
            Caption         =   "View"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblOutFileView 
            Caption         =   "(none)"
            Height          =   615
            Index           =   0
            Left            =   1200
            TabIndex        =   17
            Top             =   360
            Width           =   6135
         End
      End
      Begin VB.Frame fraGraphics 
         Caption         =   "Graphs"
         Height          =   2775
         Left            =   -66840
         TabIndex        =   13
         Top             =   600
         Width           =   2775
         Begin VB.CommandButton cmdGraph 
            Caption         =   "View"
            Height          =   255
            Left            =   1080
            TabIndex        =   40
            Top             =   2400
            Width           =   855
         End
         Begin VB.ListBox lstGraphs 
            Height          =   2010
            Left            =   120
            MultiSelect     =   1  'Simple
            TabIndex        =   14
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fraOutFileRes 
         Caption         =   "Additional Output"
         Height          =   1215
         Index           =   1
         Left            =   -74760
         TabIndex        =   10
         Top             =   2160
         Width           =   7455
         Begin VB.CommandButton cmdOutFileView 
            Caption         =   "View"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblOutFileView 
            Caption         =   "(none)"
            Height          =   615
            Index           =   1
            Left            =   1200
            TabIndex        =   12
            Top             =   360
            Width           =   6135
         End
      End
      Begin VB.Frame fraAddOut 
         Caption         =   "Additional Output"
         Height          =   1575
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   7575
         Begin VB.OptionButton optAddFormat 
            Caption         =   "Tab-Delimited"
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   42
            Top             =   480
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.OptionButton optAddFormat 
            Caption         =   "Watstore"
            Height          =   255
            Index           =   0
            Left            =   2880
            TabIndex        =   41
            Top             =   240
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CheckBox chkAddOut 
            Caption         =   "WDM"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
         Begin VB.CheckBox chkAddOut 
            Caption         =   "Text File"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   7
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpenOut 
            Caption         =   "Select"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblOutFile 
            Caption         =   "(none)"
            Height          =   735
            Index           =   1
            Left            =   1080
            TabIndex        =   9
            Top             =   720
            Width           =   5895
         End
      End
      Begin VB.Frame fraOutFile 
         Caption         =   "Output File"
         Height          =   1215
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   7575
         Begin VB.CommandButton cmdOpenOut 
            Caption         =   "Select"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblOutFile 
            Caption         =   "(none)"
            Height          =   735
            Index           =   0
            Left            =   1080
            TabIndex        =   4
            Top             =   240
            Width           =   5895
         End
      End
      Begin ATCoCtl.ATCoGrid grdSpecs 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   5530
         SelectionToggle =   0   'False
         AllowBigSelection=   -1  'True
         AllowEditHeader =   0   'False
         AllowLoad       =   0   'False
         AllowSorting    =   0   'False
         Rows            =   1
         Cols            =   2
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
   End
   Begin MSComDlg.CommonDialog cdlOpen 
      Left            =   9960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label lblInstruct 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblSpec 
      Caption         =   "PKFQWin Spec File:"
      Height          =   255
      Left            =   5040
      TabIndex        =   19
      Top             =   480
      Width           =   6135
   End
   Begin VB.Label lblData 
      Caption         =   "PEAKFQ Data File:"
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      Top             =   120
      Width           =   6135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSaveSpecs 
         Caption         =   "&Save Specs"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuFeedback 
         Caption         =   "Send &Feedback"
      End
      Begin VB.Menu mnuHelpMain 
         Caption         =   "PKFQWin Help"
      End
   End
End
Attribute VB_Name = "frmPeakfq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DefaultSpecFile As String
Const tmpSpecName As String = "PKFQWPSF.TMP"
Dim CurGraphName As String
Dim RemoveBMPs As Boolean

Private Sub chkAddOut_Click(Index As Integer)
  If Index = 1 Then 'text file additional output
    If chkAddOut(1).Value = vbChecked Then
      'expand frame to show additional output file and button to edit it
      fraAddOut.Height = 1575
      If Len(PfqPrj.AddOutFileName) = 0 Then 'set default
        lblOutFile(1).Caption = FilenameNoExt(PfqPrj.OutFile) & ".bcd"
      End If
      lblOutFile(1).Visible = vbTrue
      cmdOpenOut(1).Visible = vbTrue
      optAddFormat(0).Visible = vbTrue
      optAddFormat(1).Visible = vbTrue
    Else 'smaller frame is fine
      fraAddOut.Height = 735
      lblOutFile(1).Visible = vbFalse
      cmdOpenOut(1).Visible = vbFalse
      optAddFormat(0).Visible = vbFalse
      optAddFormat(1).Visible = vbFalse
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  Call Form_Unload(0)
End Sub

Private Sub cmdGraph_Click()
  Dim i As Long
  Dim GraphName As String
  Dim newform As frmGraph

'  For i = 1 To grdGraphs.Rows
'    If grdGraphs.Selected(i, 0) Then
'      GraphName = grdGraphs.TextMatrix(i, 0) & ".BMP"
  For i = 0 To lstGraphs.ListCount - 1
    If lstGraphs.Selected(i) Then
      GraphName = lstGraphs.List(i) & ".BMP"
      If FileExists(GraphName) Then
        Set newform = New frmGraph
        newform.Height = 7600
        newform.Width = 9700
        newform.Picture = LoadPicture(GraphName)
        newform.Show
      Else
        MsgBox "No graph available for station " & lstGraphs.List(i) & "." & vbCrLf & _
               "This station was likely skipped due to faulty data - see PeakFQ output file for details.", vbExclamation, "PeakFQ"
      End If
    End If
  Next i
End Sub

Private Sub PopulateGrid()

  Dim i As Long, j As Long, ilen As Long, ipos As Long, Ind As Long
  Dim vSta As Variant

  With grdSpecs
    .Rows = 0
    i = 0
    For Each vSta In PfqPrj.Stations
      i = i + 1
      .TextMatrix(i, 0) = vSta.id
      .col = 0
      .row = i
      .CellBackColor = grdSpecs.BackColorFixed
      If vSta.Active Then
        .TextMatrix(i, 1) = "Yes"
      Else
        .TextMatrix(i, 1) = "No"
      End If
      .TextMatrix(i, 2) = vSta.BegYear
      .TextMatrix(i, 3) = vSta.EndYear
      .TextMatrix(i, 4) = vSta.HistoricPeriod
      If vSta.SkewOpt = -1 Then
        .TextMatrix(i, 5) = "Station"
      ElseIf vSta.SkewOpt = 0 Then
        .TextMatrix(i, 5) = "Weighted"
      ElseIf vSta.SkewOpt = 1 Then
        .TextMatrix(i, 5) = "Generalized"
      End If
      .TextMatrix(i, 6) = vSta.GenSkew
      .TextMatrix(i, 7) = vSta.SESkew
      .TextMatrix(i, 8) = vSta.SESkew ^ 2
      .CellBackColor = .BackColorFixed
      .TextMatrix(i, 9) = vSta.LowHistPeak
      .CellBackColor = .BackColorFixed
      .TextMatrix(i, 10) = vSta.LowOutlier
      .TextMatrix(i, 11) = vSta.HighSysPeak
      .CellBackColor = .BackColorFixed
      .TextMatrix(i, 12) = vSta.HighOutlier
      .TextMatrix(i, 13) = vSta.GageBaseDischarge
      If vSta.UrbanRegPeaks Then
        .TextMatrix(i, 14) = "Yes"
      Else
        .TextMatrix(i, 14) = "No"
      End If
      .TextMatrix(i, 15) = vSta.Lat
      .TextMatrix(i, 16) = vSta.Lng
      .TextMatrix(i, 17) = vSta.PlotName
      ilen = Len(vSta.PlotName)
      For j = i - 1 To 1 Step -1 'look for duplicate plot names and adjust as needed
        If Left(.TextMatrix(j, 17), ilen) = vSta.PlotName Then 'duplicate found
          ipos = InStr(.TextMatrix(j, 17), "-")
          If ipos > 0 Then 'not first duplicate, increase index number
            Ind = CLng(Mid(.TextMatrix(j, 17), ipos + 1))
            .TextMatrix(i, 17) = vSta.PlotName & "-" & CStr(Ind + 1)
          Else 'first duplicate
            .TextMatrix(i, 17) = vSta.PlotName & "-1"
          End If
        End If
      Next j
    Next
    .ColsSizeByContents
  End With

End Sub

Private Sub ProcessGrid()
  
  Dim i As Long
  Dim curSta As pfqStation

  Set PfqPrj.Stations = Nothing
'  lstGraphs.Clear
  For i = 1 To grdSpecs.Rows
    Set curSta = New pfqStation
    curSta.id = grdSpecs.TextMatrix(i, 0)
    If grdSpecs.TextMatrix(i, 1) = "Yes" Then
      curSta.Active = True
'      lstGraphs.AddItem curSta.id
    Else
      curSta.Active = False
    End If
    If IsNumeric(grdSpecs.TextMatrix(i, 2)) Then _
      curSta.BegYear = grdSpecs.TextMatrix(i, 2)
    If IsNumeric(grdSpecs.TextMatrix(i, 3)) Then _
      curSta.EndYear = grdSpecs.TextMatrix(i, 3)
    If IsNumeric(grdSpecs.TextMatrix(i, 4)) Then _
      curSta.HistoricPeriod = grdSpecs.TextMatrix(i, 4)
    If grdSpecs.TextMatrix(i, 5) = "Station" Then
      curSta.SkewOpt = -1
    ElseIf grdSpecs.TextMatrix(i, 5) = "Weighted" Then
      curSta.SkewOpt = 0
    ElseIf grdSpecs.TextMatrix(i, 5) = "Generalized" Then
      curSta.SkewOpt = 1
    End If
    If IsNumeric(grdSpecs.TextMatrix(i, 6)) Then _
      curSta.GenSkew = grdSpecs.TextMatrix(i, 6)
    If IsNumeric(grdSpecs.TextMatrix(i, 7)) Then _
      curSta.SESkew = grdSpecs.TextMatrix(i, 7)
    If IsNumeric(grdSpecs.TextMatrix(i, 10)) Then _
      curSta.LowOutlier = grdSpecs.TextMatrix(i, 10)
    If IsNumeric(grdSpecs.TextMatrix(i, 12)) Then _
      curSta.HighOutlier = grdSpecs.TextMatrix(i, 12)
    If IsNumeric(grdSpecs.TextMatrix(i, 13)) Then _
      curSta.GageBaseDischarge = grdSpecs.TextMatrix(i, 13)
    If grdSpecs.TextMatrix(i, 14) = "Yes" Then
      curSta.UrbanRegPeaks = True
    Else
      curSta.UrbanRegPeaks = False
    End If
    If IsNumeric(grdSpecs.TextMatrix(i, 15)) Then _
      curSta.Lat = grdSpecs.TextMatrix(i, 15)
    If IsNumeric(grdSpecs.TextMatrix(i, 16)) Then _
      curSta.Lng = grdSpecs.TextMatrix(i, 16)
    curSta.PlotName = grdSpecs.TextMatrix(i, 17)
    PfqPrj.Stations.Add curSta
  Next

End Sub

Private Sub PopulateOutput()

  lblOutFile(0).Caption = PfqPrj.OutFile
  If PfqPrj.DataType = 0 Then 'ASCII input, can't output to WDM
    chkAddOut(0).Enabled = False
    chkAddOut(0).Value = vbUnchecked
  Else
    chkAddOut(0).Enabled = True
    If PfqPrj.AdditionalOutput Mod 2 = 1 Then
      chkAddOut(0).Value = vbChecked
    End If
  End If
  If PfqPrj.AdditionalOutput >= 2 Then
    chkAddOut(1).Value = vbChecked
    lblOutFile(1).Caption = PfqPrj.AddOutFileName
    lblOutFile(1).Visible = vbTrue
    cmdOpenOut(1).Visible = vbTrue
    optAddFormat(0).Visible = vbTrue
    optAddFormat(1).Visible = vbTrue
    If PfqPrj.AdditionalOutput < 4 Then 'watstore format
      optAddFormat(0).Value = vbTrue
    Else 'tab-separated format
      optAddFormat(1).Value = vbTrue
    End If
    fraAddOut.Height = 1575
  Else
    chkAddOut(1).Value = vbUnchecked
    lblOutFile(1).Caption = "(none)"
    lblOutFile(1).Visible = vbFalse
    cmdOpenOut(1).Visible = vbFalse
    optAddFormat(0).Visible = vbFalse
    optAddFormat(1).Visible = vbFalse
    fraAddOut.Height = 735
  End If
  If PfqPrj.IntermediateResults Then
    chkIntRes.Value = vbChecked
  Else
    chkIntRes.Value = vbUnchecked
  End If
  If PfqPrj.LinePrinter Then
    chkLinePrinter.Value = vbChecked
  Else
    chkLinePrinter.Value = vbUnchecked
  End If
  If PfqPrj.Graphic Then
    If UCase(PfqPrj.GraphFormat) = "CGM" Then
      optGraphFormat(2).Value = True
    ElseIf UCase(PfqPrj.GraphFormat) = "PS" Then
      optGraphFormat(3).Value = True
    ElseIf UCase(PfqPrj.GraphFormat) = "WMF" Then
      optGraphFormat(4).Value = True
    Else 'use BMP
      optGraphFormat(1).Value = True
    End If
  Else
    optGraphFormat(0).Value = True
  End If
  If PfqPrj.PrintPlotPos Then
    chkPlotPos.Value = vbChecked
  Else
    chkPlotPos.Value = vbUnchecked
  End If
  txtCL.Value = PfqPrj.ConfidenceLimits
  txtPlotPos.Value = PfqPrj.PlotPos
End Sub

Private Sub ProcessOutput()
  Dim i As Integer
  Dim lOutDir As String

  PfqPrj.OutFile = lblOutFile(0).Caption
  lOutDir = PathNameOnly(PfqPrj.OutFile)
  If Len(lOutDir) > 0 And lOutDir <> PfqPrj.InputDir Then PfqPrj.OutputDir = lOutDir
  lblOutFileView(0).Caption = PfqPrj.OutFile
  If chkAddOut(0).Value = vbChecked Then
    PfqPrj.AdditionalOutput = 1
  Else
    PfqPrj.AdditionalOutput = 0
  End If
  If chkAddOut(1).Value = vbChecked Then
    If optAddFormat(0).Value = vbTrue Then 'watstore format
      PfqPrj.AdditionalOutput = PfqPrj.AdditionalOutput + 2
    Else 'tab-separated format
      PfqPrj.AdditionalOutput = PfqPrj.AdditionalOutput + 4
    End If
    PfqPrj.AddOutFileName = lblOutFile(1).Caption
    lblOutFileView(1).Caption = PfqPrj.AddOutFileName
  Else
    PfqPrj.AddOutFileName = ""
    lblOutFileView(1).Caption = "(none)"
  End If
  If chkIntRes.Value = vbChecked Then
    PfqPrj.IntermediateResults = True
  Else
    PfqPrj.IntermediateResults = False
  End If
  If chkLinePrinter.Value = vbChecked Then
    PfqPrj.LinePrinter = True
  Else
    PfqPrj.LinePrinter = False
  End If
  If optGraphFormat(0).Value Then 'no graphics
    PfqPrj.Graphic = False
  Else 'get graphic format
    PfqPrj.Graphic = True
    For i = 1 To 4
      If optGraphFormat(i).Value Then
        PfqPrj.GraphFormat = optGraphFormat(i).Caption
        Exit For
      End If
    Next
  End If
  If chkPlotPos.Value = vbChecked Then
    PfqPrj.PrintPlotPos = True
  Else
    PfqPrj.PrintPlotPos = False
  End If
  PfqPrj.ConfidenceLimits = txtCL.Value
  PfqPrj.PlotPos = txtPlotPos.Value

End Sub

Private Sub cmdOpenOut_Click(Index As Integer)

  On Error GoTo FileCancel
  If Index = 0 Then
    cdlOpen.DialogTitle = "Main PeakFQ Output File"
    cdlOpen.Filter = "PeakFQ Output (*.prt)|*.prt|All Files (*.*)|*.*"
    cdlOpen.filename = PfqPrj.OutFile
  Else 'additional output file
    cdlOpen.DialogTitle = "Additional PeakFQ Output File"
    If optAddFormat(0).Value = vbTrue Then
      cdlOpen.Filter = "Watstore Output (*.bcd)|*.bcd|All Files (*.*)|*.*"
      If Len(PfqPrj.AddOutFileName) = 0 Then 'provide default file name
        PfqPrj.AddOutFileName = FilenameOnly(PfqPrj.DataFileName) & ".bcd"
      End If
    Else
      cdlOpen.Filter = "Tab-delimited Output (*.tab)|*.tab|All Files (*.*)|*.*"
      If Len(PfqPrj.AddOutFileName) = 0 Then 'provide default file name
        PfqPrj.AddOutFileName = FilenameOnly(PfqPrj.DataFileName) & ".tab"
      End If
    End If
    cdlOpen.filename = PfqPrj.AddOutFileName
  End If
  cdlOpen.ShowSave
  If FileExists(cdlOpen.filename) Then 'make sure it's OK to overwrite
    If MsgBox("File exists.  Do you want to overwrite it?", vbYesNo) = vbNo Then GoTo FileCancel
  End If
  lblOutFile(Index).Caption = cdlOpen.filename

FileCancel:
End Sub

Private Sub cmdOutFileView_Click(Index As Integer)

  If Len(lblOutFileView(Index).Caption) > 0 And _
     lblOutFileView(Index).Caption <> "(none)" Then
    Shell Chr(34) & FileViewer & Chr(34) & " " & lblOutFileView(Index).Caption, vbNormalFocus
  Else
    MsgBox "No " & fraOutFileRes(Index).Caption & " is available for viewing.", vbInformation, "PeakFQ"
  End If

End Sub

Private Function FileViewer() As String
  Static Viewer As String
  Dim fun As Long
  
  If Len(Viewer) = 0 Then
    fun = FreeFile(0)
    Open "xxx.txt" For Output As fun
    Viewer = FindAssociatedApplication("xxx.txt")
    Close fun
    Kill "xxx.txt"
  End If
  FileViewer = Viewer

End Function


Private Sub cmdRun_Click()

  Dim i As Long
  Dim s As String

  If Len(PfqPrj.SpecFileName) > 0 Then
    Me.MousePointer = vbHourglass
    lstGraphs.Clear
    ProcessGrid
    ProcessOutput
    s = PfqPrj.SaveAsString
    SaveFileString PfqPrj.SpecFileName, s
    DoEvents
    PfqPrj.RunBatchModel
    DoEvents
    If RemoveBMPs Then
'      For i = 1 To grdGraphs.Rows
'        Kill grdGraphs.TextMatrix(i, 0) & ".BMP"
      For i = 1 To lstGraphs.ListCount
        Kill lstGraphs.List(i - 1) & ".BMP"
      Next i
    End If
    If PfqPrj.Graphic Then
      SetGraphNames
      cmdGraph.Enabled = True
      If UCase(PfqPrj.GraphFormat) <> "BMP" Then
        RemoveBMPs = True
      Else
        RemoveBMPs = False
      End If
    Else
'      grdGraphs.Rows = 0
      lstGraphs.Clear
      cmdGraph.Enabled = False
      RemoveBMPs = False
    End If
    Me.MousePointer = vbDefault
    sstPfq.TabEnabled(2) = True
    sstPfq.Tab = 2
'    cmdSave.Enabled = True
'    mnuSaveSpecs.Enabled = True
  Else
    MsgBox "PeakFQ Specfication or Data File must be opened before viewing results.", vbInformation, "PeakFQ Results"
  End If

End Sub

Private Sub cmdSave_Click()
  SaveSpecFile
End Sub

Private Sub Form_Load()

  Dim i As Long

  lblInstruct.Caption = "Use File menu to Open PeakFQ data or PKFQWin spec file." & vbLf & "Update Station and Output specifications as desired." & vbLf & "Click Run PeakFQ button to generate results."

'  grdSpecs.cols = 16
  For i = 0 To 17
    grdSpecs.ColEditable(i) = False
  Next i
  grdSpecs.TextMatrix(0, 0) = "Station ID"
  grdSpecs.ColType(0) = ATCotxt
  grdSpecs.TextMatrix(-1, 1) = "Include in"
  grdSpecs.TextMatrix(0, 1) = "Analysis?"
  grdSpecs.ColType(1) = ATCotxt
  grdSpecs.TextMatrix(-1, 2) = "Beginning"
  grdSpecs.TextMatrix(0, 2) = "Year"
  grdSpecs.ColType(2) = ATCoInt
  grdSpecs.TextMatrix(-1, 3) = "Ending"
  grdSpecs.TextMatrix(0, 3) = "Year"
  grdSpecs.ColType(3) = ATCoInt
  grdSpecs.TextMatrix(-1, 4) = "Historic"
  grdSpecs.TextMatrix(0, 4) = "Period"
  grdSpecs.ColType(4) = ATCoSng
  grdSpecs.TextMatrix(-1, 5) = "Skew"
  grdSpecs.TextMatrix(0, 5) = "Option"
  grdSpecs.ColType(5) = ATCotxt
  grdSpecs.TextMatrix(-1, 6) = "Generalized"
  grdSpecs.TextMatrix(0, 6) = "Skew"
  grdSpecs.ColType(6) = ATCoSng
  grdSpecs.TextMatrix(-1, 7) = "Gen Skew"
  grdSpecs.TextMatrix(0, 7) = "Std Error"
  grdSpecs.ColType(7) = ATCoSng
  grdSpecs.TextMatrix(-1, 8) = "Mean"
  grdSpecs.TextMatrix(0, 8) = "Sqr Err"
  grdSpecs.ColType(8) = ATCoSng
  grdSpecs.TextMatrix(-1, 9) = "Low Hist"
  grdSpecs.TextMatrix(0, 9) = "Peak"
  grdSpecs.ColType(9) = ATCoSng
  grdSpecs.TextMatrix(-1, 10) = "Lo-Outlier"
  grdSpecs.TextMatrix(0, 10) = "Threshold"
  grdSpecs.ColType(10) = ATCoSng
  grdSpecs.TextMatrix(-1, 11) = "High Sys"
  grdSpecs.TextMatrix(0, 11) = "Peak"
  grdSpecs.ColType(11) = ATCoSng
  grdSpecs.TextMatrix(-1, 12) = "Hi-Outlier"
  grdSpecs.TextMatrix(0, 12) = "Threshold"
  grdSpecs.ColType(12) = ATCoSng
  grdSpecs.TextMatrix(-1, 13) = "Gage Base"
  grdSpecs.TextMatrix(0, 13) = "Discharge"
  grdSpecs.ColType(13) = ATCoSng
  grdSpecs.TextMatrix(-1, 14) = "Urban/Reg"
  grdSpecs.TextMatrix(0, 14) = "Peaks"
  grdSpecs.TextMatrix(0, 15) = "Latitude"
  grdSpecs.ColType(15) = ATCoSng
  grdSpecs.TextMatrix(0, 16) = "Longitude"
  grdSpecs.ColType(16) = ATCoSng
  grdSpecs.TextMatrix(-1, 17) = "Plot"
  grdSpecs.TextMatrix(0, 17) = "Name"
  grdSpecs.ColType(17) = ATCotxt
  grdSpecs.ColsSizeByContents
  
  sstPfq.Tab = 0
  sstPfq.TabEnabled(0) = False
  sstPfq.TabEnabled(1) = False
  sstPfq.TabEnabled(2) = False
  cmdRun.Enabled = False
  cmdSave.Enabled = False
'  grdGraphs.TextMatrix(0, 0) = "Graphs"
'  grdGraphs.ColEditable(0) = True
  grdGraphs.Visible = False
  RemoveBMPs = False

End Sub

Private Sub Form_Resize()
  Dim w As Long, h As Long
  w = Me.ScaleWidth
  h = Me.ScaleHeight
  If h < 5070 And h > 0 Then 'height too small
    Me.Height = Me.Height - h + 5070
  End If
  If w > 7300 Then
'    txtData.Width = w - txtData.Left - sstPfq.Left
'    txtSpec.Width = txtData.Width
    lblData.Width = w - lblData.Left - sstPfq.Left
    lblSpec.Width = lblData.Width
    sstPfq.Width = w - (sstPfq.Left * 2)
    fraButtons.Left = w - fraButtons.Width - 120
    Select Case sstPfq.Tab
      Case 0:   grdSpecs.Width = sstPfq.Width - (grdSpecs.Left * 2)
      Case 1:   fraOutRight.Left = sstPfq.Width - fraOutRight.Width - 120
                fraOutFile.Width = fraOutRight.Left - (fraOutFile.Left * 3)
                fraAddOut.Width = fraOutFile.Width
                lblOutFile(0).Width = fraOutFile.Width - lblOutFile(0).Left - 120
                lblOutFile(1).Width = lblOutFile(0).Width
      Case 2:   fraGraphics.Left = sstPfq.Width - fraGraphics.Width - 120
'                grdGraphs.Left = sstPfq.Width - grdGraphs.Width - 240
'                cmdGraph.Left = grdGraphs.Left + (grdGraphs.Width / 2) - (cmdGraph.Width / 2)
                fraOutFileRes(0).Width = fraGraphics.Left - (fraOutFileRes(0).Left * 3)
                fraOutFileRes(1).Width = fraOutFileRes(0).Width
                lblOutFileView(0).Width = fraOutFileRes(0).Width - lblOutFileView(0).Left - 120
                lblOutFileView(1).Width = lblOutFileView(0).Width
    End Select
  End If
  If h > 5070 Then
    fraButtons.Top = h - fraButtons.Height - 120
    sstPfq.Height = fraButtons.Top - sstPfq.Top - 120
    Select Case sstPfq.Tab
      Case 0: grdSpecs.Height = sstPfq.Height - grdSpecs.Top - 120
      Case 2: fraGraphics.Height = sstPfq.Height - fraGraphics.Top - 120
'              grdGraphs.Height = sstPfq.Height - grdGraphs.Top - cmdGraph.Height - 240
              lstGraphs.Height = fraGraphics.Height - lstGraphs.Top - cmdGraph.Height - 240
              cmdGraph.Top = lstGraphs.Top + lstGraphs.Height + 120
    End Select
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim i As Long

  On Error Resume Next

  If RemoveBMPs Then 'remove BMP graphics
    For i = 1 To lstGraphs.ListCount
      Kill lstGraphs.List(i - 1) & ".BMP"
    Next i
  End If

  gIPC.MonitorEnabled = False
  Set gIPC = Nothing

  End
End Sub

'Private Sub grdGraphs_CommitChange(ChangeFromRow As Long, ChangeToRow As Long, ChangeFromCol As Long, ChangeToCol As Long)
'  Dim NewGraphName As String
'
'  NewGraphName = FilenameNoExt(grdGraphs.TextMatrix(ChangeFromRow, ChangeFromCol)) & ".BMP"
'  RenameGraph CurGraphName, NewGraphName
'End Sub
'
'Private Sub grdGraphs_RowColChange()
'  CurGraphName = FilenameNoExt(grdGraphs.TextMatrix(grdGraphs.row, grdGraphs.col)) & ".BMP"
'End Sub

Private Sub grdSpecs_CommitChange(ChangeFromRow As Long, ChangeToRow As Long, ChangeFromCol As Long, ChangeToCol As Long)
  If ChangeFromCol = 7 Then 'changed std skew err, update mean sqr err
    grdSpecs.TextMatrix(ChangeFromRow, 8) = CSng(grdSpecs.TextMatrix(ChangeFromRow, ChangeFromCol)) ^ 2
    grdSpecs.CellBackColor = grdSpecs.BackColorFixed
  End If
End Sub

Private Sub grdSpecs_RowColChange()

  grdSpecs.ClearValues
  If grdSpecs.col = 1 Or grdSpecs.col = 14 Then
    grdSpecs.addValue "Yes"
    grdSpecs.addValue "No"
  ElseIf grdSpecs.col = 5 Then
    grdSpecs.addValue "Station"
    grdSpecs.addValue "Weighted"
    grdSpecs.addValue "Generalized"
  End If
End Sub

Private Sub lstGraphs_DblClick()
  cmdGraph_Click
End Sub

Private Sub mnuAbout_Click()
  MsgBox "Version " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation, "PKFQWin"
End Sub

Private Sub mnuExit_Click()
  Call Form_Unload(0)
End Sub

Private Sub mnuFeedback_Click()
  Dim stepname As String
  On Error GoTo errmsg
                                       stepname = "1: Dim feedback As clsATCoFeedback"
  Dim feedback As clsATCoFeedback
                                       stepname = "2: Set feedback = New clsATCoFeedback"
  Set feedback = New clsATCoFeedback
  '                                     stepname = "3: feedback.AddText"
  'feedback.AddText AboutString(False)
                                       stepname = "4: feedback.AddFile"
  feedback.AddFile Left(App.path, InStr(4, App.path, "\")) & "unins000.dat"
                                       stepname = "5: feedback.Show"
  feedback.Show App, Me.Icon
  
  Exit Sub
  
errmsg:
  MsgBox "Error opening feedback in step " & stepname & vbCr & Err.Description, _
                         "Send Feedback"
End Sub

Private Sub mnuHelpMain_Click()
  Dim s As String
  s = OpenFile(App.HelpFile, cdlOpen)
End Sub

Private Sub mnuOpen_Click()
  Dim FName As String
  Dim s As String
  On Error GoTo FileCancel

  cdlOpen.DialogTitle = "Open PeakFQ File"
  cdlOpen.Filter = "PeakFQ Watstore Data (*.pkf,*.inp,*.txt)|*.pkf;*.inp;*.txt|PeakFQ Watstore Data (*.*)|*.*|PeakFQ WDM Data (*.wdm)|*.wdm|PKFQWin Spec (*.psf)|*.psf"
  cdlOpen.ShowOpen
  FName = cdlOpen.filename
  PfqPrj.InputDir = PathNameOnly(FName)
  PfqPrj.OutputDir = PathNameOnly(FName) 'default output directory to same as input
  sstPfq.Tab = 0
  sstPfq.TabEnabled(2) = False
  Me.MousePointer = vbHourglass
  DoEvents
  If cdlOpen.FilterIndex <= 3 Then 'open data file
    PfqPrj.DataFileName = FName
    PfqPrj.BuildNewSpecFile 'build basic spec file (I/O files)
    PfqPrj.RunBatchModel 'run model to generate verbose spec file
    PfqPrj.ReadSpecFile 'read verbose spec file
    Set DefPfqPrj = PfqPrj.Copy
  Else 'open spec file
    s = WholeFileString(FName)
    'build default project from initial version of spec file
    SaveFileString tmpSpecName, s
    PfqPrj.SpecFileName = tmpSpecName 'make working verbose copy
    Set DefPfqPrj = PfqPrj.SaveDefaults(s)
  End If
  If FileExists(PfqPrj.OutFile) Then 'delete output file generated from reading data
    Kill PfqPrj.OutFile
  End If
  Me.MousePointer = vbDefault
  If PfqPrj.Stations.Count > 0 Then
'    txtData.Text = PfqPrj.DataFileName
    lblData.Caption = "PeakFQ Data File:  " & PfqPrj.DataFileName
    If cdlOpen.FilterIndex = 4 Then 'opened spec file, put name on main form
'      txtSpec.Text = fname
      lblSpec.Caption = "PKFQWin Spec File:  " & FName
    End If
    EnableGrid
    PopulateGrid
    PopulateOutput
    sstPfq.TabEnabled(0) = True
    sstPfq.TabEnabled(1) = True
    cmdRun.Enabled = True
    cmdSave.Enabled = True
    mnuSaveSpecs.Enabled = True
'    PfqPrj.SpecFileName = tmpSpecName 'use temporary name for active spec file
  End If

FileCancel:
End Sub

Private Sub EnableGrid()
  Dim i As Integer

  For i = 1 To 17
    If i <> 8 And i <> 9 And i <> 11 Then grdSpecs.ColEditable(i) = True
  Next i
'  grdSpecs.ColEditable(0) = False 'station number not editable
'  grdSpecs.ColEditable(8) = False 'low historic peak not editable
'  grdspecs.ColEditable(9) = False 'root mean square error not editable
'  grdSpecs.ColEditable(10) = False 'high historic peak not editable

End Sub

Private Sub mnuSaveSpecs_Click()
  SaveSpecFile
End Sub

Private Sub SaveSpecFile()

  Dim s As String

  On Error GoTo FileCancel
  
  cdlOpen.DialogTitle = "PKFQWin Specification File"
  cdlOpen.Filter = "PKFQWin Spec File (*.psf)|*.psf|All Files (*.*)|*.*"
  If Right(PfqPrj.SpecFileName, 12) = tmpSpecName Then 'no spec file yet
    cdlOpen.filename = FilenameOnly(PfqPrj.DataFileName) & ".psf"
  Else 'use existing spec file as default
    cdlOpen.filename = DefaultSpecFile
  End If
  cdlOpen.ShowSave
  ProcessGrid
  ProcessOutput
  s = PfqPrj.SaveAsString(DefPfqPrj)
  SaveFileString cdlOpen.filename, s 'save spec file under selected name
  lblSpec.Caption = cdlOpen.filename

FileCancel:
End Sub

Private Sub sstPfq_Click(PreviousTab As Integer)
  Form_Resize
End Sub

Private Sub SetGraphNames()
  Dim i As Long, j As Long, k As Long
  Dim ilen As Long, ipos As Long, Ind As Long
  Dim oldName As String, newName As String, GraphName As String

  On Error Resume Next

'  grdGraphs.Rows = 0
  lstGraphs.Clear
  For i = 1 To grdSpecs.Rows
    If grdSpecs.TextMatrix(i, 1) = "Yes" Then
      j = j + 1
      oldName = "PKFQ-" & j & ".BMP"
      newName = grdSpecs.TextMatrix(i, 17)
      If i > 1 Then 'look for repeating station IDs
        ilen = Len(newName)
        For k = i - 1 To 1 Step -1
'          GraphName = grdGraphs.TextMatrix(k, 0)
          GraphName = lstGraphs.List(k)
          If Left(GraphName, ilen) = newName Then
            'same station ID, add index number
            If Len(GraphName) > ilen Then 'add one to this index
              ipos = InStrRev(GraphName, "-")
              Ind = CLng(Right(GraphName, ipos - 1))
              newName = newName & CStr(Ind + 1)
            Else 'just add "-1"
              newName = newName & "-1"
            End If
          End If
        Next k
      End If
      newName = newName & ".BMP"
      RenameGraph oldName, newName
'      grdGraphs.TextMatrix(i, 0) = FilenameNoExt(newName)
      lstGraphs.AddItem FilenameNoExt(newName)
    End If
  Next i
  CurGraphName = FilenameOnly(lstGraphs.List(0)) & ".BMP"

End Sub

Private Sub RenameGraph(oldName As String, ByVal newName As String)
  'rename PeakFQ graphic files
  'always BMPs; other graphic files too if BMP is not the graphic format
  On Error Resume Next

  Kill newName
  Name oldName As newName
  If PfqPrj.GraphFormat <> "BMP" Then 'rename other graphic files too
    newName = FilenameNoExt(newName) & "." & PfqPrj.GraphFormat
    Kill newName
    Name FilenameNoExt(oldName) & "." & PfqPrj.GraphFormat As newName
  End If

End Sub
