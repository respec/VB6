VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "*\A..\ATCoCtl\ATCoCtl.vbp"
Begin VB.Form frmStreamStatsDB 
   Caption         =   "Stream Stats DB"
   ClientHeight    =   8625
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10800
   Icon            =   "frmStreamStatsDB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlgFileSel 
      Left            =   0
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   3735
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   420
      TabCaption(0)   =   "Stati&on Management"
      TabPicture(0)   =   "frmStreamStatsDB.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblState"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCategory"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblFilter"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboState"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cboCategory"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraStaSel"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboFilter"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Stat&istic Management"
      TabPicture(1)   =   "frmStreamStatsDB.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraStatistics"
      Tab(1).Control(1)=   "fraStatType"
      Tab(1).ControlCount=   2
      Begin VB.Frame fraStatType 
         Caption         =   "Statistic Type"
         Height          =   3252
         Left            =   -74880
         TabIndex        =   31
         Top             =   360
         Width           =   4390
         Begin VB.Frame fraEditStatType 
            Height          =   1800
            HelpContextID   =   26
            Left            =   960
            TabIndex        =   42
            Top             =   1340
            Visible         =   0   'False
            Width           =   3372
            Begin VB.CommandButton cmdSaveStat 
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
               HelpContextID   =   26
               Left            =   1680
               TabIndex        =   38
               Top             =   1200
               Width           =   732
            End
            Begin VB.TextBox txtStatTypeCode 
               Height          =   288
               HelpContextID   =   26
               Left            =   684
               TabIndex        =   36
               Top             =   300
               Width           =   612
            End
            Begin VB.TextBox txtStatType 
               Height          =   288
               HelpContextID   =   26
               Left            =   684
               TabIndex        =   37
               Top             =   720
               Width           =   2565
            End
            Begin VB.CommandButton cmdClose 
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
               HelpContextID   =   26
               Left            =   2520
               TabIndex        =   39
               Top             =   1200
               Width           =   732
            End
            Begin VB.Label lblAddStat 
               Alignment       =   1  'Right Justify
               Caption         =   "Code: "
               Height          =   252
               Left            =   120
               TabIndex        =   44
               Top             =   300
               Width           =   552
            End
            Begin VB.Label lblAddStatistic 
               Alignment       =   1  'Right Justify
               Caption         =   "Name: "
               Height          =   252
               Left            =   120
               TabIndex        =   43
               Top             =   720
               Width           =   552
            End
         End
         Begin VB.ComboBox cboStatTypes 
            Height          =   288
            HelpContextID   =   26
            Left            =   960
            TabIndex        =   32
            Top             =   360
            Width           =   2532
         End
         Begin VB.CommandButton cmdEditStatType 
            Caption         =   "Edit"
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
            HelpContextID   =   28
            Left            =   120
            TabIndex        =   34
            Top             =   2040
            Width           =   732
         End
         Begin VB.CommandButton cmdAddStatType 
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
            HelpContextID   =   27
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   732
         End
         Begin VB.CommandButton cmdDeleteStatType 
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
            HelpContextID   =   29
            Left            =   120
            TabIndex        =   35
            Top             =   2640
            Width           =   732
         End
         Begin VB.Label lblStatTypes 
            Alignment       =   1  'Right Justify
            Caption         =   "Selection: "
            Height          =   252
            Left            =   100
            TabIndex        =   40
            Top             =   360
            Width           =   852
         End
      End
      Begin VB.Frame fraStatistics 
         Caption         =   "Statistics"
         Height          =   3252
         Left            =   -70440
         TabIndex        =   27
         Top             =   360
         Width           =   6135
         Begin VB.ListBox lstStats 
            Height          =   2595
            HelpContextID   =   30
            Left            =   120
            MultiSelect     =   1  'Simple
            TabIndex        =   41
            Top             =   360
            Width           =   5895
         End
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         HelpContextID   =   15
         Left            =   3000
         TabIndex        =   4
         Top             =   480
         Width           =   1812
      End
      Begin VB.Frame fraStaSel 
         Caption         =   "Station Selections"
         Height          =   3252
         Left            =   4920
         TabIndex        =   25
         Top             =   360
         Width           =   5775
         Begin VB.OptionButton rdoStaOpt 
            Caption         =   "Name"
            Height          =   252
            HelpContextID   =   16
            Index           =   0
            Left            =   1440
            TabIndex        =   6
            Top             =   240
            Width           =   732
         End
         Begin VB.OptionButton rdoStaOpt 
            Caption         =   "ID"
            Height          =   252
            HelpContextID   =   16
            Index           =   1
            Left            =   2280
            TabIndex        =   7
            Top             =   240
            Width           =   852
         End
         Begin VB.OptionButton rdoStaOpt 
            Caption         =   "Both"
            Height          =   252
            HelpContextID   =   16
            Index           =   2
            Left            =   3120
            TabIndex        =   8
            Top             =   240
            Width           =   852
         End
         Begin VB.ListBox lstStations 
            Height          =   2400
            HelpContextID   =   16
            Left            =   120
            MultiSelect     =   1  'Simple
            TabIndex        =   9
            Top             =   600
            Width           =   5535
         End
         Begin VB.Label lblStaOpts 
            Caption         =   "Label Stations by:"
            Height          =   252
            Left            =   120
            TabIndex        =   26
            Top             =   264
            Width           =   1332
         End
      End
      Begin VB.ComboBox cboCategory 
         Height          =   315
         HelpContextID   =   15
         Left            =   3000
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   1812
      End
      Begin VB.ComboBox cboState 
         Height          =   315
         HelpContextID   =   14
         Left            =   510
         TabIndex        =   3
         ToolTipText     =   "Choose the State that "
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label lblFilter 
         Alignment       =   1  'Right Justify
         Caption         =   "Filter By: "
         Height          =   252
         Left            =   2040
         TabIndex        =   23
         Top             =   480
         Width           =   972
      End
      Begin VB.Label lblCategory 
         Alignment       =   1  'Right Justify
         Caption         =   "Category: "
         Height          =   252
         Left            =   2040
         TabIndex        =   24
         Top             =   960
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label lblState 
         Alignment       =   1  'Right Justify
         Caption         =   "State: "
         Height          =   252
         Left            =   36
         TabIndex        =   22
         Top             =   480
         Width           =   492
      End
   End
   Begin VB.Frame fraGenInfo 
      Caption         =   "Station Information"
      Height          =   4500
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   10815
      Begin VB.CommandButton cmdHelp 
         Caption         =   "&Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9960
         TabIndex        =   20
         Top             =   3680
         Width           =   700
      End
      Begin VB.Frame fraImport 
         Caption         =   "Import/Edit"
         Height          =   800
         HelpContextID   =   12
         Left            =   6880
         TabIndex        =   30
         Top             =   3620
         Width           =   3045
         Begin VB.CommandButton cmdXLSImport 
            Caption         =   "Ex&cel"
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
            Left            =   1560
            TabIndex        =   46
            Top             =   240
            Width           =   660
         End
         Begin VB.CommandButton cmdROISetUp 
            Caption         =   "&ROI"
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
            Left            =   2280
            TabIndex        =   45
            Top             =   240
            Width           =   660
         End
         Begin VB.CommandButton cmdBCF 
            Caption         =   "&BCF"
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
            HelpContextID   =   12
            Left            =   840
            TabIndex        =   19
            Top             =   240
            Width           =   660
         End
         Begin VB.CommandButton cmdNWIS 
            Caption         =   "N&WIS"
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
            HelpContextID   =   12
            Left            =   100
            TabIndex        =   18
            Top             =   240
            Width           =   660
         End
      End
      Begin VB.Frame fraGridButtons 
         Caption         =   "Grid Commands"
         Height          =   800
         HelpContextID   =   17
         Left            =   4240
         TabIndex        =   29
         Top             =   3620
         Width           =   2580
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Ca&ncel"
            Enabled         =   0   'False
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
            HelpContextID   =   17
            Left            =   900
            TabIndex        =   16
            Top             =   240
            Width           =   780
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
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
            HelpContextID   =   17
            Left            =   100
            TabIndex        =   15
            Top             =   240
            Width           =   732
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "C&lear"
            Enabled         =   0   'False
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
            HelpContextID   =   17
            Left            =   1760
            TabIndex        =   17
            Top             =   240
            Width           =   732
         End
      End
      Begin VB.Frame fraStatButtons 
         Caption         =   "Station Commands"
         Height          =   800
         HelpContextID   =   17
         Left            =   120
         TabIndex        =   1
         Top             =   3620
         Width           =   4040
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Enabled         =   0   'False
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
            HelpContextID   =   20
            Left            =   3200
            TabIndex        =   14
            Top             =   240
            Width           =   760
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Enabled         =   0   'False
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
            HelpContextID   =   18
            Left            =   2380
            TabIndex        =   13
            Top             =   240
            Width           =   732
         End
         Begin VB.CommandButton cmdEditStation 
            Caption         =   "&Edit"
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
            HelpContextID   =   21
            Left            =   1560
            TabIndex        =   12
            Top             =   240
            Width           =   732
         End
         Begin VB.Label lblStaSel 
            BackColor       =   &H00E0E0E0&
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   11
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label lblStaSel 
            Caption         =   "Active Station:"
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1932
         End
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9960
         TabIndex        =   21
         Top             =   4080
         Width           =   700
      End
      Begin ATCoCtl.ATCoGrid grdGenInfo 
         Height          =   3375
         HelpContextID   =   19
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   5953
         SelectionToggle =   0   'False
         AllowBigSelection=   0   'False
         AllowEditHeader =   0   'False
         AllowLoad       =   0   'False
         AllowSorting    =   -1  'True
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
         BackColorBkg    =   -2147483632
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         InsideLimitsBackground=   -2147483643
         OutsideHardLimitBackground=   8421631
         OutsideSoftLimitBackground=   8454143
         ComboCheckValidValues=   -1  'True
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuDatabase 
         Caption         =   "&Select Database"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpManual 
         Caption         =   "&Manual"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmStreamStatsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fileTitle As String
Dim Changes() As String
Dim ListedStations() As ssStation
Dim ListedStats() As ssStatLabel
Dim EditingStatType As Boolean
Dim AddingStatType As Boolean
Dim SaidYes As Boolean

Private Sub cboStatTypes_Click()
  Dim vStatLabel As Variant
  Dim statCnt&
  Dim allStatIDs$, staID$

  If cboStatTypes.ListIndex >= 0 Then
    lstStats.Clear
    Set SSDB.SelStats = Nothing
    If tabMain.Tab = 1 Then
      grdGenInfo.ClearData
      grdGenInfo.Rows = 0
    End If
    SaveSetting "StreamStatsDB", "Defaults", "ssStatType", cboStatTypes.List(cboStatTypes.ListIndex)
    Set SSDB.StatType = SSDB.StatisticTypes(cboStatTypes.ListIndex + 1)
    ReDim ListedStats(SSDB.StatType.StatLabels.Count)
    For Each vStatLabel In SSDB.StatType.StatLabels
      If vStatLabel.id > 14 Then
        statCnt = statCnt + 1
        lstStats.List(statCnt - 1) = vStatLabel.Name
        lstStats.ItemData(statCnt - 1) = vStatLabel.id
        Set ListedStats(statCnt) = vStatLabel
      End If
    Next
    allStatIDs = GetSetting("StreamStatsDB", "Defaults", cboStatTypes.ListIndex & "StatIDs")
    While Len(Trim(allStatIDs)) > 0
      staID = StrRetRem(allStatIDs)
      For statCnt = 1 To lstStats.ListCount
        If ListedStats(statCnt).id = CLng(staID) Then
          lstStats.Selected(statCnt - 1) = True
          Exit For
        End If
      Next statCnt
    Wend
  End If
End Sub

Private Sub cmdAddStatType_Click()
  AddingStatType = True
  EditingStatType = False
  fraEditStatType.Visible = True
  fraEditStatType.Caption = "Add Statistic Type"
  txtStatTypeCode.Text = ""
  txtStatType.Text = ""
End Sub

Private Sub cmdCancel_Click()
  Dim i&, j&
  If tabMain.Tab = 0 Then
    For i = 1 To SSDB.state.SelStations.Count
      If SSDB.state.SelStations(i - j).IsNew Then
        SSDB.state.SelStations.Remove (i - j)
        j = j + 1
        Set SSDB.state.SelStation = SSDB.state.SelStations(i - j)
      End If
    Next i
  ElseIf tabMain.Tab = 1 Then
    For i = 1 To SSDB.SelStats.Count
      If SSDB.SelStats(i - j).IsNew Then
        SSDB.SelStats.Remove (i - j)
        j = j + 1
      End If
    Next i
  End If
  ResetGrid
End Sub

Private Sub cmdClear_Click()
  Dim str$
  
  If tabMain.Tab = 0 Then
    If lblCategory.Visible = True Then
      str = lblCategory.Caption
      str = Mid(str, 1, Len(str) - 2)
    Else
      str = "State"
    End If
    SaveSetting "StreamStatsDB", "Defaults", SSDB.state.code & str & "Stations", ""
    cboCategory_Click
  ElseIf tabMain.Tab = 1 Then
    SaveSetting "StreamStatsDB", "Defaults", cboStatTypes.ListIndex & "StatIDs", ""
    cboStatTypes_Click
  End If
End Sub

Private Sub cmdClose_Click()
  fraEditStatType.Visible = False
  AddingStatType = False
  EditingStatType = False
End Sub

Private Sub cmdDelete_Click()
  Dim row&, col&, response&
  Dim staID As String
  
  If grdGenInfo.Rows = 0 Then Exit Sub
  
  Me.MousePointer = vbHourglass
  If tabMain.Tab = 0 Then
    If SSDB.state.SelStation Is Nothing Then
      Set SSDB.state.SelStation = SSDB.state.SelStations(grdGenInfo.row)
    End If
    If SSDB.state.SelStation.IsNew Then
      response = 2
    Else
      response = myMsgBox.Show("Are you certain you want to delete the station " & _
          SSDB.state.SelStation.Name & vbCrLf & "from the database for " & _
          SSDB.state.Name & "?", "User Action Verification", "+&Cancel", "-&Yes")
    End If
    If response = 2 Then
      With grdGenInfo
        If Not SSDB.state.SelStations(grdGenInfo.row).IsNew Then _
            SSDB.state.SelStations(.row).Delete
        ResetClass "Stations"
        .row = 1
        If .Rows > 0 Then
          Set SSDB.state.SelStation = SSDB.state.SelStations(.row)
          lblStaSel(0).Caption = SSDB.state.SelStation.id
          fraGenInfo.Caption = "Station Information - " & SSDB.state.SelStation.Name
        Else
          Set SSDB.state.SelStation = Nothing
          lblStaSel(0).Caption = ""
        End If
      End With
    End If
  ElseIf tabMain.Tab = 1 Then
    If SSDB.SelStats(grdGenInfo.row).IsNew Then
      response = 2
    Else
      response = myMsgBox.Show("Are you certain you want to delete the statistic " & _
          SSDB.SelStats(grdGenInfo.row).Name & vbCrLf & "from the database for " & _
          SSDB.state.Name & "?", "User Action Verification", "+&Cancel", "-&Yes")
    End If
    If response = 2 Then
      With grdGenInfo
        If Not SSDB.SelStats(grdGenInfo.row).IsNew Then
          row = 1
          While SSDB.SelStats(.row).id <> ListedStats(row).id
            row = row + 1
          Wend
          lstStats.RemoveItem (row - 1)
          SSDB.SelStats(.row).Delete
        End If
        SSDB.SelStats.Remove .row
        For row = .row To .Rows - 1
          For col = 0 To .Cols - 1
            .TextMatrix(row, col) = .TextMatrix(row + 1, col)
          Next col
        Next row
        .Rows = .Rows - 1
        .row = 1
        .col = 1
        If .row > 0 Then
          lblStaSel(0).Caption = .TextMatrix(grdGenInfo.row, 2)
          fraGenInfo.Caption = "Statistic Information - " & .TextMatrix(grdGenInfo.row, 2)
        Else
          lblStaSel(0).Caption = ""
          fraGenInfo.Caption = "Statistic Information"
        End If
      End With
    End If
  End If
  Me.MousePointer = vbDefault
End Sub

Private Sub cmdAdd_Click()
  Dim col&
  With grdGenInfo
    If tabMain.Tab = 0 Then
      SSDB.state.SelStations.Add New ssStation
      Set SSDB.state.SelStations(SSDB.state.SelStations.Count).DB = SSDB
      Set SSDB.state.SelStations(SSDB.state.SelStations.Count).state = SSDB.state
      SSDB.state.SelStations(SSDB.state.SelStations.Count).IsNew = True
'      cmdDelete.Enabled = True
      .Rows = .Rows + 1
      .row = .Rows
      .col = 0
      For col = 0 To .Cols - 1
        grdGenInfo.ColEditable(col) = True
      Next col
    ElseIf tabMain.Tab = 1 Then
      SSDB.SelStats.Add New ssStatLabel
      Set SSDB.SelStats(SSDB.SelStats.Count).DB = SSDB
      SSDB.SelStats(SSDB.SelStats.Count).IsNew = True
'      cmdDelete.Enabled = True
      .Rows = .Rows + 1
      .row = .Rows
      .col = 0
      For col = 0 To .Cols - 1
        grdGenInfo.ColEditable(col) = True
      Next col
    End If
  End With
  cmdDelete.Enabled = True
  cmdSave.Enabled = True
  cmdCancel.Enabled = True
  cmdClear.Enabled = True
End Sub

Private Sub cmdDeleteStatType_Click()
  Dim response&, tmpIndex&
  
  response = myMsgBox.Show("Are you certain that you want to delete " & _
      SSDB.StatType.Name & vbCrLf & " from the list of statistic types " & _
      "in the StreamStatsDB database?", "User Action Verification", "+&Cancel", "-&Yes")
  If response = 2 Then
    SSDB.StatType.Delete
  End If
  cboStatTypes.Clear
  Set SSDB.StatisticTypes = Nothing
  For tmpIndex = 1 To SSDB.StatisticTypes.Count
    cboStatTypes.List(tmpIndex - 1) = SSDB.StatisticTypes(tmpIndex).Name
    cboStatTypes.ItemData(tmpIndex - 1) = SSDB.StatisticTypes(tmpIndex).id
  Next tmpIndex
End Sub

Private Sub cmdEditStation_Click()
  
  If grdGenInfo.row > 0 Then
    If SSDB.state.SelStation Is Nothing Then
      Set SSDB.state.SelStation = SSDB.state.SelStations(grdGenInfo.row)
    End If
    frmStaData.Show vbModal, Me
  Else
    MsgBox "No stations is currently selected." & vbCrLf & _
        "Select a field on the row of the grid associated with the station you" & _
        " wish to edit" & vbCrLf & "then select the 'Edit Station' button again."
  End If
End Sub

Private Sub cmdEditStatType_Click()
  fraEditStatType.Visible = True
  fraEditStatType.Caption = "Edit Statistic Type"
  txtStatTypeCode.Text = SSDB.StatType.code
  txtStatType.Text = SSDB.StatType.Name
  EditingStatType = True
  AddingStatType = False
End Sub

Private Sub cmdExit_Click()
  ShutErDown
End Sub

Private Sub cmdNWIS_Click()
  Dim PathName$, stAbbrev$
  Dim Length&, i&
  
  On Error GoTo x
  
  PathName = GetSetting("StreamStatsDB", "Defaults", "NWISImportPath")
  With frmCDLG.CDLG
    .DialogTitle = "Select a file for import"
    If Len(PathName) > 0 Then .filename = PathName & "*.xls"
    .Filter = "(*.xls)|*.xls"
    .filterIndex = 1
    .CancelError = True
    .ShowOpen
    If Len(Dir(.filename, vbDirectory)) > 1 Then
      PathName = Left(.filename, Len(.filename) - Len(.fileTitle))
      SaveSetting "StreamStatsDB", "Defaults", "NWISImportPath", PathName
      Me.MousePointer = vbHourglass
      NWISImport .filename
      'Reassign class structure to previous selections
      cboState_Click
    End If
  End With
x:
  Unload frmCDLG
  Me.MousePointer = vbDefault
End Sub

Private Sub cmdBCF_Click()
  Dim PathName$, stName$
  Dim Length&, i&
  
  On Error GoTo x
  
  PathName = GetSetting("StreamStatsDB", "Defaults", "BCFImportPath")
  With frmCDLG.CDLG
    .DialogTitle = "Select a file for import"
    If Len(PathName) > 0 Then .filename = PathName & "*.txt"
    .Filter = "(*.txt)|*.txt"
    .filterIndex = 1
    .CancelError = True
    .ShowOpen
    If Len(Dir(.filename, vbDirectory)) > 1 Then
      PathName = PathNameOnly(.filename)
      PathName = Left(.filename, Len(.filename) - Len(.fileTitle))
      SaveSetting "StreamStatsDB", "Defaults", "BCFImportPath", PathName
      Me.MousePointer = vbHourglass
      i = InStr(1, UCase(.fileTitle), "BC")
      stName = LCase(Left(.fileTitle, i - 1))
      'Add underscore to states with double names, if necessary
      If InStr(1, stName, " ") > 0 Then
        stName = ReplaceString(stName, " ", "_")
      ElseIf InStr(1, stName, "_") = 0 Then
        Select Case Left(stName, 3)
         Case "ame": stName = Left(stName, 8) & "_" & Mid(stName, 9)
         Case "new": stName = Left(stName, 3) & "_" & Mid(stName, 4)
         Case "nor": stName = Left(stName, 5) & "_" & Mid(stName, 6)
         Case "pue": stName = Left(stName, 6) & "_" & Mid(stName, 7)
         Case "rho": stName = Left(stName, 5) & "_" & Mid(stName, 6)
         Case "sou": stName = Left(stName, 5) & "_" & Mid(stName, 6)
         Case "wes": stName = Left(stName, 4) & "_" & Mid(stName, 5)
        End Select
      End If
      For i = 1 To SSDB.States.Count
        If (stName = LCase(SSDB.States(i).Name)) Or _
           (stName = LCase(SSDB.States(i).Abbrev)) Then Exit For
      Next i
      cboState.ListIndex = i - 1
      BCFImport .filename
      cboState.ListIndex = i - 1
      Set SSDB.state.Stations = Nothing
      cboState_Click
    End If
  End With
x:
  Unload frmCDLG
  Me.MousePointer = vbDefault
End Sub

Private Sub cmdSave_Click()
  Dim row&, col&, response&
  Dim str$, allStatIDs$
  Dim madeChanges As Boolean

  'Check to ensure entries have been made for required fields
  If Not QACheck Then Exit Sub
  
  'Record new values and determine if any changes were made
  ChangesMade madeChanges
  If Not madeChanges Then
    GoTo NoChanges
  End If
  If SaidYes Then
    response = 2
  ElseIf tabMain.Tab = 0 Then
    str = "Are you certain that you want to write these new values " & _
        "to the database for " & SSDB.state.Name & "?"
    response = myMsgBox.Show(str, "User Action Verification", "+&Cancel", "-&Yes")
  Else 'saving statistic
    str = "Are you certain that you want to write these new values " & _
          "to the statistic database."
    response = myMsgBox.Show(str, "User Action Verification", "+&Cancel", "-&Yes")
  End If
  SaidYes = False
  If response = 2 Then 'Overwrite values in DB
    frmUserInfo.Show vbModal, Me
    If Not UserInfoOK Then GoTo NoChanges
    Me.MousePointer = vbHourglass
    If tabMain.Tab = 0 Then  'editing selected stations
      For row = 1 To grdGenInfo.Rows
        If SSDB.state.SelStations(row).IsNew Then
          SSDB.state.SelStations(row).Add Changes(), row
        Else
          SSDB.state.SelStations(row).Edit Changes(), row
        End If
        'Write changes to DetailedLog table
        For col = 1 To UBound(Changes, 3)
          If Changes(0, row, col) = "1" Or Changes(0, row, col) = "2" Then
            SSDB.RecordChanges TransID, "STATION", col, _
                grdGenInfo.TextMatrix(row, 0), Changes(1, row, col), Changes(2, row, col)
          End If
        Next col
        allStatIDs = allStatIDs & " " & grdGenInfo.TextMatrix(row, 0)
      Next row
      str = lblCategory.Caption
      str = Mid(str, 1, Len(str) - 2)
      SaveSetting "StreamStatsDB", "Defaults", SSDB.state.code & str & "Stations", allStatIDs
      ResetClass "Stations"
    ElseIf tabMain.Tab = 1 Then  'editing selected statistics
      For row = 1 To grdGenInfo.Rows
        If SSDB.SelStats(row).IsNew Then
          SSDB.SelStats(row).Add Changes(), row
        Else
          SSDB.SelStats(row).Edit Changes(), row
        End If
        'Write changes to DetailedLog table
        For col = 1 To UBound(Changes, 3) - 1
          If Changes(0, row, col) = "1" Or Changes(0, row, col) = "2" Then
            SSDB.RecordChanges TransID, "STATLABEL", col + 1, _
                grdGenInfo.TextMatrix(row, 1), Changes(1, row, col + 1), Changes(2, row, col + 1)
          End If
        Next col
        allStatIDs = allStatIDs & " " & SSDB.SelStats(row).id
      Next row
      SaveSetting "StreamStatsDB", "Defaults", cboStatTypes.ListIndex & "StatIDs", allStatIDs
      ResetClass "Stats"
    End If
  End If
NoChanges:
  Me.MousePointer = vbDefault
End Sub

Private Function QACheck() As Boolean
  Dim row&, col&, rowCnt&, response&, i&, j&
  Dim str$, basinID$
  Dim myStateBasin As ssStateBasin
  
  rowCnt = grdGenInfo.Rows
  If tabMain.Tab = 0 Then
    For row = 1 To rowCnt
      'Check to ensure that station code is proper length
      If Len(grdGenInfo.TextMatrix(row, 0)) < 8 Or _
          Len(grdGenInfo.TextMatrix(row, 0)) > 15 Then
        MsgBox "The length of the Station ID on row " & row & _
            " of the grid must be 8-15 digits long." & vbCrLf & _
            "Re-enter a value of this length.", _
            vbCritical, "Erroneous Code Length"
        GoTo QAproblem
      End If
      'Check to ensure that proper entries have been made for all fields
      For col = 0 To grdGenInfo.Cols - 1
        With grdGenInfo
          Select Case col
            Case 1:
              If Len(Trim(.TextMatrix(row, col))) = 0 Then
                 MsgBox "Station " & .TextMatrix(row, 0) & " " & _
                    " must have a name.", vbCritical, "Bad Data Value"
                GoTo QAproblem
              End If
           Case 7 To 8:
              If Not IsNumeric(.TextMatrix(row, col)) Then
                .TextMatrix(row, col) = ""
                .Selected(row, col) = True
                MsgBox "The value for " & .TextMatrix(-1, col) & " " & _
                    .TextMatrix(0, col) & " for station " & .TextMatrix(row, 1) & _
                    " on row " & row & " must be numeric." & vbCrLf & _
                    "Enter a positive numeric value for this field.", _
                    vbCritical, "Bad Data Value"
                GoTo QAproblem
'              ElseIf Len(Trim(.TextMatrix(row, col))) <> 6 And col = 6 Then
'                MsgBox "The value for " & .TextMatrix(-1, col) & " " & _
'                    .TextMatrix(0, col) & " for station " & .TextMatrix(row, 1) & _
'                    " on row " & row & " must be 6 numerals long [DDMMSS].", _
'                    vbCritical, "Bad Data Value"
'                GoTo QAproblem
'              ElseIf (Len(Trim(.TextMatrix(row, col))) < 6 _
'                    Or Len(Trim(.TextMatrix(row, col))) > 7) And col = 7 Then
'                MsgBox "The value for " & .TextMatrix(-1, col) & " " & _
'                    .TextMatrix(0, col) & " for station " & .TextMatrix(row, 1) & _
'                    " on row " & row & " must be 6 or 7 numerals long [DDDMMSS].", _
'                    vbCritical, "Bad Data Value"
'                GoTo QAproblem
              End If
              If col = 7 Then
                If CLng(.TextMatrix(row, col)) < 17.5 Then
                  response = myMsgBox.Show("The latitude for station " & .TextMatrix(row, 1) & _
                      " on row " & row & " is south of the Virgin Islands." & vbCrLf & _
                      "Cancel this save event then enter a value greater then 17.5 to denote" & _
                      " an area in the U.S.", "Too far south for U.S.", "+&Change Now", "-&Continue")
                  If response = 1 Then GoTo QAproblem
                ElseIf CLng(.TextMatrix(row, col)) > 72# Then
                  response = myMsgBox.Show("The latitude for station " & .TextMatrix(row, 1) & _
                      " on row " & row & " is north of Alaska." & vbCrLf & _
                      "Cancel this save event then enter a value less than 72.0 to denote " & _
                      "an area in the U.S.", "Too far north for U.S.", "+&Change Now", "-&Continue")
                  If response = 1 Then GoTo QAproblem
                End If
              ElseIf col = 8 Then
                If Abs(CLng(.TextMatrix(row, col))) < 64# Then
                  response = myMsgBox.Show("The longitude for station " & .TextMatrix(row, 1) & _
                      " on row " & row & " is east of the Virgin Islands." & vbCrLf & _
                      "To denote an area in the U.S., cancel this save" & _
                      " event then enter a value less than -64.0.", _
                      "Too far east for U.S.", "+&Change Now", "-&Continue")
                  If response = 1 Then GoTo QAproblem
                ElseIf Abs(CLng(.TextMatrix(row, col))) > 172# Then
                  response = myMsgBox.Show("The longitude for station " & .TextMatrix(row, 1) & _
                      " on row " & row & " is west of Alaska." & vbCrLf & _
                      "To denote an area in the U.S., cancel this save " & _
                      "event then enter a value greater than -172.0.", _
                      "Too far west for U.S.", "+&Change Now", "-&Continue")
                  If response = 1 Then GoTo QAproblem
                End If
              End If
            Case 14: 'State Basin
              'Separate the code and name if they've been aggregated
              i = InStr(1, .TextMatrix(row, col), "-")
              str = Mid(.TextMatrix(row, col), i + 1)
              If Len(Trim(str)) = 0 Then GoTo noEntry
              If i > 0 Then
                basinID = Left(.TextMatrix(row, col), i - 1)
              Else
                basinID = ""
              End If
              If IsNumeric(basinID) And SSDB.state.StateBasins.IndexFromKey(basinID) > 0 Then
                'State Basin with this code already exists for this state
                If Mid(.TextMatrix(row, col), i + 1) <> SSDB.state.StateBasins(basinID).Name Then
                  response = myMsgBox.Show("You have edited the state basin for station " & _
                      .TextMatrix(row, 0) & "." & vbCrLf & vbCrLf & _
                      "Would you like to overwrite the previously existing entry " & _
                      "(effectively renaming the state basin)," & vbCrLf & _
                      "or Cancel this save operation and re-select one of the existing state basins?", _
                      "User Action Verification", "-&Overwrite", "-&Cancel")
                  If response = 1 Then
                    Set myStateBasin = SSDB.state.StateBasins(basinID)
                    myStateBasin.Edit str
                    For j = row + 1 To .Rows
                      'Change State Basin of other data in grid with same source
                      If basinID = Left(.TextMatrix(j, col), i - 1) Then
                        .TextMatrix(j, col) = .TextMatrix(row, col)
                      End If
                    Next j
                    Set SSDB.state.StateBasins = Nothing
                    SaidYes = True
                  Else
                    QACheck = False
                    SaidYes = False
                    Exit Function
                  End If
                End If
              Else
                'New State Basin
                response = myMsgBox.Show("The state basin named " & _
                    .TextMatrix(row, col) & " for station " & .TextMatrix(row, 0) & _
                    " does not does not exist in the database." & _
                    vbCrLf & vbCrLf & "Would you like to save this new State " & _
                    "Basin to the database," & vbCrLf & "or cancel this save " & _
                    "operation and select one of the existing state basins?", _
                    "User Action Verification", "+&Save", "-&Cancel")
                If response = 1 Then
                  Set myStateBasin = New ssStateBasin
                  Set myStateBasin.DB = SSDB
                  myStateBasin.Add str
                  Set SSDB.state.StateBasins = Nothing
                  Set myStateBasin = _
                      SSDB.state.StateBasins(SSDB.state.StateBasins.Count)
                  .TextMatrix(row, col) = myStateBasin.code & "-" & .TextMatrix(row, col)
                  SaidYes = True
                Else
                  QACheck = False
                  SaidYes = False
                  Exit Function
                End If
              End If
noEntry:
          End Select
        End With
      Next col
    Next row
  ElseIf tabMain.Tab = 1 Then
  
  End If
  QACheck = True
QAproblem:
End Function

Private Sub cmdSaveStat_Click()
  Dim tmpIndex&
  Dim myStatType As ssStatType

  If Len(Trim(txtStatTypeCode.Text)) = 0 Or _
      Len(txtStatTypeCode.Text) > 6 Then
    MsgBox "The Statistic Type Code must be 1-6 characters long." & vbCrLf & _
        "Re-enter a Code of this length and click on 'Save' again."
    Exit Sub
  End If
  If Len(Trim(txtStatType.Text)) = 0 Or _
      Len(txtStatType.Text) > 40 Then
    MsgBox "The Statistic Type Name must be 1-40 characters long." & vbCrLf & _
        "Re-enter a Name of this length and click on 'Save' again."
    Exit Sub
  End If
  If AddingStatType Then
    Set myStatType = New ssStatType
    Set myStatType.DB = SSDB
    myStatType.Add Trim(txtStatTypeCode.Text), txtStatType.Text
  ElseIf EditingStatType Then
    SSDB.StatType.Edit Trim(txtStatTypeCode.Text), txtStatType.Text
  End If
  txtStatTypeCode.Text = ""
  txtStatType.Text = ""
  fraEditStatType.Visible = False
  Set SSDB.StatisticTypes = Nothing
  For tmpIndex = 1 To SSDB.StatisticTypes.Count
    cboStatTypes.List(tmpIndex - 1) = SSDB.StatisticTypes(tmpIndex).Name
    cboStatTypes.ItemData(tmpIndex - 1) = SSDB.StatisticTypes(tmpIndex).id
  Next tmpIndex
  If AddingStatType Then
    cboStatTypes.ListIndex = cboStatTypes.ListCount - 1
  End If
  AddingStatType = False
  EditingStatType = False
End Sub

Private Sub cmdHelp_Click()
  
  OpenHelp

End Sub

Private Sub cmdROISetUp_Click()
  If Not SSDB.state Is Nothing Then
    frmROISetUp.Show vbModal, Me
  Else
    MsgBox "No state is currently selected." & vbCrLf & _
        "Select the state for which you wish to define the ROI methodology," & _
        vbCrLf & "then select the 'ROI' button again."
  End If
End Sub

Private Sub cmdXLSImport_Click()
  frmXLImport.Tag = 0
  frmXLImport.Show 1
  If frmXLImport.Tag = 1 Then 'import succeeded
    Set SSDB.state.Stations = Nothing
    cboFilter_Click
  End If
  Unload frmXLImport

End Sub

Private Sub Form_Resize()
  If Me.Width > 8000 Then
    tabMain.Width = Me.Width - 105
    fraGenInfo.Width = Me.Width - 105
    grdGenInfo.Width = fraGenInfo.Width - 230
    fraStatistics.Width = Me.Width - 4785
    lstStats.Width = fraStatistics.Width - 240
    fraStaSel.Width = Me.Width - 5145
    lstStations.Width = fraStaSel.Width - 240
  End If
  If Me.height > 6000 Then
    fraGenInfo.height = Me.height - 4680
    grdGenInfo.height = fraGenInfo.height - 1125
    fraStatButtons.Top = grdGenInfo.height + 245
    fraGridButtons.Top = fraStatButtons.Top
    fraImport.Top = fraStatButtons.Top
    cmdHelp.Top = fraStatButtons.Top + 55
    cmdExit.Top = cmdHelp.Top + 390
  End If
End Sub

Private Sub grdGenInfo_RowColChange()
  Dim i&
  
  If tabMain.Tab = 0 Then
    If grdGenInfo.row > 0 Then
      Set SSDB.state.SelStation = SSDB.state.SelStations(grdGenInfo.row)
    Else
      Exit Sub
    End If
    With grdGenInfo
      If .row > 0 Then
        lblStaSel(0).Caption = .TextMatrix(grdGenInfo.row, 0)
        fraGenInfo.Caption = "Station Information - " & .TextMatrix(.row, 1)
      Else
        lblStaSel(0).Caption = ""
        fraGenInfo.Caption = "Station Information"
      End If
      .ClearValues
      If SSDB.state.SelStations(.row).IsNew Then
        .ColEditable(0) = True
      Else
        .ColEditable(0) = False
      End If
      If .ColWidth(5) > 4000 Then .ColWidth(5) = 4000 'limit size of Remarks columns
      Select Case .col
        Case 2:   .addValue ""
                  For i = 1 To SSDB.StationTypes.Count
                    .addValue SSDB.StationTypes(i).Name
                  Next i
                  .ComboCheckValidValues = True
        Case 3:   .addValue "Undefined"
                  .addValue "No"
                  .addValue "Yes"
                  .ComboCheckValidValues = True
        Case 13:  .addValue ""
                  For i = 1 To SSDB.state.HUCs.Count
                    .addValue SSDB.state.HUCs(i).code
                  Next i
                  .ComboCheckValidValues = True
                  .ColWidth(.col) = 1000
        Case 14:  .addValue ""
                  For i = 1 To SSDB.state.StateBasins.Count
                    .addValue SSDB.state.StateBasins(i).code & "-" & SSDB.state.StateBasins(i).Name
                  Next i
                  .ComboCheckValidValues = False
        Case 11:  .addValue ""
                  For i = 1 To SSDB.state.Counties.Count
                    .addValue SSDB.state.Counties(i).code & "-" & SSDB.state.Counties(i).Name
                  Next i
                  .ComboCheckValidValues = True
        Case 12:  .addValue ""
                  For i = 1 To SSDB.state.MCDs.Count
                    .addValue SSDB.state.MCDs(i).code & "-" & SSDB.state.MCDs(i).Name
                  Next i
                  .ComboCheckValidValues = True
        Case 9, 10: .addValue ""
                  For i = 1 To SSDB.States.Count
                    .addValue SSDB.States(i).Abbrev
                  Next i
                  .ComboCheckValidValues = True
      End Select
      .ColWidth(11) = 2500
      .ColWidth(12) = 2000
      .ColWidth(14) = 2500
    End With
  ElseIf tabMain.Tab = 1 Then
    With grdGenInfo
      If .row > 0 Then
        lblStaSel(0).Caption = .TextMatrix(grdGenInfo.row, 2)
        fraGenInfo.Caption = "Statistic Information - " & .TextMatrix(grdGenInfo.row, 2)
      Else
        lblStaSel(0).Caption = ""
        fraGenInfo.Caption = "Statistic Information"
        Exit Sub
      End If
      .ClearValues
      .ColEditable(.col) = True
      Select Case .col
        Case 0: For i = 1 To SSDB.StatisticTypes.Count
                  .addValue SSDB.StatisticTypes(i).Name
                Next i
                .ComboCheckValidValues = True
        Case 3: For i = 1 To SSDB.Units.Count
                  .addValue SSDB.Units(i).englishlabel
                Next i
                .ComboCheckValidValues = True
      End Select
    End With
  End If
End Sub

Private Sub lststations_Click()
  Dim row&, col&
  Dim allStatIDs$, str$

  On Error GoTo x

  If lstStations.Selected(lstStations.ListIndex) Then 'adding Station to grid
    SSDB.state.SelStations.Add ListedStations(lstStations.ListIndex + 1), _
        ListedStations(lstStations.ListIndex + 1).id
    ResetGrid
    cmdDelete.Enabled = True
    cmdEditStation.Enabled = True
    If grdGenInfo.Rows > 0 Then
      cmdSave.Enabled = True
      cmdCancel.Enabled = True
      cmdClear.Enabled = True
    Else
      cmdSave.Enabled = False
      cmdCancel.Enabled = False
      cmdClear.Enabled = False
    End If
  Else 'removing Station from grid
    With grdGenInfo
      row = 1
      While SSDB.state.SelStations(row).id <> ListedStations(lstStations.ListIndex + 1).id
        allStatIDs = allStatIDs & " " & SSDB.state.SelStations(row).id
        row = row + 1
      Wend
      SSDB.state.SelStations.Remove row
      For row = row To .Rows - 1
        allStatIDs = allStatIDs & " " & SSDB.state.SelStations(row).id
        For col = 0 To .Cols - 1
          .TextMatrix(row, col) = .TextMatrix(row + 1, col)
        Next col
      Next row
      .Rows = .Rows - 1
      .row = 1
      If .Rows = 0 Then
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        cmdClear.Enabled = False
        lblStaSel(0).Caption = ""
        fraGenInfo.Caption = "Station Information"
        cmdDelete.Enabled = False
        cmdEditStation.Enabled = False
      Else
        cmdSave.Enabled = True
        cmdCancel.Enabled = True
        cmdClear.Enabled = True
        Set SSDB.state.SelStation = SSDB.state.SelStations(.row)
        lblStaSel(0).Caption = .TextMatrix(.row, 0)
        fraGenInfo.Caption = "Station Information - " & .TextMatrix(.row, 1)
      End If
    End With
    If IsState Then
      str = "State"
    Else
      str = lblCategory.Caption
      str = Mid(str, 1, Len(str) - 2)
    End If
    SaveSetting "StreamStatsDB", "Defaults", SSDB.state.code & str & "Stations", allStatIDs
  End If
x:
End Sub

Private Sub Form_Load()
  Dim tmpIndex&, selState&, selStat&, response&
  Dim state$, StatType$

  Set myMsgBox = New ATCoMessage

  tabMain.Tab = 0
  state = GetSetting("StreamStatsDB", "Defaults", "ssState")
  Initialize tabMain.Tab

  InitializeFromDatabase (state)

End Sub

Private Sub cboState_Click()
  Dim stationCnt&, stateIndex&, filterIndex&
  Dim stateID
  
  On Error GoTo x
  
  If cboState.ListIndex >= 0 Then
    cmdAdd.Enabled = True
    lblStaSel(0).Caption = ""
    IsBasin = False
    IsCounty = False
    IsMCD = False
    IsHUC = False
    IsState = True
    If Not SSDB.state Is Nothing Then
      If cboState.List(cboState.ListIndex) = SSDB.state.Name Then Exit Sub
    End If
    Me.MousePointer = vbHourglass
    Initialize 0
    If Not SSDB.state Is Nothing Then
      SSDB.state.SelStations.Clear
      SSDB.state.Stations.Clear
    End If
   'Set state; populate StateBasins, County, and MCD combo boxes;
    ' and populate Stations listbox
    stateID = CStr(cboState.ItemData(cboState.ListIndex))
    If Len(stateID) = 1 Then stateID = "0" & stateID
    Set SSDB.state = SSDB.States(stateID)
    SaveSetting "StreamStatsDB", "Defaults", "SSState", SSDB.state.Name
    'Populate the "Filter By" listbox if new state selected
    cboCategory.Visible = False
    lblCategory.Visible = False
    cboFilter.Clear
    cboFilter.AddItem "Basin"
    cboFilter.AddItem "County"
    cboFilter.AddItem "MCD"
    cboFilter.AddItem "HUC"
    cboFilter.AddItem "State"
    filterIndex = GetSetting("StreamStatsDB", "Defaults", SSDB.state.code & "Filter", -1)
    cboFilter.ListIndex = filterIndex
    If IsState And filterIndex = -1 Then
      'populate Station combo box with all Stations in State
      lstStations.Clear
      grdGenInfo.ClearData
      grdGenInfo.Rows = 0
      Set SSDB.state.Stations = Nothing
      stationCnt = SSDB.state.Stations.Count
      If stationCnt = 0 Then GoTo x
      ReDim ListedStations(1 To stationCnt)
      For stateIndex = 1 To stationCnt
'        If SSDB.state.Stations(stateIndex).HasData Then
          lstStations.List(stateIndex - 1) = SSDB.state.Stations(stateIndex).label
'        Else
'          lstStations.List(stateIndex - 1) = "*" & SSDB.state.Stations(stateIndex).label
'        End If
        Set ListedStations(stateIndex) = SSDB.state.Stations(stateIndex)
      Next
      SelectStations
    End If
  End If
x:
  Me.MousePointer = vbDefault
End Sub

Private Sub cboFilter_Click()
  Dim catChoice$
  Dim catChoiceIndex&, selCatChoice&, stationCnt&, stateIndex&
  
  If cboFilter.ListIndex = -1 Then
    cboCategory.Visible = False
    lblCategory.Visible = False
  Else
    Me.MousePointer = vbHourglass
    'Clear previous selections
    Set SSDB.state.Stations = Nothing
    Set SSDB.state.SelStations = Nothing
    lstStations.Clear
    cboCategory.Clear
    grdGenInfo.ClearData
    grdGenInfo.Rows = 0
    cboCategory.Visible = True
    lblCategory.Visible = True
    'Clear the Stations collection of currently selected category
    With cboFilter
      Select Case .List(.ListIndex)
        Case "Basin":
          lblCategory.Caption = "Basin: "
          IsBasin = True
          IsCounty = False
          IsMCD = False
          IsHUC = False
          IsState = False
          catChoice = GetSetting("StreamStatsDB", "Defaults", SSDB.state.code & "ssBasin")
          For catChoiceIndex = 1 To SSDB.state.StateBasins.Count
            cboCategory.List(catChoiceIndex - 1) = _
                SSDB.state.StateBasins(catChoiceIndex).Name
            cboCategory.ItemData(catChoiceIndex - 1) = _
                SSDB.state.StateBasins(catChoiceIndex).code
            If SSDB.state.StateBasins(catChoiceIndex).Name = catChoice Then _
                selCatChoice = catChoiceIndex
          Next catChoiceIndex
          If selCatChoice > 0 Then 'select previously chosen Basin
            cboCategory.ListIndex = selCatChoice - 1
            Set SSDB.state.Statebasin = _
                SSDB.state.StateBasins(CStr(cboCategory.ItemData(cboCategory.ListIndex)))
            selCatChoice = 0
          Else
            Set SSDB.state.SelStations = Nothing
            Set SSDB.state.SelStation = Nothing
          End If
        Case "County":
          lblCategory.Caption = "County: "
          IsCounty = True
          IsBasin = False
          IsMCD = False
          IsHUC = False
          IsState = False
          catChoice = GetSetting("StreamStatsDB", "Defaults", SSDB.state.code & "ssCounty")
          For catChoiceIndex = 1 To SSDB.state.Counties.Count
            cboCategory.List(catChoiceIndex - 1) = _
                SSDB.state.Counties(catChoiceIndex).Name
            cboCategory.ItemData(catChoiceIndex - 1) = _
                SSDB.state.Counties(catChoiceIndex).code
            If SSDB.state.Counties(catChoiceIndex).Name = catChoice Then _
                selCatChoice = catChoiceIndex
          Next
          If selCatChoice > 0 Then 'select previously chosen County
            cboCategory.ListIndex = selCatChoice - 1
            catChoice = CStr(cboCategory.ItemData(cboCategory.ListIndex))
            While Len(catChoice) < 3
              catChoice = "0" & catChoice
            Wend
            Set SSDB.state.County = SSDB.state.Counties(catChoice)
            selCatChoice = 0
          Else
            Set SSDB.state.SelStations = Nothing
            Set SSDB.state.SelStation = Nothing
          End If
        Case "MCD":
          lblCategory.Caption = "MCD: "
          IsMCD = True
          IsBasin = False
          IsCounty = False
          IsHUC = False
          IsState = False
          catChoice = GetSetting("StreamStatsDB", "Defaults", SSDB.state.code & "ssMCD")
          For catChoiceIndex = 1 To SSDB.state.MCDs.Count
            cboCategory.List(catChoiceIndex - 1) = _
                SSDB.state.MCDs(catChoiceIndex).Name
            cboCategory.ItemData(catChoiceIndex - 1) = _
                SSDB.state.MCDs(catChoiceIndex).code
            If SSDB.state.MCDs(catChoiceIndex).Name = catChoice Then _
                selCatChoice = catChoiceIndex
          Next
          If selCatChoice > 0 Then 'select previously chosen MCD
            cboCategory.ListIndex = selCatChoice - 1
            catChoice = CStr(cboCategory.ItemData(cboCategory.ListIndex))
            While Len(catChoice) < 5
              catChoice = "0" & catChoice
            Wend
            Set SSDB.state.MCD = SSDB.state.MCDs(catChoice)
            selCatChoice = 0
          Else
            Set SSDB.state.SelStations = Nothing
            Set SSDB.state.SelStation = Nothing
          End If
        Case "HUC":
          lblCategory.Caption = "HUC: "
          IsHUC = True
          IsMCD = False
          IsBasin = False
          IsCounty = False
          IsState = False
          catChoice = GetSetting("StreamStatsDB", "Defaults", SSDB.state.code & "ssHUC")
          For catChoiceIndex = 1 To SSDB.state.HUCs.Count
            cboCategory.List(catChoiceIndex - 1) = _
                SSDB.state.HUCs(catChoiceIndex).code
            cboCategory.ItemData(catChoiceIndex - 1) = _
                SSDB.state.HUCs(catChoiceIndex).code
            If SSDB.state.HUCs(catChoiceIndex).code = catChoice Then _
                selCatChoice = catChoiceIndex
          Next
          If selCatChoice > 0 Then 'select previously chosen MCD
            cboCategory.ListIndex = selCatChoice - 1
            catChoice = CStr(cboCategory.ItemData(cboCategory.ListIndex))
            While Len(catChoice) < 8
              catChoice = "0" & catChoice
            Wend
            Set SSDB.state.HUC = SSDB.state.HUCs(catChoice)
            selCatChoice = 0
          Else
            Set SSDB.state.SelStations = Nothing
            Set SSDB.state.SelStation = Nothing
          End If
        Case "State":
          IsState = True
          IsMCD = False
          IsBasin = False
          IsCounty = False
          IsHUC = False
          cboCategory.Visible = False
          lblCategory.Visible = False
          stationCnt = SSDB.state.Stations.Count
          If stationCnt = 0 Then GoTo NoStations
          ReDim ListedStations(1 To stationCnt)
          catChoice = GetSetting("StreamStatsDB", "Defaults", SSDB.state.code & "StateStations")
          For stateIndex = 1 To stationCnt
'            If SSDB.state.Stations(stateIndex).HasData Then
              lstStations.List(stateIndex - 1) = SSDB.state.Stations(stateIndex).label
'            Else
'              lstStations.List(stateIndex - 1) = "*" & SSDB.state.Stations(stateIndex).label
'            End If
            Set ListedStations(stateIndex) = SSDB.state.Stations(stateIndex)
          Next
          SelectStations
      End Select
NoStations:
      SaveSetting "StreamStatsDB", "Defaults", SSDB.state.code & "Filter", .ListIndex
    End With
    Me.MousePointer = vbDefault
  End If
End Sub

Private Sub cboCategory_Click()
  Dim catCode$
  Dim stationIndex&, stationCnt&

  If cboCategory.ListIndex < 0 And Not IsState Then
    Exit Sub
  Else  'Set Basin and populate Station combo box
    Set SSDB.state.SelStations = Nothing
    grdGenInfo.Rows = 0
    lstStations.Clear
    grdGenInfo.ClearData
    If IsBasin Then
      Set SSDB.state.StateBasins = Nothing
      catCode = SSDB.state.StateBasins(cboCategory.ListIndex + 1).code
      Set SSDB.state.Statebasin = SSDB.state.StateBasins(catCode)
      stationCnt = SSDB.state.Statebasin.Stations.Count
      ReDim ListedStations(1 To stationCnt)
      For stationIndex = 1 To stationCnt
'        If SSDB.state.Statebasin.Stations(stationIndex).HasData Then
          lstStations.List(stationIndex - 1) = SSDB.state.Statebasin.Stations(stationIndex).label
'        Else
'          lstStations.List(stationIndex - 1) = "*" & SSDB.state.Statebasin.Stations(stationIndex).label
'        End If
        Set ListedStations(stationIndex) = SSDB.state.Statebasin.Stations(stationIndex)
      Next
      SaveSetting "StreamStatsDB", "Defaults", SSDB.state.code & "SSBasin", SSDB.state.Statebasin.Name
    ElseIf IsCounty Then
      Set SSDB.state.Counties = Nothing
      catCode = SSDB.state.Counties(cboCategory.ListIndex + 1).code
      Set SSDB.state.County = SSDB.state.Counties(catCode)
      stationCnt = SSDB.state.County.Stations.Count
      ReDim ListedStations(stationCnt)
      For stationIndex = 1 To SSDB.state.County.Stations.Count
'        If SSDB.state.County.Stations(stationIndex).HasData Then
          lstStations.List(stationIndex - 1) = SSDB.state.County.Stations(stationIndex).label
'        Else
'          lstStations.List(stationIndex - 1) = "*" & SSDB.state.County.Stations(stationIndex).label
'        End If
        Set ListedStations(stationIndex) = SSDB.state.County.Stations(stationIndex)
      Next
      SaveSetting "StreamStatsDB", "Defaults", SSDB.state.code & "SSCounty", SSDB.state.County.Name
    ElseIf IsMCD Then
      Set SSDB.state.MCDs = Nothing
      catCode = SSDB.state.MCDs(cboCategory.ListIndex + 1).code
      Set SSDB.state.MCD = SSDB.state.MCDs(catCode)
      stationCnt = SSDB.state.MCD.Stations.Count
      ReDim ListedStations(stationCnt)
      For stationIndex = 1 To stationCnt
'        If SSDB.state.MCD.Stations(stationIndex).HasData Then
          lstStations.List(stationIndex - 1) = SSDB.state.MCD.Stations(stationIndex).label
'        Else
'          lstStations.List(stationIndex - 1) = "*" & SSDB.state.MCD.Stations(stationIndex).label
'        End If
        Set ListedStations(stationIndex) = SSDB.state.MCD.Stations(stationIndex)
      Next
      SaveSetting "StreamStatsDB", "Defaults", SSDB.state.code & "SSMCD", SSDB.state.MCD.Name
    ElseIf IsHUC Then
      Set SSDB.state.HUCs = Nothing
      catCode = SSDB.state.HUCs(cboCategory.ListIndex + 1).code
      Set SSDB.state.HUC = SSDB.state.HUCs(catCode)
      stationCnt = SSDB.state.HUC.Stations.Count
      ReDim ListedStations(stationCnt)
      For stationIndex = 1 To stationCnt
'        If SSDB.state.HUC.Stations(stationIndex).HasData Then
          lstStations.List(stationIndex - 1) = SSDB.state.HUC.Stations(stationIndex).label
'        Else
'          lstStations.List(stationIndex - 1) = "*" & SSDB.state.HUC.Stations(stationIndex).label
'        End If
        Set ListedStations(stationIndex) = SSDB.state.HUC.Stations(stationIndex)
      Next
      SaveSetting "StreamStatsDB", "Defaults", SSDB.state.code & "SSHUC", SSDB.state.HUC.code
    ElseIf IsState Then
      stationCnt = SSDB.state.Stations.Count
      ReDim ListedStations(stationCnt)
      For stationIndex = 1 To stationCnt
'        If SSDB.state.Stations(stationIndex).HasData Then
          lstStations.List(stationIndex - 1) = SSDB.state.Stations(stationIndex).label
'        Else
'          lstStations.List(stationIndex - 1) = "*" & SSDB.state.Stations(stationIndex).label
'        End If
        Set ListedStations(stationIndex) = SSDB.state.Stations(stationIndex)
      Next
    End If
  End If
  SelectStations
  ChangeLabel
End Sub

Private Sub ChangesMade(madeChanges As Boolean)
  Dim row&, i&
  Dim OldVals() As String
  Dim myStat As ssStatistic
  Dim myStation As ssStation
  
  If tabMain.Tab = 0 Then
    ReDim OldVals(grdGenInfo.Rows, 1 To UBound(StationFields))
    For row = 1 To SSDB.state.SelStations.Count
      If Not SSDB.state.SelStations(row).IsNew Then
        Set myStation = SSDB.state.SelStations(row)
        OldVals(row, 1) = myStation.id
        OldVals(row, 2) = myStation.Name
        If Not myStation.StationType Is Nothing Then OldVals(row, 3) = myStation.StationType.Name
        OldVals(row, 4) = myStation.IsRegulated
        OldVals(row, 5) = myStation.Period
        OldVals(row, 6) = myStation.Directions
        OldVals(row, 7) = myStation.Remarks
        OldVals(row, 8) = myStation.Latitude
        OldVals(row, 9) = myStation.Longitude
        OldVals(row, 10) = SSDB.States.ItemByKey(myStation.DistrictCode).Abbrev
        OldVals(row, 11) = SSDB.States.ItemByKey(myStation.StateCode).Abbrev
        If myStation.countyCode <> "" Then
          OldVals(row, 12) = SSDB.state.Counties(myStation.countyCode).code & _
              "-" & SSDB.state.Counties(myStation.countyCode).Name
        End If
        If myStation.mcdCode <> "" Then
          OldVals(row, 13) = SSDB.state.MCDs(myStation.mcdCode).code & _
              "-" & SSDB.state.MCDs(myStation.mcdCode).Name
        End If
        OldVals(row, 14) = myStation.HUCCode
        If myStation.StatebasinCode <> "" Then
          OldVals(row, 15) = SSDB.state.StateBasins(myStation.StatebasinCode).code & _
              "-" & SSDB.state.StateBasins(myStation.StatebasinCode).Name
        End If
        Set myStation = Nothing
      End If
    Next row
  ElseIf tabMain.Tab = 1 Then
    ReDim OldVals(grdGenInfo.Rows, 1 To UBound(StatFields))
    For row = 1 To SSDB.SelStats.Count
      If Not SSDB.SelStats(row).IsNew Then
        OldVals(row, 1) = SSDB.SelStats(row).TypeName
        OldVals(row, 2) = SSDB.SelStats(row).code
        OldVals(row, 3) = SSDB.SelStats(row).Name
        OldVals(row, 4) = SSDB.Units(SSDB.SelStats(row).Units).englishlabel
        OldVals(row, 5) = SSDB.SelStats(row).Definition
        OldVals(row, 6) = SSDB.SelStats(row).Alias
      End If
    Next row
  End If
  RecordChanges OldVals(), madeChanges
End Sub

Private Sub RecordChanges(OldVals() As String, madeChanges As Boolean)
  Dim row&, col&, statCnt&, fldCnt&
  
  statCnt = grdGenInfo.Rows
  If tabMain.Tab = 0 Then fldCnt = UBound(StationFields) Else fldCnt = UBound(StatFields)
  If statCnt > 0 Then ReDim Changes(0 To 2, 1 To statCnt, 1 To fldCnt)
  For row = 1 To statCnt
    For col = 1 To fldCnt
      If grdGenInfo.TextMatrix(row, col - 1) <> OldVals(row, col) Then
        If Changes(0, row, col) <> "2" Then Changes(0, row, col) = "1"
        Changes(1, row, col) = OldVals(row, col)
        If col = 12 Or col = 13 Or col = 15 Then  'extract code from full identifier
          Changes(1, row, col) = StrSplit(OldVals(row, col), "-", "")
        End If
        madeChanges = True
      End If
      Changes(2, row, col) = grdGenInfo.TextMatrix(row, col - 1)
      If col = 12 Or col = 13 Or col = 15 Then  'extract code from full identifier
        Changes(2, row, col) = StrSplit(Changes(2, row, col), "-", "")
      End If
    Next col
  Next row
End Sub

Private Sub ResetClass(ClassName As String)
  Dim i&
  Dim catChoice$
  
  If ClassName = "Stations" Then
    If IsBasin Then
      Set SSDB.state.Statebasin.Stations = Nothing
    ElseIf IsCounty Then
      Set SSDB.state.County.Stations = Nothing
    ElseIf IsMCD Then
      Set SSDB.state.MCD.Stations = Nothing
    ElseIf IsHUC Then
      Set SSDB.state.HUC.Stations = Nothing
    End If
    Set SSDB.state.Stations = Nothing
    cboCategory_Click
  ElseIf ClassName = "Stats" Then
    Set SSDB.StatisticTypes = Nothing
    Set SSDB.SelStats = Nothing
    cboStatTypes_Click
  End If
End Sub

Private Sub SelectStations()
  Dim allStaIDs$, staID$
  Dim i&
  
  grdGenInfo.Rows = 0
  Set SSDB.state.SelStation = Nothing
  If IsBasin Then
    allStaIDs = GetSetting("StreamStatsDB", "Defaults", SSDB.state.code & "BasinStations")
  ElseIf IsCounty Then
    allStaIDs = GetSetting("StreamStatsDB", "Defaults", SSDB.state.code & "CountyStations")
  ElseIf IsMCD Then
    allStaIDs = GetSetting("StreamStatsDB", "Defaults", SSDB.state.code & "MCDStations")
  ElseIf IsHUC Then
    allStaIDs = GetSetting("StreamStatsDB", "Defaults", SSDB.state.code & "HUCStations")
  ElseIf IsState Then
    allStaIDs = GetSetting("StreamStatsDB", "Defaults", SSDB.state.code & "StateStations")
  End If
  While Len(Trim(allStaIDs)) > 0
    staID = StrRetRem(allStaIDs)
    For i = 1 To lstStations.ListCount
      If ListedStations(i).id = staID Then
        lstStations.Selected(i - 1) = True
        Exit For
      End If
    Next i
  Wend
  If SSDB.state.SelStations.Count = 0 Then
    Set SSDB.state.SelStations = Nothing
    Set SSDB.state.SelStation = Nothing
  Else
    Set SSDB.state.SelStation = SSDB.state.SelStations(1)
  End If
End Sub

Private Sub ResetGrid()
  Dim row&, col&
  Dim allStatIDs$, str$, mcdCode$, countyCode$, basinCode$
  
  If tabMain.Tab = 0 Then
    With grdGenInfo
      If SSDB.state Is Nothing Then Exit Sub
      .ClearData
      For row = 1 To SSDB.state.SelStations.Count
        If Not SSDB.state.SelStations(row).IsNew Then
          allStatIDs = allStatIDs & " " & SSDB.state.SelStations(row).id
          .TextMatrix(row, 0) = SSDB.state.SelStations(row).id
          .TextMatrix(row, 1) = SSDB.state.SelStations(row).Name
          If Not SSDB.state.SelStations(row).StationType Is Nothing Then _
              .TextMatrix(row, 2) = SSDB.state.SelStations(row).StationType.Name
          .TextMatrix(row, 3) = SSDB.state.SelStations(row).IsRegulated
          .TextMatrix(row, 4) = SSDB.state.SelStations(row).Period
          .TextMatrix(row, 5) = SSDB.state.SelStations(row).Directions
          .TextMatrix(row, 6) = SSDB.state.SelStations(row).Remarks
          If SSDB.state.SelStations(row).Latitude > 72# Then
            .TextMatrix(row, 7) = DMS2Decimal(CStr(SSDB.state.SelStations(row).Latitude))
          Else
            .TextMatrix(row, 7) = CStr(SSDB.state.SelStations(row).Latitude)
          End If
          If SSDB.state.SelStations(row).Longitude > 172# Then
            .TextMatrix(row, 8) = DMS2Decimal(CStr(SSDB.state.SelStations(row).Longitude))
          Else
            .TextMatrix(row, 8) = CStr(SSDB.state.SelStations(row).Longitude)
          End If
          .TextMatrix(row, 9) = SSDB.States.ItemByKey(SSDB.state.SelStations(row).DistrictCode).Abbrev
          .TextMatrix(row, 10) = SSDB.States.ItemByKey(SSDB.state.SelStations(row).StateCode).Abbrev
          countyCode = SSDB.state.SelStations(row).countyCode
          If SSDB.state.Counties.IndexFromKey(countyCode) > 0 Then
            .TextMatrix(row, 11) = SSDB.state.Counties(countyCode).code _
                        & "-" & SSDB.state.Counties(countyCode).Name
          End If
          mcdCode = SSDB.state.SelStations(row).mcdCode
          If SSDB.state.MCDs.IndexFromKey(mcdCode) > 0 Then
            .TextMatrix(row, 12) = SSDB.state.MCDs(mcdCode).code _
                        & "-" & SSDB.state.MCDs(mcdCode).Name
          End If
          .TextMatrix(row, 13) = SSDB.state.SelStations(row).HUCCode
          basinCode = SSDB.state.SelStations(row).StatebasinCode
          If SSDB.state.StateBasins.IndexFromKey(basinCode) > 0 Then
            .TextMatrix(row, 14) = SSDB.state.StateBasins(basinCode).code _
                & "-" & SSDB.state.StateBasins(basinCode).Name
          End If
        End If
      Next row
      .Rows = SSDB.state.SelStations.Count
      For col = 0 To .Cols - 1
        .ColEditable(col) = True
      Next col
      If .Rows > 0 Then
        cmdDelete.Enabled = True
        cmdEditStation.Enabled = True
      Else
        cmdDelete.Enabled = False
        cmdEditStation.Enabled = False
      End If
      If .row > 0 Then
        lblStaSel(0).Caption = .TextMatrix(grdGenInfo.row, 0)
        fraGenInfo.Caption = "Station Information - " & .TextMatrix(grdGenInfo.row, 1)
      Else
        lblStaSel(0).Caption = ""
        fraGenInfo.Caption = "Station Information"
      End If
      .ColsSizeByContents
    End With
    If IsState Then
      str = "State"
    Else
      str = lblCategory.Caption
      str = Mid(str, 1, Len(str) - 2)
    End If
    SaveSetting "StreamStatsDB", "Defaults", SSDB.state.code & str & "Stations", allStatIDs
  ElseIf tabMain.Tab = 1 Then
    With grdGenInfo
      .ClearData
      .Rows = SSDB.SelStats.Count
      For row = 1 To SSDB.SelStats.Count
        If Not SSDB.SelStats(row).IsNew Then
          allStatIDs = allStatIDs & " " & SSDB.SelStats(row).id
          .TextMatrix(row, 0) = SSDB.SelStats(row).TypeName
          .TextMatrix(row, 1) = SSDB.SelStats(row).code
          .TextMatrix(row, 2) = SSDB.SelStats(row).Name
          .TextMatrix(row, 3) = SSDB.Units(SSDB.SelStats(row).Units).englishlabel
          .TextMatrix(row, 4) = SSDB.SelStats(row).Definition
          .TextMatrix(row, 5) = SSDB.SelStats(row).Alias
        End If
      Next row
      .row = 1
      .col = 1
      If .Rows > 0 Then cmdDelete.Enabled = True Else cmdDelete.Enabled = False
      If .row > 0 Then
        lblStaSel(0).Caption = .TextMatrix(grdGenInfo.row, 2)
        fraGenInfo = "Statistic Information - " & .TextMatrix(grdGenInfo.row, 2)
      Else
        lblStaSel(0).Caption = ""
        fraGenInfo = "Statistic Information"
      End If
    End With
    SaveSetting "StreamStatsDB", "Defaults", cboStatTypes.ListIndex & "StatIDs", allStatIDs
  End If
  If grdGenInfo.Rows > 0 Then
    grdGenInfo.ColsSizeByContents
  Else
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdClear.Enabled = False
  End If
End Sub

Private Sub Initialize(TabIndex As Long)
  Dim fldCnt&
  Dim header$
  
  cmdDelete.Enabled = False
  cmdEditStation.Enabled = False
  
  If TabIndex = 0 Then
    StationFields(1) = "ID"
    StationFields(2) = "Name"
    StationFields(3) = "Type"
    StationFields(4) = "Regulated"
    StationFields(5) = "Period of Record"
    StationFields(6) = "Directions"
    StationFields(7) = "Remarks"
    StationFields(8) = "Latitude"
    StationFields(9) = "Longitude"
    StationFields(10) = "District"
    StationFields(11) = "State"
    StationFields(12) = "County"
    StationFields(13) = "MCD"
    StationFields(14) = "HUC"
    StationFields(15) = "Basin"
    With grdGenInfo
      .FixedRows = 2
      .Rows = lstStations.SelCount
      .ColEditable(0) = False
      .Cols = UBound(StationFields)
      For fldCnt = 1 To UBound(StationFields)
        header = StrRetRem(StationFields(fldCnt))
        If Len(StationFields(fldCnt)) > 0 Then
          .TextMatrix(-1, fldCnt - 1) = (header)
          .TextMatrix(0, fldCnt - 1) = (StationFields(fldCnt))
        Else
          .TextMatrix(-1, fldCnt - 1) = ""
          .TextMatrix(0, fldCnt - 1) = (header)
        End If
        If .Rows = 0 Then .ColWidth(fldCnt - 1) = Len(.TextMatrix(0, fldCnt - 1)) * 200
      Next fldCnt
      .ColType(7) = ATCoSng
      .ColMin(7) = 17.5
      .ColMax(7) = 72#
      .ColType(8) = ATCoSng
      .ColMin(8) = -172#
      .ColMax(8) = -64#
    End With
  ElseIf TabIndex = 1 Then
    StatFields(1) = "Type"
    StatFields(2) = "Code"
    StatFields(3) = "Full Name"
    StatFields(4) = "Units"
    StatFields(5) = "Definition"
    StatFields(6) = "Alias"
    With grdGenInfo
      .Clear
      .FixedRows = 1
      .Rows = lstStats.SelCount
      .Cols = UBound(StatFields)
      .ColEditable(0) = True
      For fldCnt = 1 To UBound(StatFields)
        .TextMatrix(0, fldCnt - 1) = StatFields(fldCnt)
        If .Rows = 0 Then .ColWidth(fldCnt - 1) = Len(.TextMatrix(0, fldCnt - 1)) * 200
      Next fldCnt
    End With
  End If
End Sub

Private Sub lstStats_Click()
  Dim row&, col&
  Dim allStatIDs$
  Dim lKey As String
  
  On Error GoTo y

  If lstStats.Selected(lstStats.ListIndex) Then   'adding Stat to grid
    lKey = CStr(ListedStats(lstStats.ListIndex + 1).id)
    If Not SSDB.SelStats.KeyExists(lKey) Then
      SSDB.SelStats.Add ListedStats(lstStats.ListIndex + 1), lKey
      ResetGrid
      If tabMain.Tab = 1 Then cmdDelete.Enabled = True
      If grdGenInfo.Rows > 0 Then
        cmdSave.Enabled = True
        cmdCancel.Enabled = True
        cmdClear.Enabled = True
      Else
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        cmdClear.Enabled = False
      End If
    End If
  ElseIf SSDB.SelStats.Count > 0 Then 'try removing Stat from grid
    With grdGenInfo
      row = 1
      While SSDB.SelStats(row).id <> ListedStats(lstStats.ListIndex + 1).id
        allStatIDs = allStatIDs & " " & SSDB.SelStats(row).id
        row = row + 1
      Wend
      SSDB.SelStats.Remove row
      For row = row To .Rows - 1
        allStatIDs = allStatIDs & " " & SSDB.SelStats(row).id
        For col = 0 To .Cols - 1
          .TextMatrix(row, col) = .TextMatrix(row + 1, col)
        Next col
      Next row
      .Rows = .Rows - 1
      .row = 1
      If .row > 0 Then
        lblStaSel(0).Caption = .TextMatrix(grdGenInfo.row, 2)
        fraGenInfo = "Statistic Information - " & .TextMatrix(grdGenInfo.row, 2)
      Else
        lblStaSel(0).Caption = ""
        fraGenInfo = "Statistic Information"
      End If
      If .Rows > 0 Then
        cmdSave.Enabled = True
        cmdCancel.Enabled = True
        cmdClear.Enabled = True
      Else
        cmdDelete.Enabled = False
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        cmdClear.Enabled = False
      End If
    End With
    SaveSetting "StreamStatsDB", "Defaults", cboStatTypes.ListIndex & "StatIDs", allStatIDs
  End If
y:
End Sub

Private Sub mnuDatabase_Click()
  Dim state$
  GetDatabaseFilename (True)
  state = GetSetting("StreamStatsDB", "Defaults", "ssState")
  InitializeFromDatabase (state)
End Sub

Private Sub mnuExit_Click()
  ShutErDown
End Sub

Private Sub mnuHelpAbout_Click()
  frmAbout.Show
End Sub

Private Sub mnuHelpManual_Click()
  OpenHelp
End Sub

Private Sub rdoStaOpt_Click(Index As Integer)
  ChangeLabel
End Sub

Private Sub ChangeLabel()
  Dim staNum&
  Dim staID$
  
  Me.MousePointer = vbHourglass
  For staNum = 1 To lstStations.ListCount
    staID = lstStations.ItemData(staNum - 1)
    If rdoStaOpt(0) Then
      ListedStations(staNum).label = ListedStations(staNum).Name
    ElseIf rdoStaOpt(1) Then
      ListedStations(staNum).label = ListedStations(staNum).id
    ElseIf rdoStaOpt(2) Then
      ListedStations(staNum).label = _
          ListedStations(staNum).id & "-" & ListedStations(staNum).Name
    End If
'    If ListedStations(staNum).HasData Then
      lstStations.List(staNum - 1) = ListedStations(staNum).label
'    Else
'      lstStations.List(staNum - 1) = "*" & ListedStations(staNum).label
'    End If
  Next staNum
  Me.MousePointer = vbDefault
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
  If PreviousTab = 1 Then  'On Stations tab
    If cboState.ListIndex >= 0 Then
      cmdAdd.Enabled = True
    Else
      cmdAdd.Enabled = False
    End If
    lblStaOpts.Visible = True
    rdoStaOpt(0).Visible = True
    rdoStaOpt(1).Visible = True
    rdoStaOpt(2).Visible = True
    cmdEditStation.Visible = True
    lblStaSel(1).Caption = "Active Station:"
    lblStaSel(0).Width = 1335
    fraGenInfo.Caption = "Station Information"
    fraStatButtons.Caption = "Station Commands"
    Initialize 0
    ResetGrid
  ElseIf PreviousTab = 0 Then  'On Statistics tab
    cmdAdd.Enabled = True
    lblStaOpts.Visible = False
    rdoStaOpt(0).Visible = False
    rdoStaOpt(1).Visible = False
    rdoStaOpt(2).Visible = False
    cmdEditStation.Visible = False
    lblStaSel(1).Caption = "Active Statistic:"
    lblStaSel(0).Width = 2175
    fraGenInfo.Caption = "Statistic Information"
    fraStatButtons.Caption = "Statistic Commands"
    Initialize 1
    ResetGrid
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  IPC.SendMonitorMessage "(EXIT)"
End Sub

Private Sub OpenHelp()
  Static HelpFilePath As String
  
  If Len(HelpFilePath) = 0 Then
    Dim ff As ATCoFindFile
    Set ff = New ATCoFindFile
    ff.SetRegistryInfo "StreamStatsDB", "files", "StreamStatsDB.chm"
    ff.SetDialogProperties "Please locate help file StreamStatsDB.chm", ExePath & "StreamStatsDB.chm"
    HelpFilePath = ff.GetName
    Set ff = Nothing
  End If
  
  If Len(HelpFilePath) > 0 Then
    If Len(Dir(HelpFilePath)) > 0 Then
      OpenFile HelpFilePath, frmCDLG.CDLG
    Else
      MsgBox "Could not find help file '" & HelpFilePath & "'"
    End If
  Else
    MsgBox "Help file not available"
  End If

End Sub

Private Sub ShutErDown()
  On Error GoTo x
  If Len(Dir(fileTitle)) > 0 Then SSDB.DB.Close
x:
  Unload Me

End Sub

Private Sub InitializeFromDatabase(aState$)
  Dim tmpIndex&, selState&, selStat&
  Dim StatType$

  'Populate state listbox on tab 1
  cboState.Clear
  For tmpIndex = 1 To SSDB.States.Count
    cboState.List(tmpIndex - 1) = SSDB.States(tmpIndex).Name
    cboState.ItemData(tmpIndex - 1) = SSDB.States(tmpIndex).code
    If SSDB.States(tmpIndex).Name = aState Then selState = tmpIndex
  Next
  If selState > 0 Then
    cboState.ListIndex = selState - 1
  End If
  rdoStaOpt(0) = True
  
  'Populate Statistic Type listbox on tab 2
  StatType = GetSetting("StreamStatsDB", "Defaults", "ssStatType")
  For tmpIndex = 1 To SSDB.StatisticTypes.Count
    cboStatTypes.List(tmpIndex - 1) = SSDB.StatisticTypes(tmpIndex).Name
    cboStatTypes.ItemData(tmpIndex - 1) = SSDB.StatisticTypes(tmpIndex).id
    If SSDB.StatisticTypes(tmpIndex).Name = StatType Then selStat = tmpIndex
  Next tmpIndex
  If selStat > 0 Then
    cboStatTypes.ListIndex = selStat - 1
  End If

End Sub
