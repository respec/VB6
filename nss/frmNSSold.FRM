VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNSSold 
   Caption         =   "National Flood Frequency Program (NSS)"
   ClientHeight    =   6480
   ClientLeft      =   936
   ClientTop       =   1608
   ClientWidth     =   7704
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmNSSold.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6480
   ScaleWidth      =   7704
   Begin VB.Frame fraManageEstimate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   252
      Index           =   1
      Left            =   3960
      TabIndex        =   32
      Top             =   570
      Width           =   3732
      Begin VB.CommandButton cmdEstimate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   ">"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   10
         Top             =   0
         Width           =   300
      End
      Begin VB.CommandButton cmdEstimate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "<"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   9
         Top             =   0
         Width           =   300
      End
      Begin VB.CommandButton cmdCompute 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "New"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   12
         Top             =   0
         Width           =   732
      End
      Begin VB.CommandButton cmdDelEstimate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Delete"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   11
         Top             =   0
         Width           =   852
      End
      Begin VB.Label lblEstimates 
         BackStyle       =   0  'Transparent
         Caption         =   "of 1"
         Height          =   252
         Index           =   5
         Left            =   1560
         TabIndex        =   36
         Top             =   30
         Width           =   732
      End
      Begin VB.Label lblEstimates 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   252
         Index           =   4
         Left            =   840
         TabIndex        =   35
         Top             =   30
         Width           =   372
      End
      Begin VB.Label lblEstimates 
         BackStyle       =   0  'Transparent
         Caption         =   "&Urban"
         Height          =   252
         Index           =   3
         Left            =   0
         TabIndex        =   34
         Top             =   30
         Width           =   612
      End
   End
   Begin VB.Frame fraBottom 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   372
      Left            =   120
      TabIndex        =   33
      Top             =   6000
      Width           =   4812
      Begin VB.CommandButton cmdFrequency 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fre&quency Plot"
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1692
      End
      Begin VB.CommandButton cmdHydrograph 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "H&ydrograph"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   0
         Width           =   1692
      End
      Begin VB.CommandButton cmdWeight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Weight"
         Height          =   375
         Left            =   3600
         TabIndex        =   19
         Top             =   0
         Width           =   1092
      End
   End
   Begin VB.ListBox lstUrbEstimate 
      Height          =   1968
      ItemData        =   "frmNSSold.frx":030A
      Left            =   3960
      List            =   "frmNSSold.frx":0311
      TabIndex        =   14
      Top             =   840
      Width           =   3732
   End
   Begin VB.Frame fraManageEstimate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   570
      Width           =   3732
      Begin VB.CommandButton cmdEstimate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "<"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   0
         Width           =   300
      End
      Begin VB.CommandButton cmdEstimate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   ">"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   6
         Top             =   0
         Width           =   300
      End
      Begin VB.CommandButton cmdDelEstimate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Delete"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   7
         Top             =   0
         Width           =   852
      End
      Begin VB.CommandButton cmdCompute 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "New"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   8
         Top             =   0
         Width           =   732
      End
      Begin VB.Label lblEstimates 
         BackStyle       =   0  'Transparent
         Caption         =   "&Rural"
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   30
         Width           =   612
      End
      Begin VB.Label lblEstimates 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   252
         Index           =   1
         Left            =   840
         TabIndex        =   38
         Top             =   30
         Width           =   372
      End
      Begin VB.Label lblEstimates 
         BackStyle       =   0  'Transparent
         Caption         =   "of 1"
         Height          =   252
         Index           =   2
         Left            =   1560
         TabIndex        =   37
         Top             =   30
         Width           =   732
      End
   End
   Begin VB.Frame sash 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   120
      MousePointer    =   7  'Size N S
      TabIndex        =   21
      Top             =   2880
      Width           =   7572
      Begin VB.Label lblUrbTab 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   4
         Left            =   5520
         TabIndex        =   31
         Top             =   630
         Width           =   252
      End
      Begin VB.Label lblUrbTab 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ft"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   3
         Left            =   5400
         TabIndex        =   30
         Top             =   720
         Width           =   252
      End
      Begin VB.Label lblRurTab 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   4
         Left            =   1320
         TabIndex        =   29
         Top             =   636
         Width           =   132
      End
      Begin VB.Label lblRurTab 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ft"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   3
         Left            =   1080
         TabIndex        =   28
         Top             =   720
         Width           =   252
      End
      Begin VB.Label lblRurTab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Recurrence    Peak,    Standard    Equivalent"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   0
         TabIndex        =   27
         Top             =   480
         Width           =   4092
      End
      Begin VB.Label lblRurTab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Interval, yrs          /s    Error, %      Years"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   2
         Left            =   0
         TabIndex        =   26
         Top             =   720
         Width           =   4092
      End
      Begin VB.Label lblUrbTab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Recurrence       Peak,      Standard"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   4200
         TabIndex        =   25
         Top             =   480
         Width           =   3252
      End
      Begin VB.Label lblUrbTab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Interval, yrs            /s        Error, %"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   2
         Left            =   4200
         TabIndex        =   24
         Top             =   720
         Width           =   3252
      End
      Begin VB.Label lblRurTab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rural flood event estimates"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Top             =   120
         Width           =   4092
      End
      Begin VB.Label lblUrbTab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Urban flood event estimates"
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   4200
         TabIndex        =   22
         Top             =   120
         Width           =   3252
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   22222
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.ListBox lstRurEstimate 
      Height          =   1968
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   3732
   End
   Begin VB.ListBox lstUrbRes 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1884
      ItemData        =   "frmNSSold.frx":0325
      Left            =   4320
      List            =   "frmNSSold.frx":032C
      TabIndex        =   16
      Top             =   3960
      Width           =   3255
   End
   Begin VB.ListBox lstRurRes 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1884
      ItemData        =   "frmNSSold.frx":033B
      Left            =   120
      List            =   "frmNSSold.frx":033D
      TabIndex        =   15
      Top             =   3960
      Width           =   4092
   End
   Begin VB.ComboBox cboState 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "-999"
      Top             =   70
      Width           =   1812
   End
   Begin VB.TextBox txtBasin 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog CMDialog2 
      Left            =   9720
      Top             =   4440
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      FontSize        =   2.54052e-29
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   9720
      Top             =   3960
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      FontSize        =   2.54052e-29
   End
   Begin VB.Label lblState 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&State:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label lblBasinName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Site &Name:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSavStatus 
         Caption         =   "Save &Status"
      End
      Begin VB.Menu mnuSavReport 
         Caption         =   "Save &Report"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuGraph 
      Caption         =   "&Graph"
      Begin VB.Menu mnuFrequency 
         Caption         =   "Fre&quency"
      End
      Begin VB.Menu mnuHydrograph 
         Caption         =   "H&ydrograph"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpCont 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpObtech 
         Caption         =   "&USGS NSS Web Page"
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "&Program Trace"
      End
      Begin VB.Menu mnuHelpSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About NSS..."
      End
   End
End
Attribute VB_Name = "frmNSSold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SashDragging As Boolean
Private lblUrbTabOffset() As Long

Public Sub DbgMsg(str$, lvl&, modul$, typ$)
  Debug.Print str, lvl, "NSS" & modul, typ
End Sub

Private Sub cboState_Click()
  Dim inifg&, stIndex&
  Dim myState As nssState

  If Project.State.id <> cboState.ListIndex Then
    If cboState.ListIndex >= 0 And _
        Project.RuralScenarios.Count + Project.UrbanScenarios.Count > 0 Then
      inifg = MsgBox("Changing the state will clear any existing estimates." & _
          Chr(13) & Chr(10) & "Are you sure you want to change the selected state?", 1)
      If inifg = 1 Then
        'change state, initialize estimates
        Set myState = Project.DB.States(cboState.ListIndex)
        Set Project.State = myState
      Else
        'keep current state
        cboState.ListIndex = Project.State.id
      End If
    Else
      'keep current state
      cboState.ListIndex = Project.State.id
    End If
  End If

End Sub

Private Sub cmdCompute_Click(Index As Integer)
    
    'Dim i&, j&, rtmp!

    'perform computations for selected regions
    'If RegUseCnt <= 0 Then
    '  MsgBox "At least one region must be selected for use in order to perform computations.", 48
    'ElseIf rurfg = False And RurReq And compfg < 1 Then
    '  MsgBox "Rural calculations must be performed prior to using this urban equation.", 48
'    ElseIf total_area <= 0 Then
'      MsgBox "A value for Basin Drainage Area must be entered before computations may be performed.", 48
    'ElseIf rurcnt >= MAX_Estimate Or urbcnt >= MAX_Estimate Then
    '  MsgBox "A maximum of " & CStr(MAX_Estimate) & " estimates may be generated.  To generate a new estimate, you must first delete an existing estimate.", 48
    'Else
      'regions selected for computations
'      If Index = 0 Then rurfg = True Else rurfg = False
      cmdWeight.Enabled = False
      frmCompute.Show 1
    'End If

End Sub

Private Sub cmdDelEstimate_Click(Index As Integer)
'  Dim i&, j&, k&
'  If Index = 0 Then rurfg = True Else rurfg = False
'  If ConfDelEstimate = 1 Then
'    'user confirmed delete
'    If rurfg = True Then 'deleting rural estimate
'      'remove next estimate info first
'      For i = 0 To rurscn(rurcnt).rcnt - 1
'        rurscn(rurcnt).reg(i) = -1
'      Next i
'      rurscn(rurcnt).rcnt = 0
'      rurcnt = rurcnt - 1
'      For i = rurind To rurcnt - 1
'        'adjust positions in estimate array
'        rurscn(i) = rurscn(i + 1)
'        'adjust positions in discharge array
'        For j = 0 To 2
'          For k = 0 To MAX_INTERVALS
'            rural_discharge(j, k, i) = rural_discharge(j, k, i + 1)
'          Next k
'        Next j
'      Next i
'      're-init last position in estimate array
'      rurscn(rurcnt).Name = ""
'      rurscn(rurcnt).rcnt = 0
'      For j = 0 To MAX_REG_USE
'        For k = 0 To MAX_REGR_COMPONENTS
'          rurscn(rurcnt).v(j).Value(k) = -999
'          rurscn(rurcnt).v(j).vname(k) = -999
'        Next k
'        rurscn(rurcnt).v(j).vcount = 0
'        rurscn(rurcnt).reg(j) = -1
'      Next j
'      're-init last position in discharge array
'      For j = 0 To 2
'        For k = 0 To MAX_INTERVALS
'          rural_discharge(j, k, rurcnt) = 0
'        Next k
'      Next j
'      If rurind >= rurcnt Then 'index too high now
'        rurind = rurcnt - 1
'      End If
'      'If rurcnt > 0 Then 'set intervals for current estimate
'      '  Call SetIntervals
'      'End If
'      'need to clear results table
'      Call InitRuralDis
'      'adjust estimate range
'      If rurcnt > 0 Then 'at least one rural estimate left
'        i = 0
'        While i < urbcnt
'          If rurind < urbscn(i).rscn Then
'            'urban estimate using rural estimate which has been renumbered,
'            'adjust urban estimate rural estimate number
'            urbscn(i).rscn = urbscn(i).rscn - 1
'            'adjust urban estimate title
'            urbscn(i).Name = "National Urban Equations (R" & urbscn(i).rscn + 1 & ")"
'          End If
'          i = i + 1
'        Wend
'      Else
'        'no estimates
'        rurind = -1
'        Call frmNSS.SetEstimate
'        'set compute flag to indicate no rural calculations
'        compfg = -1
'      End If
'    Else
'      'deleting urban estimate, urban index offset by rural count
'      urbcnt = urbcnt - 1
'      For i = urbind To urbcnt - 1
'        'adjust positions in estimate array
'        urbscn(i) = urbscn(i + 1)
'        'adjust positions in discharge array
'        For j = 0 To MAX_INTERVALS
'          urban_discharge(j, i) = urban_discharge(j, i + 1)
'        Next j
'      Next i
'      're-init last position in estimate array
'      urbscn(urbcnt).National = False
'      urbscn(urbcnt).Name = ""
'      urbscn(urbcnt).v.vcount = 0
'      urbscn(urbcnt).reg = -1
'      For j = 0 To MAX_REGR_COMPONENTS
'        urbscn(urbcnt).v.Value(j) = -999
'        urbscn(urbcnt).v.vname(j) = ""
'      Next j
'      're-init last position in discharge array
'      For j = 0 To MAX_INTERVALS
'        urban_discharge(j, urbcnt) = 0
'      Next j
'      'need to clear results table
'      Call InitUrbanDis
'      If urbind > urbcnt - 1 Then
'        'index too high now
'        urbind = urbcnt - 1
'      End If
'      'adjust estimate range
'      If urbcnt = 0 Then
'        'no urban estimates
'        urbind = -1
'        SetEstimate
'        'disable delete button
'      End If
'    End If
'    DispEstimate
'    SetCmdEnabled
'  End If
End Sub

Private Sub cmdEstimate_Click(Index As Integer)

'  If Index <= 2 Then 'a rural button was pressed
'    rurfg = True
'    If rurcnt < 1 Then
'      rurind = -1
'    Else
'      rurind = rurind + Index - 1
'      If rurind >= rurcnt Then
'        rurind = 0
'      ElseIf rurind < 0 Then
'        rurind = rurcnt - 1
'      End If
'    End If
'    Call DispRuralDis
'  Else 'an urban button was pressed
'    rurfg = False
'    If urbcnt < 1 Then
'      urbind = -1
'    Else
'      urbind = urbind + Index - 4
'      If urbind >= urbcnt Then
'        urbind = 0
'      ElseIf urbind < 0 Then
'        urbind = urbcnt - 1
'      End If
'    End If
'    Call DispUrbanDis
'  End If
'
'  Call SetEstimate

End Sub

Private Sub cmdFrequency_Click()
  frmFreq.Show
End Sub

Private Sub cmdHydrograph_Click()
  frmHyd.Show
End Sub

Private Sub cmdWeight_Click()
  frmWeight.Show 1
End Sub

Private Sub Form_Load()
'  Dim i&
'  Dim State As Variant
'  Dim EXEpath As String
'
''  Set myMsgBox = New ATCoMessage
'
'  ReDim lblUrbTabOffset(lblUrbTab.Count)
'  For i = 0 To lblUrbTab.Count - 1
'    lblUrbTabOffset(i) = lblUrbTab(i).Left - lblUrbTab(0).Left
'  Next i
'
''    Dim fname$, fpath$, ffound$
''    fpath = "c:\"  ' Set the initial path.
''    fname = "VB_NSS.DLL"
''    ffound = DirSubs(fpath, fname)
''    If ffound = "" Then
''      'did not find file, big problem
''      MsgBox "Unable to locate the file VB_NSS.DLL.  Will not be able to perform calculations.", 16
''    End If
'    App.HelpFile = "nss.hlp"
''    lstRegion.HelpContextID = National_Urban
''    cboState.HelpContextID = State_Map
'
'    EXEpath = "C:\NSS\UpdatedDB\"  ' will
'    Set nssDB = New nssDatabase
'    nssDB.Filename = EXEpath & "NSS.mdb"
'
'    frmNSS.HelpContextID = 28
'    frmWeight.HelpContextID = 30
''    Call GetStateNames
'    For Each State In nssDB.States
'      frmNSS.cboState.AddItem State.Name
'    Next State
'    stind = -1
'    InitAll
'    'default units to metric
'    metric = 1
'    'read default status file
'    Call StatusGet(0)
'    frmStart.Show 1
'    CalculateAll
'    If metric Then
'      lblRurTab(3).Caption = "m"
'      lblUrbTab(3).Caption = "m"
'    Else
'      lblRurTab(3).Caption = "ft"
'      lblUrbTab(3).Caption = "ft"
'    End If
'    'lstRegion.Height = fraRegions.Height - lstRegion.Top
'    'lstUseRegion.Height = fraUseRegion.Height - lstUseRegion.Top
'    'adjust estimate scroll bar
'    'vsbScen.Top = 95
'    'vsbScen.Height = 940
'    'adjust estimate list box height
'    'lstEstimates.Height = fraScen.Height - lstEstimates.Top
'    'default to rural calculations
'    'optTabType(0).value = True
''    ChDir "\nss\vb"
'    'cmdWeight.Enabled = True
'
End Sub

Private Sub InitAll()

'    'init national urban equation interval values
'    NatUrbEqInts(0) = 2
'    NatUrbEqInts(1) = 5
'    NatUrbEqInts(2) = 10
'    NatUrbEqInts(3) = 25
'    NatUrbEqInts(4) = 50
'    NatUrbEqInts(5) = 100
'    NatUrbEqInts(6) = 500
'    'init urban components
'    urban_comp(0).variable_name = "BDF"
'    urban_comp(0).descriptor = "Basin development factor"
'    urban_comp(0).minimum = 0
'    urban_comp(0).maximum = 12
'    urban_comp(0).units = 0
'    urban_comp(1).variable_name = "IA"
'    urban_comp(1).descriptor = "Impervious Surfaces"
'    urban_comp(1).minimum = 3
'    urban_comp(1).maximum = 50
'    urban_comp(1).units = 10
'    urban_comp(2).variable_name = "SL"
'    urban_comp(2).descriptor = "Channel slope"
'    urban_comp(2).minimum = 3
'    urban_comp(2).maximum = 70
'    urban_comp(2).units = 5
'    urban_comp(3).variable_name = "RI2"
'    urban_comp(3).descriptor = "2-hour 2-year rainfall intensity"
'    urban_comp(3).minimum = 0.2
'    urban_comp(3).maximum = 2.8
'    urban_comp(3).units = 2
'    urban_comp(4).variable_name = "ST"
'    urban_comp(4).descriptor = "Basin storage"
'    urban_comp(4).minimum = 0
'    urban_comp(4).maximum = 11
'    urban_comp(4).units = 10
'    urban_eqtn(0).standard_error = 38
'    urban_eqtn(1).standard_error = 37
'    urban_eqtn(2).standard_error = 38
'    urban_eqtn(3).standard_error = 40
'    urban_eqtn(4).standard_error = 42
'    urban_eqtn(5).standard_error = 44
'    urban_eqtn(6).standard_error = 49
'    urban_eqtn(0).regr_constant = 2.35
'    urban_eqtn(1).regr_constant = 2.7
'    urban_eqtn(2).regr_constant = 2.99
'    urban_eqtn(3).regr_constant = 2.78
'    urban_eqtn(4).regr_constant = 2.67
'    urban_eqtn(5).regr_constant = 2.5
'    urban_eqtn(6).regr_constant = 2.27
'    urban_eqtn(0).a_exponent = 0.41
'    urban_eqtn(1).a_exponent = 0.35
'    urban_eqtn(2).a_exponent = 0.32
'    urban_eqtn(3).a_exponent = 0.31
'    urban_eqtn(4).a_exponent = 0.29
'    urban_eqtn(5).a_exponent = 0.29
'    urban_eqtn(6).a_exponent = 0.29
'    urban_eqtn(0).r_exponent = 0.47
'    urban_eqtn(1).r_exponent = 0.54
'    urban_eqtn(2).r_exponent = 0.58
'    urban_eqtn(3).r_exponent = 0.6
'    urban_eqtn(4).r_exponent = 0.62
'    urban_eqtn(5).r_exponent = 0.63
'    urban_eqtn(6).r_exponent = 0.63
'    urban_eqtn(0).b_exponent = -0.32
'    urban_eqtn(1).b_exponent = -0.31
'    urban_eqtn(2).b_exponent = -0.3
'    urban_eqtn(3).b_exponent = -0.29
'    urban_eqtn(4).b_exponent = -0.28
'    urban_eqtn(5).b_exponent = -0.28
'    urban_eqtn(6).b_exponent = -0.27
'    urban_eqtn(0).regr_component(0).base_modifier = 0
'    urban_eqtn(1).regr_component(0).base_modifier = 0
'    urban_eqtn(2).regr_component(0).base_modifier = 0
'    urban_eqtn(3).regr_component(0).base_modifier = 0
'    urban_eqtn(4).regr_component(0).base_modifier = 0
'    urban_eqtn(5).regr_component(0).base_modifier = 0
'    urban_eqtn(6).regr_component(0).base_modifier = 0
'    urban_eqtn(0).regr_component(1).base_modifier = 0
'    urban_eqtn(1).regr_component(1).base_modifier = 0
'    urban_eqtn(2).regr_component(1).base_modifier = 0
'    urban_eqtn(3).regr_component(1).base_modifier = 0
'    urban_eqtn(4).regr_component(1).base_modifier = 0
'    urban_eqtn(5).regr_component(1).base_modifier = 0
'    urban_eqtn(6).regr_component(1).base_modifier = 0
'    urban_eqtn(0).regr_component(2).base_modifier = 3
'    urban_eqtn(1).regr_component(2).base_modifier = 3
'    urban_eqtn(2).regr_component(2).base_modifier = 3
'    urban_eqtn(3).regr_component(2).base_modifier = 3
'    urban_eqtn(4).regr_component(2).base_modifier = 3
'    urban_eqtn(5).regr_component(2).base_modifier = 3
'    urban_eqtn(6).regr_component(2).base_modifier = 3
'    urban_eqtn(0).regr_component(3).base_modifier = 8
'    urban_eqtn(1).regr_component(3).base_modifier = 8
'    urban_eqtn(2).regr_component(3).base_modifier = 8
'    urban_eqtn(3).regr_component(3).base_modifier = 8
'    urban_eqtn(4).regr_component(3).base_modifier = 8
'    urban_eqtn(5).regr_component(3).base_modifier = 8
'    urban_eqtn(6).regr_component(3).base_modifier = 8
'    urban_eqtn(0).regr_component(0).base_exponent = 0.15
'    urban_eqtn(1).regr_component(0).base_exponent = 0.11
'    urban_eqtn(2).regr_component(0).base_exponent = 0.09
'    urban_eqtn(3).regr_component(0).base_exponent = 0.07
'    urban_eqtn(4).regr_component(0).base_exponent = 0.06
'    urban_eqtn(5).regr_component(0).base_exponent = 0.06
'    urban_eqtn(6).regr_component(0).base_exponent = 0.05
'    urban_eqtn(0).regr_component(1).base_exponent = 0.17
'    urban_eqtn(1).regr_component(1).base_exponent = 0.16
'    urban_eqtn(2).regr_component(1).base_exponent = 0.15
'    urban_eqtn(3).regr_component(1).base_exponent = 0.15
'    urban_eqtn(4).regr_component(1).base_exponent = 0.15
'    urban_eqtn(5).regr_component(1).base_exponent = 0.15
'    urban_eqtn(6).regr_component(1).base_exponent = 0.16
'    urban_eqtn(0).regr_component(2).base_exponent = 2.04
'    urban_eqtn(1).regr_component(2).base_exponent = 1.86
'    urban_eqtn(2).regr_component(2).base_exponent = 1.75
'    urban_eqtn(3).regr_component(2).base_exponent = 1.76
'    urban_eqtn(4).regr_component(2).base_exponent = 1.74
'    urban_eqtn(5).regr_component(2).base_exponent = 1.76
'    urban_eqtn(6).regr_component(2).base_exponent = 1.86
'    urban_eqtn(0).regr_component(3).base_exponent = -0.65
'    urban_eqtn(1).regr_component(3).base_exponent = -0.59
'    urban_eqtn(2).regr_component(3).base_exponent = -0.57
'    urban_eqtn(3).regr_component(3).base_exponent = -0.55
'    urban_eqtn(4).regr_component(3).base_exponent = -0.53
'    urban_eqtn(5).regr_component(3).base_exponent = -0.52
'    urban_eqtn(6).regr_component(3).base_exponent = -0.54
'
''This array establishes the units descriptors
''and conversions used by the program for
''conversion between English and SI units
'
'    'SI units string, English units string, Multipier Code
'    units(0).SI_sym = ""
'    units(0).IP_sym = ""
'    units(0).factor = 1#
'
'    units(1).SI_sym = "km2"
'    units(1).IP_sym = "mi2"
'    units(1).factor = 2.59
'
'    units(2).SI_sym = "mm"
'    units(2).IP_sym = "in"
'    units(2).factor = 25.4
'
'    units(3).SI_sym = "deg C"
'    units(3).IP_sym = "deg F"
'    units(3).factor = 0.5555
'
'    units(4).SI_sym = "km"
'    units(4).IP_sym = "thousands of feet"
'    units(4).factor = 0.3048
'
'    units(5).SI_sym = "m/km"
'    units(5).IP_sym = "ft/mi"
'    units(5).factor = 0.1894
'
'    units(6).SI_sym = "km"
'    units(6).IP_sym = "mi"
'    units(6).factor = 1.609
'
'    units(7).SI_sym = "m"
'    units(7).IP_sym = "ft"
'    units(7).factor = 0.3048
'
'    units(8).SI_sym = "mm/h"
'    units(8).IP_sym = "in/hr"
'    units(8).factor = 25.4
'
'    units(9).SI_sym = "km/km"
'    units(9).IP_sym = "mi/mi"
'    units(9).factor = 1#
'
'    units(10).SI_sym = "%"
'    units(10).IP_sym = "%"
'    units(10).factor = 1#
'
'    units(11).SI_sym = "degrees"
'    units(11).IP_sym = "degrees"
'    units(11).factor = 1#
'
'    units(12).SI_sym = "decimal fraction"
'    units(12).IP_sym = "decimal fraction"
'    units(12).factor = 1#
'
'    units(13).SI_sym = "m3/s"
'    units(13).IP_sym = "ft3/s"
'    units(13).factor = 0.02832
'
'    units(14).SI_sym = "m/m"
'    units(14).IP_sym = "feet per foot"
'    units(14).factor = 1#
'
'    units(15).SI_sym = "# 1st-order streams/km2"
'    units(15).IP_sym = "# 1st-order streams/mi2"
'    units(15).factor = 0.386102
'
'    units(16).SI_sym = "km/km2"
'    units(16).IP_sym = "mi/mi2"
'    units(16).factor = 0.6212
'
'    units(17).SI_sym = "persons/km2"
'    units(17).IP_sym = "persons/mi2"
'    units(17).factor = 0.3861
'
'    units(18).SI_sym = "decimal degrees"
'    units(18).IP_sym = "decimal degrees"
'    units(18).factor = 1#
'
'    units(19).SI_sym = "km2/km2"
'    units(19).IP_sym = "mi2/mi2"
'    units(19).factor = 1#
'
End Sub

Private Sub Form_Resize()
  Dim y&, i&
  y = Height - 1110
  If y > sash.Top + sash.Height + 220 Then
    fraBottom.Top = y
    'moved labels into sash
    'y = sash.Top + 120
    'lblRurTab(0).Top = y
    'lblUrbTab(0).Top = y
    'y = y + 360
    'lblRurTab(0).Top = y
    'lblUrbTab(1).Top = y
    'y = y + 240
    'lblRurTab(2).Top = y
    'lblUrbTab(2).Top = y
    
    lstRurRes.Top = sash.Top + sash.Height + 60
    lstUrbRes.Top = lstRurRes.Top
    
    y = y - lstRurRes.Top - 156
    If y > 251 Then
      lstRurRes.Height = y
      lstUrbRes.Height = y
    End If
  End If
  
  If Width > 4000 Then
    sash.Width = Width - 375

    lstRurEstimate.Width = (Width - 495) / 2
    lstUrbEstimate.Width = lstRurEstimate.Width
    lstUrbEstimate.Left = lstRurEstimate.Left * 2 + lstRurEstimate.Width
    fraManageEstimate(1).Left = lstUrbEstimate.Left
    txtBasin.Width = Width - 3975
    
    lstRurRes.Width = Width * 0.52
    lstUrbRes.Width = Width * 0.42
    lstUrbRes.Left = lstRurRes.Left + lstRurRes.Width + 108
    lblUrbTab(i).Left = lstUrbRes.Left - sash.Left
    For i = 1 To lblUrbTab.Count - 1
      lblUrbTab(i).Left = lblUrbTab(0).Left + lblUrbTabOffset(i)
    Next i
    
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

  'get here when closing with the X button
  'save status
  'Call StatusPut(0)
  'exit program
  End '------------------!_!_!_!_!!_!_!!_!_!_!_!_!_!__!__!!!_!_!_!_!_!_!_!_!_!__!_!_!_!_!_!_!_!!!_!_!_

End Sub

Private Sub lblRurTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  sash_MouseDown Button, Shift, x, y
End Sub

Private Sub lblRurTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  sash_MouseMove Button, Shift, x, y
End Sub

Private Sub lblRurTab_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  sash_MouseUp Button, Shift, x, y
End Sub

Private Sub lblUrbTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  sash_MouseDown Button, Shift, x, y
End Sub

Private Sub lblUrbTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  sash_MouseMove Button, Shift, x, y
End Sub

Private Sub lblUrbTab_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  sash_MouseUp Button, Shift, x, y
End Sub


Private Sub mnuFrequency_Click()
  frmFreq.Show
End Sub

Private Sub mnuHelpObtech_Click()
  OpenFile "http://water.usgs.gov/software/nss.html"
End Sub

Private Sub mnuHydrograph_Click()
  frmHyd.Show
End Sub

Private Sub sash_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  SashDragging = True
End Sub

Private Sub sash_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  sash_MouseMove Button, Shift, x, y
  SashDragging = False
  'Dim t&, newTopHeight&
  'newTopHeight = sash.Top - 1400
  'If newTopHeight > 100 Then lstEstimates.Height = newTopHeight
End Sub

Private Sub sash_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim newTopHeight&
  If SashDragging And (sash.Top + y) > 0 And (sash.Top + y < Height) Then
    sash.Top = sash.Top + y
    newTopHeight = sash.Top - 912 '1400
    If newTopHeight > 0 Then
      lstRurEstimate.Height = newTopHeight
      lstUrbEstimate.Height = newTopHeight
      'fraManageEstimate.Top = lstEstimates.Top + newTopHeight + 96
    End If
    Form_Resize
  End If
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
  frmAbout.Show
End Sub

Private Sub mnuHelpCont_Click()

'   Const HELP_KEY = &H101
'    CMDialog2.HelpFile = "\nss\help\nss.hlp"   ' Specify the Help file to open.
'   CMDialog1.HelpCommand = HELP_KEY    ' When WINHELP.EXE is executed, Help for a specified keyword will be displayed.
'   CMDialog1.HelpKey = "MouseDown" ' Specify the keyword.
'   CMDialog2.ShowHelp .Action = 6    ' Execute WINHELP.EXE.

    Dim i&
    Dim HelpFile As String
'    HelpFile = "\nss\help\nss.hlp" & Chr$(0)
'    HelpFile = "\nsswin\nss.hlp" & Chr$(0)

    OpenFile "nss.hlp"
'    i = WinHelp(hwnd, HelpFile, HELP_INDEX, 0)
    
'    i = ShellExecute(Me.hwnd, "", _
'        "a.book_1.html", vbNullString, _
'        "c:\nss\help\html", 1)

End Sub

Private Sub mnuOpen_Click()

    'read default status file
    'Call StatusGet(1)

End Sub

Private Sub mnuPrint_Click()

    Dim NumCopies, i
    'set cancel to true
    CMDialog2.CancelError = True
    On Error GoTo ErrHandler
    'display the pring dialog box
    CMDialog2.ShowPrinter
    'get user-selected values
    NumCopies = CMDialog2.Copies
    For i = 1 To NumCopies
      PrintForm
    Next i
ErrHandler:
    Exit Sub

End Sub

Private Sub mnuSavReport_Click()
  'Call OutReport
End Sub

Private Sub mnuSavStatus_Click()
  'Call StatusPut(1)
End Sub

Private Sub txtBasin_LostFocus()
  'plot_descriptor = txtBasin.Text
End Sub

'Private Sub txtBasinArea_CommitChange()
'  Dim i&
'  If total_area > 0 And txtBasinArea.value <> total_area Then
'    'i = MsgBox("Changing the Basin Drainage Area will require clearing any existing estimates." & vbCrLf & "Are you sure you want to change its value?", 1)
'    'If i = 1 Then    'still wants to change area
'      stind = -1
'      cboState_Click
'    'End If
'  Else
'    i = 1
'  End If
'  If i = 1 Then
'    total_area = txtBasinArea.value
'  Else 'keep original value
'    txtBasinArea.value = total_area
'  End If
'End Sub

Public Sub SetEstimate()

  'set system state for a specified estimate
'  Dim i&, j&, lreg&
'  If rurcnt > 0 Then cmdDelEstimate(0).Enabled = True Else cmdDelEstimate(0).Enabled = False
'  If urbcnt > 0 Then cmdDelEstimate(1).Enabled = True Else cmdDelEstimate(1).Enabled = False
'  If rurfg And rurind >= 0 And rurcnt > 0 Then 'show current rural estimate
'    RegUseCnt = rurscn(rurind).rcnt
'    Call DispEstimate
'  ElseIf urbind >= 0 And urbcnt > 0 Then       'show current urban estimate
'    rurfg = False
'    RegUseCnt = 1
'    Call DispEstimate
'  End If
End Sub

Public Sub InitEstimate()
'  Dim Estimate&, j&, k&
'  'initialize all estimate data
'  RegUseCnt = 0
'  rurind = -1
'  rurcnt = 0
'  urbind = -1
'  urbcnt = 0
'  're-init data value arrays
'  For Estimate = 0 To MAX_Estimate
'    With urbscn(Estimate)
'      .National = False
'      .rscn = -1
'      .Name = ""
'      .v.vcount = 0
'      .reg = -1
'      For j = 0 To MAX_REGR_COMPONENTS
'        .v.Value(j) = -999
'        .v.vname(j) = ""
'      Next j
'    End With
'    With rurscn(Estimate)
'      .Name = ""
'      .rcnt = 0
'      For j = 0 To MAX_REG_USE
'        For k = 0 To MAX_REGR_COMPONENTS
'          .v(j).Value(k) = -999
'          .v(j).vname(k) = -999
'        Next k
'        .v(j).vcount = 0
'        .reg(j) = -1
'      Next j
'    End With
'  Next Estimate
'  'set to no intervals
'  NumIntrvl = 0
'  'reinit computation tables
'  Call InitRuralDis
'  Call InitUrbanDis
'  If cboState.Tag = 0 Then
'    'reset basin name and total area
'    plot_descriptor = ""
'    txtBasin.Text = ""
'    total_area = 0
''    txtBasinArea.value = 0
'  End If
'  compfg = -1
'  cboState.Tag = 0
'  SetCmdEnabled
'  'init estimate labels
'  lstRurEstimate.Clear
'  lstRurEstimate.AddItem NoEstimates
'  lstUrbEstimate.Clear
'  lstUrbEstimate.AddItem NoEstimates
'  frmNSS.lblEstimates(1) = 0
'  frmNSS.lblEstimates(2) = "of 0"
'  frmNSS.lblEstimates(4) = 0
'  frmNSS.lblEstimates(5) = "of 0"
End Sub

Public Sub SetCmdEnabled()
'  Dim enaRur As Boolean, enaUrb As Boolean
'
'  If rurcnt < 1 Or rurind < 0 Then enaRur = False Else enaRur = True
'  If urbcnt < 1 Or urbind < 0 Then enaUrb = False Else enaUrb = True
'
'  cmdWeight.Enabled = enaRur
'  cmdDelEstimate(0).Enabled = enaRur
'  cmdDelEstimate(1).Enabled = enaUrb
'
'  cmdFrequency.Enabled = enaUrb Or enaRur
'  cmdHydrograph.Enabled = enaUrb Or enaRur
'
End Sub

Private Function ConfDelEstimate()

'    Dim i&, j&, rcnt&, vcnt&, confrm&
'    Dim lstr$, NL$
'
'    NL = Chr(13) + Chr(10)
'    'assume ok
'    confrm = 1
'    If rurfg = True Then
'      'make sure no urban estimates using this rural estimate
'      i = 0
'      While i < urbcnt
'        If urbscn(i).rscn = rurind Then          'this rural estimate in use by urban
'          MsgBox "This rural estimate's results are needed for urban estimate " & i + 1 & "." & NL & "Thus it may not be deleted", 48
'          confrm = 0
'          i = urbcnt
'        Else
'          i = i + 1
'        End If
'      Wend
'    End If
'    If confrm = 1 Then
'      'ask for confirmation
'      lstr = "Confirm delete for estimate "
'      If rurfg = True Then
'        'rural estimate selected
'        lstr = lstr & "R" & rurind + 1 & ": " & rurscn(rurind).Name & NL
'        rcnt = rurscn(rurind).rcnt - 1
'      Else
'        'urban estimate selected
'        lstr = lstr & "U" & urbind + 1 & ": " & urbscn(urbind).Name & NL
'        rcnt = 0
'      End If
'      For i = 0 To rcnt
'        'summary line for each region in use
'        If rurfg = True Then
'          vcnt = rurscn(rurind).v(i).vcount - 1
'        Else
'          vcnt = urbscn(urbind).v.vcount - 1
'        End If
'        For j = 0 To vcnt
'          If rurfg = True Then
'            lstr = lstr & Trim(rurscn(rurind).v(i).vname(j)) & "=" & NumFmted(rurscn(rurind).v(i).Value(j), 7, 1)
'          Else
'            lstr = lstr & Trim(urbscn(urbind).v.vname(j)) & "=" & NumFmted(urbscn(urbind).v.Value(j), 7, 1)
'          End If
'          'pad with blanks between variables
'          lstr = lstr & "   "
'        Next j
'        If i < rcnt Then
'          'new line
'          lstr = lstr & NL
'        End If
'      Next i
'      confrm = MsgBox(lstr, 1, "NSS estimates")
'    End If
'    ConfDelEstimate = confrm

End Function
