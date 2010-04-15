VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmNSS 
   Caption         =   "National Streamflow Statistics (NSS)"
   ClientHeight    =   7440
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11190
   HelpContextID   =   12
   Icon            =   "frmNSS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   7920
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame fraBottom 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   372
      Left            =   120
      TabIndex        =   18
      Top             =   6960
      Width           =   4692
      Begin VB.CommandButton cmdFrequency 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fre&quency Plot"
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
         HelpContextID   =   25
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   1692
      End
      Begin VB.CommandButton cmdWeight 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Weight"
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
         HelpContextID   =   27
         Left            =   3600
         TabIndex        =   21
         Top             =   0
         Width           =   1092
      End
      Begin VB.CommandButton cmdHydrograph 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "H&ydrograph"
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
         HelpContextID   =   26
         Left            =   1800
         TabIndex        =   20
         Top             =   0
         Width           =   1692
      End
   End
   Begin VB.Frame fraManageEstimate 
      Caption         =   "&Urban"
      Height          =   5292
      Index           =   1
      Left            =   4920
      TabIndex        =   11
      Top             =   600
      Width           =   3852
      Begin VB.Frame fraNewEditDel 
         BorderStyle     =   0  'None
         Height          =   252
         Index           =   1
         Left            =   1920
         TabIndex        =   26
         Top             =   240
         Width           =   1812
         Begin VB.CommandButton cmdNew 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "New"
            Height          =   255
            HelpContextID   =   12
            Index           =   1
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   492
         End
         Begin VB.CommandButton cmdDelete 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Delete"
            Height          =   255
            HelpContextID   =   12
            Index           =   1
            Left            =   1200
            TabIndex        =   15
            Top             =   0
            Width           =   612
         End
         Begin VB.CommandButton cmdEdit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Edit"
            Height          =   255
            HelpContextID   =   12
            Index           =   1
            Left            =   600
            TabIndex        =   14
            Top             =   0
            Width           =   492
         End
      End
      Begin VB.TextBox txtEstimate 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2532
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Text            =   "frmNSS.frx":030A
         Top             =   2640
         Width           =   3612
      End
      Begin VB.TextBox txtSummary 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1932
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "frmNSS.frx":031B
         Top             =   600
         Width           =   3612
      End
      Begin VB.ComboBox cboScenario 
         Height          =   288
         HelpContextID   =   12
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   1692
      End
      Begin VB.Frame fraSashUpDown 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   132
         Index           =   1
         Left            =   120
         MousePointer    =   7  'Size N S
         TabIndex        =   23
         Top             =   2520
         Width           =   3612
      End
   End
   Begin VB.Frame fraManageEstimate 
      Caption         =   "&Rural"
      Height          =   5292
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4692
      Begin VB.Frame fraNewEditDel 
         BorderStyle     =   0  'None
         Height          =   252
         Index           =   0
         Left            =   2760
         TabIndex        =   25
         Top             =   240
         Width           =   1812
         Begin VB.CommandButton cmdEdit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Edit"
            Height          =   255
            HelpContextID   =   12
            Index           =   0
            Left            =   600
            TabIndex        =   7
            Top             =   0
            Width           =   492
         End
         Begin VB.CommandButton cmdNew 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "New"
            Height          =   255
            HelpContextID   =   12
            Index           =   0
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   492
         End
         Begin VB.CommandButton cmdDelete 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Delete"
            Height          =   255
            HelpContextID   =   12
            Index           =   0
            Left            =   1200
            TabIndex        =   8
            Top             =   0
            Width           =   612
         End
      End
      Begin VB.TextBox txtEstimate 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2532
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Text            =   "frmNSS.frx":032B
         Top             =   2640
         Width           =   4452
      End
      Begin VB.TextBox txtSummary 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1932
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "frmNSS.frx":033C
         Top             =   600
         Width           =   4452
      End
      Begin VB.ComboBox cboScenario 
         Height          =   288
         HelpContextID   =   12
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2532
      End
      Begin VB.Frame fraSashUpDown 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   132
         Index           =   0
         Left            =   120
         MousePointer    =   7  'Size N S
         TabIndex        =   22
         Top             =   2520
         Width           =   4452
      End
   End
   Begin VB.TextBox txtName 
      Height          =   285
      HelpContextID   =   12
      Left            =   3840
      TabIndex        =   3
      Text            =   "txtName"
      Top             =   120
      Width           =   3972
   End
   Begin VB.ComboBox cboState 
      Appearance      =   0  'Flat
      Height          =   288
      HelpContextID   =   12
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "-999"
      Top             =   120
      Width           =   1812
   End
   Begin VB.Frame fraSashLeftRight 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5172
      Left            =   4800
      MousePointer    =   9  'Size W E
      TabIndex        =   24
      Top             =   720
      Width           =   132
   End
   Begin VB.Label lblBasinName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Site &Name:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2640
      TabIndex        =   2
      Top             =   168
      Width           =   1092
   End
   Begin VB.Label lblState 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&State:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   168
      Width           =   492
   End
   Begin VB.Menu mnuTop 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Save As..."
      End
      Begin VB.Menu mnuReport 
         Caption         =   "&Report"
      End
      Begin VB.Menu mnuRecent 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "O&ptions"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuDatabase 
         Caption         =   "Select &Database"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTop 
      Caption         =   "&Graph"
      Index           =   1
      Begin VB.Menu mnuGraphFrequency 
         Caption         =   "Fre&quency"
      End
      Begin VB.Menu mnuGraphHydrograph 
         Caption         =   "H&ydrograph"
      End
   End
   Begin VB.Menu mnuTop 
      Caption         =   "&Help"
      Index           =   5
      Begin VB.Menu mnuHelpManual 
         Caption         =   "&User Manual"
      End
      Begin VB.Menu mnuWeb 
         Caption         =   "&Web Site"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmNSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private Const appName As String = "NSS"
Private Const SectionMainWin As String = "Main Window"
Private Const SectionRecentFiles As String = "Recent Files"
Private Const MaxRecentFiles As Integer = 6

Private WithEvents ProjectEvents As nssProject
Attribute ProjectEvents.VB_VarHelpID = -1
Private SashDraggingUpDown As Boolean
Private SashDraggingUpDownY As Long
Private SashDraggingLeftRight As Boolean
Private SashDraggingLeftRightX As Long
Private LeftWidthFraction As Single

Private Sub cboScenario_Click(Index As Integer)
  If Index = 0 Then Project.CurrentRuralScenario = cboScenario(0).ListIndex + 1
  If Index = 1 Then Project.CurrentUrbanScenario = cboScenario(1).ListIndex + 1
  UpdateLabels
End Sub

Private Sub cboState_Click()
  Dim answer As VbMsgBoxResult
  Dim i&

  answer = vbOK
  
  'Note: when we are not in the middle of changing the state,
  'cboState.ListIndex + 1 = Project.State.code = Project.DB.States(cboState.Code).Code
  If Project.State Is Nothing Then
    answer = vbOK
  ElseIf Project.State.code = Format(cboState.ItemData(cboState.ListIndex), "00") Then
    'already have this state selected
    answer = vbCancel
  ElseIf cboState.ListIndex >= 0 Then
    If Project.RuralScenarios.Count + Project.UrbanScenarios.Count > 0 Then
      answer = MsgBox("Changing the state will clear any existing estimates." & vbCr & _
                      "Are you sure you want to change the selected state?", vbOKCancel)
    End If
  Else
    answer = vbCancel
  End If
      
  If answer = vbOK Then
    Project.Clear
    Clear
    Set Project.State = Project.DB.States.ItemByKey(Format(cboState.ItemData(cboState.ListIndex), "00"))
  ElseIf Project.State.code <> Format(cboState.ItemData(cboState.ListIndex), "00") Then
    'keep current state
    While cboState.ItemData(i) <> CLng(Project.State.code)
      i = i + 1
    Wend
    cboState.ListIndex = i
  End If
End Sub

Private Sub cmdEdit_Click(Index As Integer)
  Dim myScenario As nssScenario
  
  If Index = 0 Then
    If Project.CurrentRuralScenario > 0 Then
      Set myScenario = Project.RuralScenarios(Project.CurrentRuralScenario).Copy
    End If
  Else
    If Project.CurrentUrbanScenario > 0 Then
      Set myScenario = Project.UrbanScenarios(Project.CurrentUrbanScenario).Copy
    End If
  End If
  
  If myScenario Is Nothing Then
    cmdNew_Click Index
  Else
    myScenario.Edit
  End If
End Sub

Private Sub cmdFrequency_Click()
  Dim i As Long, NumFloodScens As Long
  NumFloodScens = 0
  For i = 1 To Project.RuralScenarios.Count
    If Not Project.RuralScenarios(i).lowflow Then
      NumFloodScens = NumFloodScens + 1
    End If
  Next i
  For i = 1 To Project.UrbanScenarios.Count
    If Not Project.UrbanScenarios(i).lowflow Then
      NumFloodScens = NumFloodScens + 1
    End If
  Next i
  If NumFloodScens > 0 Then
    frmFreq.Show
  Else
    MsgBox "No Flood Frequency scenarios available to graph", vbOKOnly, appName
  End If
End Sub

Private Sub cmdHydrograph_Click()
  Dim i As Long, NumFloodScens As Long
  NumFloodScens = 0
  For i = 1 To Project.RuralScenarios.Count
    If Not Project.RuralScenarios(i).lowflow Then
      NumFloodScens = NumFloodScens + 1
    End If
  Next i
  For i = 1 To Project.UrbanScenarios.Count
    If Not Project.UrbanScenarios(i).lowflow Then
      NumFloodScens = NumFloodScens + 1
    End If
  Next i
  If NumFloodScens > 0 Then
    frmHyd.Show
  Else
    MsgBox "No Flood Frequency scenarios available to graph", vbOKOnly, appName
  End If
End Sub

Private Sub cmdNew_Click(Index As Integer)
  Dim myScenario As nssScenario
  Dim vScenario As Variant
  Dim NameIndex As Long, DefaultName As String, NameExists As Boolean
  
  Set myScenario = New nssScenario
  Set myScenario.Project = Project
  
  NameExists = True
  
  'Come up with a good default name for new scenario
  If Index = 1 Then
    myScenario.Urban = True
    While NameExists
      NameIndex = NameIndex + 1
      DefaultName = "Urban " & NameIndex
      NameExists = False
      For Each vScenario In Project.UrbanScenarios
        If vScenario.Name = DefaultName Then
          NameExists = True
          Exit For
        End If
      Next
    Wend
  Else
    While NameExists
      NameIndex = NameIndex + 1
      DefaultName = "Rural " & NameIndex
      NameExists = False
      For Each vScenario In Project.RuralScenarios
        If vScenario.Name = DefaultName Then
          NameExists = True
          Exit For
        End If
      Next
    Wend
  End If
  myScenario.Name = DefaultName
  
  myScenario.Edit
End Sub

Private Sub cmdDelete_Click(Index As Integer)
  If Index = 0 Then
    If Project.CurrentRuralScenario > 0 Then
      Project.RuralScenarios.Remove Project.CurrentRuralScenario
      ProjectEvents_Edited
    End If
  Else
    If Project.CurrentUrbanScenario > 0 Then
      Project.UrbanScenarios.Remove Project.CurrentUrbanScenario
      ProjectEvents_Edited
    End If
  End If
End Sub

Private Sub cmdWeight_Click()
  Dim vScenario As Variant
  Dim scen As nssScenario
  Dim NameTest As String, NameIndex As Long, NameExists As Boolean
  
  If Project.CurrentRuralScenario > 0 Then
    Set scen = Project.RuralScenarios(Project.CurrentRuralScenario)
    If scen.lowflow Then
      MsgBox "Weighting not available for Low Flow scenarios.", vbInformation, "NSS Weight"
    Else
      If scen.Weight.WeightType > 0 Then 'This is already a weighted scenario, just edit it
        scen.EditWeight
      Else 'Create new weighted scenario based on this one
        Set scen = scen.Copy
        scen.Name = scen.Name
        NameIndex = 1
        Do
          If NameIndex > 1 Then
            NameTest = scen.Name & " (weighted " & NameIndex & ")"
          Else
            NameTest = scen.Name & " (weighted)"
          End If
          NameExists = False
          For Each vScenario In Project.RuralScenarios
            If vScenario.Name = NameTest Then
              NameExists = True
              Exit For
            End If
          Next
          NameIndex = NameIndex + 1
        Loop While NameExists
        
        scen.Name = NameTest
        Project.RuralScenarios.Add scen, LCase(scen.Name)
        scen.EditWeight
        'scen.Project.CurrentRuralScenario = scen.Project.RuralScenarios.Count
        scen.Project.RaiseEdited
      End If
    End If
  End If
End Sub

Private Sub Form_Load()
  Dim i As Integer
  Dim stIndex As Long
  Dim selState As Long
  Dim progress As String
  On Error GoTo ShowProgress
  
  'progress = "Me.Show"
  'Me.Show
  progress = progress & vbCr & "RetrieveWindowSettings"
  RetrieveWindowSettings
  progress = progress & vbCr & "Setting ProjectEvents"
  Set ProjectEvents = Project
  progress = progress & vbCr & "InitDischRatio"
  InitDischRatio
  progress = progress & vbCr & "Clear"
  Clear
  progress = progress & vbCr & "cboState.Clear"
  cboState.Clear
  progress = progress & vbCr & "Adding states to cboState"
  stIndex = 0
  For i = 1 To Project.DB.States.Count
    If IsNumeric(Project.DB.States(i).code) Then
      If CInt(Project.DB.States(i).code) < 99 Then
        'weed out dummy state and ROI low flow state records
        progress = progress & vbCr & "Adding " & Project.DB.States(i).Name
        cboState.AddItem Project.DB.States(i).Name
        cboState.ItemData(stIndex) = Project.DB.States(i).code
        stIndex = stIndex + 1
        If Project.DB.States(i).Name = Project.State.Name Then selState = stIndex
      End If
    End If
  Next
  progress = progress & vbCr & "cboState.ListIndex = " & selState - 1
  cboState.ListIndex = selState - 1
  progress = progress & vbCr & "ProjectEvents_Edited"
  ProjectEvents_Edited
  Exit Sub

ShowProgress:
  MsgBox progress & vbCr & Err.Description, vbExclamation, "Error loading NSS main form"

End Sub

Private Sub Form_Unload(Cancel As Integer)
  SaveFileString DefaultSaveFile, Project.XML
  SaveWindowSettings
End Sub

Private Sub Clear()
  txtName = Project.Name
  txtSummary(0) = ""
  txtSummary(1) = ""
  txtEstimate(0) = ""
  txtEstimate(1) = ""
  cboScenario(0).Clear
  cboScenario(1).Clear
End Sub

Private Sub UpdateLabels()
  With Project
    Dim i As Long
    For i = 0 To cboState.ListCount - 1
      If cboState.List(i) = .State.Name Then
        cboState.ListIndex = i
        Exit For
      End If
    Next
    txtName = .Name
    
    cboScenario(0).ListIndex = .CurrentRuralScenario - 1
    If .CurrentRuralScenario > 0 Then
      txtSummary(0) = .RuralScenarios(.CurrentRuralScenario).Summary
    Else
      txtSummary(0) = "No Scenarios Available"
    End If
    
    cboScenario(1).ListIndex = .CurrentUrbanScenario - 1
    If .CurrentUrbanScenario > 0 Then
      txtSummary(1) = .UrbanScenarios(.CurrentUrbanScenario).Summary
    Else
      txtSummary(1) = "No Scenarios Available"
    End If
  End With
  UpdateEstimates
End Sub

Private Sub UpdateEstimates()
  Me.MousePointer = vbHourglass
  With Project
    If .CurrentRuralScenario > 0 And .RuralScenarios.Count >= .CurrentRuralScenario Then
      txtEstimate(0).Text = .RuralScenarios(.CurrentRuralScenario).EstimateString
    Else
      txtEstimate(0).Text = ""
    End If

    If .CurrentUrbanScenario > 0 And .UrbanScenarios.Count >= .CurrentUrbanScenario Then
      txtEstimate(1).Text = .UrbanScenarios(.CurrentUrbanScenario).EstimateString
    Else
      txtEstimate(1).Text = ""
    End If
  End With
  Me.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
  Dim w&, h&, contentWidth&
  On Error Resume Next
  w = Me.ScaleWidth
  h = Me.ScaleHeight
  If w > 1000 Then
    fraManageEstimate(0).Width = (w - 350) * LeftWidthFraction
    contentWidth = fraManageEstimate(0).Width - 240
    txtSummary(0).Width = contentWidth
    txtEstimate(0).Width = contentWidth
    fraSashUpDown(0).Width = contentWidth
    fraNewEditDel(0).Left = contentWidth - fraNewEditDel(0).Width + 120
    If fraNewEditDel(0).Left > 250 Then cboScenario(0).Width = fraNewEditDel(0).Left - 240
    
    fraSashLeftRight.Left = fraManageEstimate(0).Width + fraManageEstimate(0).Left
    If w > 350 + fraManageEstimate(0).Width Then
      fraManageEstimate(1).Width = w - 350 - fraManageEstimate(0).Width
      contentWidth = fraManageEstimate(1).Width - 240
      If contentWidth > 0 Then
        txtSummary(1).Width = contentWidth
        txtEstimate(1).Width = contentWidth
        fraSashUpDown(1).Width = contentWidth
      End If
    End If
    fraManageEstimate(1).Left = fraManageEstimate(0).Left + fraManageEstimate(0).Width + 108
    fraNewEditDel(1).Left = contentWidth - fraNewEditDel(1).Width + 120
    If fraNewEditDel(1).Left > 250 Then cboScenario(1).Width = fraNewEditDel(1).Left - 240
  End If
  If h > 2000 Then
    fraManageEstimate(0).Height = h - 1164
    fraManageEstimate(1).Height = fraManageEstimate(0).Height
    fraSashLeftRight.Height = fraManageEstimate(0).Height
    If fraManageEstimate(0).Height > txtEstimate(0).Top + 120 Then
      txtEstimate(0).Height = fraManageEstimate(0).Height - txtEstimate(0).Top - 120
      txtEstimate(1).Height = txtEstimate(0).Height
    End If
    fraBottom.Top = fraManageEstimate(0).Height + fraManageEstimate(0).Top + 108
  End If
End Sub

Private Sub fraSashLeftRight_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  SashDraggingLeftRight = True
  SashDraggingLeftRightX = x
End Sub
Private Sub fraSashLeftRight_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  fraSashLeftRight_MouseMove Button, Shift, x, y
  SashDraggingLeftRight = False
End Sub
Private Sub fraSashLeftRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim newLeftWidth As Long
  newLeftWidth = fraSashLeftRight.Left + x - 110
  If SashDraggingLeftRight And newLeftWidth > 0 And Me.ScaleWidth - newLeftWidth > 220 Then
    LeftWidthFraction = newLeftWidth / (Me.ScaleWidth - 350)
    Form_Resize
  End If
End Sub

Private Sub fraSashUpDown_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  SashDraggingUpDown = True
  SashDraggingUpDownY = y
End Sub
Private Sub fraSashUpDown_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  fraSashUpDown_MouseMove Index, Button, Shift, x, y
  SashDraggingUpDown = False
End Sub
Private Sub fraSashUpDown_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim newTopHeight As Long
  If SashDraggingUpDown And fraSashUpDown(0).Top + y - SashDraggingUpDownY - 588 - 300 > 0 Then
    fraSashUpDown(0).Top = fraSashUpDown(0).Top + y - SashDraggingUpDownY
    fraSashUpDown(1).Top = fraSashUpDown(0).Top
    newTopHeight = fraSashUpDown(0).Top - 588
    txtSummary(0).Height = newTopHeight
    txtSummary(1).Height = newTopHeight
    txtEstimate(0).Top = fraSashUpDown(0).Top + 120
    txtEstimate(1).Top = txtEstimate(0).Top
    If fraManageEstimate(0).Height > txtEstimate(0).Top + 120 Then
      txtEstimate(0).Height = fraManageEstimate(0).Height - txtEstimate(0).Top - 120
      txtEstimate(1).Height = txtEstimate(0).Height
    End If
  End If
End Sub

Private Sub mnuDatabase_Click()
  Dim dbPath As String
  Dim ff As New ATCoFindFile
  Dim answer As VbMsgBoxResult

FindDB:
  'Open NSS/StreamStats Database
  ff.SetDialogProperties "Please locate NSS or StreamStats database version 4", "NSSv4.mdb"
  ff.SetRegistryInfo "StreamStatsDB", "Defaults", "nssDatabaseV4"
  dbPath = ff.GetName(True)
  
  If Len(dbPath) > 0 Then
    If Not FileExists(dbPath) Then
      If MsgBox("Could not open database or project" & vbCr & vbCr _
            & Err.Description & vbCr & vbCr _
            & "Search for current database?", vbOKCancel, "NSS Database Problem") = vbOK Then
        SaveSetting "StreamStatsDB", "Defaults", "nssDatabaseV4", "NSSv4.mdb"
        GoTo FindDB
      Else
        End
      End If
    Else
      If Project.RuralScenarios.Count + Project.UrbanScenarios.Count > 0 Then
        answer = MsgBox("Changing the database will clear any existing estimates." & vbCr & _
                        "Are you sure you want to change the database?", vbOKCancel)
      Else
        answer = vbOK
      End If
      If answer = vbOK Then
        Project.Clear
        Clear
        Project.LoadNSSdatabase dbPath
      End If
    End If
  End If

End Sub

Private Sub mnuExit_Click()
  Unload Me
  DoEvents
  End
End Sub

Private Sub mnuGraphFrequency_Click()
  If Project.RuralScenarios.Count + Project.UrbanScenarios.Count > 0 Then
    frmFreq.Show
  Else
    MsgBox "No scenarios available to graph", vbOKOnly, appName
  End If
End Sub

Private Sub mnuGraphHydrograph_Click()
  If Project.RuralScenarios.Count + Project.UrbanScenarios.Count > 0 Then
    frmHyd.Show
  Else
    MsgBox "No scenarios available to graph", vbOKOnly, appName
  End If
End Sub

Private Sub mnuHelpAbout_Click()
  frmAbout.Show
End Sub

Private Sub mnuHelpManual_Click()
  App.HelpFile = OpenFile(App.HelpFile, cdlg)
End Sub

Private Sub mnuOpen_Click()
  On Error GoTo ErrExit
  With cdlg
    .DialogTitle = "Open Status File"
    .Filter = "NSS Status Files (*.nss)|*.nss|All Files|*.*"
    .FilterIndex = 0
    .ShowOpen
    Project.FileName = .FileName
    If Len(Dir(Project.FileName)) > 0 Then
      Project.XML = WholeFileString(Project.FileName)
      AddRecentFile Project.FileName
    End If
  End With

  Exit Sub

ErrExit:
  If Err.Number <> 32755 Then 'If something other than "Cancel was selected" then notify user
    MsgBox "Error opening NSS Status File '" & cdlg.FileName & "'" & vbCr _
          & Err.Description, vbCritical, appName
  End If
End Sub

Private Sub mnuOptions_Click()
  frmStart.Show
End Sub

Private Sub mnuRecent_Click(Index As Integer)
  Dim newFilePath$, tmpFilePath$
  If Index > 0 Then
    newFilePath = mnuRecent(Index).Tag
    'If UCase(Project.Filename) = UCase(newFilePath) Then
    '  If MsgBox("Discard changes and reload this project?", vbOKCancel, _
                 "Load Project") = vbCancel Then Exit Sub
    'End If
    If Len(Dir(newFilePath)) > 0 Then
      Project.FileName = newFilePath
      Project.XML = WholeFileString(newFilePath)
    Else 'status file not currently available, remove from menu
      While Index < mnuRecent.Count - 1
        tmpFilePath = mnuRecent.Item(Index + 1).Tag
        mnuRecent.Item(Index).Caption = "&" & Index & " " & FilenameOnly(tmpFilePath)
        mnuRecent.Item(Index).Tag = tmpFilePath
        Index = Index + 1
      Wend
      Unload mnuRecent.Item(Index)
      If MsgBox("Project " & newFilePath & " not found. Look for it?", vbOKCancel, _
                "Recent File Error") = vbOK Then
        mnuOpen_Click
      End If
    End If
  End If
End Sub

Private Sub mnuReport_Click()
  On Error GoTo ErrHand
  With cdlg
    .DialogTitle = "Save NSS Report"
    .Filter = "NSS Report (*.txt)|*.txt|All Files|*.*"
    .FilterIndex = 0
    .ShowSave
    SaveFileString .FileName, Project.Report
    OpenFile .FileName
  End With
  
  Exit Sub

ErrHand:
  If Err.Description <> "Cancel was selected." Then
    MsgBox "Error saving report" & vbCr & Err.Description, vbCritical, appName
  End If
End Sub

Private Sub mnuSaveAs_Click()
  On Error GoTo ErrExit
  With cdlg
    .DialogTitle = "Save NSS Status File"
    .Filter = "NSS Status Files (*.nss)|*.nss|All Files|*.*"
    .FilterIndex = 0
    .ShowSave
    SaveFileString .FileName, Project.XML
    AddRecentFile .FileName
  End With

  Exit Sub

ErrExit:
  If Err.Number <> 32755 Then 'If something other than "Cancel was selected" then notify user
    MsgBox "Error saving NSS Status File '" & cdlg.FileName & "'" _
           & vbCr & Err.Description, vbCritical, appName
  End If
End Sub

Public Sub TestAllEquations(Optional PathName As String = "")
  Dim StateIndex As Long
  Dim newScenario As nssScenario
  Dim lastRural As nssScenario
  Dim region As nssRegion
  Dim newUserRegion As userRegion
  Dim vRegion As Variant
  Dim vParam As Variant
  Dim doMax As Boolean
  Dim StateErrors As String
  
  Me.MousePointer = vbHourglass
  
  If Len(PathName) > 0 Then
    If Right(PathName, 1) <> "\" Then PathName = PathName & "\"
  End If
  
  For StateIndex = 1 To cboState.ListCount - 1 'added -1 to listcount to avoid "Dummy" state
    Project.Clear
    Clear
    Set Project.State = Project.DB.States(StateIndex)
    cboState.ListIndex = StateIndex - 1
    Set lastRural = Nothing
    StateErrors = ""
    DoEvents
    doMax = False
    For Each vRegion In Project.State.Regions
      Set region = vRegion
      If Not region.ROI Then GoSub AddRegionScenario
    Next
    Set region = Project.NationalUrban
    GoSub AddRegionScenario
    
    doMax = True
    For Each vRegion In Project.State.Regions
      Set region = vRegion
      If Not region.ROI Then GoSub AddRegionScenario
    Next
    Set region = Project.NationalUrban
    GoSub AddRegionScenario
    
    SaveFileString PathName & "NSStest_" & Project.State.Abbrev & ".txt", _
                   Project.Report & StateErrors
    Set newScenario = Nothing
    Project.Clear
  Next
  If MsgBox("Finished testing min and max values for all equations." & vbCr _
          & "Open results folder '" & PathName & "' now?", vbYesNo, "NSS Test") = vbYes Then
    OpenFile PathName
  End If
  Me.MousePointer = vbDefault
  Exit Sub
  
AddRegionScenario:
  Set newScenario = New nssScenario
  Set newScenario.Project = Project
  newScenario.Urban = region.Urban
  If Abs(region.LowFlowRegnID) > 0 Then newScenario.lowflow = True
  If doMax Then
    newScenario.Name = region.State.Abbrev & "_" & region.Name & "_Max"
  Else
    newScenario.Name = region.State.Abbrev & "_" & region.Name & "_Min"
  End If
  If region.Urban Then
    If region.UrbanNeedsRural Then
      If lastRural Is Nothing Then
        StateErrors = StateErrors & vbCrLf _
                    & "Error: No Rural scenario found for urban needing rural region " _
                    & newScenario.Name
        Return
      Else
        Set newScenario.RuralScenario = lastRural
        'Don't want area that was automatically set from rural scenario
        newScenario.SetArea 0, Project.Metric
      End If
    End If
    Project.UrbanScenarios.Add newScenario, LCase(newScenario.Name)
  Else
    If Not newScenario.lowflow Then 'make sure last rural is a peak flow scenario for urban use
      Set lastRural = newScenario
    End If
    Project.RuralScenarios.Add newScenario, LCase(newScenario.Name)
  End If
  Set newUserRegion = New userRegion
  Set newUserRegion.region = region
  newScenario.UserRegions.Add newUserRegion, region.Name
  For Each vParam In newUserRegion.UserParms
    If doMax Then
      vParam.setValue vParam.Parameter.GetMax(Project.Metric), Project.Metric
    Else
      vParam.setValue vParam.Parameter.GetMin(Project.Metric), Project.Metric
    End If
    If InStr(LCase(vParam.Parameter.Name), "area") > 0 Then
      If newScenario.getArea(Project.Metric) = 0 Then 'Or region.UrbanNeedsRural Then
        newScenario.SetArea vParam.getValue(Project.Metric), Project.Metric
      End If
    End If
  Next
  Set newUserRegion = Nothing
  Return
End Sub

Private Sub mnuWeb_Click()
  OpenFile "http://water.usgs.gov/software/nss.html"
End Sub

Private Sub ProjectEvents_Edited()
  Dim member As Variant
  
  cboScenario(0).Clear
  For Each member In Project.RuralScenarios
    cboScenario(0).AddItem member.Name
  Next
  
  cboScenario(1).Clear
  For Each member In Project.UrbanScenarios
    cboScenario(1).AddItem member.Name
  Next
  
  UpdateLabels
  
End Sub

Private Sub txtName_Change()
  Project.Name = txtName
End Sub

Private Sub AddRecentFile(FilePath As String)
  Dim rf&, rfMove&, newPath$, match As Boolean
  rf = 0
  While Not match And rf <= mnuRecent.Count - 2
    rf = rf + 1
    If UCase(mnuRecent(rf).Tag) = UCase(FilePath) Then match = True
  Wend
  If match Then 'move file to top of list
    For rfMove = rf To 2 Step -1
      mnuRecent(rfMove).Tag = mnuRecent(rfMove - 1).Tag
      mnuRecent(rfMove).Caption = "&" & rfMove & " " & FilenameOnly(mnuRecent(rfMove).Tag)
    Next rfMove
  Else 'Add file to list
    mnuRecent(0).Visible = True
    If mnuRecent.Count <= MaxRecentFiles Then Load mnuRecent(mnuRecent.Count)
    For rfMove = mnuRecent.Count - 1 To 2 Step -1
      mnuRecent(rfMove).Tag = mnuRecent(rfMove - 1).Tag
      mnuRecent(rfMove).Caption = "&" & rfMove & " " & FilenameOnly(mnuRecent(rfMove).Tag)
    Next rfMove
  End If
  mnuRecent(1).Visible = True
  mnuRecent(1).Tag = FilePath
  mnuRecent(1).Caption = "&1 " & FilenameOnly(mnuRecent(rfMove).Tag)
End Sub

Private Sub SaveWindowSettings()
  Dim rf&
  If Height > 800 And Left < Screen.Width And Top < Screen.Height Then
    SaveSetting appName, SectionMainWin, "Width", Width
    SaveSetting appName, SectionMainWin, "Height", Height
    SaveSetting appName, SectionMainWin, "Left", Left
    SaveSetting appName, SectionMainWin, "Top", Top
    SaveSetting appName, SectionMainWin, "LeftWidthFraction", LeftWidthFraction
  End If
  For rf = mnuRecent.Count - 1 To 1 Step -1
    SaveSetting appName, SectionRecentFiles, CStr(rf), mnuRecent(rf).Tag
  Next rf
  While GetSetting(appName, SectionRecentFiles, CStr(rf)) <> ""
    SaveSetting appName, SectionRecentFiles, CStr(rf), ""
    rf = rf + 1
  Wend
End Sub

Private Sub RetrieveWindowSettings()
  Dim setting As Variant, rf&
  
  setting = GetSetting(appName, SectionMainWin, "LeftWidthFraction", "0.55")
  If IsNumeric(setting) Then
    If setting > 0 Then LeftWidthFraction = setting
  End If
  
  setting = GetSetting(appName, SectionMainWin, "Left")
  If IsNumeric(setting) Then
    If setting < Screen.Width Then Left = setting
  End If
  setting = GetSetting(appName, SectionMainWin, "Top")
  If IsNumeric(setting) Then
    If setting >= 0 And setting < Screen.Height * 0.9 Then Top = setting
  End If
  setting = GetSetting(appName, SectionMainWin, "Width")
  If IsNumeric(setting) Then
    If setting > 200 And setting <= Screen.Width Then Width = setting
  End If
  setting = GetSetting(appName, SectionMainWin, "Height")
  If IsNumeric(setting) Then
    If setting > 200 And setting <= Screen.Height Then Height = setting
  End If
  
  For rf = MaxRecentFiles To 1 Step -1
    setting = GetSetting(appName, SectionRecentFiles, CStr(rf))
    If setting <> "" Then AddRecentFile CStr(setting)
  Next rf
End Sub

