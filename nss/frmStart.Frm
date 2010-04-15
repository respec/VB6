VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmStart 
   Caption         =   "NSS Options"
   ClientHeight    =   2850
   ClientLeft      =   2670
   ClientTop       =   3285
   ClientWidth     =   5505
   HelpContextID   =   11
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2850
   ScaleWidth      =   5505
   Begin VB.Frame fraVarLabel 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   1080
      Width           =   5055
      Begin VB.OptionButton optVarianceType 
         Caption         =   "Log10"
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
         Index           =   0
         Left            =   2040
         TabIndex        =   21
         Top             =   0
         Width           =   1455
      End
      Begin VB.OptionButton optVarianceType 
         Caption         =   "Percent"
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
         Index           =   1
         Left            =   3600
         TabIndex        =   20
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Variance Label Type:"
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
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.Frame fraStatLabel 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Width           =   5055
      Begin VB.OptionButton optLabelType 
         Caption         =   "Traditional (Rec Intvl)"
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
         Index           =   0
         Left            =   2040
         TabIndex        =   17
         Top             =   0
         Width           =   1335
      End
      Begin VB.OptionButton optLabelType 
         Caption         =   "Revised (% Chance)"
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
         Index           =   1
         Left            =   3600
         TabIndex        =   16
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Statistic Label Type:"
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
         Left            =   0
         TabIndex        =   18
         Top             =   60
         Width           =   2055
      End
   End
   Begin VB.Frame fraUnits 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   372
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   5055
      Begin VB.OptionButton optUnits 
         Caption         =   "&Metric"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         HelpContextID   =   11
         Index           =   1
         Left            =   3600
         TabIndex        =   14
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optUnits 
         Caption         =   "&English"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         HelpContextID   =   11
         Index           =   0
         Left            =   2040
         TabIndex        =   13
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Units for display:"
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
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.OptionButton optAnalysisType 
      Caption         =   "Other"
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
      Index           =   1
      Left            =   3840
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.OptionButton optAnalysisType 
      Caption         =   "Peak"
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
      Index           =   0
      Left            =   2280
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdBrowse 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Browse"
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
      HelpContextID   =   11
      Left            =   4320
      TabIndex        =   7
      ToolTipText     =   "Open a different project"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox txtProject 
      Height          =   288
      HelpContextID   =   11
      Left            =   1080
      TabIndex        =   3
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox txtUser 
      Height          =   288
      HelpContextID   =   11
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   1572
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   2520
      Top             =   120
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   372
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
      Begin VB.CommandButton cmdOk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&OK"
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
         HelpContextID   =   11
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   732
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Cancel"
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
         HelpContextID   =   11
         Left            =   1200
         TabIndex        =   5
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Analysis Type:"
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
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Project:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1965
      Width           =   1215
   End
   Begin VB.Label lblUserID 
      BackStyle       =   0  'Transparent
      Caption         =   "&User:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   735
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private Sub cmdBrowse_Click()
  On Error GoTo ErrExit
  With cdlg
    .DialogTitle = "Open Status File"
    .Filter = "NSS Status Files (*.nss)|*.nss|All Files|*.*"
    .FilterIndex = 0
    .ShowOpen
    txtProject.Text = .FileName
  End With

  Exit Sub

ErrExit:
  If Err.Number <> 32755 Then 'If something other than "Cancel was selected" then notify user
    MsgBox "Error opening NSS Status File '" & cdlg.FileName & "'" & vbCr _
          & Err.Description, vbCritical, "NSS"
  End If
End Sub

Private Sub cmdCancel_Click()
  'End
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim progress As String
  On Error GoTo ShowProgress
  
  progress = "Setting Hourglass"
  Me.MousePointer = vbHourglass
  
  progress = progress & vbCr & "Setting Project.Filename = " & txtProject.Text
  If Project.FileName <> txtProject.Text Then
    Project.FileName = txtProject.Text
    If Len(Dir(Project.FileName)) > 0 Then
      Project.XML = WholeFileString(Project.FileName)
    End If
  End If
  progress = progress & vbCr & "Setting Project options"
  If optUnits(0).Value Then
    Project.Metric = False
  Else
    Project.Metric = True
  End If
'  NOT HANDLING Analysis Type FOR NOW
'  If optAnalysisType(0).Value Then
'    Project.Peaks = True
'  Else
'    Project.Peaks = False
'  End If
  If optLabelType(0).Value Then
    Project.Alias = False
  Else
    Project.Alias = True
  End If
  If optVarianceType(0).Value Then
    Project.VarLog = True
  Else
    Project.VarLog = False
  End If
  progress = progress & vbCr & "Setting Project.UserName = " & txtUser.Text
  Project.UserName = txtUser.Text
  Project.RaiseEdited
  progress = progress & vbCr & "Showing frmNSS"
  'frmNSS.Show
  Me.MousePointer = vbDefault
  Unload Me
  Exit Sub

ShowProgress:
  MsgBox progress & vbCr & Err.Description, vbExclamation, "Error starting NSS"

End Sub

Private Sub Form_Load()
  Dim progress As String
  On Error GoTo ShowProgress
  
  progress = "Setting options"
  If Project.Metric Then optUnits(1).Value = True Else optUnits(0).Value = True
  If Project.Alias Then optLabelType(1).Value = True Else optLabelType(0).Value = True
  If Project.VarLog Then optVarianceType(0).Value = True Else optVarianceType(1).Value = True
  optLabelType(1).Caption = "Revised" & vbCrLf & "(% Chance)"
  progress = progress & vbCr & "txtProject = Project.Name"
  txtProject.Text = Project.FileName
  progress = progress & vbCr & "txtUser.Text = Project.UserName"
  txtUser.Text = Project.UserName
  If Len(txtUser.Text) = 0 Then txtUser.Text = APIUserName
  Exit Sub

ShowProgress:
  MsgBox progress & vbCr & Err.Description, vbExclamation, "Error loading NSS Start form"
End Sub

Private Sub Form_Resize()
  If Width > 5000 Then
    txtUser.Width = ScaleWidth - 2200
    txtProject.Width = txtUser.Width
    cmdBrowse.Left = txtUser.Left + txtUser.Width + 80
    fraButtons.Left = (ScaleWidth - fraButtons.Width) / 2
  End If
End Sub
