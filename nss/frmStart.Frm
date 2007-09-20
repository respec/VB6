VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStart 
   Caption         =   "NSS"
   ClientHeight    =   2505
   ClientLeft      =   2670
   ClientTop       =   3285
   ClientWidth     =   4245
   HelpContextID   =   11
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2505
   ScaleWidth      =   4245
   Begin VB.TextBox txtProject 
      Height          =   288
      HelpContextID   =   11
      Left            =   1080
      TabIndex        =   6
      Top             =   1440
      Width           =   1572
   End
   Begin VB.TextBox txtUser 
      Height          =   288
      HelpContextID   =   11
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   1572
   End
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
      Left            =   1560
      TabIndex        =   2
      Top             =   480
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
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   2520
      Top             =   0
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
      TabIndex        =   7
      Top             =   1920
      Width           =   2412
      Begin VB.CommandButton cmdOk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Run"
         Default         =   -1  'True
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
         TabIndex        =   10
         Top             =   0
         Width           =   732
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Cancel          =   -1  'True
         Caption         =   "&Quit"
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
         Left            =   1800
         TabIndex        =   9
         Top             =   0
         Width           =   612
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
         Height          =   375
         HelpContextID   =   11
         Left            =   840
         TabIndex        =   8
         ToolTipText     =   "Open a different project"
         Top             =   0
         Width           =   852
      End
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
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   1480
      Width           =   1212
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
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   1000
      Width           =   732
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2895
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
    txtProject.Text = .Filename
  End With

  Exit Sub

ErrExit:
  If Err.Number <> 32755 Then 'If something other than "Cancel was selected" then notify user
    MsgBox "Error opening NSS Status File '" & cdlg.Filename & "'" & vbCr _
          & Err.Description, vbCritical, "NSS"
  End If
End Sub

Private Sub cmdCancel_Click()
  End
End Sub

Private Sub cmdOK_Click()
  Dim progress As String
  On Error GoTo ShowProgress
  
  progress = "Setting Hourglass"
  Me.MousePointer = vbHourglass
  
  progress = progress & vbCr & "Setting Project.Filename = " & txtProject.Text
  If Project.Filename <> txtProject.Text Then
    Project.Filename = txtProject.Text
    If Len(Dir(Project.Filename)) > 0 Then
      Project.XML = WholeFileString(Project.Filename)
    End If
  End If
  progress = progress & vbCr & "Setting Project.metric"
  If optUnits(0).Value Then
    Project.Metric = False
  Else
    Project.Metric = True
  End If
  progress = progress & vbCr & "Setting Project.UserName = " & txtUser.Text
  Project.UserName = txtUser.Text
  progress = progress & vbCr & "Showing frmNSS"
  frmNSS.Show
  Me.MousePointer = vbDefault
  Unload Me
  Exit Sub

ShowProgress:
  MsgBox progress & vbCr & Err.Description, vbExclamation, "Error starting NSS"

End Sub

Private Sub Form_Load()
  Dim progress As String
  On Error GoTo ShowProgress
  
  progress = "Setting optUnits to Project.metric"
  If Project.Metric Then optUnits(1).Value = True Else optUnits(0).Value = True
  progress = progress & vbCr & "txtProject = Project.Name"
  txtProject.Text = Project.Filename
  progress = progress & vbCr & "txtUser.Text = Project.UserName"
  txtUser.Text = Project.UserName
  If Len(txtUser.Text) = 0 Then txtUser.Text = APIUserName
  Exit Sub

ShowProgress:
  MsgBox progress & vbCr & Err.Description, vbExclamation, "Error loading NSS Start form"
End Sub

Private Sub Form_Resize()
  If Width > 1800 Then
    txtUser.Width = ScaleWidth - 1330
    txtProject.Width = txtUser.Width
    fraButtons.Left = (ScaleWidth - fraButtons.Width) / 2
  End If
End Sub
