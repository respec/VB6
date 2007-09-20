VERSION 4.00
Begin VB.Form frmUnits 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "NSS Startup"
   ClientHeight    =   2352
   ClientLeft      =   2664
   ClientTop       =   3288
   ClientWidth     =   3744
   Height          =   2736
   Left            =   2616
   LinkTopic       =   "Form1"
   ScaleHeight     =   2352
   ScaleWidth      =   3744
   Top             =   2952
   Width           =   3840
   Begin VB.TextBox txtProjectID 
      Height          =   288
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   1932
   End
   Begin VB.TextBox txtUserID 
      Height          =   288
      Left            =   1560
      TabIndex        =   0
      Top             =   1044
      Width           =   1932
   End
   Begin VB.CommandButton cmdUnits 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdUnits 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   855
   End
   Begin VB.OptionButton optUnits 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Metric"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   1
      Left            =   1320
      TabIndex        =   5
      Top             =   480
      Width           =   732
   End
   Begin VB.OptionButton optUnits 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "English"
      ForeColor       =   &H80000008&
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   852
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Enter Project ID:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   1212
   End
   Begin VB.Label lblUserID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Enter User ID:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1092
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select units for calculations:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2052
   End
End
Attribute VB_Name = "frmUnits"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim lmetric%
Private Sub cmdUnits_Click(Index As Integer)

    If Index = 0 Then
      'set unit system
      If optUnits(0).value = True Then
        metric = False
      Else
        metric = True
      End If
      If lmetric <= 0 And lmetric <> metric Then
        'changing unit system from last run
        frmCnvrt.Show 1
      Else
        cnvrtfg = 0
      End If
      If cnvrtfg <> 2 Then
        'cnvrtfg = 2 means user cancelled
        'Convert window, so keep this one open
        UsrID = txtUserID.Text
        PrjID = txtProjectID.Text
        Hide
      End If
    ElseIf Index = 1 Then
      'exit program
      End
    End If

End Sub


Private Sub Form_Load()

    'save initial units value
    lmetric = metric
    If metric = True Then
      optUnits(1).value = True
    Else
      optUnits(0).value = True
    End If

End Sub


