VERSION 5.00
Begin VB.Form frmNSS2 
   Caption         =   "National Flood Frequency Program (NSS)"
   ClientHeight    =   5544
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   8064
   Icon            =   "frmNSS2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5544
   ScaleWidth      =   8064
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraManageEstimate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   252
      Index           =   1
      Left            =   4200
      TabIndex        =   14
      Top             =   1080
      Width           =   3732
      Begin VB.CommandButton cmdEditEstimate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Edit"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   23
         Top             =   0
         Width           =   492
      End
      Begin VB.CommandButton cmdDelEstimate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Delete"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   22
         Top             =   0
         Width           =   612
      End
      Begin VB.CommandButton cmdCompute 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "New"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   17
         Top             =   0
         Width           =   492
      End
      Begin VB.CommandButton cmdEstimate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "<"
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   16
         Top             =   0
         Width           =   300
      End
      Begin VB.CommandButton cmdEstimate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   ">"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   15
         Top             =   0
         Width           =   300
      End
      Begin VB.Label lblEstimates 
         BackStyle       =   0  'Transparent
         Caption         =   "&Urban"
         Height          =   252
         Index           =   3
         Left            =   0
         TabIndex        =   20
         Top             =   30
         Width           =   612
      End
      Begin VB.Label lblEstimates 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   252
         Index           =   4
         Left            =   840
         TabIndex        =   19
         Top             =   30
         Width           =   372
      End
      Begin VB.Label lblEstimates 
         BackStyle       =   0  'Transparent
         Caption         =   "of 1"
         Height          =   252
         Index           =   5
         Left            =   1560
         TabIndex        =   18
         Top             =   36
         Width           =   372
      End
   End
   Begin VB.TextBox txtSummary 
      BackColor       =   &H00E0E0E0&
      Height          =   2892
      Index           =   1
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frmNSS2.frx":030A
      Top             =   1440
      Width           =   3732
   End
   Begin VB.TextBox txtSummary 
      BackColor       =   &H00E0E0E0&
      Height          =   2892
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "frmNSS2.frx":0310
      Top             =   1440
      Width           =   3732
   End
   Begin VB.Frame fraManageEstimate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3732
      Begin VB.CommandButton cmdDelEstimate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Delete"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   21
         Top             =   0
         Width           =   612
      End
      Begin VB.CommandButton cmdCompute 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "New"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   8
         Top             =   0
         Width           =   492
      End
      Begin VB.CommandButton cmdEditEstimate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Edit"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   7
         Top             =   0
         Width           =   492
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
      Begin VB.Label lblEstimates 
         BackStyle       =   0  'Transparent
         Caption         =   "of 1"
         Height          =   252
         Index           =   2
         Left            =   1560
         TabIndex        =   11
         Top             =   30
         Width           =   732
      End
      Begin VB.Label lblEstimates 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   252
         Index           =   1
         Left            =   840
         TabIndex        =   10
         Top             =   30
         Width           =   372
      End
      Begin VB.Label lblEstimates 
         BackStyle       =   0  'Transparent
         Caption         =   "&Rural"
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   30
         Width           =   612
      End
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Top             =   528
      Width           =   3975
   End
   Begin VB.ComboBox cboState 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Tag             =   "-999"
      Top             =   480
      Width           =   1812
   End
   Begin VB.Label lblBasinName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Site &Name:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3000
      TabIndex        =   3
      Top             =   528
      Width           =   852
   End
   Begin VB.Label lblState 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&State:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   528
      Width           =   492
   End
End
Attribute VB_Name = "frmNSS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboState_Click()
  Dim inifg&, stIndex&
  Dim myState As nssState

  'Note: (cboState.ListIndex + 1) = myNssPrj.DB.States(cboState.Text).id
  If myNssPrj.State.id <> myNssPrj.DB.States(cboState.Text).id And cboState.ListIndex >= 0 Then
    If myNssPrj.RuralScenarios.Count + myNssPrj.urbanScenarios.Count > 0 Then
      inifg = MsgBox("Changing the state will clear any existing estimates." & _
          Chr(13) & Chr(10) & "Are you sure you want to change the selected state?", 1)
      If inifg = 1 Then
        'change state
        Set myState = myNssPrj.DB.States(cboState.ListIndex)
        Set myNssPrj.State = myState
      Else
        'keep current state
        cboState.ListIndex = myNssPrj.State.id - 1
      End If
    Else
      'change state
      Set myState = myNssPrj.DB.States(cboState.ListIndex + 1)
      Set myNssPrj.State = myState
    End If
  Else
    'keep current state
    cboState.ListIndex = myNssPrj.State.id - 1
  End If
End Sub

Private Sub cmdCompute_Click(Index As Integer)
  Dim myScenario As nssScenario
  
  Set myScenario = New nssScenario
  Set myScenario.Project = myNssPrj
  myScenario.Edit
  UpdateLabels
End Sub

Private Sub cmdEstimate_Click(Index As Integer)
  myNssPrj.CurrentScenarioIncDec (Index)
  UpdateLabels
End Sub

Private Sub Form_Load()
  Dim State As Variant
  
  Set myNssPrj = New nssProject
  
  For Each State In myNssPrj.DB.States
    cboState.AddItem State.Name
  Next State
  
  myNssPrj.Filename = "c:\vbexperimental\nss2\current.nss"  'reads file
  UpdateLabels
End Sub

Private Sub UpdateLabels()
  With myNssPrj
    cboState.Text = .State.Name
    txtName = .Name
    'write rural labels
    lblEstimates(1) = .CurrentRuralScenario
    lblEstimates(2) = "of " & .RuralScenarios.Count
    If .CurrentRuralScenario > 0 Then
      txtSummary(0) = .RuralScenarios(.CurrentRuralScenario).Summary
    Else
      txtSummary(0) = "No Scenarios Available"
    End If
    'write rural labels
    lblEstimates(4) = .CurrentUrbanScenario
    lblEstimates(5) = "of " & .urbanScenarios.Count
    If .CurrentUrbanScenario > 0 Then
      txtSummary(1) = .urbanScenarios(.CurrentUrbanScenario).Summary
    Else
      txtSummary(1) = "No Scenarios Available"
    End If
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  myNssPrj.Save ("current.nss")
End Sub

Private Sub txtName_Change()
  myNssPrj.Name = txtName
End Sub
