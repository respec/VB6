VERSION 5.00
Begin VB.Form frmLSFact 
   Caption         =   "Form1"
   ClientHeight    =   1284
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4908
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1284
   ScaleWidth      =   4908
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLSFact 
      Caption         =   "More on LS"
      Height          =   372
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   1212
   End
   Begin VB.CommandButton cmdLSFact 
      Caption         =   "OK"
      Height          =   372
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   972
   End
   Begin VB.Label lblLSFact 
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4692
   End
End
Attribute VB_Name = "frmLSFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLSFact_Click(Index As Integer)

  If Index = 0 Then
    Me.Hide
  Else
    SendKeys "{F1}"
  End If
End Sub

Private Sub Form_Load()
  lblLSFact.Caption = "No interactive estimation tools exist for the LS Factor." & vbCrLf & "Click the More on LS button for LS Factor Help."
End Sub
