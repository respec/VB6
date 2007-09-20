VERSION 5.00
Begin VB.Form frmUserInfo 
   Caption         =   "User Information"
   ClientHeight    =   2244
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   4128
   LinkTopic       =   "Form1"
   ScaleHeight     =   2244
   ScaleWidth      =   4128
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   720
      TabIndex        =   6
      Top             =   1680
      Width           =   612
   End
   Begin VB.TextBox txtUserInfo 
      Height          =   1092
      Index           =   2
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   2652
   End
   Begin VB.TextBox txtUserInfo 
      Height          =   288
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   2652
   End
   Begin VB.TextBox txtUserInfo 
      Height          =   288
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   2652
   End
   Begin VB.Label lblUserInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Reason for making changes: "
      Height          =   492
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Label lblUserInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "Company/Agency: "
      Height          =   252
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1452
   End
   Begin VB.Label lblUserInfo 
      Alignment       =   1  'Right Justify
      Caption         =   "User Name: "
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1332
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" _
      Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub cmdOK_Click()
  Dim response&, retval&
  Dim when$, logonName$, userName$, userOrg$, explanation$, where$
  Dim lpBuff As String * 25

  On Error Resume Next
  
  retval = GetUserName(lpBuff, 25)
  logonName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
  when = Date
  when = when & " at " & Time
  userName = txtUserInfo(0).Text
  userOrg = txtUserInfo(1).Text
  explanation = txtUserInfo(2).Text
  where = SSDB.state.Name
  If Len(userName) = 0 Or Len(userOrg) = 0 Or Len(explanation) = 0 Then
    response = myMsgBox.Show("Information must be entered for all three" & _
        " requested fields" & vbCrLf & "in order to process the changes" & _
        " to the database.", _
        "User Action Verification", "+&OK", "-&Cancel Changes")
    If response = 2 Then
      UserInfoOK = False
      Me.Hide
      MsgBox "No changes made to database"
    End If
    Exit Sub
  Else
    UserInfoOK = True
  End If
  'Write user info and general transaction info to TransactionLog table
  TransID = SSDB.RecordUserInfo(userName, logonName, userOrg, when, explanation, where)
  Me.Hide
End Sub

