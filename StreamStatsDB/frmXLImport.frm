VERSION 5.00
Begin VB.Form frmXLImport 
   Caption         =   "StreamStatsDB Excel Import"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboSource 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Width           =   6015
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtFile 
      BackColor       =   &H8000000B&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   6015
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select File"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblSource 
      Caption         =   "Data Source:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Select Spreadsheet File to Import; Select/Enter Data Source Reference"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6615
   End
End
Attribute VB_Name = "frmXLImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdImport_Click()

  If Len(txtFile.Text) > 0 And Len(cboSource.Text) > 0 Then
    Me.MousePointer = vbHourglass
    XLSImport txtFile.Text, cboSource.Text
    Me.Tag = 1
    Me.MousePointer = vbDefault
    Me.Hide
  Else
    MsgBox "A valid Spreadsheet file name and Data Source must be entered before Importing.", vbInformation, "StreamStatsDB Excel Import"
  End If
End Sub

Private Sub cmdSelect_Click()
  Dim PathName$, stName$
  Dim Length&, i&
  
  On Error GoTo x
  
  PathName = GetSetting("StreamStatsDB", "Defaults", "XLSImportPath")
  With frmCDLG.CDLG
    .DialogTitle = "Select an Excel spreadsheet file to import"
    If Len(PathName) > 0 Then .Filename = PathName & "*.xls"
    .Filter = "(*.xls)|*.xls"
    .filterIndex = 1
    .CancelError = True
    .ShowOpen
    If Len(Dir(.Filename, vbDirectory)) > 1 Then
      PathName = PathNameOnly(.Filename)
      PathName = Left(.Filename, Len(.Filename) - Len(.fileTitle))
      SaveSetting "StreamStatsDB", "Defaults", "XLSImportPath", PathName
      txtFile.Text = .Filename
    End If
  End With
x:
  Unload frmCDLG

End Sub

Private Sub Form_Load()
  Dim vSrc As Variant

  Me.Tag = 0
  For Each vSrc In SSDB.Sources
    If UCase(vSrc.Name) <> "NONE" Then
      cboSource.AddItem CStr(vSrc.Name)
    End If
  Next

End Sub
