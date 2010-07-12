VERSION 5.00
Begin VB.Form frmXLImport 
   Caption         =   "StreamStatsDB Excel Import"
   ClientHeight    =   2595
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
   ScaleHeight     =   2595
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboSourceURL 
      Height          =   315
      Left            =   1440
      TabIndex        =   8
      Top             =   1440
      Width           =   6015
   End
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
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2160
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
   Begin VB.Label lblSourceURL 
      Caption         =   "Source URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
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

Private Sub cboSource_Click()
  Dim i As Integer
  Dim lSrcInd As Integer
  Dim lURLInd As Integer
  
  lSrcInd = cboSource.ItemData(cboSource.ListIndex)
  lURLInd = -1
  For i = 0 To cboSourceURL.ListCount - 1
    If cboSourceURL.ItemData(i) = lSrcInd Then
      lURLInd = i
      Exit For
    End If
  Next i
  cboSourceURL.Text = cboSourceURL.List(lURLInd)

End Sub

Private Sub cboSourceURL_Click()
  Dim i As Integer
  Dim lSrcInd As Integer
  Dim lURLInd As Integer
  
  lURLInd = cboSourceURL.ItemData(cboSourceURL.ListIndex)
  lSrcInd = -1
  For i = 0 To cboSource.ListCount - 1
    If cboSource.ItemData(i) = lURLInd Then
      lSrcInd = i
      Exit For
    End If
  Next i
  cboSource.Text = cboSource.List(lSrcInd)
End Sub

Private Sub cmdCancel_Click()
  Me.Hide
End Sub

Private Sub cmdImport_Click()

  If Len(txtFile.Text) > 0 And Len(cboSource.Text) > 0 Then
    Me.MousePointer = vbHourglass
    XLSImport txtFile.Text, cboSource.Text, cboSourceURL.Text
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
    If Len(PathName) > 0 Then .filename = PathName & "*.xls"
    .Filter = "(*.xls)|*.xls"
    .filterIndex = 1
    .CancelError = True
    .ShowOpen
    If Len(Dir(.filename, vbDirectory)) > 1 Then
      PathName = PathNameOnly(.filename)
      PathName = Left(.filename, Len(.filename) - Len(.fileTitle))
      SaveSetting "StreamStatsDB", "Defaults", "XLSImportPath", PathName
      txtFile.Text = .filename
    End If
  End With
x:
  Unload frmCDLG

End Sub

Private Sub Form_Load()
  Dim vSrc As Variant
  Dim i As Integer, j As Integer

  Me.Tag = 0
  i = 0
  j = 0
  For Each vSrc In SSDB.Sources
    If UCase(vSrc.Name) <> "NONE" Then
      cboSource.AddItem CStr(vSrc.Name)
      cboSource.ItemData(i) = i
      If Len(vSrc.URL) > 0 Then
        cboSourceURL.AddItem CStr(vSrc.URL)
        cboSourceURL.ItemData(j) = i
        j = j + 1
      End If
      i = i + 1
    End If
  Next

End Sub
