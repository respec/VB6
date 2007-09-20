VERSION 5.00
Begin VB.Form frmROIImportRegions 
   Caption         =   "ROI Regions"
   ClientHeight    =   3372
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   3648
   LinkTopic       =   "Form1"
   ScaleHeight     =   3372
   ScaleWidth      =   3648
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
      Height          =   372
      Left            =   1560
      TabIndex        =   3
      Top             =   3000
      Width           =   612
   End
   Begin VB.TextBox txtRegionName 
      Height          =   288
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   360
      Width           =   2892
   End
   Begin VB.Label lblRegnName 
      Alignment       =   2  'Center
      Caption         =   "Region Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   2892
   End
   Begin VB.Label lblRegnIndex 
      Alignment       =   2  'Center
      Caption         =   "Index"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   20
      TabIndex        =   1
      Top             =   120
      Width           =   560
   End
   Begin VB.Label lblRegionID 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   252
      Index           =   0
      Left            =   20
      TabIndex        =   0
      Top             =   360
      Width           =   560
   End
End
Attribute VB_Name = "frmROIImportRegions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RegnCount As Long

Private Sub cmdok_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim i As Long
  
  RegnCount = UBound(ROIImportRegnIDs)
  ReDim ROIImportRegnNames(RegnCount)
  For i = 1 To RegnCount
    lblRegionID(i - 1) = ROIImportRegnIDs(i) & " :  "
    If i < UBound(ROIImportRegnIDs) Then
      Load txtRegionName(i)
      txtRegionName(i).Top = (i + 1) * 360
      txtRegionName(i).Visible = True
      Load lblRegionID(i)
      lblRegionID(i).Top = (i + 1) * 360
      lblRegionID(i).Visible = True
    End If
  Next i
  cmdOk.Top = 360 * RegnCount + 420
  Me.Height = 360 * RegnCount + 1200

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim i As Long
  
  For i = 1 To RegnCount
    ROIImportRegnNames(i) = txtRegionName(i - 1).Text
  Next i
End Sub

