VERSION 5.00
Begin VB.Form frmCnvrt 
   Caption         =   "Unit System Changed"
   ClientHeight    =   1896
   ClientLeft      =   5352
   ClientTop       =   3372
   ClientWidth     =   3180
   Icon            =   "frmCnvrt.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1896
   ScaleWidth      =   3180
   Begin VB.CommandButton cmdCnvrt 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Index           =   2
      Left            =   1680
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdCnvrt 
      Caption         =   "C&lear"
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
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdCnvrt 
      Caption         =   "C&onvert"
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
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "The selected unit system is different from the last session.  What should be done with the data from that session?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmCnvrt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCnvrt_Click(Index As Integer)

'  Dim i&, j&, k&
'  Dim confact!, tmp!
'
'  If Index = 0 Then
'    'convert existing data to new units system
'    For i = 0 To rurcnt - 1
'      For j = 0 To rurscn(i).rcnt - 1
'        For k = 0 To rurscn(i).v(j).vcount
'          confact = units(State.Region(rurscn(i).reg(j)).v_descriptor(k).units).factor
'          If metric = True Then 'converting to metric
'            rurscn(i).v(j).Value(k) = Signif(CDbl(ConvertVal(rurscn(i).v(j).Value(k), confact)), metric)
'          Else                  'converting to english
'            rurscn(i).v(j).Value(k) = Signif(CDbl(ConvertVal(rurscn(i).v(j).Value(k), 1 / confact)), metric)
'          End If
'        Next k
'      Next j
'    Next i
'    'convert total area
'    If metric = True Then
'      total_area = Signif(CDbl(total_area * AREA_CONVERSION), metric)
'    Else
'      total_area = Signif(CDbl(total_area / AREA_CONVERSION), metric)
'    End If
''    frmNSS.txtBasinArea.value = total_area
'    Call frmNSS.SetEstimate
'  ElseIf Index = 1 Then
'    'clear existing data for this run
'    i = frmNSS.cboState.ListIndex
'    frmNSS.cboState.ListIndex = -1
'    frmNSS.cboState.ListIndex = i
'  End If
'  cnvrtfg = Index
'  Hide

End Sub
