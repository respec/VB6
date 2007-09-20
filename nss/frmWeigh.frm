VERSION 5.00
Object = "{872F11D5-3322-11D4-9D23-00A0C9768F70}#1.6#0"; "ATCOCTL.OCX"
Begin VB.Form frmWeight 
   Caption         =   "Weight"
   ClientHeight    =   4788
   ClientLeft      =   912
   ClientTop       =   1428
   ClientWidth     =   5988
   Icon            =   "frmWeigh.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4788
   ScaleWidth      =   5988
   Begin ATCoCtl.ATCoText txtObs 
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   24
      Top             =   1680
      Width           =   975
      _ExtentX        =   1715
      _ExtentY        =   445
      InsideLimitsBackground=   16777215
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      HardMax         =   -999
      HardMin         =   0
      SoftMax         =   -999
      SoftMin         =   -999
      MaxWidth        =   5
      Alignment       =   1
      DataType        =   2
      DefaultValue    =   "0"
      Value           =   "0"
      Enabled         =   -1  'True
   End
   Begin ATCoCtl.ATCoText txtYears 
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   975
      _ExtentX        =   1715
      _ExtentY        =   445
      InsideLimitsBackground=   16777215
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      HardMax         =   -999
      HardMin         =   0
      SoftMax         =   -999
      SoftMin         =   -999
      MaxWidth        =   5
      Alignment       =   1
      DataType        =   2
      DefaultValue    =   "0"
      Value           =   "0"
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdWeight 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   4680
      TabIndex        =   12
      Top             =   4320
      Width           =   972
   End
   Begin VB.CommandButton cmdWeight 
      Caption         =   "&Apply"
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
      Left            =   3240
      TabIndex        =   11
      Top             =   4320
      Width           =   972
   End
   Begin VB.CommandButton cmdWeight 
      Caption         =   "C&ompute"
      Default         =   -1  'True
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
      Left            =   240
      TabIndex        =   10
      Top             =   4320
      Width           =   972
   End
   Begin ATCoCtl.ATCoText txtObs 
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   25
      Top             =   1920
      Width           =   975
      _ExtentX        =   1715
      _ExtentY        =   445
      InsideLimitsBackground=   16777215
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      HardMax         =   -999
      HardMin         =   0
      SoftMax         =   -999
      SoftMin         =   -999
      MaxWidth        =   5
      Alignment       =   1
      DataType        =   2
      DefaultValue    =   "0"
      Value           =   "0"
      Enabled         =   -1  'True
   End
   Begin ATCoCtl.ATCoText txtObs 
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   26
      Top             =   2160
      Width           =   975
      _ExtentX        =   1715
      _ExtentY        =   445
      InsideLimitsBackground=   16777215
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      HardMax         =   -999
      HardMin         =   0
      SoftMax         =   -999
      SoftMin         =   -999
      MaxWidth        =   5
      Alignment       =   1
      DataType        =   2
      DefaultValue    =   "0"
      Value           =   "0"
      Enabled         =   -1  'True
   End
   Begin ATCoCtl.ATCoText txtObs 
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   27
      Top             =   2400
      Width           =   975
      _ExtentX        =   1715
      _ExtentY        =   445
      InsideLimitsBackground=   16777215
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      HardMax         =   -999
      HardMin         =   0
      SoftMax         =   -999
      SoftMin         =   -999
      MaxWidth        =   5
      Alignment       =   1
      DataType        =   2
      DefaultValue    =   "0"
      Value           =   "0"
      Enabled         =   -1  'True
   End
   Begin ATCoCtl.ATCoText txtObs 
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   28
      Top             =   2640
      Width           =   975
      _ExtentX        =   1715
      _ExtentY        =   445
      InsideLimitsBackground=   16777215
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      HardMax         =   -999
      HardMin         =   0
      SoftMax         =   -999
      SoftMin         =   -999
      MaxWidth        =   5
      Alignment       =   1
      DataType        =   2
      DefaultValue    =   "0"
      Value           =   "0"
      Enabled         =   -1  'True
   End
   Begin ATCoCtl.ATCoText txtObs 
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   29
      Top             =   2880
      Width           =   975
      _ExtentX        =   1715
      _ExtentY        =   445
      InsideLimitsBackground=   16777215
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      HardMax         =   -999
      HardMin         =   0
      SoftMax         =   -999
      SoftMin         =   -999
      MaxWidth        =   5
      Alignment       =   1
      DataType        =   2
      DefaultValue    =   "0"
      Value           =   "0"
      Enabled         =   -1  'True
   End
   Begin ATCoCtl.ATCoText txtObs 
      Height          =   255
      Index           =   6
      Left            =   3120
      TabIndex        =   30
      Top             =   3120
      Width           =   975
      _ExtentX        =   1715
      _ExtentY        =   445
      InsideLimitsBackground=   16777215
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      HardMax         =   -999
      HardMin         =   0
      SoftMax         =   -999
      SoftMin         =   -999
      MaxWidth        =   5
      Alignment       =   1
      DataType        =   2
      DefaultValue    =   "0"
      Value           =   "0"
      Enabled         =   -1  'True
   End
   Begin ATCoCtl.ATCoText txtObs 
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   31
      Top             =   3360
      Width           =   975
      _ExtentX        =   1715
      _ExtentY        =   445
      InsideLimitsBackground=   16777215
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      HardMax         =   -999
      HardMin         =   0
      SoftMax         =   -999
      SoftMin         =   -999
      MaxWidth        =   5
      Alignment       =   1
      DataType        =   2
      DefaultValue    =   "0"
      Value           =   "0"
      Enabled         =   -1  'True
   End
   Begin VB.Label lblCol 
      BackStyle       =   0  'Transparent
      Caption         =   "Weighted Flow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   35
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblCol 
      BackStyle       =   0  'Transparent
      Caption         =   "Observed Flow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3170
      TabIndex        =   34
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblCol 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimated Flow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   33
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblCol 
      BackStyle       =   0  'Transparent
      Caption         =   "Recurrence Interval"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   32
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   4440
      TabIndex        =   23
      Top             =   3360
      Width           =   1092
   End
   Begin VB.Label lblWeight 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   9
      Left            =   240
      TabIndex        =   22
      Top             =   3360
      Width           =   2892
   End
   Begin VB.Label lblCmd 
      BackStyle       =   0  'Transparent
      Caption         =   "Apply replaces Estimated values with Weighted values."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   21
      Top             =   3960
      Width           =   5415
   End
   Begin VB.Label lblCmd 
      BackStyle       =   0  'Transparent
      Caption         =   "Compute generates weighted values above."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   20
      Top             =   3720
      Width           =   5415
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   4440
      TabIndex        =   19
      Top             =   3120
      Width           =   1092
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   4440
      TabIndex        =   18
      Top             =   2880
      Width           =   1092
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   4440
      TabIndex        =   17
      Top             =   2640
      Width           =   1092
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   4440
      TabIndex        =   16
      Top             =   2400
      Width           =   1092
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   4440
      TabIndex        =   15
      Top             =   2160
      Width           =   1092
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   4440
      TabIndex        =   14
      Top             =   1920
      Width           =   1092
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   4440
      TabIndex        =   13
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Label lblWeight 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   8
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   2892
   End
   Begin VB.Label lblWeight 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   2892
   End
   Begin VB.Label lblWeight 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   2892
   End
   Begin VB.Label lblWeight 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   2892
   End
   Begin VB.Label lblWeight 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   2892
   End
   Begin VB.Label lblWeight 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   2892
   End
   Begin VB.Label lblWeight 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2892
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter observed data for each interval:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   4212
   End
   Begin VB.Label lblYears 
      BackStyle       =   0  'Transparent
      Caption         =   "Years of observed data:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3132
   End
End
Attribute VB_Name = "frmWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim nobs!, obs!(MAX_INTERVALS), wgt!(MAX_INTERVALS)
'
'
'Private Sub cmdWeight_Click(Index As Integer)
'
'    Dim i&
'
'    If Index = 0 Then
'      'perform weighted calculations
'      nobs = txtYears.Value
'      If nobs > 0 Then
'        Call CalcWeight
'      Else
'        MsgBox "A value for the number of years of observed data must be entered to generate a weighted estimate.", 48, "NSS Weight"
'      End If
'    ElseIf Index = 1 Then
'      If wgt(0) > 0 Then
'        frmNSS.cmdWeight.Enabled = False
'        'replace estimated values w/weighted values
'        For i = 0 To NumIntrvl - 1
'          rural_discharge(0, i, rurind) = wgt(i)
'          rural_discharge(1, i, rurind) = 0
'          rural_discharge(2, i, rurind) = rural_discharge(2, i, rurind) + nobs
'        Next i
'        If InStr(rurscn(rurind).Name, "(Weighted)") = 0 Then
'          'indicate that estimate results are weighted
'          rurscn(rurind).Name = Trim(rurscn(rurind).Name) & " (Weighted)"
'          Call DispEstimate
'        End If
'        Call DispRuralDis
'      Else
'        'no weighted values
'        MsgBox "No Weighted values have been calculated.  Click the Compute to generate Weighted values.", 48, "NSS Weight"
'        Index = 0
'      End If
'    End If
'    If Index > 0 Then
'      Hide
'    End If
'
'End Sub
'
'Private Sub Form_Activate()
'
'  Dim i&, lstr$
'  Call SetIntervals
'  For i = 0 To NumIntrvl - 1
'    'display discharge values for each interval
'    lstr = NumFmtI(CInt(Interval(i)), 5)
'    lblWeight(i + 2) = lstr & NumFmted(Signif(CDbl(rural_discharge(0, i, rurind)), metric), 16, 0)
'    lblRes(i) = ""
'    txtObs(i).Visible = True
'  Next i
'  For i = NumIntrvl To MAX_INTERVALS
'    'disable unused interval entry fields
'    txtObs(i).Visible = False
'  Next i
'
'End Sub
'
'Public Sub CalcWeight()
'
'    Dim i&, tmp!
'
'    For i = 0 To NumIntrvl - 1
'      obs(i) = txtObs(i).Value
'      If rural_discharge(0, i, rurind) > 0 And obs(i) > 0 Then
'        'estimated and observed values ok
'        tmp = (rural_discharge(2, i, rurind) * Log10(CDbl(rural_discharge(0, i, rurind))) + nobs * Log10(CDbl(obs(i)))) / ((rural_discharge(2, i, rurind) + nobs))
'        wgt(i) = 10# ^ tmp
'      Else
'        'estimated or observed value invalid
'        wgt(i) = 0
'      End If
'      lblRes(i).Caption = NumFmted(Signif(CDbl(wgt(i)), metric), 8, 0)
'    Next i
'
'End Sub
