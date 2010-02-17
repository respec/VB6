VERSION 5.00
Object = "{872F11D5-3322-11D4-9D23-00A0C9768F70}#1.10#0"; "ATCoCtl.ocx"
Begin VB.Form frmEditWeight 
   Caption         =   "Edit Weight"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   8355
   HelpContextID   =   27
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optWeight 
      Caption         =   "Weight for gaged site using observed data and variance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   3120
      TabIndex        =   38
      Top             =   120
      Width           =   2295
   End
   Begin ATCoCtl.ATCoGrid grdWgt 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3836
      SelectionToggle =   0   'False
      AllowBigSelection=   0   'False
      AllowEditHeader =   0   'False
      AllowLoad       =   0   'False
      AllowSorting    =   0   'False
      Rows            =   2
      Cols            =   4
      ColWidthMinimum =   1000
      gridFontBold    =   0   'False
      gridFontItalic  =   0   'False
      gridFontName    =   "MS Sans Serif"
      gridFontSize    =   8
      gridFontUnderline=   0   'False
      gridFontWeight  =   400
      gridFontWidth   =   0
      Header          =   ""
      FixedRows       =   2
      FixedCols       =   0
      ScrollBars      =   0
      SelectionMode   =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorBkg    =   -2147483632
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      InsideLimitsBackground=   -2147483643
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      ComboCheckValidValues=   0   'False
   End
   Begin VB.OptionButton optWeight 
      Caption         =   "Weight for ungaged site using weighted gaged values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   6120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.OptionButton optWeight 
      Caption         =   "Weight for gaged site using observed data and equivalent years"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   3480
      TabIndex        =   6
      Top             =   4320
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4800
      TabIndex        =   7
      Top             =   4320
      Width           =   972
   End
   Begin VB.TextBox txtYears 
      Alignment       =   1  'Right Justify
      Height          =   252
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   972
   End
   Begin VB.ComboBox cboWgtSelect 
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   2172
   End
   Begin VB.Label lblWgtSelect 
      BackStyle       =   0  'Transparent
      Caption         =   "Select scenario containing weighted gaged values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   37
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label lblEstimated 
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
      Height          =   255
      Index           =   8
      Left            =   1560
      TabIndex        =   36
      Top             =   3960
      Width           =   1290
   End
   Begin VB.Label lblEstimated 
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
      Height          =   255
      Index           =   7
      Left            =   1560
      TabIndex        =   35
      Top             =   3720
      Width           =   1290
   End
   Begin VB.Label lblEstimated 
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
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   34
      Top             =   3480
      Width           =   1290
   End
   Begin VB.Label lblEstimated 
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
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   33
      Top             =   3240
      Width           =   1290
   End
   Begin VB.Label lblEstimated 
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
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   32
      Top             =   3000
      Width           =   1290
   End
   Begin VB.Label lblEstimated 
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
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   31
      Top             =   2760
      Width           =   1290
   End
   Begin VB.Label lblEstimated 
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
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   30
      Top             =   2520
      Width           =   1290
   End
   Begin VB.Label lblInterval 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   29
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lblInterval 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   28
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblInterval 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   27
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lblInterval 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   26
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblInterval 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   25
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblInterval 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   24
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblInterval 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   23
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblYears 
      BackStyle       =   0  'Transparent
      Caption         =   "Years of observed data:"
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
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter observed data for each interval:"
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
      Left            =   120
      TabIndex        =   21
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label lblInterval 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblEstimated 
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
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   19
      Top             =   2280
      Width           =   1290
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
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   18
      Top             =   2280
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
      Height          =   255
      Index           =   2
      Left            =   4320
      TabIndex        =   17
      Top             =   2520
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
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   16
      Top             =   2760
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
      Height          =   255
      Index           =   4
      Left            =   4320
      TabIndex        =   15
      Top             =   3000
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
      Height          =   255
      Index           =   5
      Left            =   4320
      TabIndex        =   14
      Top             =   3240
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
      Height          =   255
      Index           =   6
      Left            =   4320
      TabIndex        =   13
      Top             =   3480
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
      Height          =   255
      Index           =   7
      Left            =   4320
      TabIndex        =   12
      Top             =   3720
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
      Height          =   255
      Index           =   8
      Left            =   4320
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblCol 
      BackStyle       =   0  'Transparent
      Caption         =   "Recurrence Interval (years)"
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
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblCol 
      BackStyle       =   0  'Transparent
      Caption         =   "Estimated Flow"
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
      Index           =   1
      Left            =   1560
      TabIndex        =   9
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblCol 
      BackStyle       =   0  'Transparent
      Caption         =   "Observed Flow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3045
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblCol 
      BackStyle       =   0  'Transparent
      Caption         =   "Weighted Flow"
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
      Index           =   3
      Left            =   4320
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "frmEditWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private pScenario As nssScenario
Private pSaveWeight As nssWeight 'Original pScenario.Weight in case we want to cancel
Private pCurrentControl As Long
Private pFirstReturns As FastCollection
Private pFinishedInit As Boolean

Public Property Set Scenario(newValue As nssScenario)
  Dim ReturnIndex As Long
  pFinishedInit = False
  Set pScenario = newValue
  Set pSaveWeight = pScenario.Weight
  Set pScenario.Weight = pSaveWeight.Copy
  If pScenario.Weight.WeightType = 0 Then
    pScenario.Weight.WeightType = 1
  End If
  Me.Caption = "Editing weight for """ & pScenario.Name & """"
  PopulateIntervals
  pFinishedInit = True
  For ReturnIndex = 1 To pFirstReturns.Count
    If IsNumeric(txtYears.Text) Then
      pScenario.Weight.SetGagedYears pFirstReturns(ReturnIndex).Name, txtYears.Text
    End If
    If IsNumeric(grdWgt.TextMatrix(ReturnIndex, 2)) Then
      pScenario.Weight.SetGagedValue pFirstReturns(ReturnIndex).Name, grdWgt.TextMatrix(ReturnIndex, 2)
    End If
    If pScenario.Weight.WeightType = 2 And IsNumeric(grdWgt.TextMatrix(ReturnIndex, 3)) Then 'include variance
      pScenario.Weight.SetGagedVariance pFirstReturns(ReturnIndex).Name, grdWgt.TextMatrix(ReturnIndex, 3)
    End If
  Next
  PopulateResults
End Property

Private Sub PopulateIntervals()
  Dim ReturnIndex As Long
  Dim Interval As String
  Dim d() As Double
  Dim v() As Double
  Dim lCol As Integer

  If pScenario.Weight.WeightType > 0 Then
    optWeight(pScenario.Weight.WeightType - 1).Value = True
  Else
    optWeight(0).Value = True
  End If
  If pScenario.UserRegions.Count = 0 Then
    'Can't do anything with no user regions
  Else
    Set pFirstReturns = pScenario.UserRegions(1).Region.DepVars
    d = pScenario.Discharges
    v = pScenario.Variances
    For ReturnIndex = 1 To pFirstReturns.Count
      Interval = pFirstReturns(ReturnIndex).Name
      grdWgt.TextMatrix(ReturnIndex, 0) = Interval
      'quirky feature of grid requires setting back color after setting text
      grdWgt.col = 0
      grdWgt.row = ReturnIndex
      grdWgt.CellBackColor = &HE0E0E0
      grdWgt.TextMatrix(ReturnIndex, 1) = StrPad(SignificantDigits(d(ReturnIndex), 3), 9)
      lCol = 1
      grdWgt.col = lCol
      grdWgt.CellBackColor = &HE0E0E0
      If pScenario.Weight.WeightType = 2 Then
        grdWgt.TextMatrix(ReturnIndex, 2) = StrPad(SignificantDigits(v(ReturnIndex), 3), 9)
        lCol = 2
        grdWgt.col = lCol
        grdWgt.CellBackColor = &HE0E0E0
      End If
      grdWgt.TextMatrix(ReturnIndex, lCol + 1) = pScenario.Weight.GetGagedValue(Interval)
      grdWgt.col = lCol + 1
      grdWgt.CellBackColor = grdWgt.BackColor
      If pScenario.Weight.WeightType = 1 Then 'populate equiv years text box
        txtYears.Text = pScenario.Weight.GetGagedYears(Interval)
        grdWgt.col = 3
        grdWgt.CellBackColor = &HE0E0E0
      ElseIf pScenario.Weight.WeightType = 2 Then 'populate variance values
        grdWgt.TextMatrix(ReturnIndex, 4) = pScenario.Weight.GetGagedVariance(Interval)
        grdWgt.col = 5
      End If
    Next
    grdWgt.Height = 242 * (pFirstReturns.Count + 2)
    cmdApply(1).Top = grdWgt.Top + grdWgt.Height + 120
    cmdCancel.Top = cmdApply(1).Top
    frmEditWeight.Height = cmdCancel.Top + cmdCancel.Height + 440
  End If
End Sub

Private Sub cboWgtSelect_Click()
  Dim i&, d() As Double
  Dim e() As Double
  Dim WgtScenario As nssScenario
  Dim ScenarioAreaStateUnits As Double

  pFinishedInit = False
  For i = 1 To pScenario.Project.RuralScenarios.Count
    If pScenario.Project.RuralScenarios(i).Name = cboWgtSelect.Text Then
      'found selected scenario name
      Set WgtScenario = pScenario.Project.RuralScenarios(i)
    End If
  Next i
  pScenario.Weight.AreaGaged = WgtScenario.GetArea(WgtScenario.Project.State.Metric)
  ScenarioAreaStateUnits = pScenario.GetArea(WgtScenario.Project.State.Metric)
  If ScenarioAreaStateUnits < pScenario.Weight.AreaGaged / 2 Then
    ssMessageBox "Drainage area of ungaged site is less than half the area of the gaged site." & vbCr _
         & "This weighting method is not reccommended for this condition.", vbCritical, "NSS Weight"
  ElseIf ScenarioAreaStateUnits > pScenario.Weight.AreaGaged * 1.5 Then
    ssMessageBox "Drainage area of ungaged site is more than 1.5 times the area of the gaged site." & vbCr _
         & "This weighting method is not reccommended for this condition.", vbCritical, "NSS Weight"
  End If
  d = WgtScenario.WeightedDischarges
  e = WgtScenario.EquivalentYears
  For i = 1 To pScenario.UserRegions(1).Region.DepVars.Count
    grdWgt.TextMatrix(i, 2) = StrPad(SignificantDigits(d(i), 3), 9)
    If IsNumeric(grdWgt.TextMatrix(i, 2)) Then
      pScenario.Weight.SetGagedValue pScenario.UserRegions(1).Region.DepVars(i).Name, grdWgt.TextMatrix(i, 2)
      pScenario.Weight.SetGagedYears pScenario.UserRegions(1).Region.DepVars(i).Name, e(i)
    End If
  Next i
  pFinishedInit = True
  PopulateResults

End Sub

Private Sub cboWgtSelect_GotFocus()
  pCurrentControl = 0
End Sub

Private Sub cmdApply_Click(Index As Integer)
  Dim i As Integer
  Set pSaveWeight = Nothing
  
  For i = 1 To pScenario.Project.RuralScenarios.Count
    If pScenario.Project.RuralScenarios.ItemByIndex(i).Name = pScenario.Name Then
      pScenario.Project.CurrentRuralScenario = i
    End If
  Next
  pScenario.Project.RaiseEdited
  Unload Me
End Sub

Private Sub cmdApply_GotFocus(Index As Integer)
  pCurrentControl = 0
End Sub

Private Sub cmdCancel_Click()
  Set pScenario.Weight = Nothing
  Set pScenario.Weight = pSaveWeight
  Set pSaveWeight = Nothing
  Unload Me
End Sub

Private Sub cmdCancel_GotFocus()
  pCurrentControl = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn And pCurrentControl <> 0 Then
    If pCurrentControl = -1 Then pCurrentControl = 0
    If pCurrentControl + 1 > grdWgt.rows Then 'txtObs.Count Then
      cmdApply(1).SetFocus
    Else
      grdWgt.SetFocus
      grdWgt.row = pCurrentControl + 1
    End If
    KeyCode = 0 'Don't let individual controls see this key if we used it
  End If
End Sub

Private Sub Form_Load()
  Dim i As Integer

  grdWgt.TextMatrix(-1, 0) = "Recurrence"
  grdWgt.TextMatrix(0, 0) = "Interval (yrs)"
  grdWgt.ColType(0) = ATCoTxt
  grdWgt.TextMatrix(-1, 1) = "Estimated"
  grdWgt.TextMatrix(0, 1) = "Flow"

End Sub

Private Sub grdWgt_GotFocus()
  grdWgt.row = 1
  grdWgt.col = 2
End Sub

Private Sub grdWgt_RowColChange()
  pCurrentControl = grdWgt.row
End Sub

Private Sub grdWgt_TextChange(ChangeFromRow As Long, ChangeToRow As Long, ChangeFromCol As Long, ChangeToCol As Long)
  Dim ReturnIndex As Long
  Dim lCol As Long

  If pFinishedInit Then
    If pScenario.Weight.WeightType = 1 Then
      lCol = 2
    ElseIf pScenario.Weight.WeightType = 2 Then
      lCol = 3
    End If
    For ReturnIndex = 1 To pFirstReturns.Count
      If IsNumeric(grdWgt.TextMatrix(ReturnIndex, lCol)) Then
        pScenario.Weight.SetGagedValue pFirstReturns(ReturnIndex).Name, grdWgt.TextMatrix(ReturnIndex, lCol)
        If pScenario.Weight.WeightType = 2 And IsNumeric(grdWgt.TextMatrix(ReturnIndex, lCol + 1)) Then
          pScenario.Weight.SetGagedVariance pFirstReturns(ReturnIndex).Name, grdWgt.TextMatrix(ReturnIndex, lCol + 1)
        End If
      End If
    Next
    PopulateResults
  End If

End Sub

Private Sub optWeight_Click(Index As Integer)
  Dim i As Long, j As Long
  Dim lLastCol As Long
  Dim lWid As Long

  If Index = 1 Then 'need extra field for Variance weighting
    lLastCol = 5
  Else
    lLastCol = 3
  End If
  grdWgt.cols = lLastCol + 1
  
  If Index = 0 Then 'weight by Equivalent Years
    lblWgtSelect.Visible = False
    cboWgtSelect.Visible = False
    lblYears.Visible = True
    txtYears.Visible = True
    Label1.Visible = True
    grdWgt.TextMatrix(-1, 2) = "Observed"
    grdWgt.TextMatrix(0, 2) = "Flow"
    grdWgt.ColEditable(2) = True
    pScenario.Weight.WeightType = 1
  ElseIf Index = 1 Then 'weight by Variance
    lblWgtSelect.Visible = False
    cboWgtSelect.Visible = False
    lblYears.Visible = False
    txtYears.Visible = False
    Label1.Visible = True
    grdWgt.TextMatrix(-1, 2) = "Estimated"
    grdWgt.TextMatrix(0, 2) = "Variance"
    grdWgt.ColEditable(2) = False
    grdWgt.TextMatrix(-1, 3) = "Observed"
    grdWgt.TextMatrix(0, 3) = "Flow"
    grdWgt.ColEditable(3) = True
    grdWgt.TextMatrix(-1, 4) = "Observed"
    grdWgt.TextMatrix(0, 4) = "Variance"
    grdWgt.ColEditable(4) = True
    grdWgt.ColAlignment(4) = 7
    grdWgt.TextMatrix(-1, lLastCol + 1) = "Weighted"
    grdWgt.TextMatrix(0, lLastCol + 1) = "Variance"
    grdWgt.ColEditable(lLastCol + 1) = False
    grdWgt.ColAlignment(lLastCol + 1) = 7
    grdWgt.TextMatrix(-1, lLastCol + 2) = "Weighted"
    grdWgt.TextMatrix(0, lLastCol + 2) = "Std Error, %"
    grdWgt.ColEditable(lLastCol + 2) = False
    grdWgt.ColAlignment(lLastCol + 2) = 7
    pScenario.Weight.WeightType = 2
  Else 'weight from previously weighted estimate
    lblWgtSelect.Visible = True
    cboWgtSelect.Visible = True
    lblYears.Visible = False
    txtYears.Visible = False
    Label1.Visible = False
    grdWgt.TextMatrix(-1, 2) = "Weighted"
    grdWgt.TextMatrix(0, 2) = "Gaged Flow"
    grdWgt.ColEditable(2) = False
    pScenario.Weight.WeightType = 3
    cboWgtSelect.Clear
    If pScenario.Project.RuralScenarios.Count > 0 Then
      'look for other scenarios that have gaged weighted results
      For i = 1 To pScenario.Project.RuralScenarios.Count
        If pScenario.Project.RuralScenarios(i).Weight.WeightType = 1 Or _
           pScenario.Project.RuralScenarios(i).Weight.WeightType = 2 Then
          'add to combo box list
          cboWgtSelect.AddItem pScenario.Project.RuralScenarios(i).Name
        End If
      Next i
    End If
    If cboWgtSelect.ListCount = 0 Then 'no gaged weighted scenarios available
      ssMessageBox "There are no scenarios with gaged weighted values that use the same region as the scenario being weighted." & vbCrLf & _
             "Thus, this type of weighting may not be performed currently.", vbExclamation, "Weighting Problem"
      optWeight(0).Value = True
      Exit Sub
    Else
      cboWgtSelect.ListIndex = 0
    End If
  End If
  For i = 1 To lLastCol
    grdWgt.ColAlignment(i) = 7
  Next i
  grdWgt.ColEditable(lLastCol) = False
  grdWgt.TextMatrix(-1, lLastCol) = "Weighted"
  grdWgt.TextMatrix(0, lLastCol) = "Flow"
  
  If pScenario.Weight.WeightType = 2 Then
    lWid = 980
  Else
    lWid = 2000
  End If
'  grdWgt.colWidth(0) = lWid
'  grdWgt.colWidth(1) = lWid
  For j = 0 To lLastCol
    grdWgt.colWidth(j) = lWid
    grdWgt.col = j
    For i = 1 To grdWgt.rows
      grdWgt.row = i
      If (Index = 0 And j = 2) Or (Index = 1 And j = 3) Then 'set observed flow column to white background for editing
        grdWgt.CellBackColor = grdWgt.BackColor
      ElseIf Index = 1 And j = 4 Then 'set observed variance column to white background for editing
        grdWgt.CellBackColor = grdWgt.BackColor
      Else 'column not editable
        grdWgt.CellBackColor = &HE0E0E0
        If Index = 1 And j = lLastCol Then 'grey out two add'l fields for variance and SE
          grdWgt.col = j + 1
          grdWgt.CellBackColor = &HE0E0E0
          grdWgt.col = j + 2
          grdWgt.CellBackColor = &HE0E0E0
          grdWgt.col = j
        End If
      End If
    Next i
  Next j
  PopulateIntervals
  PopulateResults
End Sub

Private Sub optWeight_GotFocus(Index As Integer)
  pCurrentControl = 0
End Sub

Private Sub txtYears_Change()
  Dim ReturnIndex As Long
  Debug.Print "txtYears_Change " & txtYears.Text
  If pFinishedInit And IsNumeric(txtYears.Text) Then
    For ReturnIndex = 1 To pFirstReturns.Count
      pScenario.Weight.SetGagedYears pFirstReturns(ReturnIndex).Name, txtYears.Text
    Next
    PopulateResults
  End If
End Sub

Private Sub PopulateResults(Optional WhichResult As Long = -1)
  Dim ReturnIndex As Long
  Dim d() As Double
  Dim lLastCol As Long

  If pScenario.Weight.WeightType = 2 Then 'weighting by variance, fill 4th column
    lLastCol = 5
  Else 'fill 3rd column
    lLastCol = 3
  End If
  If pScenario.Weight.AreaGaged < 0.001 And pScenario.Weight.WeightType = 3 Then
    'Can't compute weighted extimate of this type with zero area gaged
  ElseIf pFinishedInit Then 'And txtYears.Value > 0.001 Then
    d = pScenario.WeightedDischarges
    For ReturnIndex = 1 To pFirstReturns.Count
      grdWgt.TextMatrix(ReturnIndex, lLastCol) = StrPad(SignificantDigits(d(ReturnIndex), 3), 9)
      If pScenario.Weight.WeightType = 2 Then
        grdWgt.TextMatrix(ReturnIndex, lLastCol + 1) = StrPad(SignificantDigits(pScenario.Weight.Variance(grdWgt.TextMatrix(ReturnIndex, 0)), 4), 9)
        grdWgt.TextMatrix(ReturnIndex, lLastCol + 2) = StrPad(SignificantDigits(pScenario.Weight.StandardError(grdWgt.TextMatrix(ReturnIndex, 0)), 4), 9)
      End If
    Next
  End If
  
  Exit Sub

ClearVals:
  For ReturnIndex = 1 To pFirstReturns.Count
    grdWgt.TextMatrix(ReturnIndex, lLastCol) = ""
  Next
End Sub

Private Sub txtYears_GotFocus()
  pCurrentControl = -1
  txtYears.SelStart = 0
  txtYears.SelLength = Len(txtYears.Text)
End Sub

Private Sub txtYears_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyDown
      If pScenario.Weight.WeightType = 1 Then
        grdWgt.row = 1
        grdWgt.col = 2
        grdWgt.SetFocus
      Else
        cmdApply(1).SetFocus
      End If
    Case vbKeyUp
      If optWeight(0).Visible And optWeight(0).Enabled And optWeight(0).Value Then
        optWeight(0).SetFocus
      End If
  End Select
End Sub
