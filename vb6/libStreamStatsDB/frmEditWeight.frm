VERSION 5.00
Object = "{872F11D5-3322-11D4-9D23-00A0C9768F70}#1.10#0"; "ATCoCtl.ocx"
Begin VB.Form frmEditWeight 
   Caption         =   "Edit Weight"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5670
   HelpContextID   =   27
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   5670
   StartUpPosition =   3  'Windows Default
   Begin ATCoCtl.ATCoGrid grdWgt 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3836
      SelectionToggle =   0   'False
      AllowBigSelection=   0   'False
      AllowEditHeader =   0   'False
      AllowLoad       =   0   'False
      AllowSorting    =   0   'False
      Rows            =   2
      Cols            =   4
      ColWidthMinimum =   300
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
      Height          =   372
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   3012
   End
   Begin VB.OptionButton optWeight 
      Caption         =   "Weight for gaged site using observed data"
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2412
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
      Left            =   1680
      TabIndex        =   6
      Top             =   4080
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
      Left            =   3000
      TabIndex        =   7
      Top             =   4080
      Width           =   972
   End
   Begin VB.TextBox txtYears 
      Alignment       =   1  'Right Justify
      Height          =   252
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   972
   End
   Begin VB.ComboBox cboWgtSelect 
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   720
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
      Height          =   492
      Left            =   120
      TabIndex        =   37
      Top             =   720
      Width           =   2892
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
      Height          =   252
      Index           =   8
      Left            =   1560
      TabIndex        =   36
      Top             =   3600
      Width           =   1296
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
      Height          =   252
      Index           =   7
      Left            =   1560
      TabIndex        =   35
      Top             =   3360
      Width           =   1296
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
      Height          =   252
      Index           =   6
      Left            =   1560
      TabIndex        =   34
      Top             =   3120
      Width           =   1296
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
      Height          =   252
      Index           =   5
      Left            =   1560
      TabIndex        =   33
      Top             =   2880
      Width           =   1296
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
      Height          =   252
      Index           =   4
      Left            =   1560
      TabIndex        =   32
      Top             =   2640
      Width           =   1296
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
      Height          =   252
      Index           =   3
      Left            =   1560
      TabIndex        =   31
      Top             =   2400
      Width           =   1296
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
      Height          =   252
      Index           =   2
      Left            =   1560
      TabIndex        =   30
      Top             =   2160
      Width           =   1296
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
      Height          =   252
      Index           =   8
      Left            =   240
      TabIndex        =   29
      Top             =   3600
      Width           =   612
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
      Height          =   252
      Index           =   7
      Left            =   240
      TabIndex        =   28
      Top             =   3360
      Width           =   612
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
      Height          =   252
      Index           =   6
      Left            =   240
      TabIndex        =   27
      Top             =   3120
      Width           =   612
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
      Height          =   252
      Index           =   5
      Left            =   240
      TabIndex        =   26
      Top             =   2880
      Width           =   612
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
      Height          =   252
      Index           =   4
      Left            =   240
      TabIndex        =   25
      Top             =   2640
      Width           =   612
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
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   24
      Top             =   2400
      Width           =   612
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
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   23
      Top             =   2160
      Width           =   612
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
      Height          =   252
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   3132
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
      Height          =   252
      Left            =   120
      TabIndex        =   21
      Top             =   1080
      Width           =   4212
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
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   20
      Top             =   1920
      Width           =   612
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
      Height          =   252
      Index           =   1
      Left            =   1560
      TabIndex        =   19
      Top             =   1920
      Width           =   1296
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
      Left            =   4320
      TabIndex        =   18
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
      Index           =   2
      Left            =   4320
      TabIndex        =   17
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
      Index           =   3
      Left            =   4320
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
      Index           =   4
      Left            =   4320
      TabIndex        =   15
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
      Index           =   5
      Left            =   4320
      TabIndex        =   14
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
      Index           =   6
      Left            =   4320
      TabIndex        =   13
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
      Index           =   7
      Left            =   4320
      TabIndex        =   12
      Top             =   3360
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
      Index           =   8
      Left            =   4320
      TabIndex        =   11
      Top             =   3600
      Width           =   1092
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
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1332
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
      Height          =   372
      Index           =   1
      Left            =   1560
      TabIndex        =   9
      Top             =   1440
      Width           =   1092
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
      Height          =   492
      Index           =   2
      Left            =   3048
      TabIndex        =   8
      Top             =   1440
      Width           =   1092
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
      Height          =   372
      Index           =   3
      Left            =   4320
      TabIndex        =   5
      Top             =   1440
      Width           =   1092
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
  Select Case pScenario.Weight.WeightType
    Case 0, 1
      pScenario.Weight.WeightType = 1
    Case 2
      pScenario.Weight.WeightType = 2
  End Select
  Me.Caption = "Editing weight for """ & pScenario.Name & """"
  PopulateIntervals
  pFinishedInit = True
  For ReturnIndex = 1 To pFirstReturns.Count
    If IsNumeric(txtYears.Text) Then
      pScenario.Weight.SetGagedYears pFirstReturns(ReturnIndex).Name, txtYears.Text
    End If
'    If IsNumeric(txtObs(ReturnIndex).Text) Then
'      pScenario.Weight.SetGagedValue pFirstReturns(ReturnIndex).Name, txtObs(ReturnIndex).Text
    If IsNumeric(grdWgt.TextMatrix(ReturnIndex, 2)) Then
      pScenario.Weight.SetGagedValue pFirstReturns(ReturnIndex).Name, grdWgt.TextMatrix(ReturnIndex, 2)
    End If
  Next
  PopulateResults
End Property

Private Sub PopulateIntervals()
  Dim ReturnIndex As Long
  Dim Interval As String
  Dim d() As Double
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
    grdWgt.ColEditable(2) = True
    For ReturnIndex = 1 To pFirstReturns.Count
      Interval = pFirstReturns(ReturnIndex).Name
'      lblInterval(ReturnIndex).Caption = Interval
'      lblEstimated(ReturnIndex).Caption = StrPad(SignificantDigits(d(ReturnIndex), 3), 9)
'      txtObs(ReturnIndex).Visible = True
'      txtObs(ReturnIndex).Text = CDbl(pScenario.Weight.GetGagedValue(Interval))
      grdWgt.TextMatrix(ReturnIndex, 0) = Interval
      grdWgt.col = 0
      grdWgt.row = ReturnIndex
      grdWgt.CellBackColor = &HE0E0E0
      grdWgt.TextMatrix(ReturnIndex, 1) = StrPad(SignificantDigits(d(ReturnIndex), 3), 9)
      grdWgt.col = 1
      grdWgt.CellBackColor = &HE0E0E0
      grdWgt.TextMatrix(ReturnIndex, 2) = pScenario.Weight.GetGagedValue(Interval)
      txtYears.Text = pScenario.Weight.GetGagedYears(Interval)
      grdWgt.col = 3
      grdWgt.CellBackColor = &HE0E0E0
    Next
    grdWgt.Height = 242 * (pFirstReturns.Count + 2)
    cmdApply(1).Top = grdWgt.Top + grdWgt.Height + 120
    cmdCancel.Top = cmdApply(1).Top
    frmEditWeight.Height = cmdCancel.Top + cmdCancel.Height + 440
'    While ReturnIndex <= txtObs.Count
'      txtObs(ReturnIndex).Visible = False
'      ReturnIndex = ReturnIndex + 1
'    Wend
  End If
End Sub

'Private Sub chkArea_Click()
'  If chkArea.Value = vbChecked Then
'    txtGagedArea.Enabled = True
'    pScenario.Weight.WeightType = 2
'  Else
'    txtGagedArea.Enabled = False
'    pScenario.Weight.WeightType = 1
'  End If
'  PopulateResults
'End Sub

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
'    txtObs(i).Text = StrPad(SignificantDigits(d(i), 3), 9)
    grdWgt.TextMatrix(i, 2) = StrPad(SignificantDigits(d(i), 3), 9)
'    If IsNumeric(txtObs(i).Text) Then
    If IsNumeric(grdWgt.TextMatrix(i, 2)) Then
'      pScenario.Weight.SetGagedValue pScenario.UserRegions(1).Region.DepVars(i).Name, txtObs(i).Text
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
'    ElseIf Not (txtObs(pCurrentControl + 1).Visible And txtObs(pCurrentControl + 1).Visible) Then
'      cmdApply(1).SetFocus
    Else
      grdWgt.SetFocus
      grdWgt.row = pCurrentControl + 1
      'txtObs(pCurrentControl + 1).SetFocus
    End If
    KeyCode = 0 'Don't let individual controls see this key if we used it
  End If
End Sub

Private Sub Form_Load()

  grdWgt.TextMatrix(-1, 0) = "Recurrence"
  grdWgt.TextMatrix(0, 0) = "Interval (years)"
  grdWgt.ColType(0) = ATCoSng
  grdWgt.TextMatrix(-1, 1) = "Estimated"
  grdWgt.TextMatrix(0, 1) = "Flow"
  grdWgt.ColType(1) = ATCoSng
  grdWgt.ColType(2) = ATCoSng
  grdWgt.ColAlignment(3) = 7
  grdWgt.TextMatrix(-1, 3) = "Weighted"
  grdWgt.TextMatrix(0, 3) = "Flow"

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
  If pFinishedInit Then
    For ReturnIndex = 1 To pFirstReturns.Count
      If IsNumeric(grdWgt.TextMatrix(ReturnIndex, 2)) Then
        pScenario.Weight.SetGagedValue pFirstReturns(ReturnIndex).Name, grdWgt.TextMatrix(ReturnIndex, 2)
      End If
    Next
    PopulateResults
  End If

End Sub

Private Sub optWeight_Click(Index As Integer)
  Dim i%

  If Index = 0 Then
    lblWgtSelect.Visible = False
    cboWgtSelect.Visible = False
    lblYears.Visible = True
    txtYears.Visible = True
    Label1.Visible = True
'    lblCol(2).Caption = "Observed Flow"
    grdWgt.TextMatrix(-1, 2) = "Observed"
    grdWgt.TextMatrix(0, 2) = "Flow"
'    For i = 1 To pScenario.UserRegions(1).Region.DepVars.Count
'      txtObs(i).Enabled = True
'    Next i
    grdWgt.ColEditable(2) = True
    pScenario.Weight.WeightType = 1
  Else
    lblWgtSelect.Visible = True
    cboWgtSelect.Visible = True
    lblYears.Visible = False
    txtYears.Visible = False
    Label1.Visible = False
'    lblCol(2).Caption = "Weighted Gaged Flow"
    grdWgt.TextMatrix(-1, 2) = "Weighted"
    grdWgt.TextMatrix(0, 2) = "Gaged Flow"
'    For i = 1 To pScenario.UserRegions(1).Region.DepVars.Count
'      txtObs(i).Enabled = False
'    Next i
    grdWgt.ColEditable(2) = False
    pScenario.Weight.WeightType = 2
    cboWgtSelect.Clear
    If pScenario.Project.RuralScenarios.Count > 0 Then
      'look for other scenarios that have gaged weighted results
      For i = 1 To pScenario.Project.RuralScenarios.Count
        If pScenario.Project.RuralScenarios(i).Weight.WeightType = 1 Then
          'add to combo box list
          cboWgtSelect.AddItem pScenario.Project.RuralScenarios(i).Name
        End If
      Next i
    End If
    If cboWgtSelect.ListCount = 0 Then 'no gaged weighted scenarios available
      ssMessageBox "There are no scenarios with gaged weighted values that use the same region as the scenario being weighted." & vbCrLf & _
             "Thus, this type of weighting may not be performed currently.", vbExclamation, "Weighting Problem"
      optWeight(0).Value = True
    Else
      cboWgtSelect.ListIndex = 0
    End If
  End If
  grdWgt.col = 2
  For i = 1 To grdWgt.rows
    grdWgt.row = i
    If Index = 0 Then
      grdWgt.CellBackColor = grdWgt.BackColor
    Else
      grdWgt.CellBackColor = &HE0E0E0
    End If
  Next i
End Sub

Private Sub optWeight_GotFocus(Index As Integer)
  pCurrentControl = 0
End Sub

'Private Sub txtGagedArea_Change()
'  pScenario.Weight.AreaGaged = txtGagedArea.Value
'  PopulateResults
'End Sub

'Private Sub txtObs_Change(Index As Integer)
'  Dim ReturnIndex As Long
'  If pFinishedInit Then
'    For ReturnIndex = 1 To pFirstReturns.Count
'      If IsNumeric(txtObs(ReturnIndex).Text) Then
'        pScenario.Weight.SetGagedValue pFirstReturns(ReturnIndex).Name, txtObs(ReturnIndex).Text
'      End If
'    Next
'    PopulateResults
'  End If
'End Sub
'
'Private Sub txtObs_GotFocus(Index As Integer)
'  pCurrentControl = Index
'  txtObs(Index).SelStart = 0
'  txtObs(Index).SelLength = Len(txtObs(Index).Text)
'End Sub
'
'Private Sub txtObs_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'  Dim newIndex As Integer
'  Select Case KeyCode
'    Case vbKeyReturn, vbKeyDown, vbKeyUp
'      If KeyCode = vbKeyUp Then newIndex = Index - 1 Else newIndex = Index + 1
'      If newIndex < 1 Then
'        If cboWgtSelect.Visible And cboWgtSelect.Enabled Then
'          cboWgtSelect.SetFocus
'        ElseIf txtYears.Visible And txtYears.Enabled Then
'          txtYears.SetFocus
'        End If
'      ElseIf newIndex <= txtObs.Count Then
'        If txtObs(newIndex).Visible Then
'          txtObs(newIndex).SetFocus
'        Else
'          cmdApply(1).SetFocus
'        End If
'      Else
'        cmdApply(1).SetFocus
'      End If
'  End Select
'End Sub

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
  If pScenario.Weight.AreaGaged < 0.001 And pScenario.Weight.WeightType = 2 Then
    'Can't compute weighted extimate of this type with zero area gaged
  ElseIf pFinishedInit Then 'And txtYears.Value > 0.001 Then
    d = pScenario.WeightedDischarges
    For ReturnIndex = 1 To pFirstReturns.Count
'      lblRes(ReturnIndex).Caption = StrPad(SignificantDigits(d(ReturnIndex), 3), 9)
      grdWgt.TextMatrix(ReturnIndex, 3) = StrPad(SignificantDigits(d(ReturnIndex), 3), 9)
    Next
  End If
  
  Exit Sub

ClearVals:
  For ReturnIndex = 1 To pFirstReturns.Count
'    lblRes(ReturnIndex).Caption = ""
    grdWgt.TextMatrix(ReturnIndex, 3) = ""
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
'      If txtObs(1).Visible And txtObs(1).Enabled Then
'        txtObs(1).SetFocus
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
