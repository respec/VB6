VERSION 5.00
Begin VB.Form frmFreq 
   Caption         =   "Frequency"
   ClientHeight    =   1785
   ClientLeft      =   3195
   ClientTop       =   2085
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   25
   Icon            =   "frmFreq.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1785
   ScaleWidth      =   7410
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7215
      Begin VB.CommandButton cmdList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&List"
         Default         =   -1  'True
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdPlot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Plot"
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   5760
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.CheckBox chkEstimate 
      Caption         =   "Rural  1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   6975
   End
End
Attribute VB_Name = "frmFreq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private Sub PlotFreq(agr As ATCoGraph) ', lst As ATCoGrid)

  'calculate values for flood frequency plot
  Dim i&, pflg&, icrv&, ivar&, ipos&, iret&, cnt&, lNumIntervals As Long
  Dim qmin!, qmax!, tmin!, tmax!
  Dim plmn!(10), plmx!(10), vmin!(40), vmax!(40)
  Dim Ntics&(3), xlab$, ylab$, titl$, capt$
  Dim t() As Double
  Dim q() As Double
  Dim which&(40) 'which axis for each variable
  Dim tran&(40)  'transformation (1=none - arithmetic, 2=log)
  Dim vlab$(40)  'variable label
  Dim clab$(20)  'curve label
  Dim ctype&(20), ltype&(20), stype&(20), lthick&(20), lcolor&(20)
  Dim eqnMetric As Boolean
  Dim IntervalName As String
  
  Dim chkIndex As Long
  Dim Scenario As nssScenario
'  Dim graphing As Boolean
'  If agr Is Nothing Then graphing = False Else graphing = True
  
  
  qmin = 100000
  qmax = 0
  tmin = 500
  tmax = -500
  ivar = -1
  icrv = -1
  
'  If graphing Then
    agr.init
'  Else
'    lst.cols = 2
'    lst.rows = 2
'    lst.FixedRows = 1
'    lst.col = 0
'    lst.row = 1
'    lst.ColTitle(0) = "Years"
'  End If
  For chkIndex = 0 To chkEstimate.Count - 1
    If chkEstimate(chkIndex).Value = vbChecked Then
      If chkIndex < Project.RuralScenarios.Count Then
        Set Scenario = Project.RuralScenarios(chkIndex + 1)
      Else
        Set Scenario = Project.UrbanScenarios(chkIndex - Project.RuralScenarios.Count + 1)
      End If
'      If Not graphing Then
'        If lst.col > 0 Then lst.cols = lst.cols + 1
'        lst.col = lst.col + 1
'        lst.ColTitle(lst.col) = Scenario.Name
'      End If
      lNumIntervals = Scenario.UserRegions(1).region.DepVars.Count
      q = Scenario.WeightedDischarges
      ReDim t(LBound(q) To UBound(q))
      icrv = icrv + 1
      ivar = ivar + 1
      'set x-axis values (probabilities)
      For i = 1 To lNumIntervals
        IntervalName = ReplaceString(Scenario.UserRegions(1).region.DepVars(i).Name, "_", ".")
        If Left(IntervalName, 2) = "PK" Then
          IntervalName = Mid(IntervalName, 3)
        End If
        t(i) = gausex(1 / CDbl(IntervalName))
'        If Not graphing Then
'          lst.TextMatrix(i, 0) = Scenario.UserRegions(1).region.DepVars(i).Name 't(i)
'          lst.TextMatrix(i, lst.col) = q(i)
'        End If
      Next i
'      If graphing Then
        If t(1) < tmin Then tmin = t(1)
        If t(lNumIntervals) > tmax Then tmax = t(lNumIntervals)
  
        'put x values in plot buffer
        agr.SetData ivar, ipos, lNumIntervals, t(), iret
        vmin(ivar) = t(lNumIntervals)
        vmax(ivar) = t(1)
        which(ivar) = 4
        tran(ivar) = 1
        vlab(ivar) = "Years"
        'update buffer position
        ipos = ipos + lNumIntervals
  '      'set min/max X-axis range
  '      plmn(3) = t(1)
  '      plmx(3) = t(lNumIntervals)
        'put flow values in plot buffer
        If q(1) < qmin Then qmin = q(1)
        If q(lNumIntervals) > qmax Then qmax = q(lNumIntervals)
  
        ivar = ivar + 1
        agr.SetData ivar, ipos, lNumIntervals, q(), iret
        vmin(ivar) = q(1)
        vmax(ivar) = q(lNumIntervals)
        which(ivar) = 1
        tran(ivar) = 2
        vlab(ivar) = "Discharge"
        'update buffer position
        ipos = ipos + lNumIntervals
        clab(icrv) = Scenario.Name
        ctype(icrv) = 7
        ltype(icrv) = 1
        stype(icrv) = 0
        lthick(icrv) = 1
        lcolor(icrv) = (icrv + 9) Mod 15
        If lcolor(icrv) = 7 Or lcolor(icrv) = 15 Then lcolor(icrv) = 8 'White -> Gray
'      End If
'!!!  moved following line inside End If (4/29/03 - prh) (corrected subscpt rng error that didn't reproduce in production mode)
      agr.SetVars icrv, ivar, ivar - 1
    End If
    'If graphing Then
  Next
  xlab = "Recurrence Interval, in years"
  If Project.Metric = True Then
    ylab = "Peak Discharge, in cubic meters per second"
  Else
    ylab = "Peak Discharge, in cubic feet per second"
  End If
'  If graphing Then
    'set axes types and labels
    'set x-axis to probability, y-axis to log
    agr.SetAxesInfo 4, 2, 0, 0, xlab, ylab, "", ""
    titl = "Flood Frequency Plot"
    capt = "Frequency Plot"
    agr.SetTitles titl, capt
    'set min/max Y-axis range
    Call Scalit(2, qmin, qmax, plmn(0), plmx(0))
    Call Scalit(4, tmin, tmax, plmn(3), plmx(3))
    agr.SetNumVars icrv + 1, ivar + 1
    agr.SetScale plmn(), plmx(), Ntics()
    agr.SetCurveInfo ctype, ltype, lthick, stype, lcolor, clab
    agr.SetVarInfo vmin, vmax, which, tran, vlab
    agr.ShowIt
'  Else
'    lst.header = ylab
'    lst.Parent.Caption = "Flood Frequency"
'    lst.Parent.Show
'    lst.ColsSizeByContents
'    lst.Colsfit to width
'  End If
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdPlot_Click()
  Dim agr As ATCoGraph
  Set agr = New ATCoGraph
  PlotFreq agr ', Nothing
End Sub

Private Sub Form_Activate()
  Dim Scenario As nssScenario
  Dim ScenIndex As Long
  Dim chkIndex As Long
  
  chkIndex = 0
  
  For ScenIndex = 1 To Project.RuralScenarios.Count
    Set Scenario = Project.RuralScenarios(ScenIndex)
    
    If chkIndex > 0 Then
      If chkIndex < chkEstimate.Count Then Unload chkEstimate(chkIndex)
      Load chkEstimate(chkIndex)
      chkEstimate(chkIndex).Top = chkEstimate(chkIndex - 1).Top _
                                + chkEstimate(chkIndex - 1).Height _
                                + 108
      chkEstimate(chkIndex).Visible = True
    End If
    chkEstimate(chkIndex).Caption = Scenario.Name
    If Scenario.lowflow Then 'can't graph lowflow results
      chkEstimate(chkIndex).Value = vbUnchecked
      chkEstimate(chkIndex).Enabled = False
    Else
      chkEstimate(chkIndex).Value = vbChecked
      chkEstimate(chkIndex).Enabled = True
    End If
    'If ScenIndex = Project.CurrentRuralScenario Then
    chkIndex = chkIndex + 1
  Next

  For ScenIndex = 1 To Project.UrbanScenarios.Count
    Set Scenario = Project.UrbanScenarios(ScenIndex)
    
    If chkIndex > 0 Then
      If chkIndex < chkEstimate.Count Then Unload chkEstimate(chkIndex)
      Load chkEstimate(chkIndex)
      chkEstimate(chkIndex).Top = chkEstimate(chkIndex - 1).Top _
                                + chkEstimate(chkIndex - 1).Height _
                                + 108
      chkEstimate(chkIndex).Visible = True
    End If
    chkEstimate(chkIndex).Caption = Scenario.Name
    If Scenario.lowflow Then 'can't graph lowflow results
      chkEstimate(chkIndex).Value = vbUnchecked
      chkEstimate(chkIndex).Enabled = False
    End If
    'If ScenIndex = Project.CurrentRuralScenario Then
    chkIndex = chkIndex + 1
  Next

  'adjust command button positions
  fraButtons.Top = chkEstimate(chkEstimate.Count - 1).Top _
                 + chkEstimate(chkEstimate.Count - 1).Height _
                 + 150
'  cmdPlot.Top = chkEstimate(chkEstimate.Count - 1).Top
'              + chkEstimate(chkEstimate.Count - 1).Height + 150
'  cmdClose.Top = cmdPlot.Top
  frmFreq.Height = fraButtons.Top + fraButtons.Height + 500

End Sub

