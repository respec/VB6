VERSION 5.00
Object = "*\A..\ATCoCtl\ATCoCtl.vbp"
Begin VB.Form frmHyd 
   Caption         =   "Hydrograph"
   ClientHeight    =   3780
   ClientLeft      =   3165
   ClientTop       =   2070
   ClientWidth     =   8265
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
   HelpContextID   =   26
   Icon            =   "frmHyd.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3780
   ScaleWidth      =   8265
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   8055
      Begin VB.CommandButton cmdHyd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&Plot"
         Default         =   -1  'True
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdHyd 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "&List"
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   10
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   6720
         TabIndex        =   9
         Top             =   0
         Width           =   1215
      End
   End
   Begin ATCoCtl.ATCoText txtLagtime 
      Height          =   252
      Index           =   0
      Left            =   7320
      TabIndex        =   7
      Top             =   1080
      Width           =   732
      _ExtentX        =   1296
      _ExtentY        =   450
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
   Begin ATCoCtl.ATCoText txtBasinLength 
      Height          =   252
      Left            =   7320
      TabIndex        =   2
      ToolTipText     =   "Units are miles even when metric is used for calculations"
      Top             =   240
      Width           =   732
      _ExtentX        =   1296
      _ExtentY        =   450
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
   Begin VB.CheckBox chkBasLen 
      Caption         =   "&Estimate Lag Time using Basin Length  (National Urban only)"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5532
   End
   Begin VB.ComboBox cboInterval 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.CheckBox chkEstimate 
      Caption         =   "Rural  1"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Value           =   1  'Checked
      Width           =   5895
   End
   Begin VB.Label lblBasinLength 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Basin Length (mi):"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   5640
      TabIndex        =   1
      ToolTipText     =   "Units are miles even when metric is used for calculations"
      Top             =   264
      Width           =   1572
   End
   Begin VB.Label lblInterval 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Recurrence Interval:"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   640
      Width           =   1812
   End
   Begin VB.Label lblLagtime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Lag Time (hrs)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6840
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmHyd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private basin_length!

Private Sub EstimateLag()
  Dim j&, lag_time!
  Dim userRegion1 As userRegion
  
  basin_length = txtBasinLength.Value
  
  If chkBasLen.Value = 1 And basin_length > 0 Then
    For j = Project.RuralScenarios.Count To chkEstimate.Count - 1
      'If chkEstimate(j).Value = vbChecked Then
        Set userRegion1 = Project.UrbanScenarios(LCase(chkEstimate(j).Caption)).UserRegions(1)
        If userRegion1.region.State.code = "00" Then 'National
          lag_time = 0.003 * (basin_length ^ 0.71)
          lag_time = lag_time * ((13# - userRegion1.UserParms(5).getValue(False)) ^ 0.34)
          lag_time = lag_time * ((userRegion1.UserParms(4).getValue(False) + 10#) ^ 2.53)
          lag_time = lag_time * (userRegion1.UserParms(3).getValue(False) ^ -0.44)
          lag_time = lag_time * (userRegion1.UserParms(1).getValue(False) ^ -0.2)
          lag_time = lag_time * (userRegion1.UserParms(2).getValue(False) ^ -0.14)
          txtLagtime(j).Value = lag_time
          txtLagtime(j).Enabled = False
          chkEstimate(j).Value = vbChecked
        End If
      'End If
    Next j
  End If
End Sub
'          lag_time = 0.003 * (basin_length ^ 0.71)
'          lag_time = lag_time * ((13# - userRegion1.UserParms(1).getValue(False)) ^ 0.34)
'          lag_time = lag_time * ((userRegion1.UserParms(5).getValue(False) + 10#) ^ 2.53)
'          lag_time = lag_time * (userRegion1.UserParms(4).getValue(False) ^ -0.44)
'          lag_time = lag_time * (userRegion1.UserParms(2).getValue(False) ^ -0.2)
'          lag_time = lag_time * (userRegion1.UserParms(3).getValue(False) ^ -0.14)

Private Sub CalcHyd(useGraph As Boolean)
  'calculate values for hydrograph plot
  Dim i&, j&, k&, icrv&, ipos&, iret&, cnt&, ivar&, tvar&
  Dim peak!, time_ratio!, qmax!, lt!, llt!, tmin!, tmax!, lint!
  Dim plmn!(3), plmx!(3), Ntics&(3)
  Dim vmin!(40), vmax!(40), vlab$(40), tran&(40), which&(40)
  Dim clab$(20), ctype&(20), ltype&(20), stype&(20), lthick&(20), lcolor&(20)
  Dim xlab$, ylab$, titl$, capt$
  Static t(1 To HYDRO_SIZE) As Double
  Static q(1 To HYDRO_SIZE) As Double
  Dim eqnMetric As Boolean
  Dim agr As ATCoGraph
  Dim lst As frmList
  Dim Scenario As nssScenario
  Dim curDischarges() As Double
  Dim IntervalIndex As Long
  Dim StepName As String
  Dim ScenarioIndex As Long

  On Error GoTo ErrHand
  
  If Not IsNumeric(cboInterval.Text) Then
    MsgBox "Must select a return interval before plotting", vbOKOnly, "NSS Hydrograph"
    Exit Sub
  End If
  
  xlab = "Time (hours)"
  If Project.Metric Then
    ylab = "Discharge (cubic meters per second)"
  Else
    ylab = "Discharge (cubic feet per second)"
  End If
  
  StepName = "CSng(cboInterval.Text)"
  lint = CSng(cboInterval.Text)
  llt = 0
  ivar = -1
  ipos = 1
  tmin = 1000000#
  tmax = 0
  qmax = 0
  icrv = -1

  If useGraph Then
    StepName = "Set agr = New ATCoGraph"
    Set agr = New ATCoGraph
    StepName = "agr.init"
    agr.init
  Else
    Set lst = New frmList
    lst.agd.cols = 1
  End If
  
  For j = 0 To chkEstimate.Count - 1
    'assume we're not plotting this one
    peak = 0
    StepName = "If chkEstimate(" & j & ").Value = vbChecked"
    If chkEstimate(j).Value = vbChecked Then
      If j < Project.RuralScenarios.Count Then       'rural hydrograph
        StepName = "Scenario = Project.RuralScenarios(" & j + 1 & ")"
        Set Scenario = Project.RuralScenarios(j + 1)
      Else                                           'urban hydrograph
        ScenarioIndex = j - Project.RuralScenarios.Count + 1
        StepName = "Set Scenario = Project.UrbanScenarios(" & ScenarioIndex & ")"
        Set Scenario = Project.UrbanScenarios(ScenarioIndex)
      End If
      lt = txtLagtime(j).Value
      curDischarges = Scenario.WeightedDischarges
      For IntervalIndex = 1 To Scenario.UserRegions(1).region.DepVars.Count
        StepName = "CSng(" & Scenario.UserRegions(1).region.DepVars(IntervalIndex).Name & ")"
        If CSng(Scenario.UserRegions(1).region.DepVars(IntervalIndex).Name) = lint Then
          'interval available for this Estimate
          peak = curDischarges(IntervalIndex)
          Exit For
        End If
      Next
      If peak = 0 Then
        MsgBox "Unable to generate hydrograph for " & Scenario.Name & vbCr _
             & "The specified interval is not available for this estimate.", 64, "Hydrograph"
      Else
        'calculate hydrograph
        icrv = icrv + 1
        If lt <> llt Then
          'need to calc time values
          StepName = "calculating time values"
          ivar = ivar + 1
          time_ratio = 0.25
          For i = 1 To HYDRO_SIZE
            t(i) = time_ratio * lt
            time_ratio = time_ratio + 0.05
          Next i
          If useGraph Then
            'put time values in plot buffer
            StepName = "putting time values in plot buffer"
            agr.SetData ivar, ipos, HYDRO_SIZE, t(), iret
            which(ivar) = 4
            tran(ivar) = 1
            vmin(ivar) = t(1)
            vmax(ivar) = t(HYDRO_SIZE)
            vlab(ivar) = "Hours"
            If vmin(ivar) < tmin Then tmin = vmin(ivar) 'new minimum time value
            If vmax(ivar) > tmax Then tmax = vmax(ivar) 'new maximum time value
          Else
            
  '              'put time values in list buffer
  '              Call frmList.lpdatr(ivar, ipos, HYDRO_SIZE, t(), iret)
  '              'save time label for listing
  '              Call frmList.lplbvr(ivar, "Time (hrs)")
          End If
          'update buffer position
          ipos = ipos + HYDRO_SIZE
          'save time variable position
          tvar = ivar
        End If
        'save this lag time value
        llt = lt
        'update variable count
        ivar = ivar + 1
        which(ivar) = 1
        tran(ivar) = 1
        vmin(ivar) = 1000000
        vmax(ivar) = -1000000
        vlab(ivar) = "Discharge"
        StepName = "For i = 1 To HYDRO_SIZE"
        For i = 1 To HYDRO_SIZE
          q(i) = disch_ratio(i - 1) * peak
          If q(i) > qmax Then qmax = q(i)
          If q(i) > vmax(ivar) Then vmax(ivar) = q(i)
          If q(i) < vmin(ivar) Then vmin(ivar) = q(i)
        Next i
        If useGraph Then           'put flow values in plot buffer
          StepName = "agr.SetData " & ivar & ", " & ipos & ", " & HYDRO_SIZE & ", " & iret
          agr.SetData ivar, ipos, HYDRO_SIZE, q(), iret
          clab(icrv) = Scenario.Name
          ctype(icrv) = 7
          ltype(icrv) = 1
          stype(icrv) = 0
          lthick(icrv) = 1
          lcolor(icrv) = (icrv + 9) Mod 15
          If lcolor(icrv) = 7 Or lcolor(icrv) = 15 Then lcolor(icrv) = 8 'White -> Gray
  
          StepName = "agr.SetVars " & icrv & ", " & ivar & ", " & tvar
          agr.SetVars icrv, ivar, tvar 'set variables for this curve
        Else                        'put flow values in list buffer
          With lst.agd
            .cols = .cols + 1
            If .cols > 2 Then .cols = .cols + 1
            .col = .cols - 1
            .ColTitle(.col - 1) = xlab
            .ColTitle(.col) = Scenario.Name
            For i = 1 To HYDRO_SIZE
              .TextMatrix(i, .col - 1) = NumFmted(t(i), 5, 2)
              .TextMatrix(i, .col) = NumFmted(Signif(q(i), Scenario.Project.Metric), 11, 0)
            Next
          End With
  '            Call frmList.lpdatr(ivar, ipos, HYDRO_SIZE, q(), iret)
  '            If k = 1 Then             'set legend labels
  '              Call frmList.lplbvr(ivar, "R" & j + 1)
  '            Else
  '              Call frmList.lplbvr(ivar, "U" & j + 1)
  '            End If
  '            Call frmList.lpwcxy(icrv, ivar, tvar) 'set variables for this listing
        End If
        'update buffer position
        ipos = ipos + HYDRO_SIZE
      End If
    End If
  Next j
  If useGraph Then 'plotting
    StepName = "agr.SetNumVars " & icrv + 1 & ", " & ivar + 1
    agr.SetNumVars icrv + 1, ivar + 1 'set number of curves and variables
    StepName = "Scalit 1"
    Call Scalit(1, tmin, tmax, plmn(3), plmx(3)) 'set min/max X-axis range
    StepName = "Scalit 2"
    Call Scalit(1, 0, qmax, plmn(0), plmx(0))    'set min/max Y-axis range
    Ntics(0) = 5  'Y
    Ntics(3) = 10 'X
    StepName = "agr.SetScale"
    agr.SetScale plmn(), plmx(), Ntics()
    StepName = "agr.SetCurveInfo"
    agr.SetCurveInfo ctype, ltype, lthick, stype, lcolor, clab
    StepName = "agr.SetVarInfo"
    agr.SetVarInfo vmin, vmax, which, tran, vlab
    'set axes types and labels
    titl = "Hydrograph for " & cboInterval.Text & "-yr interval&" & _
           "NOTE:  May not represent actual hydrograph"
    capt = "Hydrograph Plot"
    'set x-axis to arithmetic as time axis not yet implemented
    StepName = "agr.SetTitles"
    agr.SetTitles titl, capt
    StepName = "agr.SetAxesInfo"
    agr.SetAxesInfo 1, 1, 0, 0, xlab, ylab, "", ""
    StepName = "agr.ShowIt"
    agr.ShowIt
  Else
    lst.agd.ColsSizeByContents
    lst.agd.ColsSizeToWidth
    lst.agd.header = ylab
    lst.Caption = "Hydrograph listing for " & cboInterval.Text & "-year interval"
    lst.Show
'      'listing output, set title
'      titl = "Hydrograph listing for " & cboInterval.Text & "-year interval"
'      Call frmList.lptitl(titl)
'      'set number of variables to list
'      Call frmList.lpnlis(icrv)
  End If

  Exit Sub

ErrHand:
  MsgBox "Error Calculating Hydrograph" & vbCr _
        & Err.Description & vbCr & StepName, vbCritical, "NSS"
End Sub

Private Sub chkBasLen_Click()
  Dim RuralCount As Long
  Dim UrbanCount As Long
  Dim j As Long
  If chkBasLen.Value = 1 Then
    lblBasinLength.Enabled = True
    txtBasinLength.Enabled = True
    EstimateLag
  Else
    RuralCount = Project.RuralScenarios.Count
    UrbanCount = Project.UrbanScenarios.Count
    For j = RuralCount - 1 To RuralCount + UrbanCount - 1
      txtLagtime(j).Enabled = True
    Next j
    lblBasinLength.Enabled = False
    txtBasinLength.Enabled = False
  End If
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdHyd_Click(Index As Integer)
  Dim chkIndex As Long
  Dim ZeroLag As Boolean
  Dim StepName As String
  
  On Error GoTo ErrHand
  
  ZeroLag = False
  chkIndex = 0
  StepName = "Checking for zero lag times in selected curves"
  While chkIndex < chkEstimate.Count And Not ZeroLag
    If chkEstimate(chkIndex).Value = vbChecked Then
      If txtLagtime(chkIndex).Value <= 0 Then ZeroLag = True
    End If
    chkIndex = chkIndex + 1
  Wend
  If ZeroLag Then
    MsgBox "A valid Lag Time value (> 0) must be entered" & vbCr _
         & "for each estimate to generate the hydrograph.", 48, "NSS Hydrograph"
  ElseIf Index = 0 Then
    StepName = "CalcHyd(True)"
    Call CalcHyd(True)  'generate hydrograph values, set plot info
  ElseIf Index = 1 Then
    StepName = "CalcHyd(False)"
    Call CalcHyd(False) 'generate hydrograph values, set list info
  End If
  
  Exit Sub
  
ErrHand:
  MsgBox "Error Generating Hydrograph" & vbCr _
        & Err.Description & vbCr & StepName, vbCritical, "NSS"
End Sub

Private Sub Form_Load()

  Dim i& ', j&, ip&, tnumint&
  Dim chkIndex As Long
  Dim ScenIndex As Long
  Dim Scenario As nssScenario
  Dim HaveNationalUrban As Boolean
  Dim allint() As Single, nAllInt As Long

  'init basin length and lag time
  If basin_length > 0 Then txtBasinLength.Value = basin_length

  HaveNationalUrban = False
  For ScenIndex = 1 To Project.UrbanScenarios.Count
    'Used to also do test And urban_discharge(0, i) > 0
    If Project.UrbanScenarios(ScenIndex).UserRegions(1).region.State.code = 0 Then
      HaveNationalUrban = True
    End If
  Next
  If HaveNationalUrban Then
    'national urban estimate exist, make basin
    'length available for lag time estimation
    chkBasLen.Enabled = True
    'chkBasLen.value = 1
  Else
    'no national urbans in use, disable basin length stuff
    chkBasLen.Value = 0
    chkBasLen.Enabled = False
    lblBasinLength.Enabled = False
    txtBasinLength.Enabled = False
  End If

  'get all available intervals for estimates
  Call AllIntervals(nAllInt, allint())
  'put available intervals in combo box
  cboInterval.Clear
  For i = 0 To nAllInt - 1
    cboInterval.AddItem CStr(allint(i))
  Next i
  cboInterval.ListIndex = 0
  'load rural scenarios as available plots
  For ScenIndex = 1 To Project.RuralScenarios.Count
    Set Scenario = Project.RuralScenarios(ScenIndex)
    
    If chkIndex > 0 Then
      If chkIndex < chkEstimate.Count Then
        Unload chkEstimate(chkIndex)
        Unload txtLagtime(chkIndex)
      End If
      Load chkEstimate(chkIndex)
      Load txtLagtime(chkIndex)
    End If
    chkEstimate(chkIndex).Caption = Scenario.Name
    If Scenario.lowflow Then 'can't graph lowflow results
      chkEstimate(chkIndex).Value = vbUnchecked
      chkEstimate(chkIndex).Enabled = False
      txtLagtime(chkIndex).Enabled = False
    Else
      chkEstimate(chkIndex).Value = vbChecked
      chkEstimate(chkIndex).Enabled = True
      txtLagtime(chkIndex).Enabled = True
    End If
    'If ScenIndex = Project.CurrentRuralScenario Then
    chkIndex = chkIndex + 1
  Next
  
  'load urban scenarios as available plots
  For ScenIndex = 1 To Project.UrbanScenarios.Count
    Set Scenario = Project.UrbanScenarios(ScenIndex)
    
    If chkIndex > 0 Then
      If chkIndex < chkEstimate.Count Then
        Unload chkEstimate(chkIndex)
        Unload txtLagtime(chkIndex)
      End If
      Load chkEstimate(chkIndex)
      Load txtLagtime(chkIndex)
    End If
    chkEstimate(chkIndex).Caption = Scenario.Name
    If Scenario.lowflow Then 'can't graph lowflow results
      chkEstimate(chkIndex).Value = vbUnchecked
      chkEstimate(chkIndex).Enabled = False
      txtLagtime(chkIndex).Enabled = False
    Else
      chkEstimate(chkIndex).Value = vbChecked
      chkEstimate(chkIndex).Enabled = True
      txtLagtime(chkIndex).Enabled = True
    End If
    chkIndex = chkIndex + 1
  Next

  'Position checkboxes and lag time textboxes
  For chkIndex = 1 To chkEstimate.Count - 1
    chkEstimate(chkIndex).Top = chkEstimate(chkIndex - 1).Top _
                              + chkEstimate(chkIndex).Height + 30
    chkEstimate(chkIndex).Visible = True
    chkEstimate(chkIndex).Value = 0
    txtLagtime(chkIndex).Top = chkEstimate(chkIndex).Top _
                             + txtLagtime(0).Top _
                             - chkEstimate(0).Top
    txtLagtime(chkIndex).Visible = True
  Next

'      'assume no values for this estimate, disable box
'      chkEstimate(i).Value = 0
'      chkEstimate(i).Enabled = False
'      txtLagtime(i).Enabled = False
'      j = 0
'      While j < rurscn(i).numint
'        If rural_discharge(0, j, i) > 0 Then
'          'values exist for this estimate, enable box
'          chkEstimate(i).Value = 1
'          chkEstimate(i).Enabled = True
'          txtLagtime(i).Enabled = True
'          j = rurscn(i).numint
'        Else
'          j = j + 1
'        End If
'      Wend
'    Next i
  
  'adjust command button positions
  fraButtons.Top = chkEstimate(chkEstimate.Count - 1).Top _
                 + chkEstimate(chkEstimate.Count - 1).Height _
                 + 150
'  cmdHyd(0).Top = chkEstimate(chkEstimate.Count - 1).Top
'                + chkEstimate(chkEstimate.Count - 1).Height + 150
'  cmdHyd(1).Top = cmdHyd(0).Top
'  cmdClose.Top = cmdHyd(0).Top
  frmHyd.Height = fraButtons.Top + fraButtons.Height + 500
End Sub

Private Sub txtBasinLength_Change()
  If txtBasinLength.Value > 0 Then
    chkBasLen.Value = vbChecked
    EstimateLag
  End If
End Sub

