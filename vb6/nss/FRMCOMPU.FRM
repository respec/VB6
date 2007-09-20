VERSION 5.00
Object = "{872F11D5-3322-11D4-9D23-00A0C9768F70}#1.6#0"; "ATCOCTL.OCX"
Begin VB.Form frmCompute 
   Caption         =   "Compute"
   ClientHeight    =   3480
   ClientLeft      =   1800
   ClientTop       =   3000
   ClientWidth     =   10128
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
   Icon            =   "FRMCOMPU.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3480
   ScaleWidth      =   10128
   Begin VB.Frame fraTotalArea 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   360
      Left            =   2880
      TabIndex        =   14
      Top             =   0
      Width           =   3972
      Begin ATCoCtl.ATCoText txtBasinArea 
         Height          =   300
         Left            =   2520
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   735
         _ExtentX        =   1291
         _ExtentY        =   529
         InsideLimitsBackground=   16777215
         OutsideHardLimitBackground=   255
         OutsideSoftLimitBackground=   65535
         HardMax         =   -999
         HardMin         =   0
         SoftMax         =   -999
         SoftMin         =   -999
         MaxWidth        =   5
         Alignment       =   1
         DataType        =   2
         DefaultValue    =   -999
         Value           =   "-999"
         Enabled         =   0   'False
      End
      Begin VB.Label lblBasinArea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Basin Drainage Area:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   50
         Width           =   2415
      End
      Begin VB.Label lblUnits 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "mi"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   3360
         TabIndex        =   5
         Top             =   50
         Width           =   252
      End
      Begin VB.Label lblUnitExponent 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   3620
         TabIndex        =   6
         Top             =   0
         Width           =   372
      End
   End
   Begin VB.Frame fraCribu 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   350
      Left            =   2880
      TabIndex        =   8
      Top             =   2280
      Width           =   5172
      Begin VB.CommandButton cmdMap 
         Caption         =   "&Map..."
         Height          =   252
         Left            =   4080
         TabIndex        =   11
         Top             =   0
         Width           =   972
      End
      Begin VB.ComboBox comboRegion 
         Height          =   288
         ItemData        =   "FRMCOMPU.frx":030A
         Left            =   3000
         List            =   "FRMCOMPU.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Crippen and Bue region for optional maximum flood-envelope computation"
         Top             =   0
         Width           =   975
      End
      Begin VB.Label lblRegion 
         BackStyle       =   0  'Transparent
         Caption         =   "&Crippen && Bue (1997) flood region:"
         Height          =   252
         Left            =   0
         TabIndex        =   9
         Top             =   50
         Width           =   3252
      End
   End
   Begin VB.ListBox lstRegion 
      Height          =   1008
      ItemData        =   "FRMCOMPU.frx":030E
      Left            =   120
      List            =   "FRMCOMPU.frx":0310
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      Top             =   2880
      Width           =   1095
   End
   Begin ATCoCtl.ATCoGrid agd 
      Height          =   1692
      Left            =   2880
      TabIndex        =   7
      Top             =   360
      Width           =   7092
      _ExtentX        =   12510
      _ExtentY        =   2985
      SelectionToggle =   0   'False
      AllowBigSelection=   0   'False
      AllowEditHeader =   0   'False
      AllowLoad       =   0   'False
      AllowSorting    =   -1  'True
      Rows            =   1
      Cols            =   2
      ColWidthMinimum =   300
      gridFontBold    =   0   'False
      gridFontItalic  =   0   'False
      gridFontName    =   "MS Sans Serif"
      gridFontSize    =   8
      gridFontUnderline=   0   'False
      gridFontWeight  =   400
      gridFontWidth   =   0
      Header          =   ""
      FixedRows       =   1
      FixedCols       =   1
      ScrollBars      =   3
      SelectionMode   =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorBkg    =   -2147483632
      BackColorSel    =   -2147483634
      ForeColorSel    =   16777215
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      InsideLimitsBackground=   16777215
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      ComboCheckValidValues=   -1  'True
   End
   Begin VB.Frame sashV 
      BorderStyle     =   0  'None
      Height          =   1692
      Left            =   2760
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      Top             =   360
      Width           =   60
   End
   Begin VB.Label lblRegions 
      BackStyle       =   0  'Transparent
      Caption         =   "&Regions"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmCompute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
''Dim arpos%, vind%, regind%, region%
'Dim rval(MAX_REGR_COMPONENTS, MAX_REG_USE) As Single
''conversion factor array for metrication
'Dim confact(MAX_REGR_COMPONENTS, MAX_REG_USE) As Single
''variable index or cell in agd(row, col)
'Dim agdVarInd() As Long '(MAX_REG_USE * MAX_REGR_COMPONENTS, MAX_REG_USE + 1)
'
'Private SashVdragging As Boolean
'Private InAgd As Boolean
'Private lastHeight&
'
'Private Sub ChkUrbEstimate(imatch&)
'  'look for a match to an existing urban Estimate
'  Dim i&, j&
'
'  If urbcnt > 0 Then 'Estimates exist to compare against
'    i = 0
'    While i < urbcnt
'      'assume a Estimate match
'      imatch = i
'      If urbscn(i).reg = urbscn(urbcnt).reg Then 'same equation in use
'        If urbscn(i).rscn = urbscn(urbcnt).rscn Or Not urbscn(i).National Then
'          'same rural equation in use or not using national equation
'          For j = 0 To urbscn(i).v.vcount - 1 'check value of each urban parameter
'            If urbscn(i).v.Value(j) <> rval(j, 0) Then 'values don't match
'              imatch = -1
'            End If
'          Next j
'        Else              'different rural equations in use
'          imatch = -1
'        End If
'      Else                'different equation type
'        imatch = -1
'      End If
'      If imatch >= 0 Then 'same equation definition, don't look any more
'        i = urbcnt
'      Else                'keep looking for match
'        i = i + 1
'      End If
'    Wend
'  Else                    'no Estimates exist, can't have a match
'    imatch = -1
'  End If
'End Sub
'
'Private Function InRange(Ind&, irg&) As Boolean
'
'  'check validity of an entered parameter value
'  'ind - variable index for this region
'  'irg - region index among region(s) in use
'
'  Dim rmin!, rmax!
'
'  InRange = True                   'assume valid value
'  If rurfg Then             'rural equation in use
'    rmin = ConvertVal(State.Region(rurscn(rurcnt).reg(irg)).v_descriptor(Ind).minimum, confact(Ind, irg))
'    rmax = ConvertVal(State.Region(rurscn(rurcnt).reg(irg)).v_descriptor(Ind).maximum, confact(Ind, irg))
'  Else                             'urban calcualtions
'    If urbscn(urbcnt).reg < 0 Then 'national urban equation in use
'      rmin = ConvertVal(urban_comp(Ind).minimum, confact(Ind, irg))
'      rmax = ConvertVal(urban_comp(Ind).maximum, confact(Ind, irg))
'    Else                           'state urban equation in use
'      rmin = ConvertVal(State.Region(urbscn(urbcnt).reg).v_descriptor(Ind).minimum, confact(Ind, irg))
'      rmax = ConvertVal(State.Region(urbscn(urbcnt).reg).v_descriptor(Ind).maximum, confact(Ind, irg))
'    End If
'  End If
'  If rval(Ind, irg) < rmin Or rval(Ind, irg) > rmax Then 'value out of range
'    InRange = False
'  End If
'End Function
'
'Private Sub agd_CommitChange(ChangeFromRow As Long, ChangeToRow As Long, ChangeFromCol As Long, ChangeToCol As Long)
'  Dim row&, col&, ival&, Region&
'  Dim oldrow&, oldcol&
'  If Not InAgd Then
'    InAgd = True
'    oldrow = agd.row
'    oldcol = agd.col
'    For row = ChangeFromRow To ChangeToRow
'      agd.row = row
'      For col = ChangeFromCol To ChangeToCol
'        agd.col = col
'        Region = col - 1
'        ival = agdVarInd(row, col)
'        If ival >= 0 And IsNumeric(agd.TextMatrix(row, col)) Then
'          rval(ival, Region) = CSng(agd.TextMatrix(row, col))
'          If InRange(ival, Region) Then
'            agd.CellBackColor = agd.InsideLimitsBackground
'          Else
'            agd.CellBackColor = agd.OutsideSoftLimitBackground
'          End If
'          frmNSS.DbgMsg "agd_CommitChange Rgn " & Region & ", Var " & ival & " = " & rval(ival, Region), 3, "Compute", "P"
'        End If
'      Next col
'    Next row
'    'If rurfg Then
'    txtBasinArea_UpdateValue
'    agd.row = oldrow
'    agd.col = oldcol
'    InAgd = False
'  End If
'End Sub
'
'Private Sub agd_RowColChange()
'  Dim curReg&, curVar&, irg&
'  frmNSS.DbgMsg "agd_RowColChange (" & agd.row & ", " & agd.col & ")", 3, "Compute", "K"
'  If Not InAgd Then
'    InAgd = True
'    curVar = agdVarInd(agd.row, agd.col)
'    If curVar < 0 Then
'      agd.ColEditable(agd.col) = False
'      agd.ColSoftMin(agd.col) = NONE
'      agd.ColSoftMax(agd.col) = NONE
'    Else
'      agd.ColEditable(agd.col) = True
'      curReg = agd.col - 1
'      If rurfg And rurcnt >= 0 And curReg >= 0 Then 'rural
'        irg = rurscn(rurcnt).reg(curReg)
'        If irg >= 0 Then
'          agd.ColSoftMin(agd.col) = ConvertVal(State.Region(irg).v_descriptor(curVar).minimum, confact(curVar, curReg))
'          agd.ColSoftMax(agd.col) = ConvertVal(State.Region(irg).v_descriptor(curVar).maximum, confact(curVar, curReg))
'        End If
'      ElseIf Not rurfg And urbcnt >= 0 Then 'urban
'        irg = urbscn(urbcnt).reg
'        If irg < 0 Then     'national urban
'          agd.ColSoftMin(agd.col) = ConvertVal(urban_comp(curVar).minimum, confact(curVar, curReg))
'          agd.ColSoftMax(agd.col) = ConvertVal(urban_comp(curVar).maximum, confact(curVar, curReg))
'        Else                'state urban
'          agd.ColSoftMin(agd.col) = ConvertVal(State.Region(irg).v_descriptor(curVar).minimum, confact(curVar, curReg))
'          agd.ColSoftMax(agd.col) = ConvertVal(State.Region(irg).v_descriptor(curVar).maximum, confact(curVar, curReg))
'        End If
'      End If
'    End If
'    InAgd = False
'  End If
'End Sub
'
'Private Sub agd_TextChange(ChangeFromRow As Long, ChangeToRow As Long, ChangeFromCol As Long, ChangeToCol As Long)
'  'If rurfg Then txtBasinArea_UpdateValue
'End Sub
'
'Private Sub txtBasinArea_UpdateValue()
'  Dim tarea!, regn&, regnarea$
'  tarea = 0
'  'If rurfg Then
'    For regn = 1 To agd.cols - 1
'      regnarea = agd.TextMatrix(1, regn)
'      If IsNumeric(regnarea) Then tarea = tarea + CSng(regnarea)
'    Next
'  'End If
'  If tarea > 0 And txtBasinArea.Value <> tarea Then txtBasinArea.Value = tarea
'End Sub
'
'Private Sub cmdOK_Click()
'
'  Dim regn&, var&, intervl&, closefg&, vcnt&, imatch&, wgtpos&
'  Dim tarea!, rtmp!(MAX_REGR_COMPONENTS), rmin!, rmax!
'  Dim estr$
'  Dim total_basin!(2, MAX_INTERVALS)
'  Dim nOutOfRange
'  Dim eqnMetric As Boolean
'
'  'If txtBasinArea.Enabled And IsNumeric(txtBasinArea.value) Then
'  '  total_area = txtBasinArea.value
'  'End If
'
''  On Error GoTo ErrHandler
'  nOutOfRange = 0
'  closefg = 1
'  tarea = 0
'
'  If RegUseCnt > 0 Then 'If any regions were selected use them to build an estimate
'    For regn = 0 To RegUseCnt - 1
'      If rurfg Then
'        vcnt = State.Region(rurscn(rurcnt).reg(regn)).v_count - 1
'      Else
'        vcnt = urbscn(urbcnt).v.vcount - 1
'      End If
'      tarea = tarea + rval(0, regn) 'sum total area to compare against main window value
'      For var = 0 To vcnt
'        If Not InRange(var, regn) Then
'          nOutOfRange = nOutOfRange + 1
'        End If
'      Next var
'    Next regn
'
'    If tarea <> total_area Then total_area = tarea
'
'    If nOutOfRange > 0 Then 'some values out of range
'      If nOutOfRange = 1 Then estr = "One value is" Else estr = nOutOfRange & " values are"
'      estr = estr & " outside the expected range." & vbCrLf
'      closefg = MsgBox(estr & "Perform computation anyway?", vbOKCancel + vbExclamation + vbDefaultButton2, "NSS Compute")
'    End If
'
'    If closefg = 1 Then
'      If rurfg Then
'        'ok to do rural computations
'        If State.unitsys = 2 Then eqnMetric = True Else eqnMetric = False
'        For intervl = 0 To MAX_INTERVALS
'          total_basin(0, intervl) = 0
'          total_basin(1, intervl) = 0
'          total_basin(2, intervl) = 0
'        Next intervl
'        'look through rural Estimates for match
'        Call ChkRurEstimate(imatch)
'        For regn = 0 To RegUseCnt - 1
'          If imatch < 0 Then           'no match found, new Estimate
'            rurind = rurcnt
'            For var = 0 To State.Region(rurscn(rurind).reg(regn)).v_count - 1
'              'assign values for this Estimate
'              rurscn(rurind).v(regn).Value(var) = rval(var, regn)
'            Next var
'            If regn = 0 Then              'set intervals for new Estimate
'              Call SetIntervals
'              rurscn(rurind).numint = NumIntrvl
'              For intervl = 0 To NumIntrvl - 1
'                rurscn(rurind).intrvl(intervl) = Interval(intervl)
'              Next intervl
'              'set Estimate name to first region descriptor
'              rurscn(rurind).Name = Trim(State.Region(rurscn(rurind).reg(regn)).descriptor)
'            Else 'add other region descriptors to Estimate name
'              rurscn(rurind).Name = Trim(rurscn(rurind).Name) & ", " & Trim(State.Region(rurscn(rurind).reg(regn)).descriptor)
'            End If
'          Else
'            rurind = imatch
'          End If
'          For var = 0 To State.Region(rurscn(rurind).reg(regn)).v_count - 1
'            'set calculation parameters to english units
'            rtmp(var) = ConvertVal(rurscn(rurind).v(regn).Value(var), 1 / confact(var, regn))
'          Next var
'          rtmp(0) = total_area / confact(0, regn)
'          Call CompRuralDis(rurscn(rurind).reg(regn), regn, rtmp())
'          For intervl = 0 To NumIntrvl - 1
'            If metric And Not eqnMetric Then      'using metric and eqtn english, convert
'              rural_discharge(0, intervl, rurind) = rural_discharge(0, intervl, rurind) * FLOW_CONVERSION
'            ElseIf Not metric And eqnMetric Then
'              rural_discharge(0, intervl, rurind) = rural_discharge(0, intervl, rurind) / FLOW_CONVERSION
'            End If
'            total_basin(0, intervl) = total_basin(0, intervl) + rural_discharge(0, intervl, rurind) * rurscn(rurind).v(regn).Value(0) / total_area
'            total_basin(1, intervl) = total_basin(1, intervl) + State.Region(rurscn(rurind).reg(regn)).Interval(intervl).standard_error
'            total_basin(2, intervl) = total_basin(2, intervl) + State.Region(rurscn(rurind).reg(regn)).Interval(intervl).eq_years_of_record
'          Next intervl
'        Next regn
'        For intervl = 0 To NumIntrvl - 1
'          rural_discharge(0, intervl, rurind) = total_basin(0, intervl)
'          If nOutOfRange = 0 Then 'only display std error and equivalent years if vars are in range
'            rural_discharge(1, intervl, rurind) = total_basin(1, intervl) / RegUseCnt
'            rural_discharge(2, intervl, rurind) = total_basin(2, intervl) / RegUseCnt
'          Else
'            rural_discharge(1, intervl, rurind) = 0
'            rural_discharge(2, intervl, rurind) = 0
'          End If
'        Next intervl
'        'Call DispRuralDis
'        rurscn(rurind).regCribu = comboRegion.ListIndex
'        'rurscn(rurind).tarea = total_area
'        compfg = 1
'        wgtpos = InStr(rurscn(rurind).Name, "(Weighted)")
'        If wgtpos > 0 Then      'remove weighted indicator from Estimate name
'          rurscn(rurind).Name = Left(rurscn(rurind).Name, wgtpos - 1)
'        End If
'        If imatch < 0 Then 'new Estimate, update count of rural equations
'          rurcnt = rurcnt + 1
'        End If
'        Call DispEstimate
'        'frmNSS.vsbScen.Max = rurcnt - 1
'        'frmNSS.vsbScen.value = rurind
'        'If rurcnt > 1 Then 'enable scrolling
'        '  frmNSS.vsbScen.Enabled = True
'        'Else
'        '  frmNSS.vsbScen.Enabled = False
'        'End If
'      Else   'urban equation in use
'        'check range of area term
'        If urbscn(urbcnt).reg >= 0 Then
'          'state-specific urban equation
'          urbscn(urbcnt).National = False
'          rmin = State.Region(urbscn(urbcnt).reg).v_descriptor(0).minimum
'          rmax = State.Region(urbscn(urbcnt).reg).v_descriptor(0).maximum
'        Else
'          'national urban equation
'          urbscn(urbcnt).National = True
'          rmin = 0.2
'          rmax = 100#
'        End If
'        If State.unitsys = 2 And Not urbscn(urbcnt).National Then eqnMetric = True Else eqnMetric = False
'        If metric And Not eqnMetric Then      'using metric and eqtn english, convert
'          rmin = rmin * AREA_CONVERSION
'          rmax = rmax * AREA_CONVERSION
'        ElseIf Not metric And eqnMetric Then  'using english and eqtn metric, convert
'          rmin = rmin / AREA_CONVERSION
'          rmax = rmax / AREA_CONVERSION
'        End If
'        If total_area < rmin Or total_area > rmax Then
'          'warn user that area term is out of range
'          MsgBox "Warning: Basin Drainage Area (" & Format(total_area, "#.##") & ") is outside the acceptable range (" & Format$(rmin, "#.##") & "-" & Format$(rmax, "#.##") & ") for this equation.", 48, "NSS Compute"
'        End If
'
'        Call ChkUrbEstimate(imatch) ' see if this Estimate already in use
'        If imatch < 0 Then 'no match found, new Estimate
'          urbind = urbcnt
'          urbcnt = urbcnt + 1
'          For var = 0 To urbscn(urbind).v.vcount - 1
'            'assign urban parameters for this Estimate
'            urbscn(urbind).v.Value(var) = rval(var, 0)
'          Next var
'          If urbscn(urbind).National Then
'            urbscn(urbind).Name = "National Urban Equations (R" & rurind + 1 & ")"
'          Else
'            urbscn(urbind).Name = Trim(State.Region(urbscn(urbind).reg).descriptor)
'            If Left(urbscn(urbind).Name, 1) = "$" Then urbscn(urbind).Name = Mid(urbscn(urbind).Name, 2)
'          End If
'          'set intervals for this Estimate
'          SetIntervals
'          urbscn(urbind).numint = NumIntrvl
'          For intervl = 0 To NumIntrvl - 1
'            urbscn(urbind).intrvl(intervl) = Interval(intervl)
'          Next intervl
'        Else 'use existing Estimate for calculations
'          urbind = imatch
'          SetIntervals
'        End If
'        For var = 0 To urbscn(urbind).v.vcount - 1
'          'set calculation parameters to english units
'          rtmp(var) = ConvertVal(rval(var, 0), 1 / confact(var, 0))
'        Next var
'        'do urban calculations
'        CompUrbanDis urbscn(urbind).reg, rtmp()
'        DispUrbanDis
'        DispEstimate
'      End If
'    End If
'  End If
'  GoTo EndErr
'ErrHandler:
'  estr = "Unable to perform calculations due to invalid input." & vbCrLf & "Check parameter values entered on Compute window."
'  MsgBox estr, 48
'  closefg = 0
'EndErr:
'  If closefg = 1 Then Unload frmCompute '.Hide
'End Sub
'
'Private Sub cmdCancel_Click()
'  Unload frmCompute '.Hide
'End Sub
'
'Sub SetComputeGrid()
'  Dim row&, col&, irg&, vcnt&, UnitIndex&, desiredHeight&, ival&
'  Dim curReg&, curVar&, curVarName$, curVarAbbrev, UnitString$
'  Dim eqnMetric As Boolean
'
'  For row = 0 To UBound(agdVarInd, 1)
'    For col = 0 To UBound(agdVarInd, 2)
'      agdVarInd(row, col) = -1
'    Next col
'  Next row
'  If rurfg Then
'    SetRurEstimate
'  Else
'    SetUrbEstimate  'init urban info
'  End If
'
'  agd.rows = 1 'will be expanded as necessary below
'  agd.TextMatrix(agd.rows, 0) = ""
'  agd.cols = RegUseCnt + 1
'  agd.ColType(0) = ATCoTxt
'  agd.ColTitle(0) = "Variable"
'
'  'RegUseCnt = 0
'  'For curReg = 0 To lstRegion.ListCount - 1
'  '  If lstRegion.Selected(curReg) Then
'  '    If rurfg And rurcnt > 0 Then
'  '      rurscn(rurcnt).reg(RegUseCnt) = curReg
'  '    Else
'  '      urbscn(urbcnt).reg = curReg
'  '    End If
'  '    RegUseCnt = RegUseCnt + 1
'  '  End If
'  'Next curReg
'  If rurfg And rurcnt > -1 Then rurscn(rurcnt).rcnt = RegUseCnt
'  For curReg = 0 To RegUseCnt - 1
'    col = curReg + 1
'    agd.ColType(col) = ATCoSng
'    agd.ColEditable(col) = True
'    If rurfg And rurcnt > -1 Then
'      irg = rurscn(rurcnt).reg(curReg)
'      If irg >= 0 Then
'        agd.ColTitle(col) = State.Region(irg).descriptor
'        vcnt = State.Region(irg).v_count
'      End If
'    ElseIf Not rurfg And urbcnt > -1 Then 'urban calculations in use
'      irg = urbscn(urbcnt).reg
'      If irg >= 0 Then 'state urban equation
'        If Left(State.Region(irg).descriptor, 1) = "$" Then 'strip leading "$"
'          agd.ColTitle(col) = Mid(State.Region(irg).descriptor, 2)
'        Else
'          agd.ColTitle(col) = State.Region(irg).descriptor
'        End If
'        'vcnt = State.Region(irg).v_count
'      Else 'national urban equations
'        agd.ColTitle(col) = "National Urban Equations"
'        'vcnt = URBAN_COMPONENTS + 1
'      End If
'      vcnt = urbscn(urbcnt).v.vcount
'    Else 'Can't set rural or urban
'      Exit Sub
'    End If
''    Stop
'    For curVar = 0 To vcnt - 1
'      confact(curVar, curReg) = 1
'      If rurfg Then
'        curVarName = State.Region(irg).v_descriptor(curVar).descriptor
'        UnitIndex = State.Region(irg).v_descriptor(curVar).units
'        curVarAbbrev = State.Region(irg).v_descriptor(curVar).variable_name
'        If RegUseCnt = 1 And curVar = 0 And rurscn(rurcnt).v(curReg).Value(curVar) <= 0 Then
'          'only one region and area term is undefined, default to total area from main window
'          rurscn(rurcnt).v(curReg).Value(curVar) = total_area
'        End If
'        rval(curVar, curReg) = rurscn(rurcnt).v(curReg).Value(curVar)
'        If State.unitsys = 2 Then eqnMetric = True Else eqnMetric = False
'      Else 'urban calculations
'        If irg < 0 Then 'national urban equations
'          curVarName = urban_comp(curVar).descriptor
'          UnitIndex = urban_comp(curVar).units
'          curVarAbbrev = urban_comp(curVar).variable_name
'        Else 'state urban equation
'          curVarName = State.Region(irg).v_descriptor(curVar).descriptor
'          UnitIndex = State.Region(irg).v_descriptor(curVar).units
'          curVarAbbrev = State.Region(irg).v_descriptor(curVar).variable_name
'        End If
'        rval(curVar, curReg) = urbscn(urbcnt).v.Value(curVar)
'        If State.unitsys = 2 And Not urbscn(urbcnt).National Then eqnMetric = True Else eqnMetric = False
'      End If
'
'      If metric And Not eqnMetric Then      'using metric and eqtn english, convert
'        confact(curVar, curReg) = units(UnitIndex).factor
'      ElseIf Not metric And eqnMetric Then  'using english and eqtn metric, convert
'        confact(curVar, curReg) = 1 / units(UnitIndex).factor
'      Else                                  'no conversion needed
'        confact(curVar, curReg) = 1
'      End If
'
'      If metric Then
'        UnitString = Trim(units(UnitIndex).SI_sym)
'      Else
'        UnitString = Trim(units(UnitIndex).IP_sym)
'      End If
'
'      'curVarName = Trim(curVarName)
'      'curVarAbbrev = Trim(curVarAbbrev)
'      If Len(UnitString) > 0 Then curVarName = curVarName & " (" & UnitString & ")"
'      If Len(curVarAbbrev) > 0 Then curVarName = curVarName & " " & curVarAbbrev
'
'      'Find row that already has this variable if it already exists
'      row = 1
'      While row < agd.rows And agd.TextMatrix(row, 0) <> curVarName
'        row = row + 1
'      Wend
'      If agd.TextMatrix(row, 0) <> curVarName Then
'        While agd.TextMatrix(row, 0) <> ""
'          row = row + 1
'        Wend
'        agd.TextMatrix(row, 0) = curVarName
'      End If
'      'On Error GoTo RedimagdVarInd
'      agdVarInd(row, col) = curVar
'      'On Error GoTo 0
'    Next curVar
'  Next curReg
'
'  For row = 1 To agd.rows
'    For col = 1 To agd.cols - 1
'      agd.row = row
'      agd.col = col
'      ival = agdVarInd(row, col)
'      If ival >= 0 Then
'        If rval(ival, col - 1) = -999 Then
'          agd.Text = ""
'          agd.CellBackColor = agd.OutsideSoftLimitBackground
'        Else
'          agd.Text = rval(ival, col - 1)
'          If InRange(ival, col - 1) Then
'            agd.CellBackColor = agd.InsideLimitsBackground
'          Else
'            agd.CellBackColor = agd.OutsideSoftLimitBackground
'          End If
'        End If
'      Else
'        agd.CellBackColor = agd.BackColorFixed
'        agd.Text = ""
'      End If
'    Next col
'  Next row
'  txtBasinArea_UpdateValue
'  agd.row = 1
'  agd.col = 1
'  'If agd.rows > UBound(agdVarInd, 1) Or agd.cols > UBound(agdVarInd, 2) Then
'  '  ReDim Preserve agdVarInd(agd.rows, agd.cols)
'  'End If
'  agd.ColsSizeByContents
'  agd_RowColChange
'  Me.Width = agd.gridWidth + agd.Left + 100
'  If rurfg Then
'    desiredHeight = 2100 + 253 * (agd.rows + 1)
'  Else
'    desiredHeight = 1655 + 253 * (agd.rows + 1)
'  End If
'  If desiredHeight > Height Then Height = desiredHeight
''  Exit Sub
''RedimagdVarInd:
''  ReDim Preserve agdVarInd(agd.rows, agd.cols)
''  Resume Next
'End Sub
'
'Private Sub cmdMap_Click()
'  frmCrippenBue.Show vbModal 'MsgBox "fixme: open help tp Crippen & Bue map"
'End Sub
'
'Private Sub Form_Load()
'  ReDim agdVarInd(MAX_REG_USE * MAX_REGR_COMPONENTS, MAX_REG_USE + 1) '(2, 2)
'  lstRegion.Clear
'  'If rurfg Then rurcnt = rurcnt + 1 Else urbcnt = urbcnt + 1
'  Call ListRegions
'  fraCribu.Visible = rurfg
'  fraTotalArea.Visible = rurfg
'  'If rurfg Then txtBasinArea.Enabled = False
'  If metric Then lblUnits.Caption = "km" Else lblUnits.Caption = "mi"
'
'  HelpContextID = 29
'  txtBasinArea.Value = total_area
'  'SetComputeGrid
'  'txtBasinArea_UpdateValue
'
'  If rurfg And comboRegion.ListCount < 10 Then
'    comboRegion.Clear
'    comboRegion.AddItem "None"
'    Dim rgn&
'    For rgn = 1 To 17
'      comboRegion.AddItem rgn
'    Next rgn
'    comboRegion.ListIndex = 0
'  End If
'
'  lastHeight = 0
'  Form_Resize
'End Sub
'
'Private Sub Form_Resize()
'  Dim fh%, fw%, newRightWidth%
'
'  fw = Width
'  fh = Height
'  'sashV.Left = fw - agd.Width - sashV.Width - 100
'  'If sashV.Left < 200 Then
'  '  sashV.Left = 100
'  '  sashV_MouseMove 0, 0, 0, 0
'  'Else
'  'If Width > 3500 Then agd.Width = Width - 3150
'  newRightWidth = fw - (sashV.Left + sashV.Width + 100)
'  If newRightWidth > 0 And agd.Width <> newRightWidth Then agd.Width = newRightWidth
'
'  agd.Left = sashV.Left + sashV.Width
'  fraTotalArea.Left = agd.Left
'  fraCribu.Left = agd.Left
'  If Height > 2000 Then
'    If rurfg Then
'      agd.Height = Height - 1800
'      fraCribu.Top = agd.Top + agd.Height + comboRegion.Height / 2 ' comboRegion.Top = agd.Top + agd.Height + comboRegion.Height / 2
'      cmdOk.Top = fraCribu.Top + fraCribu.Height + cmdOk.Height / 3
'    Else
'      agd.Height = Height - 1655
'      cmdOk.Top = agd.Top + agd.Height + cmdOk.Height / 3
'    End If
'    If agd.Height <> lastHeight Then 'eliminate flashing because list height snaps to full line heights
'      lastHeight = agd.Height
'      lstRegion.Height = lastHeight
'      sashV.Height = lastHeight
'    End If
'    cmdCancel.Left = Width / 2 + cmdCancel.Width / 4
'    cmdOk.Left = cmdCancel.Left - cmdOk.Width * 1.5
'  End If
'
'  'lblRegion.Top = comboRegion.Top
'  cmdCancel.Top = cmdOk.Top
'End Sub
'
''list available regions for a selected state
'Private Sub ListRegions()
'  Dim reg&, nextReg&, regdesc$
'  Dim intrvl&, comp&
'  Dim okToAdd As Boolean
'  Dim NotAdded$, plural$
'
'  plural = " was"
'
'  If Not rurfg Then 'first value in list is national urban equations
'    If rurind < 0 Or rurcnt < 1 Then
'      NotAdded = "National Urban Equations" & vbCr
'    Else
'      lstRegion.AddItem "National Urban Equations"
'      RegAvl(0) = -1
'      nextReg = 1
'    End If
'  Else
'    nextReg = 0
'  End If
'  For reg = 0 To State.regcount - 1
'    regdesc = State.Region(reg).descriptor
'    If Left(State.Region(reg).descriptor, 1) = "$" Then
'      If Not rurfg Then
'        okToAdd = True
'        If rurind < 0 Or rurcnt < 1 Then
'          With State.Region(reg)
'            For intrvl = 0 To .i_count - 1
'              For comp = 0 To .Interval(intrvl).comp_count - 1
'                If .Interval(intrvl).regr_component(comp).base_variable < 0 Then
'                  'indicates to use rural discharge or area in urban computations
'                  okToAdd = False
'                End If
'              Next comp
'            Next intrvl
'          End With
'        End If
'
'        If okToAdd Then
'          lstRegion.AddItem Mid(regdesc, 2)
'          RegAvl(nextReg) = reg
'          nextReg = nextReg + 1
'        Else
'          If NotAdded <> "" Then plural = "s were"
'          NotAdded = NotAdded & Mid(regdesc, 2) & vbCr
'        End If
'      End If
'    Else
'      If rurfg Then
'        lstRegion.AddItem regdesc
'        RegAvl(nextReg) = reg
'        nextReg = nextReg + 1
'      End If
'    End If
'  Next reg
'  If NotAdded <> "" Then MsgBox "A rural result was not selected, so the following urban equation" & plural & " not available" & vbCr & NotAdded
'  If lstRegion.ListCount = 1 Then lstRegion.Selected(0) = True ': lstRegion_Click
'End Sub
'
'Private Sub SetUrbEstimate()  'set state for next urban Estimate
'  Dim var&, var2& ', newV As vsav_type
'  If urbind < 0 Then urbind = 0
'  With urbscn(urbcnt)
'    If .reg < 0 Then 'national urban equation in use
'      .National = True
'      .rscn = rurind
'      .v.vcount = URBAN_COMPONENTS + 1
'      For var = 0 To .v.vcount - 1
'        .v.vname(var) = urban_comp(var).variable_name
'      Next var
'    Else                           'state urban equation in use
'      .National = False
'      .v.vcount = State.Region(.reg).v_count
'      For var = 0 To .v.vcount - 1
'        .v.vname(var) = State.Region(.reg).v_descriptor(var).variable_name
'      Next var
'    End If
'
'    For var = 0 To .v.vcount - 1
'      .v.Value(var) = -999
'      For var2 = 0 To urbscn(urbind).v.vcount - 1
'        If .v.vname(var) = urbscn(urbind).v.vname(var2) Then
'          .v.Value(var) = urbscn(urbind).v.Value(var2)
'          var2 = urbscn(urbind).v.vcount
'        End If
'      Next var2
'    Next var
'
'    '.v.vcount = newV.vcount
'    'For var = 0 To newV.vcount - 1
'    '  .v.vname(var) = newV.vname(var)
'    '  .v.value(var) = newV.value(var)
'    'Next var
'    If .v.vname(0) = "A" Then .v.Value(0) = total_area
'  End With
'End Sub
'
'Public Sub SetRurEstimate()    'set state for next rural Estimate
'  Dim regn&, regn2&, var&, var2&, newV As vsav_type
'
'  If rurind < 0 Then rurind = 0
'  If rurscn(rurcnt).rcnt = 0 Then    'use same count of regions as current Estimate
'    rurscn(rurcnt).rcnt = rurscn(rurind).rcnt
'  End If
'  For regn = 0 To RegUseCnt - 1
'    If rurscn(rurcnt).reg(regn) < 0 Then      'use same region as current Estimate
'      rurscn(rurcnt).reg(regn) = rurscn(rurind).reg(regn)
'    End If
'  Next regn
'  For regn = 0 To RegUseCnt - 1
'    If rurscn(rurcnt).reg(regn) >= 0 Then
'      rurscn(rurcnt).v(regn).vcount = State.Region(rurscn(rurcnt).reg(regn)).v_count
'      For var = 0 To rurscn(rurcnt).v(regn).vcount - 1
'        newV.vname(var) = State.Region(rurscn(rurcnt).reg(regn)).v_descriptor(var).variable_name
'        'rurscn(rurcnt).v(regn).value(var) = rurscn(rurind).v(regn).value(var)
'      Next var
'      For var = 0 To rurscn(rurcnt).v(regn).vcount - 1
'        newV.Value(var) = -999
'        For regn2 = 0 To regn
'          For var2 = 1 To rurscn(rurcnt).v(regn2).vcount - 1 ' start at 1 to skip area
'            If newV.vname(var) = rurscn(rurind).v(regn2).vname(var2) Then
'              newV.Value(var) = rurscn(rurind).v(regn2).Value(var2)
'              var2 = rurscn(rurcnt).v(regn2).vcount
'            End If
'          Next var2
'        Next regn2
'      Next var
'      For var = 0 To rurscn(rurcnt).v(regn).vcount - 1
'        rurscn(rurcnt).v(regn).vname(var) = newV.vname(var)
'        rurscn(rurcnt).v(regn).Value(var) = newV.Value(var)
'      Next var
'    End If
'  Next regn
'End Sub
'
'Public Sub ChkRurEstimate(imatch&)
'
'  'look for a match to an existing rural Estimate
'  Dim i&, j&, k&
'
'  If rurcnt > 0 Then      'rural Estimates exist to compare against
'    i = 0
'    While i < rurcnt      'assume a Estimate match
'      imatch = i
'      If rurscn(i).rcnt = rurscn(rurcnt).rcnt Then 'same number of regions
'        For j = 0 To rurscn(i).rcnt - 1
'          If rurscn(i).reg(j) = rurscn(rurcnt).reg(j) Then  'same region
'            For k = 0 To rurscn(i).v(j).vcount - 1 'check value of each parameter
'              If rurscn(i).v(j).Value(k) <> rval(k, j) Then 'values don't match
'                imatch = -1
'              End If
'            Next k
'          Else            'different regions in use
'            imatch = -1
'          End If
'        Next j
'      Else                'different number of regions
'        imatch = -1
'      End If
'      If imatch >= 0 Then 'same equation definition, don't look any more
'        i = rurcnt
'      Else                'keep looking for match
'        i = i + 1
'      End If
'    Wend
'  Else                    'no rural Estimates to check
'    imatch = -1
'  End If
'
'End Sub
'
'Private Sub lstRegion_Click()
'  Static InClick As Boolean
'  Debug.Print Time, "lstRegion_Click"
'  If Not InClick Then
'    InClick = True
'    If Not rurfg Then
'      Dim i&
'      For i = 0 To lstRegion.ListCount - 1
'        If i = lstRegion.ListIndex Then lstRegion.Selected(i) = True Else lstRegion.Selected(i) = False
'      Next i
'    End If
'    SetSelectedRegionsFromList 'AddRemoveRegion lstRegion.ListIndex, lstRegion.Selected(lstRegion.ListIndex)
'    InClick = False
'  End If
'End Sub
'
'Private Sub SetSelectedRegionsFromList() 'AddRemoveRegion(ind&, Add As Boolean) 'add or remove an available region for use in analysis
'  Static SettingSelected As Boolean
'  Dim i&, j&, lstIndx&, Ind&
'  Dim RurReq As Boolean 'does this urban equation require rural results?
'
'  If lstRegion.SelCount > MAX_REG_USE And rurfg Then
'    MsgBox "A maximum of " & MAX_REG_USE & " regions may be used for rural calculations.", 48, "NSS Limit"
'  ElseIf Not rurfg And lstRegion.SelCount > 1 Then
'    MsgBox "Only one equation may be used for urban calculations.", 48, "NSS Limit"
'  Else
'    'ok to add the region
'    If rurfg Then
'      RegUseCnt = 0 'lstRegion.SelCount
'      For lstIndx = 0 To lstRegion.ListCount - 1
'        If lstRegion.Selected(lstIndx) Then
'          rurscn(rurcnt).reg(RegUseCnt) = RegAvl(lstIndx)
'          RegUseCnt = RegUseCnt + 1
'        End If
'      Next lstIndx
'      rurscn(rurcnt).rcnt = RegUseCnt
'    Else      'urban
'      RegUseCnt = 1
'      For lstIndx = 0 To lstRegion.ListCount - 1
'        If lstRegion.Selected(lstIndx) Then Ind = RegAvl(lstIndx)
'      Next lstIndx
'      urbscn(urbcnt).reg = Ind
'      'does this urban equation require rural results?
'    End If
'  End If
'  If Not rurfg And RurReq Then
'    MsgBox "Before this urban equation can be used, a rural equation for the area must be computed."
'  Else
'    SetComputeGrid
'  End If
'End Sub
'
''Sub RemoveRegion()
''
''    Dim ind&, pos&, i&
''
'    'ind = lstUseRegion.ListIndex
'    'If ind >= 0 Then
'      'return item to available list
''      If rurfg = True Then
'        'adjust available rural regions
''        RegAvl(pos) = rurscn(rurcnt).reg(ind)
''        rurscn(rurcnt).reg(RegUseCnt - 1) = 0
''      End If
''      RegUseCnt = lstRegion.SelCount
''      rurscn(rurcnt).rcnt = RegUseCnt
'    'End If
'
''End Sub
'
'Private Sub sashV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  SashVdragging = True
'End Sub
'
'Private Sub sashV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  If SashVdragging And (sashV.Left + X) > 100 And (sashV.Left + X < Width - 100) Then
'    Dim newLeftWidth&
'    sashV.Left = sashV.Left + X
'    If sashV.Left < lstRegion.Left + 200 Then sashV.Left = lstRegion.Left + 200
'    newLeftWidth = sashV.Left - lstRegion.Left
'    If newLeftWidth > 0 And lstRegion.Width <> newLeftWidth Then lstRegion.Width = newLeftWidth
'    Form_Resize
'  End If
'End Sub
'
'Private Sub sashV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  SashVdragging = False
'End Sub
'
Private Sub Form_Load()

End Sub
