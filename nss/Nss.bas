Attribute VB_Name = "NSS1"
'Option Explicit
'
'Global nssDB As nssDatabase
'Global nssProj As nssProject
'Global curState As nssState
'
'Global Const MAX_RGR_DES_LEN = 41  'max regression component desc. length
'Global Const MAX_VAR_LEN = 10  'max var name length
'Global Const MAX_VAR_LENP1 = 11  'same as above plus 1
''Global Const MAX_REG_DES_LEN = 45  'max regional desc. length
'Global Const MAX_REG_DES_LENP1 = 46  'same as above plus 1
'Global Const MAX_STA_DES_LEN = 21  'max state desc length
'Global Const ID_CODE_LEN = 4  'state id code length
'Global Const MAX_REGR_COMPONENTS = 11  'max # regression formula components
'Global Const MAX_REGIONS = 19  'max number of regions in each state
'Global Const MAX_INTERVALS = 7  'max # recurrence intervals
'Global Const MAX_STATES = 60  'max # states
'Global Const URBAN_COMPONENTS = 4  'max # urban components
'Global Const MAX_Estimate = 8  'max # allowed urban or rural discharge calculations
'Global Const MAX_HYDRO_PLOTS = 3  'max # hydrographs to plot at once
'Global Const HYDRO_SIZE = 45  '# elements in dimensionless hydrograph
'Global Const URBAN_ROFFSET = 0  'urban discharge table row offset
'Global Const URBAN_COFFSET = 46  'urban discharge table col offset
'Global Const RURAL_ROFFSET = 0  'rural discharge table row offset
'Global Const RURAL_COFFSET = 0  'rural discharge table col offset
'Global Const MAX_REG_USE = 10  'max # regions in use at a time
'Global Const MAX_UNIT_TYPES = 19  'maximum types of units allowed
'Global Const MAX_SYM_LEN = 24  'maximum length of units symbol string
'Global Const FLOW_CONVERSION = 0.028317
'Global Const AREA_CONVERSION = 2.589988
'Global Const NoEstimates = "No Estimates"
''#define FLOW_FORMAT                     (metric ? %ll.2f : %ll.0)
''#define C_2_F(a)                        (a * 1.8f + 32.0f)
''#define F_2_C(a)                        ((a - 32.0f) / 1.8f)
''#define FLOW_UNITS                      (metric ? "m^3/s" : "ft3/s")
''#define AREA_UNITS                      (metric ? "km^2 " : "sq mi")
'
''context sensitive help IDs
'Global Const Area_1 = 10
'Global Const Area_2 = 11
'Global Const Area_3 = 12
'Global Const Area_4 = 13
'Global Const Area_5 = 14
'Global Const Area_6 = 15
'Global Const IDH_CONTENTS = 1
'Global Const National_Urban = 4
'Global Const Procedure = 8
'Global Const State_Map = 5
'Global Const STATEWIDE_RURAL = 6
'Global Const Statewide_Urban = 3
'Global Const Summary = 7
'
'Type regr_type
'    base_variable As Integer  'index into variable table of base variable
'    base_modifier As Single  'modifier to add to base_variable
'    base_multiplier As Single  'value to multiply base_variable by
'    base_exponent As Single  'exponent of component
'    exp_variable As Integer  'index into variable table of exponent variable
'    exp_modifier As Single  'modifier to add to exp variable
'    exp_exponent As Single  'exponent exponent
'End Type
'
'Type vari_type
'    variable_name As String '* MAX_VAR_LENP1  'name of variable
'    descriptor As String '* MAX_RGR_DES_LEN  'variable descriptor
'    minimum As Single  'variable minimum
'    maximum As Single  'variable maximum
'    units As Integer  'units of measure for variable
'End Type
'
'Type inter_type
'    intval As Single          'numeric value for this interval
'    standard_error As Single  '% standard error for data
'    eq_years_of_record As Single  'equivalent yrs of record for data
'    regr_constant As Single  'regression constant for interval
'    comp_count As Integer  'number of components in equation
'    regr_component(MAX_REGR_COMPONENTS) As regr_type  'regression components for interval
'End Type
'
'Type region_type
'    descriptor As String '* MAX_REG_DES_LENP1  'region descriptor
'    v_count As Integer  'count of components of equation
'    i_count As Integer  'count of intervals for this region
'    v_descriptor(MAX_REGR_COMPONENTS) As vari_type  'descriptors of components
'    Interval(MAX_INTERVALS) As inter_type  'const, mod, and exp for interval
'End Type
'
'Type state_type
'    id_code As String    '* ID_CODE_LEN  'state id code
'    descriptor As String '* MAX_STA_DES_LEN  'state descriptor
'    regcount As Integer  'number of regions in state
'    unitsys As Integer   'unit system for state's equations
'    Region(MAX_REGIONS) As region_type  'regional data
'End Type
'
'Type uregr_type
'    base_modifier As Single  'modifier to add to base variable
'    base_exponent As Single  'exponent of component
'End Type
'
'Type urban_type
'    standard_error As Single
'    regr_constant As Single  'regression constant
'    a_exponent As Single  'area exponent
'    r_exponent As Single  'rural exponent
'    b_exponent As Single  'bdf exponent
'    regr_component(URBAN_COMPONENTS) As uregr_type  'component modifiers and exponents
'End Type
'
'Type vsav_type
'    vcount As Integer
'    vname(MAX_REGR_COMPONENTS) As String '* MAX_VAR_LEN
'    Value(MAX_REGR_COMPONENTS) As Single
'End Type
'
'Type rursav_type
'    Name As String '* 80
'    rcnt As Integer
'    regCribu As Integer
'    'tarea As Single
'    reg(MAX_REG_USE) As Integer
'    numint As Integer
'    intrvl(MAX_INTERVALS) As Single
'    v(MAX_REG_USE) As vsav_type
'End Type
'
'Type urbsav_type
'    National As Boolean ', Whether or not a national urban calculation
'    Name As String '* MAX_REG_DES_LEN
'    reg As Integer
'    rscn As Integer
'    numint As Integer
'    intrvl(MAX_INTERVALS) As Single
'    v As vsav_type
'End Type
'
'Type units_type
'    'SI -- inch-pound conversion information
'    SI_sym As String '* MAX_SYM_LEN
'    IP_sym As String '* MAX_SYM_LEN
'    factor As Single
'End Type
'
'Global plot_descriptor As String '* 80
'Global State As state_type
'Global stpos(MAX_STATES) As Long
'Global RegAvl(MAX_REGIONS) As Integer
'Global rural_discharge(2, MAX_INTERVALS, MAX_Estimate) As Single
'Global urban_discharge(MAX_INTERVALS, MAX_Estimate) As Single
'Global urban_comp(URBAN_COMPONENTS) As vari_type
'Global urban_eqtn(MAX_INTERVALS) As urban_type
'Global total_area As Single
'Global RegUseCnt  As Integer
'Global Interval(MAX_INTERVALS) As Single
'Global NumIntrvl As Integer
'Global NatUrbEqInts(MAX_INTERVALS) As Single
'Global urbscn(MAX_Estimate) As urbsav_type
'Global rurscn(MAX_Estimate) As rursav_type
'Global urbcnt&, urbind&, rurcnt&, rurind&
'Global disch_ratio(HYDRO_SIZE) As Single
'Global units(MAX_UNIT_TYPES) As units_type
'Dim b0, b1, b2 As Double
'Global UsrID$, PrjID$
'Global compfg&, metric As Boolean, stind&, rurfg As Boolean, cnvrtfg&
'
''stats routines in DLL used by NSS
''Declare Sub VB_SLREG Lib "VB_NSS" (r!, r!, L&, r!, r!)
''Declare Sub VB_SLREG2 Lib "VB_NSS" (r!, r!, L&, r!, r!, r!)
''Declare Function VB_HARTRG Lib "VB_NSS" (r!) As Single
''Declare Function VB_WILFRT Lib "VB_NSS" (r!, r!, L&) As Single
'
''Help engine declarations
''Commands to pass WinHelp()
'Global Const HELP_CONTEXT = &H1     ' Display topic identified by number in Data
'Global Const HELP_QUIT = &H2        ' Terminate help
'Global Const HELP_INDEX = &H3       ' Display index
'Global Const HELP_HELPONHELP = &H4  ' Display help on using help
'Global Const HELP_SETINDEX = &H5    ' Set an alternate Index for help file with more than one index
'Global Const HELP_KEY = &H101       ' Display topic for keyword in Data
'Global Const HELP_MULTIKEY = &H201  ' Lookup keyword in alternate table and display topic
''Windows Help function
'Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
'
''declaration to execute html files
'Declare Function ShellExecute Lib _
'    "shell32.dll" Alias "ShellExecuteA" _
'    (ByVal hwnd As Long, _
'    ByVal lpOperation As String, _
'    ByVal lpFile As String, _
'    ByVal lpParameters As String, _
'    ByVal lpDirectory As String, _
'    ByVal nShowCmd As Long) As Long
'
'
'Sub BldEqtn(irg%, intrvl!, ostr$)
'
'    'build equation for main output file
'    Dim i&, ivl&, Ind&, tstr$
'
'    ivl = -1
'    i = 0
'    While i < State.Region(irg).i_count
'      If State.Region(irg).Interval(i).intval = intrvl Then
'        'intervals match, use this one
'        ivl = i
'        i = State.Region(irg).i_count
'      End If
'      i = i + 1
'    Wend
'    If ivl >= 0 Then
'      If State.Region(irg).Interval(ivl).comp_count > 0 Then
'        ostr = Format$(State.Region(irg).Interval(ivl).regr_constant, "#.###")
'        For i = 0 To State.Region(irg).Interval(ivl).comp_count - 1
'          If State.Region(irg).Interval(ivl).regr_component(i).base_variable <> 0 Then
'            Ind = Abs(State.Region(irg).Interval(ivl).regr_component(i).base_variable) - 1
'            tstr = State.Region(irg).v_descriptor(Ind).variable_name
'            If State.Region(irg).Interval(ivl).regr_component(i).base_modifier <> 0 Then
'              'include base modifier
'              tstr = "(" & tstr & "+" & Format$(State.Region(irg).Interval(ivl).regr_component(i).base_modifier, "#.###") & ")"
'            End If
'            ostr = ostr & "*" & tstr
'          End If
'          ostr = ostr & "^" & Format$(State.Region(irg).Interval(ivl).regr_component(i).base_exponent, "#.###")
'        Next i
'      Else
'        'no equation for this interval
'        ostr = "N/A"
'      End If
'    Else
'      'this interval not available for this region
'      ostr = "N/A"
'    End If
'
'End Sub
'
'Sub CompRuralDis(ByVal Region&, ByVal regind&, rval() As Single)
'
'    Dim i&, j&, Index&
'    Dim bvtmp#, evtmp#
'    Static lrdis(MAX_INTERVALS) As Single
'
'    'compute rural discharge
'    For i = 0 To NumIntrvl - 1
'      rural_discharge(0, i, rurind) = State.Region(Region).Interval(i).regr_constant
'      If State.Region(Region).Interval(i).comp_count > 0 Then
'        For j = 0 To State.Region(Region).Interval(i).comp_count - 1
'          With State.Region(Region).Interval(i).regr_component(j)
'            'determine entire exponent's value
'            If .exp_variable <> 0 Then
'              Index = Abs(.exp_variable) - 1
'              evtmp = Abs(rval(Index) + .exp_modifier) 'evtmp = Abs(rurscn(rurind).v(regind).value(Index) + .exp_modifier)
'              If .exp_variable < 0 Then
'                evtmp = Log10(evtmp)
'              End If
'              evtmp = evtmp ^ .exp_exponent
'            Else
'              evtmp = 1#
'            End If
'            evtmp = evtmp * CDbl(.base_exponent)
'            'determine entire base variable's value
'            If .base_variable > 0 Then
'              Index = .base_variable - 1
'              bvtmp = CDbl(rval(Index))           'bvtmp = CDbl(rurscn(rurind).v(regind).value(Index))
'            Else
'              bvtmp = 0#
'            End If
'            bvtmp = Abs(bvtmp + .base_modifier)
'            bvtmp = bvtmp * CDbl(.base_multiplier)
'            'determine discharge value
'            rural_discharge(0, i, rurind) = rural_discharge(0, i, rurind) * bvtmp ^ evtmp
'          End With
'        Next j
'        lrdis(i) = rural_discharge(0, i, rurind)
'      Else
'        'no components, so force discharge to zero
'        lrdis(i) = 0#
'        If i = NumIntrvl - 1 Then 'if we're dealing w/ Q500...extrapolate Q500
'          Call nss500(lrdis())
'          'assign eq_years_of_record from Q100 to Q500 for calculation of
'          'the weighted average of observed and regression estimates
'          With State.Region(Region)
'            If .Interval(i - 1).eq_years_of_record > 0# Then .Interval(i).eq_years_of_record = .Interval(i - 1).eq_years_of_record
'          End With
'        End If
'        rural_discharge(0, i, rurind) = lrdis(i)
'      End If
'    Next i
'
'End Sub
'
'Private Function GetRuralDischarge(ruralIndex&, YearInterval As Single)
'  Dim i&
'  GetRuralDischarge = 0
'  For i = 0 To rurscn(ruralIndex).numint - 1
'    If Abs(rurscn(ruralIndex).intrvl(i) - YearInterval) < 0.5 Then 'floating point sucks
'      GetRuralDischarge = rural_discharge(0, i, rurind)
'    End If
'  Next i
'End Function
'
'Sub CompUrbanDis(Region As Integer, rval() As Single)
'
'    Dim i&, j&, Index&
'    Dim bvtmp#, evtmp#, utmp!
'    Static ludis(MAX_INTERVALS) As Single
'
'    On Error GoTo CompErr
'
'    'compute urban discharge
'    For i = 0 To NumIntrvl - 1
'      If urbscn(urbind).National Then
'        'national urban equation in use
'        utmp = urban_eqtn(i).regr_constant
'        utmp = utmp * total_area ^ urban_eqtn(i).a_exponent
'        utmp = utmp * GetRuralDischarge(rurind, urbscn(urbind).intrvl(i)) ^ urban_eqtn(i).r_exponent
'        utmp = utmp * (13 - rval(0)) ^ urban_eqtn(i).b_exponent
'        For j = 0 To URBAN_COMPONENTS - 1
'          bvtmp = rval(j + 1) + urban_eqtn(i).regr_component(j).base_modifier
'          evtmp = urban_eqtn(i).regr_component(j).base_exponent
'          utmp = utmp * bvtmp ^ evtmp
'        Next j
'        urban_discharge(i, urbind) = utmp
'      Else
'        'one of the state urban equations in use
'        urban_discharge(i, urbind) = State.Region(Region).Interval(i).regr_constant
'        If State.Region(Region).Interval(i).comp_count <> 0 Then
'          For j = 0 To State.Region(Region).Interval(i).comp_count - 1
'            With State.Region(Region).Interval(i).regr_component(j)
'              'determine entire exponent's value
'              If .exp_variable <> 0 Then
'                Index = Abs(.exp_variable) - 1
'                evtmp = Abs(rval(Index - 1) + .exp_modifier)
'                If .exp_variable < 0 Then
'                  evtmp = Log10(evtmp)
'                End If
'                evtmp = evtmp ^ .exp_exponent
'              Else
'                evtmp = 1#
'              End If
'              evtmp = evtmp * CDbl(.base_exponent)
'              'determine entire base variable's value
'              If .base_variable > 0 Then
'                Index = .base_variable - 1
'                If Index < 0 Then 'use area term
'                  bvtmp = CDbl(total_area)
'                Else 'use value input from compute window
'                  bvtmp = CDbl(rval(Index))
'                End If
'              Else
'                If .base_variable < 0 Then
'                  If .base_variable = -1 Then
'                    bvtmp = CDbl(total_area)
'                  ElseIf .base_variable = -2 Then
'                    bvtmp = GetRuralDischarge(rurind, urbscn(urbind).intrvl(i))
'                  End If
'                Else
'                  bvtmp = 0#
'                End If
'              End If
'              bvtmp = Abs(bvtmp + .base_modifier)
'              bvtmp = bvtmp * CDbl(.base_multiplier)
'              'determine discharge value
'              urban_discharge(i, urbind) = urban_discharge(i, urbind) * bvtmp ^ evtmp
'            End With
'          Next j
'          ludis(i) = urban_discharge(i, urbind)
'        Else
'          'no components, so force discharge to zero
'          ludis(i) = 0#
'          If i = NumIntrvl - 1 Then
'            'if we're dealing w/ Q500...extrapolate Q500
'            Call nss500(ludis())
'          End If
'          urban_discharge(i, urbind) = ludis(i)
'        End If
'      End If
'    Next i
'CompErr:
'  Err.Clear
'  utmp = 0
'  Resume Next
'End Sub
'
'Sub DispRuralDis()
'
'  'display rural discharge results
'  Dim i&, j&, Index&
'  Dim lstr$
'
'  If rurcnt < 1 Or rurind < 0 Then
'    InitRuralDis
'  ElseIf rural_discharge(0, 0, rurind) <= 0 Then
'    InitRuralDis
'  ElseIf rurcnt > 0 And rurind >= 0 Then
'    Call SetIntervals
'    compfg = 1
'    frmNSS.lstRurRes.Clear
'    'total_area = rurscn(rurind).tarea
'    For i = 0 To NumIntrvl - 1
'      'display discharge values for each interval
'      lstr = NumFmted(Interval(i), 7, 1)
'      If rural_discharge(0, i, rurind) > 0 Then
'        'valid rural discharge was calculated
'        lstr = lstr & NumFmted(Signif(CDbl(rural_discharge(0, i, rurind)), metric), 10, 0)
'        If rural_discharge(1, i, rurind) <> 0 Then
'          lstr = lstr & NumFmted(rural_discharge(1, i, rurind), 8, 1)
'        Else
'          lstr = lstr & "   N/A  "
'        End If
'        If rural_discharge(2, i, rurind) <> 0 And rural_discharge(1, i, rurind) <> 0 Then
'          lstr = lstr & NumFmted(rural_discharge(2, i, rurind), 8, 1)
'        Else
'          lstr = lstr & "   N/A  "
'        End If
'      End If
'      frmNSS.lstRurRes.AddItem lstr
'    Next i
'    DispMaxFloodEnvelope total_area, rurscn(rurind).regCribu
'  End If
'End Sub
'
''  abstract:    envelope-curve value of max flood discharge for given
''               drainage area and flood region
''
''  input arguments      : drainage_area - drainage area of basin
''                         flood_region - US flood region
''
''  subfunction(s) called: clrbox(), locate()
''
''  libfunction(s) called: fopen(), disp_printf(), scanf()
''
'
'Sub DispMaxFloodEnvelope(ByVal drainage_area#, ByVal flood_region&)
'  Static k1!(17), k2!(17), k3!(17)
'  Dim tmp#, envelope#, i& ' i = index in k arrays (flood_region-1)
'  If drainage_area > 0 And flood_region >= 1 And flood_region <= 17 Then
'    If k1(0) <> 23200! Then
'      k1(0) = 23200!
'      k1(1) = 28000!
'      k1(2) = 54400!
'      k1(3) = 42600!
'      k1(4) = 121000!
'      k1(5) = 70500!
'      k1(6) = 49100!
'      k1(7) = 43800!
'      k1(8) = 75000!
'      k1(9) = 62500!
'      k1(10) = 40800!
'      k1(11) = 89900!
'      k1(12) = 64500!
'      k1(13) = 10000!
'      k1(14) = 116000!
'      k1(15) = 98900!
'      k1(16) = 80500!
'
'      k2(0) = 0.895
'      k2(1) = 0.77
'      k2(2) = 0.924
'      k2(3) = 0.938
'      k2(4) = 0.838
'      k2(5) = 0.937
'      k2(6) = 0.883
'      k2(7) = 0.954
'      k2(8) = 0.849
'      k2(9) = 1.116
'      k2(10) = 0.919
'      k2(11) = 0.935
'      k2(12) = 0.873
'      k2(13) = 0.71
'      k2(14) = 1.059
'      k2(15) = 1.029
'      k2(16) = 1.024
'
'      k3(0) = -1.082
'      k3(1) = -0.897
'      k3(2) = -1.373
'      k3(3) = -1.327
'      k3(4) = -1.354
'      k3(5) = -1.297
'      k3(6) = -1.352
'      k3(7) = -1.357
'      k3(8) = -1.368
'      k3(9) = -1.371
'      k3(10) = -1.352
'      k3(11) = -1.304
'      k3(12) = -1.338
'      k3(13) = -0.844
'      k3(14) = -1.572
'      k3(15) = -1.341
'      k3(16) = -1.461
'    End If
'    tmp = 5# + drainage_area ^ 0.5
'    i = flood_region - 1
'    envelope = k1(i) * drainage_area ^ k2(i) * tmp ^ k3(i)
'
'    frmNSS.lstRurRes.AddItem "maximum envelope: " & NumFmted((envelope), 8, 1)
'  End If
'End Sub
'
'Sub DispUrbanDis()
'
'  'display urban discharge results
'  Dim i&, j&, Index&
'  Dim lstr$, tmpstr$
'  Dim standard_error As Single
'  Dim eqnMetric As Boolean
'
'  InitUrbanDis
'
'  If urbcnt < 1 Or urbind < 0 Then
'
'  ElseIf urban_discharge(0, urbind) <= 0 Then
'
'  Else
'    Call SetIntervals
'    For i = 0 To NumIntrvl - 1
'      If Abs(urban_discharge(i, urbind)) > 0.1 Then
'        'display discharge values for each interval
'        lstr = NumFmted(Interval(i), 7, 1)
'        If State.unitsys = 2 And Not urbscn(urbind).National Then eqnMetric = True Else eqnMetric = False
'        If metric And Not eqnMetric Then      'using metric and eqtn english, convert
'          lstr = lstr & NumFmted(Signif(CDbl(urban_discharge(i, urbind) * FLOW_CONVERSION), metric), 12, 2)
'        ElseIf Not metric And eqnMetric Then  'using english and eqtn metric, convert
'          lstr = lstr & NumFmted(Signif(CDbl(urban_discharge(i, urbind) / FLOW_CONVERSION), metric), 12, 0)
'        Else
'          lstr = lstr & NumFmted(Signif(CDbl(urban_discharge(i, urbind)), metric), 12, 0)
'        End If
'        If urbscn(urbind).National Then
'          standard_error = urban_eqtn(i).standard_error
'        Else
'          standard_error = State.Region(urbscn(urbind).reg).Interval(i).standard_error
'        End If
'        If standard_error <> 0 Then
'          lstr = lstr & NumFmted(standard_error, 8, 1)
'        Else
'          lstr = lstr & "   N/A  "
'        End If
'        frmNSS.lstUrbRes.AddItem lstr
'      End If
'    Next i
'  End If
'End Sub
'
'Sub DispEstimate()
'
'  'display the regions/parameter values for current estimate
'  Dim i&, j&
'  Dim rcnt& 'region count minus one
'  Dim vcnt&
'  Dim tstr$, lstr$, NL$, lst As ListBox
'  Dim Area!
'
'  Debug.Print "DispEstimate rurfg=" & rurfg
'  frmNSS.lblEstimates(1) = rurind + 1
'  frmNSS.lblEstimates(2) = "of " & rurcnt
'  frmNSS.lblEstimates(4) = urbind + 1
'  frmNSS.lblEstimates(5) = "of " & urbcnt
'
'  If rurfg And rurind < 0 Then
'    Debug.Print "DispEstimate Clear Rural"
'    frmNSS.lstRurEstimate.Clear
'    frmNSS.lstRurRes.Clear
'  ElseIf Not rurfg And urbind < 0 Then
'    Debug.Print "DispEstimate Clear Urban"
'    frmNSS.lstUrbEstimate.Clear
'    frmNSS.lstUrbRes.Clear
'  Else
'    NL = Chr(13) + Chr(10)
'
'    If rurfg Then
'      Set lst = frmNSS.lstRurEstimate
'      lst.Clear
'      'If rural_discharge(0, 0, rurind) > 0 Then 'calculations performed for this estimate, display results
'        Call DispRuralDis
'        If rural_discharge(0, 0, rurind) < 0.00001 Then compfg = 0 Else compfg = 1
'      'Else            'clear table
'       ' Call InitRuralDis
'        'compfg = 0
'      'End If
'      rcnt = rurscn(rurind).rcnt - 1
'      lst.AddItem "R" & rurind + 1 & ":    " & Trim(rurscn(rurind).Name)
'      'rurfg = False
'    Else
'      'If urbscn(urbind).rscn >= 0 Then
'        If rurind <> urbscn(urbind).rscn Then
'          Dim saveUrbind&
'          rurind = urbscn(urbind).rscn
'          rurfg = True
'          saveUrbind = urbind
'          DispEstimate
'          urbind = saveUrbind 'DispEstimate may have changed it
'          rurfg = False
'        End If
'      'End If
'      Set lst = frmNSS.lstUrbEstimate
'      lst.Clear
'      If urban_discharge(0, urbind) > 0 Then 'calculations performed for this estimate, display results
'        DispUrbanDis
'      Else        'clear table
'        InitUrbanDis
'      End If
'      rcnt = 0
'      lst.AddItem "U" & urbind + 1 & ":    " & Trim(urbscn(urbind).Name)
'    End If
'    tstr = ""
'    Area = 0
'    For i = 0 To rcnt
'      'summary line for each region in use
'      If rurfg Then
'        vcnt = rurscn(rurind).v(i).vcount - 1
'      Else
'        vcnt = urbscn(urbind).v.vcount - 1
'      End If
'      tstr = ""
'      If rurfg Then Area = Area + rurscn(rurind).v(i).Value(0)
'      For j = 0 To vcnt
'        If rurfg Then
'          lstr = Trim(rurscn(rurind).v(i).vname(j))
'          tstr = tstr & lstr & "=" & CStr(rurscn(rurind).v(i).Value(j))
'        Else
'          lstr = Trim(urbscn(urbind).v.vname(j))
'          If lstr = "A" Then Area = Area + urbscn(urbind).v.Value(j)
'          tstr = tstr & lstr & " = " & CStr(urbscn(urbind).v.Value(j))
'        End If
'        'pad with blanks between variables
'        tstr = tstr & "   "
'      Next j
'      lst.AddItem tstr
'    Next i
'    If Area > 0 Then
'      lst.AddItem "Total Area=" & Area
'      If rurfg Then total_area = Area
'    End If
'    If rurfg Then
'      If urbcnt > 0 Then
'        Dim HaveUrbanMatch As Boolean, urbMatch&
'        HaveUrbanMatch = False
'        urbMatch = -1
'        If urbind >= 0 Then
'          If urbscn(urbind).rscn = rurind Then HaveUrbanMatch = True: urbMatch = urbind
'        End If
'        While Not HaveUrbanMatch And urbMatch < urbcnt
'          urbMatch = urbMatch + 1
'          If urbscn(urbMatch).rscn = rurind Then HaveUrbanMatch = True
'        Wend
'        If Not HaveUrbanMatch Then urbMatch = -1
'        If urbMatch <> urbind Then
'          urbind = urbMatch
'          rurfg = False
'          DispEstimate
'        End If
'      End If
'    Else
'    End If
'  End If
'  frmNSS.SetCmdEnabled
'End Sub
'
'Function ecribu(drainage_area As Double, flood_region As Integer) As Single
'
'    Dim tmp#, envelope#
'    Static k1(17), k2(17), k3(17) As Single
'
'    k1(0) = 23200
'    k1(1) = 28000
'    k1(2) = 54400
'    k1(3) = 42600
'    k1(4) = 121000
'    k1(5) = 70500
'    k1(6) = 49100
'    k1(7) = 43800
'    k1(8) = 75000
'    k1(9) = 62500
'    k1(10) = 40800
'    k1(11) = 89900
'    k1(12) = 64500
'    k1(13) = 10000
'    k1(14) = 116000
'    k1(15) = 98900
'    k1(16) = 80500
'    k2(0) = 0.895
'    k2(1) = 0.77
'    k2(2) = 0.924
'    k2(3) = 0.938
'    k2(4) = 0.838
'    k2(5) = 0.937
'    k2(6) = 0.883
'    k2(7) = 0.954
'    k2(8) = 0.849
'    k2(9) = 1.116
'    k2(10) = 0.919
'    k2(11) = 0.935
'    k2(12) = 0.873
'    k2(13) = 0.71
'    k2(14) = 1.059
'    k2(15) = 1.029
'    k2(16) = 1.024
'    k3(0) = -1.082
'    k3(1) = -0.897
'    k3(2) = -1.373
'    k3(3) = -1.327
'    k3(4) = -1.354
'    k3(5) = -1.297
'    k3(6) = -1.352
'    k3(7) = -1.357
'    k3(8) = -1.368
'    k3(9) = -1.371
'    k3(10) = -1.352
'    k3(11) = -1.304
'    k3(12) = -1.338
'    k3(13) = -0.844
'    k3(14) = -1.572
'    k3(15) = -1.341
'    k3(16) = -1.461
'
'    'abstract:    envelope-curve value of max flood discharge for given
'    '             drainage area and flood region
'    'function return value: maximum envelope                                    *
'    If drainage_area <= 0# Then
'      ecribu = -1#
'    ElseIf flood_region < 1 Or flood_region > 17 Then
'      ecribu = -1#
'    Else
'      tmp = 5# + (drainage_area ^ 0.5)
'      envelope = k1(flood_region - 1) * (drainage_area ^ k2(flood_region - 1)) * (tmp ^ k3(flood_region - 1))
'      ecribu = envelope
'    End If
'
'End Function
'
'Sub GetStateData(istate&)
'
'    'read data values from database for selected state
'    Dim i&, j&, k&
'
'    'find starting offset in data file for this state
'    Seek #1, stpos(istate) + 1
'
'    'get state id and description
'    State.id_code = Str_Read()
'    State.descriptor = Trim(Str_Read())
'    Get #1, , State.regcount
'    Get #1, , State.unitsys
'    For i = 0 To State.regcount - 1
'      With State.Region(i)
'        .descriptor = Trim(Str_Read())
'        Get #1, , .v_count
'        Get #1, , .i_count
'        For j = 0 To .v_count - 1
'          .v_descriptor(j).variable_name = Trim(Str_Read())
'          .v_descriptor(j).descriptor = Trim(Str_Read())
'          Get #1, , .v_descriptor(j).minimum
'          Get #1, , .v_descriptor(j).maximum
'          Get #1, , .v_descriptor(j).units
'        Next j
'      End With
''      For j = 0 To MAX_INTERVALS
'      For j = 0 To State.Region(i).i_count - 1
'        With State.Region(i).Interval(j)
'          Get #1, , .intval
'          Get #1, , .standard_error
'          Get #1, , .eq_years_of_record
'          Get #1, , .regr_constant
'          Get #1, , .comp_count
'          If .comp_count <> 0 Then
'            For k = 0 To .comp_count - 1
'              Get #1, , .regr_component(k).base_variable
'              Get #1, , .regr_component(k).base_modifier
'              Get #1, , .regr_component(k).base_multiplier
'              Get #1, , .regr_component(k).base_exponent
'              Get #1, , .regr_component(k).exp_variable
'              Get #1, , .regr_component(k).exp_modifier
'              Get #1, , .regr_component(k).exp_exponent
'            Next k
'          End If
'        End With
'      Next j
'    Next i
'
'End Sub
'
''Sub GetStateNames()
''  Dim State As Variant
''
''  For Each State In States
''    frmNSS.cboState.AddItem State.Name
''  Next State
''
''    'read available state names from database
''    Dim st As String * ID_CODE_LEN, stnam$, tmp$
''    Dim cnt&
''    Open "state.bin" For Binary Access Read As #1
''    Open "state.ndx" For Binary Access Read As #2
''
''    On Error GoTo ErrHandler
''    cnt = 0
''    Do While Not EOF(2)
''      st = Input(ID_CODE_LEN, #2)
''      If st <> "    " Then
''        'state found, get pointer for data file
''        Get #2, , stpos(cnt)
''        Seek #1, stpos(cnt) + 4
''        stnam = ""
''        Do
''          tmp = Input$(1, #1)
''          If Asc(tmp$) = 0 Then
''            Exit Do
''          Else
''            stnam = stnam & tmp
''          End If
''        Loop
''        frmNSS.cboState.AddItem stnam
''        cnt = cnt + 1
''      End If
''    Loop
''ErrHandler:
''End Sub
'
'Sub InitRuralDis()
'
'  'init display table for rural discharge values
'  Dim i&
'  compfg = 0
'  frmNSS.lstRurRes.Clear
'  SetIntervals
'  'For i = 0 To NumIntrvl - 1      'display each interval
'  '  frmNSS.lstRurRes.AddItem NumFmtI(CInt(Interval(i)), 4)
'  'Next i
'
'End Sub
'
'Sub InitUrbanDis()
'
'  'init display table for urban discharge values
'  'Dim i&
'
'  frmNSS.lstUrbRes.Clear
'  SetIntervals
'  'For i = 0 To NumIntrvl - 1      'display each interval
'  '  frmNSS.lstUrbRes.AddItem NumFmtI(CInt(Interval(i)), 4)
'  'Next i
'
'End Sub
'
'Sub nss500(tempq() As Single)
'
'    'Perform extrapolation of state-equation flood freq. curves
'    'to 500 year flood along Log-Pearson Type III curves.
'
'    Dim zx!, wx#, bt0!, bt1!, bt2!
'    Dim npts&, errflg&
'    Static t(MAX_INTERVALS), z(MAX_INTERVALS) As Single
'    Static w(MAX_INTERVALS) As Single
'    Static ql(MAX_INTERVALS) As Single
'    Dim sku!, tempa!
'    Const one# = 1#
'    Const zero# = 0#
'    Static w2, w10, w100, qx, q(MAX_INTERVALS) As Double
'    Dim i, k As Integer
'
'    b0 = 0#
'    b1 = 0#
'    b2 = 0#
'
'    npts = -1
'    For i = 0 To NumIntrvl - 1
'      If Abs(tempq(i)) > 0.001 Then
'        npts = npts + 1
'        q(npts) = tempq(i)
'        t(npts) = Interval(i)
'      End If
'    Next i
'    If npts > 1 Then
'      For i = 0 To npts
'        ql(i) = Log10(q(i))
'        tempa = one / t(i)
'        z(i) = gausex(tempa)
'      Next i
'      'Fit quadratic to freq curve on log-probability coordinates
''      bt0 = CSng(b0)
''      bt1 = CSng(b1)
''      bt2 = CSng(b2)
'      Call slreg2(ql(), z(), npts + 1) ', bt0, bt1, bt2)
''      b0 = CDbl(bt0)
''      b1 = CDbl(bt1)
''      b2 = CDbl(bt2)
'
'      'Determine Pearson Type III skew 2, 10, 100-year points on curve
'      w2 = b0
'      w10 = nsspoly(CDbl(gausex(0.1)))
'      w100 = nsspoly(CDbl(gausex(0.01)))
'      sku = hartrg((w100 - w10) / (w10 - w2))
'
'      'Transform to Pearson/Wilfrt probability scale
'      For i = 0 To npts
'        w(i) = wilfrt(sku, z(i), errflg)
'      Next i
'
'      'Fit straight line to freq curve in log-Pearson coordinates */
''      bt0 = b0
''      bt1 = b1
'      Call slreg(ql(), w(), npts + 1) ' bt1, bt0)
''      b0 = bt0
''      b1 = bt1
'      b2 = zero
'
'      'Extrapolate straight line in log-Pearson-III coordinates                */
'      zx = gausex(1# / 500#)
'      wx = wilfrt(sku, zx, errflg)
'      qx = nsspoly(wx)
'      q(npts + 1) = 10# ^ qx
'      tempq(NumIntrvl - 1) = q(npts + 1)
'
'      'For i = 0 To NumIntrvl - 1
'      '  If Abs(tempq(i)) > 0.001 Then
'      '    tempq(i) = q(k)
'      '    k = k + 1
'      '  End If
'      'Next i
'    End If
'
'End Sub
'
'Function nsspoly(z As Double) As Double
'
'    nsspoly = b0 + z * (b1 + z * b2)
'
'End Function
'
'Sub OutReport()
'
'    'generate main output file
'    Dim i&, j&, k&, ErrRet&, ostr$, fname$, lint!
'    Dim tnumint&, allint!(2 * MAX_INTERVALS)
'    Dim fil% ' file handle
'    Dim standard_error As Single
'    Dim eqnMetric As Boolean
'
''    ChDir "\NSS"
''    ChDir "\NSSWIN"
'    On Error GoTo ErrHandler2
'    ErrRet = 0
'    'get name from user
'    With frmNSS.CMDialog1
'      .DialogTitle = "Save Report File"
'      .Filter = "Report Files (*.out)|*.out|All Files|*.*"
'      .FilterIndex = 0
'      .Flags = &H2&
'      .CancelError = True
'      .ShowSave
'      fname = .Filename
'    End With
'    ErrRet = 1
'    If FileLen(fname) > 0 Then
'      'get rid of existing file
'      Kill fname
'    End If
'    ErrRet = 0
'BackFromErr5:
'    'log output file (empty)
'    fil = FreeFile
'    Open fname For Output As fil
'
'    Print #fil, "National Flood Frequency Program"
'    Print #fil, "Version " & App.Major & "." & App.Minor '& "." & App.Revision
'    Print #fil,
'    Print #fil, "Based on Water-Resources Investigations Report 94-4002"
'    Print #fil, "Equations from NSS data base version 1.0"
'    Print #fil, "Equations developed using English units"
'    Print #fil,
'    Print #fil, "Project ID:  " & PrjID
'    Print #fil, "User ID:  " & UsrID
'    Print #fil, "Date: " & Date$ & " "; Time$
'    Print #fil,
'    ostr = Trim(plot_descriptor)
'    Print #fil, "Site Name:  " & ostr & ", " & State.descriptor
'    If metric Then
'      Print #fil, "Basin Drainage Area (km^2): " & NumFmted(total_area, 8, 0)
'    Else
'      Print #fil, "Basin Drainage Area (mi^2): " & NumFmted(total_area, 8, 0)
'    End If
'    Print #fil,
'    For k = 0 To rurcnt - 1
'      Print #fil, "Rural Estimate " & k + 1 & ":  " & rurscn(k).Name
'      For i = 0 To rurscn(k).rcnt - 1
'        Print #fil, "  Region: " & State.Region(rurscn(k).reg(i)).descriptor
'        Print #fil,
'        'We used to try printing out equation, but it didn't work well for all cases
'        'Print #fil, "    Interval  Equation"
'        'For j = 0 To rurscn(k).numint - 1
'        '  lint = rurscn(k).intrvl(j)
'        '  Call BldEqtn(rurscn(k).reg(i), lint, ostr)
'        '  Print #fil, NumFmted(lint, 8, 1), ostr
'        'Next j
'        'Print #fil,
'        Print #fil, "  Variable", "Value", "Units", "Definition"
'        'Print #fil, "  Variable    Value         Units                       Definition"
'        For j = 0 To State.Region(rurscn(k).reg(i)).v_count - 1
'          Print #fil, "   " & Left(State.Region(rurscn(k).reg(i)).v_descriptor(j).variable_name, 9), rurscn(k).v(i).Value(j),
'          If metric Then
'            ostr = units(State.Region(rurscn(k).reg(i)).v_descriptor(j).units).SI_sym
'          Else
'            ostr = units(State.Region(rurscn(k).reg(i)).v_descriptor(j).units).IP_sym
'          End If
'          Print #fil, ostr, State.Region(rurscn(k).reg(i)).v_descriptor(j).descriptor
'        Next j
'        Print #fil,
'      Next i
'    Next k
'    If urbcnt > 0 Then
'      'now do urban stuff
'      Print #fil,
'      For i = 0 To urbcnt - 1
'        Print #fil, "Urban Estimate " & i + 1 & ":  " & urbscn(i).Name
'        Print #fil,
'        If urbscn(i).National Then
'          Print #fil, "    Interval  Equation"
'          Print #fil, NumFmted(NatUrbEqInts(0), 8, 1), "2.35*A^.41*SL^.17*(RI2+3)^2.04*(ST+8)^-.65*(13-BDF)^-.32*IA^.15*RQ2^.47"
'          Print #fil, NumFmted(NatUrbEqInts(1), 8, 1), "2.70*A^.35*SL^.16*(RI2+3)^1.86*(ST+8)^-.59*(13-BDF)^-.31*IA^.11*RQ5^.54"
'          Print #fil, NumFmted(NatUrbEqInts(2), 8, 1), "2.99*A^.32*SL^.15*(RI2+3)^1.75*(ST+8)^-.57*(13-BDF)^-.30*IA^.09*RQ10^.58"
'          Print #fil, NumFmted(NatUrbEqInts(3), 8, 1), "2.78*A^.31*SL^.15*(RI2+3)^1.76*(ST+8)^-.55*(13-BDF)^-.29*IA^.07*RQ25^.6"
'          Print #fil, NumFmted(NatUrbEqInts(4), 8, 1), "2.67*A^.29*SL^.15*(RI2+3)^1.74*(ST+8)^-.53*(13-BDF)^-.28*IA^.06*RQ50^.62"
'          Print #fil, NumFmted(NatUrbEqInts(5), 8, 1), "2.5*A^.29*SL^.15*(RI2+3)^1.76*(ST+8)^-.52*(13-BDF)^-.28*IA^.06*RQ100^.63"
'          Print #fil, NumFmted(NatUrbEqInts(6), 8, 1), "2.27*A^.29*SL^.16*(RI2+3)^1.86*(ST+8)^-.54*(13-BDF)^-.27*IA^.05*RQ500^.63"
'        Else
'        'We used to try printing out equation, but it didn't work well for all cases
'        '  Print #fil, "    Interval  Equation*"
'        '  For j = 0 To urbscn(i).numint - 1
'        '    lint = urbscn(i).intrvl(j)
'        '    Call BldEqtn(urbscn(i).reg, lint, ostr)
'        '    Print #fil, NumFmted(lint, 8, 1), ostr
'        '  Next j
'        End If
'        Print #fil,
'        Print #fil, "  Variable", "Value", "Units", "Definition"
'        For j = 0 To urbscn(i).v.vcount - 1
'          Print #fil, "   " & Left(urbscn(i).v.vname(j), 9), urbscn(i).v.Value(j),
'          If urbscn(i).National Then
'            If metric Then
'              ostr = units(urban_comp(j).units).SI_sym
'            Else
'              ostr = units(urban_comp(j).units).IP_sym
'            End If
'            Print #fil, ostr, urban_comp(j).descriptor
'          Else
'            If metric Then
'              ostr = units(State.Region(urbscn(i).reg).v_descriptor(j + 1).units).SI_sym
'            Else
'              ostr = units(State.Region(urbscn(i).reg).v_descriptor(j + 1).units).IP_sym
'            End If
'            Print #fil, ostr, State.Region(urbscn(i).reg).v_descriptor(j + 1).descriptor
'          End If
'        Next j
'        If urbscn(i).National Then
'          If metric Then
'            Print #fil, "    A", total_area, units(1).SI_sym, "Basin Drainage Area"
'            Print #fil, "    RQT*", "-", units(14).SI_sym, "Peak Discharge for equivalent Rural basin"
'          Else
'            Print #fil, "    A", total_area, units(1).IP_sym, "Basin Drainage Area"
'            Print #fil, "    RQT*", "-", units(14).IP_sym, "Peak Discharge for equivalent Rural basin"
'          End If
'        End If
'        'Print #fil, "   *Area (A) term defined by Basin Drainage Area on main window"
'        If urbscn(i).National Then
'          Print #fil, "  *Flow (RQT) terms defined by corresponding Rural Flow values for the interval T"
'        End If
'        Print #fil,
'      Next i
'    End If
'    Print #fil,
'    If metric Then
'      Print #fil, "Flood Peak Discharges, in cubic meters per second"
'    Else
'      Print #fil, "Flood Peak Discharges, in cubic feet per second"
'    End If
'    Print #fil, "                     Recurrence  Peak       Standard  Equivalent"
'    Print #fil, "Estimate             Interval    Flow       Error %   Years of Record"
'    Print #fil, "___________________  __________  ________   ________  _______________"
'    'get all available intervals in Estimates
'    Call AllIntervals(tnumint, allint())
'    'output the rural results
'    For j = 0 To rurcnt - 1
'      Print #fil,
'      Print #fil, "R" & j + 1 & " " & rurscn(j).Name
'      k = 0
'      For i = 0 To tnumint - 1
'        If rurscn(j).intrvl(k) = allint(i) Then
'          ostr = "                        " & NumFmted(allint(i), 7, 1)
'          'valid interval for this Estimate
'          If rural_discharge(0, k, j) > 0.001 Then
'            ostr = ostr & NumFmted(Signif(CDbl(rural_discharge(0, k, j)), metric), 10, 0)
'          Else
'            ostr = ostr & "  N/A  "
'          End If
'          If rural_discharge(1, k, j) <> 0 Then
'            ostr = ostr & NumFmted(rural_discharge(1, k, j), 11, 1)
'          Else
'            ostr = ostr & "     N/A   "
'          End If
'          If rural_discharge(2, k, j) <> 0 And rural_discharge(1, k, j) <> 0 Then
'            ostr = ostr & NumFmted(rural_discharge(2, k, j), 8, 1)
'          Else
'            ostr = ostr & "   N/A  "
'          End If
'          Print #fil, ostr
'          k = k + 1
'        End If
'      Next i
'    Next j
'    'output any urban results
'    For j = 0 To urbcnt - 1
'      Print #fil,
'      Print #fil, "U" & j + 1 & " " & Left$(urbscn(j).Name, 28) & " "
'      k = 0
'      For i = 0 To tnumint - 1
'        If urbscn(j).intrvl(k) = allint(i) Then
'          ostr = "                        " & NumFmted(allint(i), 7, 1)
'          'valid interval for this Estimate
'          If Abs(urban_discharge(k, j)) > 0.001 Then
'            If State.unitsys = 2 And Not urbscn(j).National Then eqnMetric = True Else eqnMetric = False
'            If metric And Not eqnMetric Then      'using metric and eqtn english, convert
'              ostr = ostr & NumFmted(urban_discharge(k, j) * FLOW_CONVERSION, 10, 0)
'            ElseIf Not metric And eqnMetric Then  'using english and eqtn metric, convert
'              ostr = ostr & NumFmted(urban_discharge(k, j) / FLOW_CONVERSION, 10, 0)
'            Else 'using same units as equation
'              ostr = ostr & NumFmted(urban_discharge(k, j), 10, 0)
'            End If
'          Else
'            ostr = ostr & "  N/A  "
'          End If
'          If urbscn(j).National Then
'            standard_error = urban_eqtn(i).standard_error
'          Else
'            standard_error = State.Region(urbscn(j).reg).Interval(i).standard_error
'          End If
'          If standard_error > 0.0001 Then
'            ostr = ostr & NumFmted(standard_error, 11, 1)
'          Else
'            ostr = ostr & "   N/A  "
'          End If
'          Print #fil, ostr
'          k = k + 1
'        End If
'      Next i
'    Next j
'
'    Close fil
'BackFromErr4:
'    Exit Sub
'ErrHandler2:
'    If ErrRet = 0 Then
'      Resume BackFromErr4
'    Else
'      Resume BackFromErr5
'    End If
'    Resume BackFromErr4
'
'End Sub
'
'
'Sub StatusGet(opt&)
'
'    'get the state of the system from a status file
'    Dim buff$, fname$, fil%
'    Dim i&, j&, k&, ipos&, ist&
'    Dim ErrRet&, InitDone As Boolean
'
'    On Error GoTo ErrHandler1
'
'    If opt = 1 Then
'      'get name from user
'BackFromErr2:
'      ErrRet = 1
'      With frmNSS.CMDialog1
'        .DialogTitle = "Open Status File"
'        .Filter = "Status Files (*.sta)|*.sta|All Files|*.*"
'        .FilterIndex = 0
'        .CancelError = True
'        .ShowOpen
'        fname = .Filename
'      End With
'    Else
'      'use default
'      ErrRet = 2
'      fname = "NSS.STA"
'    End If
'    'open status file
'    fil = FreeFile
'    Open fname For Input As fil
'    rurcnt = 0
'
'    Do While Not EOF(fil)
'      Line Input #fil, buff
'      If Left$(buff, 6) = "METRIC" Then            'units flag
'        metric = Trim(Mid(buff, 8))
'      ElseIf Left$(buff, 5) = "BASIN" Then         'basin name
'        plot_descriptor = Mid(buff, 15, Len(buff) - 14)
'        frmNSS.txtBasin.Text = plot_descriptor
'      ElseIf Left$(buff, 7) = "ST/AREA" Then       'state and basin area
'        frmNSS.cboState.Tag = 1
'        stind = -1
'        frmNSS.cboState.ListIndex = CInt(Mid(buff, 15, 3))
'        total_area = CSng(Mid(buff, 29, 14))
''        frmNSS.txtBasinArea.value = total_area
'      ElseIf Left$(buff, 7) = "RSCN/NR" Then       'rural Estimate number and number of regions
'        rurcnt = CInt(Mid(buff, 15, 3))
'        k = rurcnt - 1
'        rurscn(k).rcnt = CInt(Mid(buff, 29, 3))
'        rurscn(k).Name = Mid(buff, 43)
'        j = -1
'      ElseIf Left$(buff, 7) = "NI/INTS" Then       'number of intervals and interval values
'        rurscn(k).numint = CInt(Mid(buff, 15, 3))
'        ipos = 15
'        For i = 0 To rurscn(k).numint - 1
'          ipos = ipos + 14
'          rurscn(k).intrvl(i) = CSng(Mid(buff, ipos, 10))
'        Next i
'      ElseIf Left$(buff, 6) = "REGSEL" Then        'selected region
'        j = j + 1
'        rurscn(k).reg(j) = CInt(Mid(buff, 15, 3))
'      ElseIf Left$(buff, 6) = "VALUES" Then        'region values
'        rurscn(k).v(j).vcount = State.Region(rurscn(k).reg(j)).v_count
'        ipos = 1
'        For i = 0 To rurscn(k).v(j).vcount - 1
'          rurscn(k).v(j).vname(i) = State.Region(rurscn(k).reg(j)).v_descriptor(i).variable_name
'          ipos = ipos + 14
'          rurscn(k).v(j).Value(i) = CSng(Mid(buff, ipos, 10))
'        Next i
'      ElseIf Left$(buff, 6) = "URBSCN" Then        'urban Estimate
'        urbcnt = CInt(Mid(buff, 15, 3))
'        k = urbcnt - 1
'        urbscn(k).rscn = CInt(Mid(buff, 29, 3))
'      ElseIf Left$(buff, 8) = "UNI/INTS" Then      'number of intervals and interval values
'        urbscn(k).numint = CInt(Mid(buff, 15, 3))
'        ipos = 15
'        For i = 0 To urbscn(k).numint - 1
'          ipos = ipos + 14
'          urbscn(k).intrvl(i) = CSng(Mid(buff, ipos, 10))
'        Next i
'      ElseIf Left$(buff, 10) = "URBSEL/CNT" Then   'urban equation id
'        urbscn(k).reg = CInt(Mid(buff, 15, 3))
'        If urbscn(k).reg < 0 Then
'          'urbscn(k).reg = 0
'          urbscn(k).National = True
'          urbscn(k).Name = "National Urban Equations (R" & urbscn(k).rscn + 1 & ")"
'        Else
'          urbscn(k).National = False
'          urbscn(k).Name = Trim(State.Region(urbscn(k).reg).descriptor)
'        End If
'        'count of parameters
'        urbscn(k).v.vcount = CInt(Mid(buff, 29, 3))
'      ElseIf Left$(buff, 7) = "UVALUES" Then       'urban equation parameter values
'        ipos = 1
'        For i = 0 To urbscn(k).v.vcount - 1
'          If urbscn(k).National Then    'national equation
'            urbscn(k).v.vname(i) = urban_comp(i).variable_name
'          Else                                 'state urban equation
'            urbscn(k).v.vname(i) = State.Region(urbscn(k).reg).v_descriptor(i).variable_name
'          End If
'          ipos = ipos + 14
'          urbscn(k).v.Value(i) = CDbl(Mid(buff, ipos, 10))
'        Next i
'      End If
'    Loop
'    Close fil 'close the status file
'BackFromErr1:
'    InitDone = True
'    CalculateAll
'    Exit Sub
'ErrHandler1:
'    Debug.Print Err.Description
'    If ErrRet = 2 Then Exit Sub
'
'    If Err <> 32755 Then
'      MsgBox "Problem Reading Status File " & fname & " " & Err
'    End If
'    If ErrRet = 1 Then
'      Resume BackFromErr1
'    Else
'      Resume BackFromErr2
'    End If
'
'End Sub
'
'Sub CalculateAll()
'  Dim confact(MAX_REGR_COMPONENTS, MAX_REG_USE) As Single
'  Dim total_basin!(2, MAX_INTERVALS)
'  Dim rtmp!(MAX_REGR_COMPONENTS)
'  Dim var&, regn&, intervl&
'  Dim eqnMetric As Boolean
'
'  If rurcnt > 0 Then 'Estimates exist, do calculations and set to first one
'    rurfg = True
'    Call SetIntervals
''    If State.unitsys = 2 Then eqnMetric = True Else eqnMetric = False
'     For rurind = 0 To rurcnt - 1
'      RegUseCnt = rurscn(rurind).rcnt
'      total_area = 0
'      For regn = 0 To RegUseCnt - 1
'        total_area = total_area + rurscn(rurind).v(regn).Value(0)
'      Next regn
'      For regn = 0 To RegUseCnt - 1
'        For var = 0 To State.Region(rurscn(rurind).reg(regn)).v_count - 1
'          If metric And Not eqnMetric Then      'using metric and eqtn english, convert
'            confact(var, regn) = units(State.Region(regn).v_descriptor(var).units).factor
'          ElseIf Not metric And eqnMetric Then
'            confact(var, regn) = 1 / units(State.Region(regn).v_descriptor(var).units).factor
'          Else
'            confact(var, regn) = 1
'          End If
'          'convert calculation parameters to correct units
'          rtmp(var) = ConvertVal(rurscn(rurind).v(regn).Value(var), 1 / confact(var, regn))
'        Next var
'        rtmp(0) = total_area / confact(0, regn)
'        Call CompRuralDis(rurscn(rurind).reg(regn), regn, rtmp())
'        For intervl = 0 To NumIntrvl - 1
'          If metric And Not eqnMetric Then      'using metric and eqtn english, convert
'            rural_discharge(0, intervl, rurind) = rural_discharge(0, intervl, rurind) * FLOW_CONVERSION
'          ElseIf Not metric And eqnMetric Then
'            rural_discharge(0, intervl, rurind) = rural_discharge(0, intervl, rurind) / FLOW_CONVERSION
'          End If
'          If total_area > 0 Then total_basin(0, intervl) = total_basin(0, intervl) + rural_discharge(0, intervl, rurind) * rurscn(rurind).v(regn).Value(0) / total_area
'          total_basin(1, intervl) = total_basin(1, intervl) + State.Region(rurscn(rurind).reg(regn)).Interval(intervl).standard_error
'          total_basin(2, intervl) = total_basin(2, intervl) + State.Region(rurscn(rurind).reg(regn)).Interval(intervl).eq_years_of_record
'        Next intervl
'      Next regn
'      For intervl = 0 To NumIntrvl - 1
'        rural_discharge(0, intervl, rurind) = total_basin(0, intervl)
'        'If nOutOfRange = 0 Then 'only display std error and equivalent years if vars are in range
'          rural_discharge(1, intervl, rurind) = total_basin(1, intervl) / RegUseCnt
'          rural_discharge(2, intervl, rurind) = total_basin(2, intervl) / RegUseCnt
'        'Else
'        '  rural_discharge(1, intervl, rurind) = 0
'        '  rural_discharge(2, intervl, rurind) = 0
'        'End If
'      Next intervl
'
'    Next rurind
'    rurind = 0
'    RegUseCnt = rurscn(rurind).rcnt
'    Call frmNSS.SetEstimate
'    'If rurcnt > 1 Then
'    '  'multiple Estimates, enable scroll bar
'    '  frmNSS.vsbScen.Max = rurcnt - 1
'    '  frmNSS.vsbScen.value = rurind
'    '  frmNSS.vsbScen.Enabled = True
'    'End If
'  Else
'    'no Estimates
'    frmNSS.lstRurEstimate.Clear
'    frmNSS.lstRurEstimate.AddItem NoEstimates
'    rurind = -1
'  End If
'
'
'  If urbcnt > 0 Then
'    Dim irg&
'    rurfg = False
'    Call SetIntervals
'    RegUseCnt = 1
'    For urbind = 0 To urbcnt - 1
'      If State.unitsys = 2 And Not urbscn(urbind).National Then eqnMetric = True Else eqnMetric = False
'      irg = urbscn(urbind).reg
'
'      For var = 0 To urbscn(urbind).v.vcount - 1
'        If metric And Not eqnMetric Then      'using metric and eqtn english, convert
'          confact(var, 0) = units(State.Region(regn).v_descriptor(var).units).factor
'        ElseIf Not metric And eqnMetric Then  'using english and eqtn metric, convert
'          confact(var, 0) = 1 / units(State.Region(regn).v_descriptor(var).units).factor
'        Else 'Added 9-10-99 not tested yet
'          confact(var, 0) = 1
'        End If
'        rtmp(var) = ConvertVal(urbscn(urbind).v.Value(var), 1 / confact(var, 0))
'      Next var
'      Call CompUrbanDis(urbscn(urbind).reg, rtmp())
'    Next urbind
'    urbind = 0 'set initial urban Estimate
'    Call DispUrbanDis
'    Call DispEstimate
'  Else
'    urbind = -1
'    frmNSS.lstUrbEstimate.Clear
'    frmNSS.lstUrbEstimate.AddItem NoEstimates
'  End If
'
'End Sub
'
'Sub StatusPut(opt&)
'
'    'output the state of the system to a status file
'    Dim i&, j&, k&, ipos&, ErrRet&, lnsub&
'    Dim fname$, updfname$, buff$
'    Dim fil% 'file handle
''    ChDir "\NSS"
''    ChDir "\NSSWIN"
'    On Error GoTo ErrHandler
'    ErrRet = 0
'    If opt = 1 Then
'      'get name from user
'      With frmNSS.CMDialog1
'        .DialogTitle = "Save Status File"
'        .Filter = "Status Files (*.sta)|*.sta|All Files|*.*"
'        .FilterIndex = 0
'        .CancelError = True
'        .ShowSave
'        fname = .Filename
'      End With
'    Else
'      'use default
'      fname = "NSS.STA"
'    End If
'
'    ErrRet = 1
'    If FileLen(fname) > 0 Then
'      'get rid of existing file
'      Kill fname
'    End If
'    ErrRet = 0
'BackFromErr3:
'    'log output file (empty)
'    fil = FreeFile
'    Open fname For Output As fil
'    'units flag
'    Print #fil, "METRIC:", metric
'    'basin name and area
'    Print #fil, "BASIN:", plot_descriptor
'    'state and area of basin
'    Print #fil, "ST/AREA:", frmNSS.cboState.ListIndex, total_area
'    For j = 0 To rurcnt - 1
'      'Estimate #, region count, Estimate name
'      Print #fil, "RSCN/NR: ", j + 1, rurscn(j).rcnt, rurscn(j).Name
'      'number of intervals and interval values
'      Print #fil, "NI/INTS:", rurscn(j).numint,
'      For i = 0 To rurscn(j).numint - 1
'        Print #fil, rurscn(j).intrvl(i),
'      Next i
'      Print #fil,
'      'region and its variable values
'      For i = 0 To rurscn(j).rcnt - 1
'        Print #fil, "REGSEL: ", rurscn(j).reg(i)
'        Print #fil, "VALUES: ",
'        For k = 0 To State.Region(rurscn(j).reg(i)).v_count - 1
'          Print #fil, rurscn(j).v(i).Value(k),
'        Next k
'        Print #fil,
'      Next i
'    Next j
'    For i = 0 To urbcnt - 1
'      'Estimate number, include rural Estimate used
'      Print #fil, "URBSCN: ", i + 1, urbscn(i).rscn
'      'number of intervals and interval values
'      Print #fil, "UNI/INTS:", urbscn(i).numint,
'      For j = 0 To urbscn(i).numint - 1
'        Print #fil, urbscn(i).intrvl(j),
'      Next j
'      Print #fil,
'      'urban equation used and number of variables
'      Print #fil, "URBSEL/CNT: ", urbscn(i).reg, urbscn(i).v.vcount
'      Print #fil, "UVALUES:",
'      For j = 0 To urbscn(i).v.vcount - 1
'        Print #fil, urbscn(i).v.Value(j),
'      Next j
'      Print #fil,
'    Next i
'    Close fil
'BackFromErr:
'    Exit Sub
'ErrHandler:
'    If ErrRet = 0 Then
'      Resume BackFromErr
'    Else
'      ErrRet = 0
'      Resume BackFromErr3
'    End If
'
'End Sub
'
'Public Sub slreg(Y!(), X!(), n&)
'
'    'simple linear (straight line) regression of Y on X.  WK 750212.
'    Dim i&, sy!, sx!, sxy!, sxx!, varx!
'    sy = 0#
'    sx = 0#
'    sxy = 0#
'    sxx = 0#
'    For i = 0 To n - 1
'      sy = sy + Y(i)
'      sx = sx + X(i)
'      sxy = sxy + X(i) * Y(i)
'      sxx = sxx + X(i) ^ 2
'    Next i
'    sy = sy / n
'    sx = sx / n
'    sxy = sxy / n
'    sxx = sxx / n
'    varx = sxx - sx ^ 2
'    b1 = 0#
'    If varx > 0 Then
'      b1 = (sxy - sx * sy) / varx
'    End If
'    b0 = sy - b1 * sx
'
'End Sub
'
'Public Sub slreg2(Y!(), X!(), n&)
'
'    'SLREG2 - simple quadratic regression of Y on X
'    '         Y = B0 + B1*X + B2*X**2
'    Dim i&, ybar!, xbar!, xi!, yi!
'    Dim sxx#, sx3#, sx4#, sxy#, sxxy#, d#
'
'    sxx = 0#
'    sx3 = 0#
'    sx4 = 0#
'    sxy = 0#
'    sxxy = 0#
'    ybar = 0#
'    xbar = 0#
'    For i = 0 To n - 1
'      xbar = xbar + X(i)
'      ybar = ybar + Y(i)
'    Next i
'    xbar = xbar / n
'    ybar = ybar / n
'    For i = 0 To n - 1
'      xi = X(i) - xbar
'      yi = Y(i) - ybar
'      sxx = sxx + xi ^ 2
'      sx3 = sx3 + xi ^ 3
'      sx4 = sx4 + xi ^ 4
'      sxy = sxy + xi * yi
'      sxxy = sxxy + yi * xi ^ 2
'    Next i
'    d = sx4 - sxx ^ 2 / n
'    b2 = (sxxy * sxx - sx3 * sxy) / (sxx * d - sx3 ^ 2)
'    b1 = (sxy - sx3 * b2) / sxx
'    b0 = -b2 * sxx / n
'    b0 = b0 + ybar - b1 * xbar + b2 * xbar ^ 2
'    b1 = b1 - b2 * xbar * 2#
'
'End Sub
'
'
'Public Function hartrg(r!)
'
'    'COMPUTES SKEW COEFF OF PEARSON TYPE III  DISTN,
'    'GIVEN THE RATIO (Q.100 - Q.10)/(Q.10 - Q.2),
'    'WHERE Q.T IS THE T-YEAR (1.-1/T - PROBABILITY)
'    'QUANTILE.   THE EQUATIONS WERE FOUND BY POLYNOMIAL
'    'REGRESSION, ETC., OF SKEW VS RATIO, WHERE THE RATIOS
'    'WERE LOOKED UP IN HARTER'S TABLES FOR GIVEN SKEWS.
'    '       WK 11/80.  FOR WRC BULL 17-B.
'
'    Dim rtmp!, rt1!
'
'    If r < 0.243 Then
'      If r > 0# Then
'        rtmp = -6# + 10# ^ (0.72609 + 0.15397 * Log10(CDbl(r)))
'      Else
'        rtmp = -4.8
'      End If
'    ElseIf r > 1.6 Then
'      rtmp = 7.1 + 1.6 * (r - 2.4) - 1.4 * ((r - 2.4) ^ 2 + 5.1888) ^ 0.5
'    Else
'      rt1 = 2.35713 + r * (-0.7387)
'      rtmp = -2.51898 + r * (3.82069 + r * (-2.3196 + r * (rt1)))
'    End If
'    hartrg = rtmp
'
'End Function
'
'
'
'Public Function wilfrt(sku!, zeta!, errflg&) As Single
'
'    'WILFRT -- WILSON-HILFERTY REVISED TRANSFORM
'    'PURPOSE -- APPROXIMATE TRANSFORMATION OF GAUSSIAN PERCENTAGE POINT
'    '   INTO STANDARDIZED PEARSON TYPE III.   THIS VERSION REPRODUCES
'    '   CORRECT MEAN, VARIANCE, SKEW AND LOWER BOUND OF STANDARDIZED
'    '   PEARSON-III AT SKEWS UP TO 9.0 AT LEAST.  DIFFERENCES BETWEEN
'    '   WILFRT PERCENTAGE POINTS AND HARTERS TABLES ARE OF THE ORDER OF
'    '   A FEW HUNDREDTHS OF A STD. DEVIATION, EXCEPT IN EXTREME POSITIV
'    '   TAIL (95% OR SO) WHERE ERROR IS OF ORDER OF TENTHS IN MAGNI-
'    '   TUDE BUT ABOUT 3% IN RELATIVE MAG.
'    'USAGE --      X=WILFRT(SKEW,ZETA)*STDDEV+AMEAN
'    '   SKEW IS INPUT SKEW, MAY BE ZERO OR NEGATIVE OR POSITIVE.
'    '         IF ABS(SKEW) IS GREATER THAN 9.75, 9.75 IS USED.
'    '   ZETA IS STANDARD GAUSSIAN VARIATE.   FOR EXAMPLE, GAUSSB(IRAN)
'    '         YIELDS RANDOM NOS WHILE GAUSAB(PROB) YIELDS THE
'    '         PROB-TH QUANTILE.
'    '   STDDEV AND AMEAN ARE DESIRED VALUES OF STD DEVIATION AND
'    '         MEAN, IF DIFFERENT FROM ONE AND ZERO.
'    'NOTE -- EACH INPUT SKEW VALUE IS COMPARED WITH PREVIOUS INPUT
'    '   VALUE. IF DIFFERENT BY MORE THAN 0.0003, TABLE LOOKUP OF NEW
'    '   PARAMETERS TAKES PLACE.  THEREFORE, CHAGE THE INPUT SKEW
'    '   AS SELDOM AS POSSIBLE.
'    'WKIRBY  72-02-25
'    'REVISED 73-02-09   TO ACCEPT ZERO SKEW.
'    'REF -- W.KIRBY, COMPUTER-ORIENTED WILSON-HILFERTY TRANSFORMATION..
'    '    WATER RESOUR RESCH 8(5)1251-4, OCT 72.
'    'REV 6/83 WK FOR PRIME ---- SAVE STTMNT ----
'    'REV 7/86 BY AML TO OSW CODING CONVENTION
'    'rev 9/96 by PRH for VB
'
'    Dim ask!, a!, b!, G!, h!, z!, sig!, fmu!
'    Const skutol = 0.0003
'
'    'first time thru or new sku (skew)
'    ask = Abs(sku)
'    If ask >= skutol Then
'      'nonzero skew
'      Call wilfrs(ask, G, h, a, b, errflg)
'      sig = G * 0.1666667
'      fmu = 1# - sig * sig
'      If sku < 0# Then
'        sig = -sig
'        a = -a
'      End If
'      z = fmu + sig * zeta
'      If z < h Then
'        z = h
'      End If
'      wilfrt = a * (z * z * z - b)
'    Else
'      'zero skew
'      wilfrt = zeta
'    End If
'
'End Function
'
'Public Sub wilfrs(sk!, G!, h!, a!, b!, errflg&)
'
'    'COMPUTES PARAMETERS USED BY WILFRT TRANSFORMATIN
'    'USES APPROX FORMULA AND CORRECTION TERMS PREPARED FROM
'    'ROUTINE WHMPP (E443-5).  WKIRBY  FEB72
'    'PARAMETERS RETURNED TO WILFRT ARE INTENDED TO MAKE
'    'WILFRT A STANDARDIZED R.V.  (MEAN=0,STDEV=1) WITH
'    'SPECIFIED SKEW AND CORRECT LOWER BOUND
'    'REVISED CALC OF CORRECTION TABLE  72-03-03 WK
'
'    Const nroz = 40
'    Dim flag&, i&, k&
'    Dim row!(1 To 4), table!(1 To 40, 1 To 4)
'    Dim s!, q!, p!, tog!
'
'    table(1, 1) = 0#
'    table(2, 1) = 0.25
'    table(3, 1) = 0.5
'    table(4, 1) = 0.75
'    table(5, 1) = 1#
'    table(6, 1) = 1.25
'    table(7, 1) = 1.5
'    table(8, 1) = 1.75
'    table(9, 1) = 2#
'    table(10, 1) = 2.25
'    table(11, 1) = 2.5
'    table(12, 1) = 2.75
'    table(13, 1) = 3#
'    table(14, 1) = 3.25
'    table(15, 1) = 3.5
'    table(16, 1) = 3.75
'    table(17, 1) = 4#
'    table(18, 1) = 4.25
'    table(19, 1) = 4.5
'    table(20, 1) = 4.75
'    table(21, 1) = 5#
'    table(22, 1) = 5.25
'    table(23, 1) = 5.5
'    table(24, 1) = 5.75
'    table(25, 1) = 6#
'    table(26, 1) = 6.25
'    table(27, 1) = 6.5
'    table(28, 1) = 6.75
'    table(29, 1) = 7#
'    table(30, 1) = 7.25
'    table(31, 1) = 7.5
'    table(32, 1) = 7.75
'    table(33, 1) = 8#
'    table(34, 1) = 8.25
'    table(35, 1) = 8.5
'    table(36, 1) = 8.75
'    table(37, 1) = 9#
'    table(38, 1) = 9.25
'    table(39, 1) = 9.5
'    table(40, 1) = 9.75
'    table(1, 2) = 0#
'    table(2, 2) = -0.000144
'    table(3, 2) = -0.001137
'    table(4, 2) = -0.003762
'    table(5, 2) = -0.008674
'    table(6, 2) = -0.011555
'    table(7, 2) = -0.010076
'    table(8, 2) = -0.006049
'    table(9, 2) = -0.000921
'    table(10, 2) = 0.004189
'    table(11, 2) = 0.008515
'    table(12, 2) = 0.011584
'    table(13, 2) = 0.013139
'    table(14, 2) = 0.013122
'    table(15, 2) = 0.010945
'    table(16, 2) = 0.007546
'    table(17, 2) = 0.002767
'    table(18, 2) = -0.003181
'    table(19, 2) = -0.010089
'    table(20, 2) = -0.017528
'    table(21, 2) = -0.025476
'    table(22, 2) = -0.033609
'    table(23, 2) = -0.042434
'    table(24, 2) = -0.050525
'    table(25, 2) = -0.058192
'    table(26, 2) = -0.065221
'    table(27, 2) = -0.07141
'    table(28, 2) = -0.076638
'    table(29, 2) = -0.080655
'    table(30, 2) = -0.083349
'    table(31, 2) = -0.084584
'    table(32, 2) = -0.084203
'    table(33, 2) = -0.082089
'    table(34, 2) = -0.078126
'    table(35, 2) = -0.072165
'    table(36, 2) = -0.064188
'    table(37, 2) = -0.054059
'    table(38, 2) = -0.041633
'    table(39, 2) = -0.027005
'    table(40, 2) = -0.010188
'    table(1, 3) = 0#
'    table(2, 3) = 0.004614
'    table(3, 3) = 0.009159
'    table(4, 3) = 0.013553
'    table(5, 3) = 0.017753
'    table(6, 3) = 0.021764
'    table(7, 3) = 0.025834
'    table(8, 3) = 0.030406
'    table(9, 3) = 0.03571
'    table(10, 3) = 0.04173
'    table(11, 3) = 0.048321
'    table(12, 3) = 0.055309
'    table(13, 3) = 0.062538
'    table(14, 3) = 0.069873
'    table(15, 3) = 0.077334
'    table(16, 3) = 0.084682
'    table(17, 3) = 0.091926
'    table(18, 3) = 0.099028
'    table(19, 3) = 0.105967
'    table(20, 3) = 0.112695
'    table(21, 3) = 0.119245
'    table(22, 3) = 0.106551
'    table(23, 3) = 0.095488
'    table(24, 3) = 0.085671
'    table(25, 3) = 0.07699
'    table(26, 3) = 0.06929
'    table(27, 3) = 0.062443
'    table(28, 3) = 0.056349
'    table(29, 3) = 0.050908
'    table(30, 3) = 0.046047
'    table(31, 3) = 0.041702
'    table(32, 3) = 0.037815
'    table(33, 3) = 0.034339
'    table(34, 3) = 0.031229
'    table(35, 3) = 0.028445
'    table(36, 3) = 0.025964
'    table(37, 3) = 0.023753
'    table(38, 3) = 0.021782
'    table(39, 3) = 0.020043
'    table(40, 3) = 0.018528
'    table(1, 4) = 0#
'    table(2, 4) = 0#
'    table(3, 4) = -0.000001
'    table(4, 4) = -0.000004
'    table(5, 4) = -0.000021
'    table(6, 4) = -0.000075
'    table(7, 4) = -0.00019
'    table(8, 4) = -0.000326
'    table(9, 4) = -0.000317
'    table(10, 4) = 0.000116
'    table(11, 4) = 0.000434
'    table(12, 4) = 0.000116
'    table(13, 4) = -0.000464
'    table(14, 4) = -0.000981
'    table(15, 4) = -0.001165
'    table(16, 4) = -0.000743
'    table(17, 4) = 0.000435
'    table(18, 4) = 0.002479
'    table(19, 4) = 0.005462
'    table(20, 4) = 0.009353
'    table(21, 4) = 0.014206
'    table(22, 4) = 0.019964
'    table(23, 4) = 0.026829
'    table(24, 4) = 0.034307
'    table(25, 4) = 0.042495
'    table(26, 4) = 0.051293
'    table(27, 4) = 0.060593
'    table(28, 4) = 0.070324
'    table(29, 4) = 0.080332
'    table(30, 4) = 0.090532
'    table(31, 4) = 0.100831
'    table(32, 4) = 0.111114
'    table(33, 4) = 0.121283
'    table(34, 4) = 0.131245
'    table(35, 4) = 0.140853
'    table(36, 4) = 0.15012
'    table(37, 4) = 0.158901
'    table(38, 4) = 0.167085
'    table(39, 4) = 0.174721
'    table(40, 4) = 0.181994
'
'    s = sk
'    k = 1
'    flag = 0
'    errflg = 0
'    i = 1
'    Do
'      i = i + 1
'      If table(i, 1) > s Then flag = 1
'      k = i - 1
'    Loop While i < nroz And flag = 0
'
'    If flag = 0 Then
'      errflg = 1
'      For i = 1 To 4
'        row(i) = table(nroz, i)
'      Next i
''     replace "row" equivalence
''      s = table(nroz, 1)
''      q = table(nroz, 2)
''      p = table(nroz, 3)
''      tog = table(nroz, 4)
'    Else
'      p = (s - table(k, 1)) / (table(k + 1, 1) - table(k, 1))
'      q = 1# - p
'      For i = 2 To 4
'        row(i) = q * table(k, i) + p * table(k + 1, i)
'      Next i
''     replace "row" equivalence
''      q = q * table(k, 2) + p * table(k + 1, 2)
''      p = q * table(k, 3) + p * table(k + 1, 3)
''      tog = q * table(k, 4) + p * table(k + 1, 4)
'    End If
'
'    G = s + row(2)
''   replace "row" equivalence
''    g = s + q
'    If s > 1# Then G = G - 0.063 * (s - 1#) ^ 1.85
'    tog = 2# / s
'    q = tog
'    If q < 0.4 Then q = 0.4
'    a = q + row(3)
''   replace "row" equivalence
''    a = q + p
'    q = 0.12 * (s - 2.25)
'    If q < 0# Then q = 0#
'    b = 1# + q * q + row(4)
''   replace "row" equivalence
''    b = 1# + q * q + tog
'    If (b - tog / a) < 0# Then
'      'Stop WILFRS
'      MsgBox "Very serious problem in routine WILFRS.  Contact software distributor", 16
'    End If
'    h = (b - tog / a) ^ 0.3333333
'
'End Sub
'
'Public Function gausex(exprob!) As Single
'
'    'GAUSSIAN PROBABILITY FUNCTIONS   W.KIRBY  JUNE 71
'       'GAUSEX=VALUE EXCEEDED WITH PROB EXPROB
'       'GAUSAB=VALUE (NOT EXCEEDED) WITH PROBCUMPROB
'       'GAUSCF=CUMULATIVE PROBABILITY FUNCTION
'       'GAUSDY=DENSITY FUNCTION
'    'SUBPGMS USED -- NONE
'
'    'GAUSCF MODIFIED 740906 WK -- REPLACED ERF FCN REF BY RATIONAL APPRX N
'    'ALSO REMOVED DOUBLE PRECISION FROM GAUSEX AND GAUSAB.
'    '76-05-04 WK -- TRAP UNDERFLOWS IN EXP IN GUASCF AND DY.
'
'    'rev 8/96 by PRH for VB
'
'    Const c0! = 2.515517
'    Const c1! = 0.802853
'    Const c2! = 0.010328
'    Const d1! = 1.432788
'    Const d2! = 0.189269
'    Const d3! = 0.001308
'    Dim pr!, rtmp!, p!, t!, numerat!, denom!
'
'    p = exprob
'    If p >= 1# Then
'      'set to minimum
'      rtmp = -10#
'    ElseIf p <= 0# Then
'      'set at maximum
'      rtmp = 10#
'    Else
'      'compute value
'      pr = p
'      If p > 0.5 Then pr = 1# - pr
'      t = (-2# * Log(pr)) ^ 0.5
'      numerat = (c0 + t * (c1 + t * c2))
'      denom = (1# + t * (d1 + t * (d2 + t * d3)))
'      rtmp = t - numerat / denom
'      If p > 0.5 Then rtmp = -rtmp
'    End If
'    gausex = rtmp
'
'End Function
'
'Public Sub SetIntervals() 'set the recurrence intervals for current Estimate
'  Dim i&, j&, k&, l&, lcnt&
''Debug.Print "SetIntervals was:" & NumIntrvl
'  NumIntrvl = 0                'init number of intervals
'  If rurfg Then
'    If rurind < 0 Then
'      NumIntrvl = 0
'    Else
'      For i = 0 To rurscn(rurind).rcnt - 1
''        With State.Region(rurscn(rurind).reg(i))
'        With curState.Regions(rurscn(rurind).reg(i))
'          If .Returns.Count > NumIntrvl Then 'look through all intervals for this region
'            lcnt = .Returns.Count
'          Else                         'look through all current intervals
'            lcnt = NumIntrvl
'          End If
'          k = 0
''          For j = 0 To lcnt - 1
''            If i = 0 Then          'first time through, use this regions intervals
''              Interval(j) = .Returns(j + 1).Interval
''              NumIntrvl = .i_count
''            ElseIf Interval(j) < .Interval(k).intval Or .Interval(j).intval = 0 Then
''              'region's interval is larger than current (or 0), remove current interval value
''              For L = j To State.Region(rurscn(rurind).reg(i - 1)).i_count - 2
''                Interval(L) = Interval(L + 1)
''              Next L
''              'init last position to 0
''              Interval(State.Region(rurscn(rurind).reg(i - 1)).i_count - 1) = 0
''              NumIntrvl = NumIntrvl - 1 'decrement count of current intervals
''            ElseIf Interval(j) > .Interval(k).intval Then
''              'region's interval is smaller than current, skip it
''              k = k + 1
''            Else
''              k = k + 1
''            End If
''          Next j
'        End With
'      Next i
'    End If
'  ElseIf urbind >= 0 And urbcnt > 0 Then
'    If urbscn(urbind).National Then 'assign national urban equation intervals
'      NumIntrvl = 7
'      For j = 0 To NumIntrvl - 1
'        Interval(j) = NatUrbEqInts(j)
'      Next j
'    Else                            'assign state urban equation intervals
'      NumIntrvl = State.Region(urbscn(urbind).reg).i_count
'      For j = 0 To NumIntrvl - 1
'        Interval(j) = State.Region(urbscn(urbind).reg).Interval(j).intval
'      Next j
'    End If
'  End If
''Debug.Print "SetIntervals now:" & NumIntrvl
'End Sub
'
'Public Function ConvertVal!(ByVal val!, ByVal ConvFact!)
'  Dim tmp!
'  tmp = val
'  If ConvFact = 0.5555 Then tmp = tmp - 32 'temperature F to C
'  tmp = tmp * ConvFact
'  If Abs(ConvFact - (1 / 0.5555)) < 0.001 Then tmp = tmp + 32 'temperature C to F
'  ConvertVal = tmp
'End Function
