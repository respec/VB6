Attribute VB_Name = "modROI"
Option Explicit

Private UserRegressVars() As Single  'user-entered values, not including lat/long
Private RegVarCnt As Long       'number of elements used in regression
Private StaLats() As Single     'Station latitudes
Private StaLngs() As Single     'Station longitudes
Private Sum() As Single         'Used in calcing SimVarSDs
Private INDX() As Integer       'dim(StaCnt) ranks Distance Distance(INDX(1)) = shortest distance
                                'Vars(INDX(1), 13) = drainage area of closest station
Private Mcon As Variant         '2-d from StateMatrices table
Private Rhoc As Variant         '2-d from RHO table for NC, calculated for TX
Private CF As Variant           '2-d from ClimateFactor table
Private Const MaxSites& = 50    'maximum number of sites that can be selected
Private Const MaxInd& = 10      'maximum number of independent variables including the regression constant
Private AkColl As FastCollection 'of constants that apply to return periods
Private PkLabColl As FastCollection 'of labels for return periods
Private Cf2!, Cf25!, Cf100!     'climate factors calculated in subroutine CFX
Private Ak() As Single          'labels for return periods
Private PkLab() As String         'labels for return periods
Private Nsites As Long          'number of stations to be used in final flood frequency calculations
Private OutFile As Long         'used for dummy outfile in current stand-alone program

'Remaining variables are used in statistical calculations
Private Const alpha_area = 1.5
Private Yvar As Single, sig As Single, Alpha As Single, Theta As Single
Private Sta(MaxSites) As Single             'time sampling error for stations
Private X(MaxSites, MaxInd) As Single       'calculated x-coordinate for # of closest sites
Private Y(MaxSites, 1) As Single            'calculated y-coordinate for # of closest sites
Private Xt(MaxInd, MaxSites) As Single
Private XtXinv(MaxInd, MaxInd) As Single
Private Cov(MaxSites, MaxSites) As Single   'covariance array
Private Wt(MaxSites, MaxSites) As Single
Private e(MaxSites, 1) As Single
Private ET(1, MaxSites) As Single
Private c1(1, 1) As Single
Private Bhat(MaxInd, 1) As Single
Private Gamasq!, Atse!, Atscov!

Public Function ComputeROIdischarge(Incoming As nssScenario, EquivYears() As Double, _
                stdErrMinus() As Double, stdErrPlus() As Double, PredInts() As Double) As Double()
'Program to estimate flood frequency in North Carolina
  Dim i&, j&, k&, iStep&, iSave&, roiRegion&, dimDist&, dimCF&
  Dim RHO As Single, ss As Single, Years As Single, yhat As Single, ysav As Single
  Dim tSave As Single, sdbeta As Single, tbeta As Single
  Dim icall As Integer, jpeak As Integer, ireg As Integer
  Dim sql$, str$, thisRetPd$
  Dim pMetric As Boolean
  Dim roiRegions() As Long
  Dim retval() As Double
  Dim Correlation() As Single, CorrelationLimit As Single
  Dim uParm As userParameter, vParm As Variant, tmpParm As nssParameter
  Dim uRegion As userRegion
  Dim myStation As ssStation
  Dim myDB As nssDatabase
  Dim Scenario As nssScenario
  
  Dim Lat As Double           'Ungaged station latitude
  Dim Lng As Double           'Ungaged station longitude
  Dim StaName$                'ungaged station name
  Dim RegName$                'name of user-selected region
  Dim UserSimVars() As Single 'ungaged station values used in similarity calcs
  Dim NumPeaks As Long        'number of return periods (NC=8, TX=6)
  Dim StaCnt As Long          'number of stations in state
  Dim SimVarCnt As Long       'number of parameters used in similarity calcs
  Dim StaIDs() As String      'Station ID numbers
  Dim RegVars() As Single     'gaged station values used in regression analysis
  Dim SimVars() As Single     'gaged station values used in similarity calcs
  Dim SimVarSDs() As Single   'Standard Deviation of SimVars
  Dim Distance() As Single    'dim(StaCnt) "distance" measurement calculated to ungaged station
  Dim RegParms As FastCollection 'nssParameters
  Dim SimParms As FastCollection 'nssParameters
  Dim Flows() As Single       'Flow values for each station

  Set Scenario = Incoming
  Set uRegion = Scenario.UserRegions(1)
  '!!! FOR NOW ASSUME ALL ROI APPS ARE IN ENGLISH (resolves AR metric confusion)
  pMetric = False ' uRegion.Region.State.Metric
  StaName = Scenario.Name
  RegName = uRegion.Region.Name
  Set myDB = uRegion.Region.DB

  'Check whether enough stations to perform calc
  StaCnt = uRegion.Region.State.ROIStations.Count
  Nsites = uRegion.Region.State.ROISimStations
  If StaCnt < Nsites Then
    ssMessageBox "Not enough stations in this ROI region to perform calculation"
    Exit Function
  End If
  
  SimVarCnt = 0
  RegVarCnt = 0
  Set RegParms = New FastCollection 'nssParameters
  Set SimParms = New FastCollection 'nssParameters

  For Each vParm In uRegion.Region.ROIParameters
    Set tmpParm = vParm
    If tmpParm.SimulationVar Then
      SimParms.Add vParm, tmpParm.Abbrev
      SimVarCnt = SimVarCnt + 1
    End If
    If tmpParm.RegressionVar Then
      RegParms.Add vParm, tmpParm.Abbrev
      RegVarCnt = RegVarCnt + 1
    End If
  Next vParm
  RegVarCnt = RegVarCnt + 1
  
  If uRegion.Region.State.ROIClimateFactor Then SimVarCnt = SimVarCnt + 1
  If uRegion.Region.State.ROIDistance Then SimVarCnt = SimVarCnt + 1
  'Set dimensions for Climate Factor and Distance in SimVars array
  If uRegion.Region.State.ROIClimateFactor Then
    dimCF = SimVarCnt
    If uRegion.Region.State.ROIDistance Then dimDist = SimVarCnt - 1
  Else
    If uRegion.Region.State.ROIDistance Then dimDist = SimVarCnt
  End If
  
  'Initialize constants and labels
  Init
  
  'Set NumPeaks = # of peak-flow return periods
  NumPeaks = Scenario.Project.State.ROIPeakFlows.Count
  
  'Size following arrays to number of Peak Flows for this state
  ReDim retval(NumPeaks)      'contains output for NSS interface
  ReDim EquivYears(NumPeaks)  'equivalent years statistic
  ReDim stdErrMinus(NumPeaks) 'standard error of estimation (low)
  ReDim stdErrPlus(NumPeaks)  'standard error of estimation (high)
  ReDim PredInts(2, NumPeaks) 'low/high 90% prediction intervals
  ReDim PkLab(NumPeaks)
  ReDim Ak(NumPeaks)
  For i = 1 To NumPeaks
    j = InStr(Scenario.Project.State.ROIPeakFlows(i).Name, "_")
    If j > 0 Then
      str = Left(Scenario.Project.State.ROIPeakFlows(i).Name, j - 1)
      PkLab(i) = PkLabColl(str)
      Ak(i) = AkColl(str)
    End If
  Next i
  
  'Redimension arrays for user-entered variables
  ReDim UserRegressVars(1 To RegVarCnt)   'user-entered variables, except lat/long
  ReDim Correlation(1 To RegVarCnt)        'variable correlation types for backward step regression
  ReDim UserSimVars(1 To SimVarCnt)       'ungaged variables used in similarity calcs
  'Redimension arrays for station variables
  ReDim StaLats(1 To StaCnt)              'station latitudes
  ReDim StaLngs(1 To StaCnt)              'station longitudes
  ReDim StaIDs(1 To StaCnt)               'station ID numbers
  ReDim roiRegions(1 To StaCnt)           'station ID numbers
  ReDim Flows(1 To StaCnt, 1 To NumPeaks) 'flows at stations
  ReDim FlowSDs(1 To StaCnt)              'SD over range of flows at stations
  ReDim SimVars(1 To StaCnt, 1 To SimVarCnt)  'variables used for simulations
  ReDim RegVars(1 To StaCnt, 1 To RegVarCnt - 1) 'variables used for simulations
  ReDim Sum(0 To SimVarCnt)                   'running total of data for stat analysis
  ReDim SimVarSDs(1 To SimVarCnt)             'std dev of Vars() array
  ReDim Distance(1 To StaCnt)             'similarity b/t stations
  ReDim INDX(1 To StaCnt)                 'station similarity rankings

  roiRegion = uRegion.Region.ROIRegnID
  'Assign log of user-entered values to UserRegressVars/UserSimVars arrays
  i = 0
  For Each vParm In SimParms
    Set uParm = uRegion.UserParms(vParm.Name)
    i = i + 1
    UserSimVars(i) = Log10(uParm.GetValue(pMetric))
  Next
  i = 1
  UserRegressVars(i) = 1#
  For Each vParm In RegParms
    Set uParm = uRegion.UserParms(vParm.Name)
    i = i + 1
    UserRegressVars(i) = Log10(uParm.GetValue(pMetric))
  Next
  If uRegion.Region.State.ROIDistance Or uRegion.Region.State.ROIClimateFactor Then 'need lat/lng
    Lat = uRegion.UserParms("Latitude").GetValue(False)
    Lng = uRegion.UserParms("Longitude").GetValue(False)
  End If
  If uRegion.Region.State.ROIDistance Then 'distance to self is 0
    UserSimVars(dimDist) = 0
  End If
  
  'Read in State Matrix, Climate Factor arrays; and for NC, RHO matrix from DB
  Mcon = Scenario.Matrix  '~4 secs for NC, ~15 secs for TX
  If uRegion.Region.State.ROIClimateFactor Then CF = Scenario.CF

  'Read in station attributes from STATION/STATISTIC tables: ~5 secs for NC
  For i = 1 To StaCnt
    Set myStation = uRegion.Region.State.ROIStations(i)
    StaIDs(i) = myStation.ID
    StaLats(i) = myStation.Latitude
    StaLngs(i) = myStation.Longitude
    If myStation.Statistics.KeyExists("25") Then
      roiRegions(i) = myStation.Statistics("25").Value
    Else 'just assign to user region
      roiRegions(i) = roiRegion
    End If
    'Read in the peak-flow periods for this station
    j = 0
    For Each vParm In Scenario.Project.State.ROIPeakFlows
      j = j + 1
      Flows(i, j) = Log10(myStation.Statistics(CStr(vParm.ID)).Value)
    Next vParm
    j = 0
    'Read in current station's stat values used in similarity calcs
    For Each vParm In SimParms
      j = j + 1
      SimVars(i, j) = Log10(myStation.Statistics(CStr(vParm.LabelCode)).Value)
      'Keep running tally of vars for ensuing stat calcs
      Sum(j) = Sum(j) + SimVars(i, j)
    Next vParm
    j = 0
    'Read in current station's stat values used in regression analysis
    For Each vParm In RegParms
      j = j + 1
      RegVars(i, j) = Log10(myStation.Statistics(CStr(vParm.LabelCode)).Value)
    Next vParm
    If uRegion.Region.State.ROIDistance Then 'read in distance from site to station
      SimVars(i, dimDist) = TASKER_DISTANCE(Lat, StaLats(i), Lng, StaLngs(i))
      'Convert miles into kilometers if necessary
      If pMetric Then SimVars(i, dimDist) = SimVars(i, dimDist) * 1.609344
      Sum(dimDist) = Sum(dimDist) + SimVars(i, dimDist)
    End If
    If uRegion.Region.State.ROIClimateFactor Then 'read in climate factors
      SimVars(i, dimCF) = Log10(myStation.Statistics("68").Value)
      Sum(dimCF) = Sum(dimCF) + SimVars(i, dimCF)
    End If
    'Read in SD across flows for each station, or calc if not stored
    If myStation.Statistics.IndexFromKey("227") > 0 Then 'SD is in DB
      FlowSDs(i) = myStation.Statistics("227").Value
    Else
      'Calc estimate of SD of peaks if not kept on DB
      FlowSDs(i) = (Flows(i, 6) - Flows(i, 1)) / Ak(6)
    End If
  Next i

  'Compute St. Dev. of independent variables used in similarity calcs
  For i = 1 To SimVarCnt
    SimVarSDs(i) = Sum(i) / StaCnt  'actually calcing avg of variable here
    Sum(i) = 0
    For j = 1 To StaCnt
      Sum(i) = Sum(i) + (SimVars(j, i) - SimVarSDs(i)) ^ 2
    Next j
    SimVarSDs(i) = (Sum(i) / (StaCnt - 1)) ^ 0.5
  Next i

  'Compute climate factor from lat and long coordinates
  icall = 0
  If uRegion.Region.State.ROIClimateFactor Then
    CFX icall, Lat, Lng  'calculates climate factor at user's site
    If Cf25 > 0 Then UserSimVars(dimCF) = Log10(CDbl(Cf25))
  End If
  icall = 1

  'Calculate similarities between ungaged station and gaged stations
  For i = 1 To StaCnt
    Distance(i) = 0
    For j = 1 To SimVarCnt
      Distance(i) = Distance(i) + ((UserSimVars(j) - SimVars(i, j)) / SimVarSDs(j)) ^ 2
    Next j
    Distance(i) = Distance(i) ^ 0.5
    'Discount stations in other ROI Regions
    Distance(i) = Distance(i) + Abs((roiRegion - roiRegions(i))) * 1000
  Next i

  'Rank distances
  INDEXX StaCnt, Distance(), INDX()

  ysav = 0#
  sig = 0#

  'open file for output
  OutFile = FreeFile
  Open "NSS_Output.txt" For Output As OutFile
  'write header to file
  str = "REGION OF INFLUENCE METHOD" & vbCrLf & vbCrLf & _
      "Flood frequency estimates for station [" & StaName & _
      "] in region " & RegName & ", " & uRegion.Region.State.Abbrev & vbCrLf & vbCrLf
  For Each vParm In uRegion.UserParms
    Set uParm = vParm
    str = str & uParm.Parameter.Name & ": " & uParm.GetValue(False) & vbCrLf
  Next vParm
  str = str & vbCrLf & "Data used for ROI Method:" & vbCrLf
  Print #OutFile, str

  str = "StaID" & vbTab _
        & "Dist" & vbTab _
        & "LAT" & vbTab _
        & "LNG" & vbTab
  For i = 1 To NumPeaks
    str = str & "LOG(" & Scenario.Project.State.ROIPeakFlows(i).Code & ")" & vbTab
  Next i
  For Each vParm In SimParms
    str = str & "LOG(" & vParm.Abbrev & ")" & vbTab
  Next vParm
  If uRegion.Region.State.ROIDistance Then 'write distance header
    str = str & "Distance" & vbTab
  End If
  If uRegion.Region.State.ROIClimateFactor Then 'write climate factor header
    str = str & "LOG(CF25)" & vbTab
  End If
  Print #OutFile, str
  
  'loop thru sites - calc avg StdDev for peak flows across all sites
  For i = 1 To Nsites
    'write values for this station to outfile
    str = StaIDs(INDX(i)) & vbTab & Distance(INDX(i)) & vbTab & StaLats(INDX(i)) & vbTab & StaLngs(INDX(i)) & vbTab
    For j = 1 To NumPeaks
      str = str & Format(Flows(INDX(i), j), "#0.0000") & vbTab
    Next j
    For j = 1 To SimVarCnt
      str = str & Format(SimVars(INDX(i), j), "####0.0000") & vbTab
    Next j
    Print #OutFile, str
    sig = sig + FlowSDs(INDX(i))
  Next i
  sig = sig / Nsites  ' = avg StDev of peak flows across all stations
  
  If RegParms.KeyExists("CONTDA") Then
    str = vbCrLf & "area = " & uRegion.UserParms("Contributing_Drainage_Area").GetValue(False)
  ElseIf RegParms.KeyExists("DRNAREA") Then
    str = vbCrLf & "area = " & uRegion.UserParms("Drainage_Area").GetValue(False)
  Else
    str = ""
  End If
  str = vbCrLf & "For " & StaName & str
  If uRegion.Region.State.ROIClimateFactor Then _
      str = str & "    : cf25 = " & Format(Cf25, "#0.00")
  str = str & vbCrLf & vbCrLf & "RI" & vbTab & " PREDICTED(cfs)" & _
        vbTab & "- SE (%)" & vbTab & "+ SE (%)" & vbTab & "90% PRED INT (cfs)" & vbCrLf
  Print #OutFile, str
        
  If Scenario.Project.State.Code = "48" Then 'build matrix on the fly for Texas
    BuildMatrix
  Else 'most states have their own matrix
    Rhoc = Scenario.RHO
  End If

  For jpeak = 1 To NumPeaks
    'Reset user-entered parameters in case of backward-step regression
    RegVarCnt = RegParms.Count + 1
    i = 1
    UserRegressVars(i) = 1#
    For Each vParm In RegParms
      Set uParm = uRegion.UserParms(vParm.Name)
      i = i + 1
      UserRegressVars(i) = Log10(uParm.GetValue(pMetric))
      Correlation(i) = vParm.CorrelationType
    Next
    'Reset station parameters in case of backward-step regression
    For i = 1 To Nsites
      'build the X matrix
      X(i, 1) = 1#
      For j = 2 To RegVarCnt
        X(i, j) = RegVars(INDX(i), j - 1)
      Next j
      'build the Xt-transpose matrix
      For j = 1 To RegVarCnt
        Xt(j, i) = X(i, j)
      Next j
    Next i
    
    'Compute regional average standard deviation for each return period
    Sum(0) = 0#
    ss = 0#
    For i = 1 To Nsites
      Y(i, 1) = Flows(INDX(i), jpeak)
      Sum(0) = Sum(0) + Y(i, 1)
      ss = ss + Y(i, 1) ^ 2
    Next i
    Yvar = (ss - Sum(0) ^ 2 / Nsites) / (Nsites - 1#)
    'Compute time sampling error, sta(i), for each site
    Atse = 0#
    Atscov = 0#
    For i = 1 To Nsites
      For j = 1 To i
        Years = Mcon(INDX(i), INDX(j)) _
                / (Mcon(INDX(i), INDX(i)) * Mcon(INDX(j), INDX(j)))
        If Scenario.Project.State.Code = "48" Then 'use run-time RHO matrix
          RHO = Rhoc(i, j)
        Else
          RHO = Rhoc(INDX(i), INDX(j))
        End If
        Cov(i, j) = RHO * sig ^ 2 * (1 + RHO * 0.5 * Ak(jpeak) ^ 2) * Years
        If Cov(i, j) < 0 Then
          ssMessageBox "ERROR--Cholesky Decomp requires >= 0", vbCritical, "Bad Value"
          Exit Function
        End If
        Cov(j, i) = Cov(i, j)
        If (i = j) Then
          Sta(i) = Cov(i, i)
          Atse = Atse + Sta(i)
        Else
          Atscov = Atscov + Cov(i, j)
        End If
      Next j
    Next i
    
    iStep = 0
regress:
    'do regression
    Secant
    If Scenario.Project.State.ROIRegress Then
      iStep = iStep + 1
      
      OutPut iStep, jpeak, yhat, OutFile, retval(jpeak), _
             stdErrMinus(jpeak), stdErrPlus(jpeak), EquivYears(jpeak), _
             PredInts(1, jpeak), PredInts(2, jpeak)
  
     'Compute T for each Beta in model
     '   do step-backward by dropping all variables with T < 2,
     '   which is a p-value of about 0.30 or smaller depending upon
     '   degrees of freedom on the T-distribution.  Tsave is used to insure
     '   that in the case where two or more variables have T < 2, then the
     '   variable having the smallest T is preferentially dropped first.
      tSave = 2#
      iSave = 100
      'Loop through all variables successive to drainage area
      For i = 2 To RegVarCnt  'check through variables for correlations
        If Correlation(i) <> 0 Then 'variable has a correlation of some kind
          sdbeta = XtXinv(i, i) ^ 0.5
          If Correlation(i) > 100 Then 'pos or negative correlation
            tbeta = Abs(Bhat(i, 1) / sdbeta)
            CorrelationLimit = Abs(Correlation(i) / 1000)
          Else
            If Correlation(i) < 0 Then
              tbeta = -Bhat(i, 1) / sdbeta
            ElseIf Correlation(i) > 0 Then
              tbeta = Bhat(i, 1) / sdbeta
            End If
            CorrelationLimit = Abs(Correlation(i))
          End If
          'Besides looking to drop Tbetas less than 2, we also want the smallest
          'Tbeta less than 2 for each step in the step-wise regression.
          If tbeta < CorrelationLimit And tbeta < tSave Then
            tSave = tbeta
            iSave = i
          End If
        End If
      Next i
  
      'Exclude variable with lowest Tbeta, if value is less than 2
      If iSave = RegVarCnt Then
        'want to drop last independent variable - ignore last column of X
        RegVarCnt = RegVarCnt - 1
        GoTo regress
      ElseIf iSave < RegVarCnt Then
        '# of independent variable decrements by one and shifts
        RegVarCnt = RegVarCnt - 1
        'remove dropped variable from X matrix and its transpose
        For i = 1 To Nsites
          For j = iSave To RegVarCnt
            X(i, j) = X(i, j + 1)
            Xt(j, i) = X(i, j)
          Next j
        Next i
        'Shift ungaged site parameters to account for dropped parameter
        For j = iSave To RegVarCnt
          UserRegressVars(j) = UserRegressVars(j + 1)
          Correlation(j) = Correlation(j + 1)
        Next j
        GoTo regress
      End If
    End If

    'output final regression summary
    iStep = 99
    OutPut iStep, jpeak, yhat, OutFile, retval(jpeak), _
           stdErrMinus(jpeak), stdErrPlus(jpeak), EquivYears(jpeak), _
           PredInts(1, jpeak), PredInts(2, jpeak)
    'check to see if predicted value is greater than previous prediction
    If (yhat < ysav) Then
      ssMessageBox "CAUTION: Predicted T-year flow is smaller" & vbCrLf & _
             "than T-Year flow with lower recurrence interval." & vbCrLf & _
             "See output."
    End If
    ysav = yhat
  Next jpeak
  Close OutFile
  ComputeROIdischarge = retval
End Function

Private Sub mltply(Prod() As Single, Xmat() As Single, Ymat() As Single, k1&, k2&, k3&, N1&, N2&, N3&)

  Dim i&, j&, k&
  Dim Sum!
' --------------------------------------------------------------
'  Xmat IS K1*K2 MATRIX
'  Ymat IS K2*K3 MATRIX
'  PROD = Xmat*Ymat IS A K1*K3 MATRIX
' --------------------------------------------------------------
  For i = 1 To k1
    For k = 1 To k3
      Sum = 0#
      For j = 1 To k2
        Sum = Sum + Xmat(i, j) * Ymat(j, k)
      Next j
      Prod(i, k) = Sum
    Next k
  Next i

End Sub

Private Sub invert(n&, Ndim&, det!, CovInv() As Single, Cov() As Single)

  Dim i&, im&, j&, k&
  Dim detl!, Sum!, temp!
  Dim B() As Single, A() As Single
  '--------------------------------------------------------------
  '  COV IS AN N*N MATRIX
  '  SUBROUTINE COMPUTES DETERMINANT OF COV AS COVINV
  '  B IS THE LOWER TRIANGULAR DECOMPOSITION OF COV
  '--------------------------------------------------------------
  ReDim B(n, n)
  ReDim A(n, n)
  If n = 2 Then  'dimension of array is only 2
    det = Cov(1, 1) * Cov(2, 2) - Cov(1, 2) ^ 2
    temp = Cov(1, 1) / det
    CovInv(1, 1) = Cov(2, 2) / det
    CovInv(2, 2) = temp
    CovInv(1, 2) = -Cov(1, 2) / det
    CovInv(2, 1) = CovInv(1, 2)
  Else
    decomp n, Ndim, Cov(), B()
    detl = B(1, 1)
    For i = 2 To n
      If detl > 5E+19 Then
        ssMessageBox "ERROR--Numerical overflow on B(i,i) product expansion series."
'        Stop
      End If
      detl = detl * B(i, i)
    Next i
   'Following if statement is a questionable fix
    If detl > 5E+19 Then
      ssMessageBox "ERR0R--Determinant is too large." & vbCrLf & _
             "Numerical overflow." & vbCrLf & _
             "Bad fix, but try fewer stations."
      Stop
    Else
      det = detl ^ 2
    End If
    
    A(1, 1) = 1# / B(1, 1)
    A(2, 2) = 1# / B(2, 2)
    A(2, 1) = -B(2, 1) * A(1, 1) * A(2, 2)

    For i = 3 To n
      A(i, i) = 1# / B(i, i)
      im = i - 1
      For k = 1 To im
        Sum = 0#
        For j = k To im
          Sum = Sum + B(i, j) * A(j, k)
        Next j
        A(i, k) = -Sum * A(i, i)
      Next k
    Next i

    For i = 1 To n
      For j = 1 To i
        Sum = 0#
        For k = i To n
          Sum = Sum + A(k, i) * A(k, j)
        Next k
        CovInv(i, j) = Sum
        CovInv(j, i) = Sum
      Next j
    Next i
  End If

End Sub

Private Sub decomp(n&, Ndim&, XLAM() As Single, B() As Single)

  Dim iis&, ism&, js&, jsm&, ks&
  Dim bh!, bn!
  '--------------------------------------------------------------
  ' CHOLESKY DECOMPOSITION  BB-TRANSPOSE = XLAM
  '--------------------------------------------------------------
  If XLAM(1, 1) <= 0# Or XLAM(2, 2) <= 0# Then
    ssMessageBox "IN DECOMP/ NDIM,XLAM 1-1,2-1,2-2,1-2 = " _
           & Ndim & XLAM(1, 1) & XLAM(2, 1) & XLAM(2, 2) & XLAM(1, 2) & vbCrLf & _
           " COVARIANCE MATRIX NOT POSITIVE DEFINITE"
  End If
  B(1, 1) = XLAM(1, 1) ^ 0.5
  B(1, 2) = 0#
  B(2, 1) = XLAM(2, 1) / B(1, 1)
  B(2, 2) = (XLAM(2, 2) - B(2, 1) ^ 2) ^ 0.5

  If n <= 2 Then '2x2 or 1x1 matrix, exit early
    Exit Sub
  End If
  'Main decomposition algorithm
  For iis = 3 To n
    B(iis, 1) = XLAM(iis, 1) / B(1, 1)
    bn = XLAM(iis, iis) - B(iis, 1) ^ 2
    ism = iis - 1
    For js = 2 To ism
      jsm = js - 1
      bh = XLAM(iis, js)
      For ks = 1 To jsm
        bh = bh - B(iis, ks) * B(js, ks)
      Next ks
      B(iis, js) = bh / B(js, js)             '*******OVERFLOW!!!!!!!!!!
      bn = bn - B(iis, js) ^ 2
    Next js
    If bn <= 0# Then
      ssMessageBox "COVARIANCE MATRIX NOT POSITIVE DEFINITE BN=" & bn
    End If
    'b(iis, iis) = (AMAX1(bn, 0#)) ^ 0.5
    If bn > 0 Then
      B(iis, iis) = bn ^ 0.5
    Else
      B(iis, iis) = 0
    End If
  Next iis

End Sub

Private Function STUTP(xx!, n&) As Single

  Dim A!, B!, T!, yy!, z!
  Dim j&, nn&, keeplooping%
  Const rhpi = 0.63661977
  'STUDENT T PROBABILITY
  'STUTP = PROB( STUDENT T WITH N DEG FR  .LT.  xx )
  'NOTE  -  PROB(ABS(T).GT.xx) = 2.*STUTP(-xx,N) (FOR xx .GT. 0.)
  'SUBPGM USED - GAUSCF
  'REF - G.W. HILL, ACM ALGOR 395, OCTOBER 1970.
  'USGS - WK 12/79.
  STUTP = 0.5
  If n < 1 Then Exit Function
  nn = n
  z = 1#
  T = xx ^ 2
  yy = T / nn
  B = 1# + yy
  If Not (nn >= 20 And T < nn Or nn > 200) Then
    If nn < 20 And T < 4# Then 'nested summation of cosine series
      yy = yy ^ 0.5
      A = yy
      If nn = 1 Then A = 0#
    Else 'tail series for large t
      A = B ^ 0.5
      yy = A * nn
      j = 0
      keeplooping = 1
      While keeplooping = 1
        j = j + 2
        If (A = z) Then
          nn = nn + 2
          z = 0#
          yy = 0#
          A = -A
          keeplooping = 0
        Else
          z = A
          yy = yy * (j - 1) / (B * j)
          A = A + yy / (nn + j)
        End If
      Wend
    End If
    keeplooping = 1
    While keeplooping = 1
      nn = nn - 2
      If nn <= 1 Then
        If nn = 0 Then A = A / (B ^ 0.5)
        If nn <> 0 Then A = (Atn(yy) + A / B) * rhpi
        STUTP = 0.5 * (z - A)
        If xx > 0# Then STUTP = 1# - STUTP
        Exit Function
      Else
        A = (nn - 1) / (B * nn) * A + yy
      End If
    Wend
  End If

  'asymptotic series for large or noninteger N
  If yy > 0.000001 Then yy = Log(B)
  A = nn - 0.5
  B = 48# * A ^ 2
  yy = A * yy
  yy = (((((-0.4 * yy - 3.3) * yy - 24#) * yy - 85.5) / _
      (0.8 * yy ^ 2 + 100# + B) + yy + 3#) / B + 1#) * yy ^ 0.5
  STUTP = gauscf(-yy)
  If xx > 0# Then STUTP = 1# - STUTP
End Function

Private Sub Secant()
'function WLS was passed as argument in original Fortran

  Dim f1!, f2!, f3!, fnew!, X1!, X2!, x3!, xnew!
  Dim i&, j&
 
  X1 = 0#
  x3 = 0#
  X2 = Yvar * 2#
  f2 = WLS(X2)
  f1 = WLS(X1)
  If f1 < 0# Then 'midpoint good starting point for search
    For j = 1 To 3
      xnew = (X1 + X2) / 2#
      fnew = WLS(xnew)
      If fnew < 0# Then
        X1 = xnew
        fnew = fnew
      Else
        X2 = xnew
        f2 = fnew
      End If
    Next j

    'Search for gama sq using secant search
    For i = 1 To 30
      If (f2 - f1) = 0# Then
        ssMessageBox "STATUS: F2-F1 = 0, about to divide by 0"
'        Stop
      End If
      x3 = X1 - f1 * (X2 - X1) / (f2 - f1)
      If x3 < 0# Then
        'x3 = AMIN1(X2, X1) / 2#
        If X1 < X2 Then
          x3 = X1 / 2#
        Else
          x3 = X2 / 2#
        End If
        f3 = WLS(x3)
      Else
        f3 = WLS(x3)
      End If
      If Abs(f3) < 0.0001 Then Exit For
      If Abs(f1) < Abs(f2) Then
        X2 = x3
        f2 = f3
      Else
        X1 = x3
        f1 = f3
      End If
    Next i
  End If
  Gamasq = x3

End Sub

Private Function WLS(GAMA2!) As Single
  'computes the regression coefficients and
  'the weighted SSE (Sum of Square Error)

  Dim det!, xtx!(MaxInd, MaxInd), work!(MaxInd, MaxSites), _
      work2!(MaxInd, MaxSites), work1!(1, MaxSites)
  Dim k&
  
  'Compute the weighting matrix Wt by first generating the estimated
  '  GLS covariance matrix, Cov(k,k) by updating the diagonal with the
  '  sum of the model error GAMA2 estimated from first call of SECANT and
  '  the station time sampling error, Sta(k).
  For k = 1 To Nsites
    Cov(k, k) = GAMA2 + Sta(k)
    If Cov(k, k) <= 0# Then
      ssMessageBox "Diagonal of variance-covariance matrix is <= 0" & vbCrLf & _
             "Location (k,k) of " & k & ", " & k
'      Stop
    End If
  Next k
  'Inverting the Variance-Covariance Matrix
  Call invert(Nsites, MaxSites, det, Wt(), Cov())
  'Multiply Xt-transpose by the inverted covariance array (Wt) and produce work
  Call mltply(work(), Xt(), Wt(), RegVarCnt, Nsites, Nsites, MaxInd, MaxInd, MaxSites)
  'Multiply Xt*Wt (work) by X to produce XtX
  Call mltply(xtx(), work(), X(), RegVarCnt, Nsites, RegVarCnt, MaxInd, MaxInd, MaxSites)

  'COMPUTE THE REGRESSION COEFFICIENTS
  'Invert the Xt*Wt-1*X matrix
  Call invert(RegVarCnt, MaxInd, det, XtXinv, xtx)
  'Multiply the inverted variance-covariance matrix by Xt
  Call mltply(work, XtXinv, Xt, RegVarCnt, RegVarCnt, Nsites, MaxInd, MaxInd, MaxInd)
  'Multiply the results of above the Wt matrix
  Call mltply(work2, work, Wt, RegVarCnt, Nsites, Nsites, MaxInd, MaxInd, MaxSites)
  'Estimate the coefficients Bhat
  Call mltply(Bhat, work2, Y, RegVarCnt, Nsites, 1, MaxInd, MaxInd, MaxSites)

  'ERROR ESTIMATION
  'Determine the Error matrix E from by first estimating E as
  '  product of the independent variables and the coefficients
  Call mltply(e, X, Bhat, Nsites, RegVarCnt, 1, MaxSites, MaxSites, MaxInd)
  'The fill the E matrix with the residuals and make the transpose
  For k = 1 To Nsites
    e(k, 1) = Y(k, 1) - e(k, 1)
    ET(1, k) = e(k, 1)
  Next k
  'Multiply Et by Wt to produce a vector of error weights
  Call mltply(work1, ET, Wt, 1, Nsites, Nsites, 1, 1, MaxSites)
  'Multiply work1 by the Errors to produce a single value of error
  Call mltply(c1, work1, e, 1, Nsites, 1, 1, 1, MaxSites)
  'WLS is the weighted SSE sum of square error
  WLS = (Nsites - RegVarCnt) / c1(1, 1) - 1#

End Function

Private Sub CFX(icall, Latitude As Double, Longitude As Double)
  Dim cfRec As Recordset
  Dim sql$
  Dim C2a(30, 26) As Single, C25a(30, 26) As Single, C100a(30, 26) As Single, _
      sx As Single, sy As Single
  Dim q As Double, q0 As Double, q1 As Double, q2 As Double
  Dim phi As Double, phi0 As Double, phi1 As Double, phi2 As Double
  Dim e As Double, e2 As Double
  Dim m1 As Double, m2 As Double, n As Double
  Dim RHO As Double, rho0 As Double
  Dim A As Double, c As Double
  Dim theta1 As Double
  Dim lam As Double, lam0 As Double
  Dim X As Double, Y As Double
  Dim xoff As Double, yoff As Double
  Dim dp As Double, mp As Double, sp As Double
  Dim dm As Double, mm As Double, sm As Double
  Dim one As Double, two As Double, tpi As Double
  Dim i As Integer, ix As Integer, iy As Integer, j As Integer
  
  dp = Latitude
  dm = Longitude
  
' subroutine computes climate factors from lat and long

   '  COMPUTES X,Y COORDINATES(KILOMETERS) BASED ON LAT AND LONG,
   '  USING ALBERS EQUAL-AREA CONIC PROJECTION--WITH EARTH AS ELLIPSOID.
   '  EQUATIONS FROM BULLETIN 1532, PAGES 96-99.
   '  A = a, equatorial radius of ellipsoid
   '  E2 = e*2, square of eccentricity of ellipsoid (0.006768658).
   '  PHI1,PHI2 = standard parallels of latitude (29.5,45.5).
   '  PHI0 = middle latitude, or latitude chosen as the origin of y-coordinate (
   '  38.0).
   '  LAM0 = central meridian of the map, or longitude chosen as the origin
   '         of x-coordinate (96.0).
   '  PHI = latitude of desired y-coordinate.
   '  LAM = longitude of desired x-coordinate.
   '  XOFF = offset to x-coordinate to make all positive for KRIGING.
   '  YOFF = offset to y-coordinate to make all positive for KRIGING.

  A = 6378206.4
  e2 = 0.006768658
  tpi = 6.2831853
  phi1 = 29.5
  phi2 = 45.5
  phi0 = 38#
  lam0 = 96#
  
  xoff = 1000#
  yoff = 1500#
  one = 1#
  two = 2#
  e = e2 ^ 0.5
  phi0 = phi0 * tpi / 360#
  phi1 = phi1 * tpi / 360#
  phi2 = phi2 * tpi / 360#
  q0 = (one - e2) * (Sin(phi0) / (one - e2 * Sin(phi0) * Sin(phi0)) _
      - (one / (two * e)) * Log((one - e * Sin(phi0)) / (one + e * Sin(phi0))))
  q1 = (one - e2) * (Sin(phi1) / (one - e2 * Sin(phi1) * Sin(phi1)) _
      - (one / (two * e)) * Log((one - e * Sin(phi1)) / (one + e * Sin(phi1))))
  q2 = (one - e2) * (Sin(phi2) / (one - e2 * Sin(phi2) * Sin(phi2)) _
      - (one / (two * e)) * Log((one - e * Sin(phi2)) / (one + e * Sin(phi2))))
  m1 = Cos(phi1) / (one - e2 * Sin(phi1) * Sin(phi1)) ^ 0.5
  m2 = Cos(phi2) / (one - e2 * Sin(phi2) * Sin(phi2)) ^ 0.5
  n = (m1 * m1 - m2 * m2) / (q2 - q1)
  c = m1 * m1 + n * q1
  rho0 = A * (c - n * q0) ^ 0.5 / n

  If icall = 0 Then 'init and read in Kriged climate factors
    For i = 1 To 30
      For j = 1 To 26
        C2a(i, j) = 0#
        C25a(i, j) = 0#
        C100a(i, j) = 0#
      Next j
    Next i
    For i = 1 To UBound(CF, 1)
      ix = CF(i, 1)
      iy = CF(i, 2)
      C2a(ix, iy) = CF(i, 3)
      C25a(ix, iy) = CF(i, 4)
      C100a(ix, iy) = CF(i, 5)
    Next i
  End If
  
  'dp = CDbl(StaLat(1))
  'mp = CDbl(StaLat(2))
  'sp = CDbl(StaLat(3))
  'dm = CDbl(StaLong(1))
  'mm = CDbl(StaLong(2))
  'sm = CDbl(StaLong(3))
  mp = mp + sp / 60#
  phi = (dp + mp / 60#) * tpi / 360#
  mm = mm + sm / 60#
  lam = (dm + mm / 60#)
  q = (one - e2) * (Sin(phi) / (one - e2 * Sin(phi) * Sin(phi)) - (one / (two * e)) _
     * Log((one - e * Sin(phi)) / (one + e * Sin(phi))))
  theta1 = n * (lam0 - lam) * tpi / 360#
  RHO = A * (c - n * q) ^ 0.5 / n
  X = RHO * Sin(theta1) / 1000# + xoff
  Y = (rho0 - RHO * Cos(theta1)) / 1000# + yoff
  sx = CSng(X)
  sy = CSng(Y)
  
  wgt sx, sy, C2a(), C25a(), C100a()
      
End Sub

Private Sub wgt(X As Single, Y As Single, C2a() As Single, C25a() As Single, C100a() As Single)
  Dim cp As Single, r1 As Single, r2 As Single, r3 As Single, _
      r4 As Single, rx As Single, ry As Single, wr1 As Single, wr2 As Single, _
      wr3 As Single, wr4 As Single, xr As Single, yr As Single
  Dim ix As Integer, ixp As Integer, iy As Integer, iyp As Integer
      
  wr1 = 0#
  wr2 = 0#
  wr3 = 0#
  wr4 = 0#
'  ESTIMATE CLIMATE FACTOR FROM INTERPOLATION OF KRIGED VALUES AT 4 POINTS
'  SURROUNDING X,Y LOCATION OF INTEREST.  HOWEVER, USE KRIGED VALUE IF
'  DISTANCE TO NEAREST POINT IS 10KM OR LESS.
  ix = Fix(X / 100#)
  rx = ix
  xr = X - 100# * rx
  iy = Fix(Y / 100#)
  ry = iy
  yr = Y - 100# * ry
  ix = ix + 1
  iy = iy + 1
  r1 = (xr ^ 2 + yr ^ 2) ^ 0.5
  If r1 <= 10# Then
    wr1 = 1#
    GoTo 100
  End If
  r2 = (xr ^ 2 + (100# - yr) ^ 2) ^ 0.5
  If r2 <= 10# Then
    wr2 = 1#
    GoTo 100
  End If
  r3 = ((100# - xr) ^ 2 + (100# - yr) ^ 2) ^ 0.5
  If r3 <= 10# Then
    wr3 = 1#
    GoTo 100
  End If
  r4 = ((100# - xr) ^ 2 + yr ^ 2) ^ 0.5
  If r4 <= 10# Then
    wr4 = 1#
    GoTo 100
  End If
  cp = 1# / (1# / r1 + 1# / r2 + 1# / r3 + 1# / r4)
  wr1 = cp / r1
  wr2 = cp / r2
  wr3 = cp / r3
  wr4 = cp / r4
100:
  ixp = ix + 1
  iyp = iy + 1
  Cf2 = wr1 * C2a(ix, iy) + wr2 * C2a(ix, iyp) + _
       wr3 * C2a(ixp, iyp) + wr4 * C2a(ixp, iy)
  Cf25 = wr1 * C25a(ix, iy) + wr2 * C25a(ix, iyp) + _
        wr3 * C25a(ixp, iyp) + wr4 * C25a(ixp, iy)
  Cf100 = wr1 * C100a(ix, iy) + wr2 * C100a(ix, iyp) + _
         wr3 * C100a(ixp, iyp) + wr4 * C100a(ixp, iy)
End Sub

Private Sub INDEXX(n As Long, ARRIN() As Single, INDX() As Integer)
'Subroutine INDEXX indexs an array ARRIN of length N, outputs the
' array INDX such that ARRIN(INDX(J)) is in ascending order for
' J=1,2,..,N. The input quantities ARRIN and N are not changed
'     (ref. Numerical Recipes, p. 233)
  Dim q As Double
  Dim i&, indxt&, ir&, j&, l&

  For j = 1 To n
    INDX(j) = j
  Next j
  l = n / 2 + 1
  ir = n
20:
  If (l > 1) Then
    l = l - 1
    indxt = INDX(l)
    q = ARRIN(indxt)
  Else
    indxt = INDX(ir)
    q = ARRIN(indxt)
    INDX(ir) = INDX(1)
    ir = ir - 1
    If (ir = 1) Then
      INDX(1) = indxt
      Exit Sub
    End If
  End If
  i = l
  j = l + l
30:
  If (j <= ir) Then
    If (j < ir) Then
      If (ARRIN(INDX(j)) < ARRIN(INDX(j + 1))) Then j = j + 1
    End If
    If (q < ARRIN(INDX(j))) Then
      INDX(i) = INDX(j)
      i = j
      j = j + j
    Else
      j = ir + 1
    End If
    GoTo 30
  End If
  INDX(i) = indxt
  GoTo 20
End Sub

Private Function gauscf(xx!) As Single
  'cumulative probability function
  Dim ax!, T!, d!
  Const xlim! = 18.3

  ax = Abs(xx)
  gauscf = 1#
  If ax <= xlim Then
    T = 1# / (1# + 0.2316419 * ax)
    d = 0.3989423 * Exp(-xx * xx * 0.5)
    gauscf = 1# - _
             d * T * ((((1.330274 * T - 1.821256) * T + 1.781478) * T - 0.3565638) _
             * T + 0.3193815)
  End If
  If xx < 0 Then gauscf = 1# - gauscf
End Function

Private Function gausdy(xx!) As Single
  'cumulative probability function
  Const xlim! = 18.3

  gausdy = 0#
  If (Abs(xx) <= xlim) Then
    gausdy = 0.3989423 * Exp(-0.5 * xx * xx)
  End If
End Function

Private Sub OutPut(IOUT As Long, IPK As Integer, pru As Single, _
  OutFile As Long, ByRef Flow As Double, ByRef StdErrorMinus As Double, _
  ByRef StdErrorPlus As Double, ByRef EquivYrs As Double, _
  cl90 As Double, cu90 As Double)
  'subroutine outputs results to screen and file
  Dim eqyrs As Single, cookd As Single, _
      delres As Single, errmod As Single, hatdig As Single, hmax As Single, _
      press As Single, prx As Single, maxlev As Single, _
      pv4 As Single, resid As Single, samerr As Single, sdbeta As Single, _
      sepc As Single, sepu As Single, arhoc As Single, MEV As Single
  Dim ssam As Single, stdres As Single, Sum As Single, tbeta As Single, _
      hatdigtest As Single, cooktest As Single, tstat As Single, tv4 As Single, _
      varres As Single, vpu As Single, asep As Single, splus As Single, _
      sminu As Single
  Dim work1(50, 1) As Single, work2(10, 1) As Single, work3(10, 50) As Single, _
      xo(1, 10) As Single, xot(10, 1) As Single, ccc(1, 1) As Single, _
      hat(50, 50) As Single, hats(50, 50) As Single
  Dim i As Integer, j As Integer, l As Integer, ndf As Long

  If (IOUT > 0) Then
    'Variance of Betas sits along diagonal of XtXinv
    sdbeta = XtXinv(1, 1) ^ 0.5
    'Divide Beta by standard deviation to "Studentize" or standardize
    tbeta = Bhat(1, 1) / sdbeta
'   write(17,9020) Bhat(1,1),sdbeta,Tbeta
    'Loop thru remaining BETAS
    For i = 2 To RegVarCnt
      sdbeta = XtXinv(i, i) ^ 0.5
      tbeta = Bhat(i, 1) / sdbeta
      ndf = Nsites - RegVarCnt
      tv4 = Abs(tbeta)
      pv4 = 2# * STUTP(-tv4, ndf)
      If (pv4 < 0.0001) Then pv4 = 0.0001
'     write(17,9025) Ylabel(i-1),Bhat(i,1),sdbeta,Tbeta,pv4
    Next i
    
    If (IOUT < 99) Then Exit Sub
    
    Call mltply(work1(), X(), Bhat(), Nsites, RegVarCnt, 1, MaxSites, MaxSites, MaxInd)
    ssam = 0
    press = 0
    hmax = 0
    maxlev = 0
    'ensuing 10 and 50 are max # of ind vars and stations, respectively
    Call mltply(work3, XtXinv, Xt, RegVarCnt, RegVarCnt, Nsites, MaxInd, MaxInd, MaxInd)
    Call mltply(hats, X, work3, Nsites, RegVarCnt, Nsites, MaxSites, MaxSites, MaxInd)
    Call mltply(hat, hats, Wt, Nsites, Nsites, Nsites, MaxSites, MaxSites, MaxSites)
    For i = 1 To Nsites
      pru = work1(i, 1)
      resid = Y(i, 1) - work1(i, 1)
      samerr = hats(i, i)
      ssam = ssam + samerr
      If (samerr > hmax) Then hmax = samerr
      hatdig = hat(i, i)
      If hatdig >= maxlev Then maxlev = hatdig
      delres = resid / (1# - hatdig)
      press = press + delres ^ 2
      errmod = Gamasq
      varres = Gamasq + Sta(i) - samerr
      stdres = resid / (varres ^ 0.5)
      cookd = stdres ^ 2 * samerr / (RegVarCnt * varres)
      'The following 2 measures ID outliers, or "Highly Influencial Observations"
      hatdigtest = 2# * RegVarCnt / Nsites
      cooktest = 4# / Nsites
'      if ( hatdig.GT.hatdigtest .OR. cookd.GT.cooktest ) then
'        write(17,9035) Stano(i),Y(i,1),log_pred,stdres,hatdig,cookd
'      endif
    Next i

    'TOTAL REGRESSION DIAGNOSTICS
    'Compute the 'total' regression errors etc . . .
    'ssam   = Average Sampling Error Variance
    'GamaSq = Model Error Variance
    'Atse   = Average Time-Sampling Error Variance
    'Atscov = Average Cross-Covariance
    'Arhoc  = Average Cross-Correlation Coefficient
    'MEV    = Mean error variance of total regression
    ssam = ssam / Nsites    'is calcd but not used
    press = press / Nsites  'is calcd but not used
    Atse = Atse / Nsites    'calcd in main module, used for arhoc
    Atscov = Atscov / (Nsites / 2# * (Nsites - 1))  'calcd in main module, used for arhoc
    arhoc = Atscov / Atse   'average cross correlation, not used
    MEV = ssam + Gamasq

    'Calculate the ungaged site estimation, log10 then real-space
    pru = 0#
    For i = 1 To RegVarCnt
      pru = pru + Bhat(i, 1) * UserRegressVars(i)
    Next i
    prx = 10 ^ pru

    'Calculate the error variance specific to the ungaged site
    ' in relation to the other sites in the regression model
    For i = 1 To RegVarCnt
      Sum = 0#
      For j = 1 To RegVarCnt
        Sum = Sum + XtXinv(i, j) * UserRegressVars(j)
      Next j
      work2(i, 1) = Sum
    Next i
    Sum = 0#
    For j = 1 To RegVarCnt
      Sum = Sum + UserRegressVars(j) * work2(j, 1)
    Next j
    sepu = Sum
    vpu = Gamasq + sepu
    eqyrs = sig ^ 2 * (1# + Ak(IPK) ^ 2 / 2#) / vpu
    
    'Convert to cfs and percent error
    sepc = 100# * (Exp(vpu * 5.302) - 1#) ^ 0.5
    asep = vpu ^ 0.5
    splus = 100# * (10 ^ (asep) - 1#)
    sminu = 100# * (10 ^ (-asep) - 1#)
    tstat = vpu ^ 0.5 * STUTX(1 - (0.1 / 2), ndf)
    '"STUTX(1 - (0.1 / 2), ndf)" => 1.65 as ndf => infinity (=1.70 when ndf=28)
    'NC program used 1.65, but calling STUTX is more accurate
    cu90 = 10 ^ (tstat + pru)
    cl90 = 10 ^ (pru - tstat)
    Call Round(prx)
    Call Round(cl90)
    Call Round(cu90)
    
    'set the vars sent as arguments
    Flow = prx
    StdErrorPlus = splus
    StdErrorMinus = sminu
    EquivYrs = eqyrs
    
    Print #OutFile, PkLab(IPK) _
          & vbTab & Format(prx, "#####0") _
          & vbTab & vbTab & Format(sminu, "###.0") _
          & vbTab & vbTab & Format(splus, "###.0") _
          & vbTab & vbTab & Trim(Format(cl90, "######0")) & " - " & Trim(Format(cu90, "######0"))

    If (sepu > hmax) Then
      Print #OutFile, "WARNING: Prediction is outside range of observed data" & vbCrLf
      ssMessageBox "WARNING: Prediction is an extrapolation beyond'" & vbCrLf & _
          "observed data. Check for errors in input basin" & vbCrLf & _
          "characteristics. If no errors use results with caution."
    End If
  End If
End Sub

Private Sub Init()
  Dim sql$
  Dim i&, j&

  Set PkLabColl = New FastCollection
  PkLabColl.Add "2-year", "2"
  PkLabColl.Add "5-year", "5"
  PkLabColl.Add "10-year", "10"
  PkLabColl.Add "25-year", "25"
  PkLabColl.Add "50-year", "50"
  PkLabColl.Add "100-year", "100"
  PkLabColl.Add "200-year", "200"
  PkLabColl.Add "500-year", "500"

  Set AkColl = New FastCollection
  AkColl.Add 0#, "2"
  AkColl.Add 0.84162, "5"
  AkColl.Add 1.28155, "10"
  AkColl.Add 1.75069, "25"
  AkColl.Add 2.05375, "50"
  AkColl.Add 2.32635, "100"
  AkColl.Add 2.575, "200"
  AkColl.Add 2.87816, "500"

  'the following 2 values apply only to TX
  Alpha = 0.0025
  Theta = 0.983
  
  'Nsites = 30
  
End Sub

Private Function TASKER_DISTANCE(ByVal latxin As Single, _
                                 ByVal latyin As Single, _
                                 ByVal longxin As Single, _
                                 ByVal longyin As Single) As Single
' The x latitude and longitude are inputted in decimal degrees
' The y latitude and longitude are inputted in decimal degrees
' Distance is returned in miles
' The projection is Albers Equal-Area Conic Projection (with Earth as an
' elipsoid) Ref. Snyder, J.P., 1982, "Map Projections Used by The U.S.
' Geological Survey", U. S. Geological Survey Bulletin 1532, p. 96-99.
  Dim radian As Single, A As Single, e2 As Single, latx As Single, _
      laty As Single, longx As Single, longy As Single, _
      latdiff As Single, longdiff As Single, latave As Single, _
      op1 As Single, op2 As Single, op3 As Single

  'Convert station coords to decimal degrees, if necessary
  If longxin > 999999 Then
    longxin = Left(longxin, 3) + Mid(longxin, 4, 2) / 60 + Mid(longxin, 6) / 3600
  ElseIf longxin > 99999 Then
    longxin = Left(longxin, 2) + Mid(longxin, 3, 2) / 60 + Mid(longxin, 5) / 3600
  End If
  If latxin > 99999 Then latxin = Left(latxin, 2) + Mid(latxin, 3, 2) / 60 + Mid(latxin, 5) / 3600
  If longyin > 999999 Then
    longyin = Left(longyin, 3) + Mid(longyin, 4, 2) / 60 + Mid(longyin, 6) / 3600
  ElseIf longyin > 99999 Then
    longyin = Left(longyin, 2) + Mid(longyin, 3, 2) / 60 + Mid(longyin, 5) / 3600
  End If
  If latyin > 99999 Then latyin = Left(latyin, 2) + Mid(latyin, 3, 2) / 60 + Mid(latyin, 5) / 3600
  
  A = 6378206.4
  e2 = 0.00676866
  radian = 3.14159265 / 180
  latx = latxin * radian
  laty = latyin * radian
  longx = longxin * radian
  longy = longyin * radian

  latdiff = Abs(latx - laty)
  longdiff = Abs(longx - longy)
  latave = (latx + laty) / 2#
  op1 = A / (1# - e2 * Sin(latave) ^ 2) ^ 0.5
  op2 = ((1# - e2) ^ 2 * (latdiff) ^ 2)
  op2 = op2 / ((1# - e2 * Sin(latave) ^ 2) ^ 2)
  op3 = Cos(latave) ^ 2 * longdiff ^ 2
' Convert distance to miles by dividing the meters by 1000 to get
' kilometeres and then convert to miles by multiplying by 0.621
  TASKER_DISTANCE = 0.62137119 * (op1 * (op2 + op3) ^ 0.5) / 1000#
End Function

Private Sub BuildMatrix()
'-- Build the cross-correlation matrix on the number of stations selected
'--  This subroutine uses the arrays lat and long as basis
  Dim dist As Single, latx As Single, laty As Single, longx As Single, longy As Single
  Dim i&, j&

  If (Theta < 0 Or Theta >= 1#) Then
     ssMessageBox "ERROR-- Theta must be between 0 and 1"
     Exit Sub
  End If

  ReDim Rhoc(Nsites, Nsites) As Single
  For i = 1 To Nsites
    For j = 1 To Nsites
      latx = StaLats(INDX(i))
      laty = StaLats(INDX(j))
      longx = StaLngs(INDX(i))
      longy = StaLngs(INDX(j))
      Rhoc(i, j) = TASKER_DISTANCE(latx, laty, longx, longy)
    Next j
  Next i

  For i = 1 To Nsites
    For j = 1 To i
      If (i = j) Then
        Rhoc(i, j) = 1#
      Else
        Rhoc(i, j) = Theta ^ (Rhoc(i, j) / (Alpha * Rhoc(i, j) + 1#))
        If (Rhoc(i, j) < 0#) Then
          ssMessageBox "rhoc(i,j) = " & i & " " & j & " " & Rhoc(i, j) & vbCrLf & _
                "ERROR--Cholesky Decomp requires >= 0"
          Exit Sub
        End If
      End If
    Next j
  Next i
End Sub

Private Function STUTX(p As Single, n As Long) As Single
  'STUDENT T QUANTILES --
  ' STUTX(P,N) = X SUCH THAT PROB(STUDENT T WITH N D.F. .LE. X) = P.
  'NOTE - ABS(T) HAS PROB Q OF EXCEEDING STUTX( 1.-Q/2., N ).
  ' NOTE -      IER - ERROR FLAG --  1 = F.LT.1.,
  '                                  2 = P NOT  IN (0,1),  3 = 1+2
  ' SUBPGMS USED -- GAUSAB     (GAUSSIAN ABSCISSA)
  'REF - G. W. HILL (1970) ACM ALGO 396.  COMM ACM 13(10)619-20.
  '           REV BY WKIRBY 10/76. 2/79.  10/79.
  Dim q As Single, A As Single, B As Single, c As Single, d As Single, _
      HPI As Single, FN As Single, X As Single, Y As Single
  Dim sign As Long, IER As Long

      HPI = 1.5707963268
      sign = 1#
      If p < 0.5 Then sign = -1#
      q = 2# * p
      If q > 1# Then q = 2# * (1# - p)
      If q < 1# Then GoTo Next1
      STUTX = 0#
      Exit Function
Next1:
      FN = n
      If (n >= 1 And q > 0# And q < 1#) Then GoTo Next2
      IER = 3
      If n >= 1 Then IER = 2
      STUTX = sign * 1E+38
      Exit Function
Next2:
      If n <> 1 Then GoTo Next3
'  -- 1 DEG FR - EXACT
      STUTX = sign / Tan(HPI * q)
      Exit Function
Next3:
      If n <> 2 Then GoTo Next4
'  -- 2 DEG FR - EXACT
      STUTX = ((2# / (q * (2# - q)) - 2#) ^ 0.5) * sign
      Exit Function
Next4:
'  -- EXPANSION FOR N .GT. 2
      A = 1# / (FN - 0.5)
      B = 48# / (A * A)
      c = ((20700# * A / B - 98#) * A - 16#) * A + 96.36
      d = ((94.5 / (B + c) - 3#) / B + 1#) * ((A * HPI) ^ 0.5) * FN
      X = d * q
      Y = X ^ (2# / FN)
      If Y > (A + 0.05) Then GoTo Next5
      Y = ((1# / (((FN + 6#) / (FN * Y) - 0.089 * d - 0.822) * (FN + 2#) * 3#) + _
          0.5 / (FN + 4#)) * Y - 1#) * (FN + 1#) / (FN + 2#) + 1# / Y
      STUTX = ((FN * Y) ^ 0.5) * sign
      Exit Function
Next5:
'   -- ASYMPTOTIC INVERSE EXPANSION ABOUT NORMAL
      X = GausAB(0.5 * q)
      Y = X * X
      If FN < 5# Then c = c + 0.3 * (FN - 4.5) * (X + 0.6)
      c = (((0.05 * d * X - 5#) * X - 7#) * X - 2#) * X + B + c
      Y = (((((0.4 * Y + 6.3) * Y + 36#) * Y + 94.5) / c - Y - 3#) / B + 1#) * X
      X = A * Y ^ 2
      Y = X + 0.5 * X ^ 2
      If X > 0.002 Then Y = Exp(X) - 1#
      STUTX = ((FN * Y) ^ 0.5) * sign
  End Function
  
  Private Function GausAB(CUMPRB As Single) As Single
  'GAUSSIAN PROBABILITY FUNCTIONS   W.KIRBY  JUNE 71
  'GAUSEX=VALUE EXCEEDED WITH PROB EXPROB
  'GAUSCF MODIFIED 740906 WK -- REPLACED ERF FCN REF BY RATIONAL APPR
  'ALSO REMOVED DOUBLE PRECISION FROM GAUSEX AND GAUSAB.
  '76-05-04 WK -- TRAP UNDERFLOWS IN EXP IN GUASCF AND DY.
  '02-07-17 R.Dusenbury -- converted from FORTRAN to VB (DOES NOT WORK PROPERLY!!!)

    Dim p As Single, pr As Single, T As Single, c0 As Single, c1 As Single, _
        C2 As Single, d1 As Single, d2 As Single, d3 As Single
    Dim numerat!, denom!

    On Error GoTo 0
    
    c0 = 2.515517
    c1 = 0.802853
    C2 = 0.010328
    d1 = 1.432788
    d2 = 0.189269
    d3 = 0.001308
    GausAB = 0#
    p = 1# - CUMPRB
    If p >= 1# Then 'set to minimum
      GausAB = -10#
    ElseIf p <= 0# Then 'set to maximum
      GausAB = 10#
    Else 'compute value
      pr = p
      If p > 0.5 Then pr = 1# - pr
      T = (-2# * Log10(pr)) ^ 0.5
      T = (-2# * Log(pr)) ^ 0.5
      numerat = (c0 + T * (c1 + T * C2))
      denom = (1# + T * (d1 + T * (d2 + T * d3)))
      GausAB = T - numerat / denom
      If p > 0.5 Then GausAB = -GausAB
    End If
  End Function
