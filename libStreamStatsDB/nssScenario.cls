VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nssScenario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private pName As String
Private pArea As Double
Private pRegCrippenBue As Long
Private pUrban As Boolean
Private pROI As Boolean
Private pLowFlow As Boolean
Private pRevEqtn As Boolean
Private pUsePredInts As Boolean
Private pProject As nssProject
Private pUserRegions As FastCollection 'of userRegions
Private pRuralScenario As nssScenario
Private pWeight As nssWeight
Private pSaveEquivYears() As Double
Private pSaveStdError() As Double
Private pSaveStdErrorPlus() As Double
Private pSavePredInts() As Double
Private pMatrix As Variant
Private pRHO As Variant
Private pCF As Variant
Private pParmOutOfRange As Boolean
Private pErrorDescription As String
Private pExtrapolated As Boolean

Public Function ErrorDescription() As String
  ErrorDescription = pErrorDescription
  pErrorDescription = ""
End Function

Public Function Copy() As nssScenario
  Dim oldUserRegion As Variant
  Dim newUserRegion As userRegion
  Dim retval As nssScenario
  Set retval = New nssScenario
  With retval
    .Name = pName
    Set .Project = pProject
    .SetArea pArea, pProject.State.Metric
    .RegCrippenBue = pRegCrippenBue
    .Urban = pUrban
    .ROI = pROI
    .LowFlow = pLowFlow
    .UsePredInts = pUsePredInts
    If Not IsEmpty(pRHO) Then .RHO = pRHO
    If Not IsEmpty(pCF) Then .CF = pCF
    Set .Project = pProject
    Set .Weight = pWeight
    If Not pRuralScenario Is Nothing Then Set .RuralScenario = pRuralScenario
    For Each oldUserRegion In UserRegions
      Set newUserRegion = New userRegion
      'newUserRegion.Area = oldUserRegion.Area
      Set newUserRegion.Region = oldUserRegion.Region
      Set newUserRegion.UserParms = oldUserRegion.CopyUserParms
      .UserRegions.Add newUserRegion, newUserRegion.Region.Name
      Set newUserRegion = Nothing
    Next
  End With
  Set Copy = retval
  Set retval = Nothing
End Function

Public Property Get Name() As String
  Name = pName
End Property
Public Property Let Name(ByVal newValue As String)
  pName = newValue
End Property

Public Function GetArea(WantMetric As Boolean) As Double
  GetArea = pProject.DB.AreaUnits.Convert(pArea, pProject.State.Metric, WantMetric)
End Function
Public Sub SetArea(ByVal newValue As Double, IsMetric As Boolean)
  pArea = pProject.DB.AreaUnits.Convert(newValue, IsMetric, pProject.State.Metric)
End Sub

Public Property Get RegCrippenBue() As Long
  RegCrippenBue = pRegCrippenBue
End Property
Public Property Let RegCrippenBue(ByVal newValue As Long)
  pRegCrippenBue = newValue
End Property

Public Property Get Urban() As Boolean
  Urban = pUrban
End Property
Public Property Let Urban(ByVal newValue As Boolean)
  pUrban = newValue
End Property

Public Property Get ROI() As Boolean
  ROI = pROI
End Property
Public Property Let ROI(ByVal newValue As Boolean)
  pROI = newValue
End Property

Public Property Get LowFlow() As Boolean
  LowFlow = pLowFlow
End Property
Public Property Let LowFlow(ByVal newValue As Boolean)
  pLowFlow = newValue
End Property

Public Property Get RevEqtn() As Boolean
  RevEqtn = pRevEqtn
End Property
Public Property Let RevEqtn(ByVal newValue As Boolean)
  pRevEqtn = newValue
End Property

Public Property Get UsePredInts() As Boolean
  UsePredInts = pUsePredInts
End Property
Public Property Let UsePredInts(ByVal newValue As Boolean)
  pUsePredInts = newValue
End Property

Public Property Get Project() As nssProject
  Set Project = pProject
End Property
Public Property Set Project(ByVal newValue As nssProject)
  Set pProject = newValue
End Property

Public Property Get Weight() As nssWeight
  Set Weight = pWeight
End Property
Public Property Set Weight(ByVal newValue As nssWeight)
  Set pWeight = newValue
End Property

Public Property Get ParmOutOfRange() As Boolean
  ParmOutOfRange = pParmOutOfRange
End Property
Public Property Let ParmOutOfRange(ByVal newValue As Boolean)
  pParmOutOfRange = newValue
End Property

Public Property Get Summary() As String
  Dim Region As Variant 'userRegion
  Dim Param As Variant 'userParameter
  Dim retval As String
  Dim WeightText As String
  Dim GagedYears As String
  Dim Val As Double, Max As Double, Min As Double
  retval = pName
  pParmOutOfRange = False 'assume all parameter values are within range
  'If Me.Urban Then retval = retval & " (urban)" Else retval = retval & " (rural)"
  retval = retval & vbCrLf & "Basin Drainage Area: " & SignificantDigits(GetArea(pProject.Metric), 3) & " " & pProject.DB.AreaUnits.Label(pProject.Metric)
  
  retval = retval & vbCrLf & UserRegions.Count & " Region"
  If UserRegions.Count > 1 Then retval = retval & "s"
  For Each Region In pUserRegions
    retval = retval & vbCrLf & "Region: " & Region.Region.Name
    For Each Param In Region.UserParms
      Val = Param.GetValue(Me.Project.Metric)
      Max = Param.Parameter.GetMax(Me.Project.Metric)
      Min = Param.Parameter.GetMin(Me.Project.Metric)
      retval = retval & vbCrLf & "   " & Param.Parameter.NSSName & " = " _
                                       & SignificantDigits(Val, 3) _
                                       & " " & Param.Parameter.Units.Label(Me.Project.Metric)
      If Abs(Max + 999) > 1 Then
        If Val > Max Then
          retval = retval & " (above max value " & SignificantDigits(Max, 3) & ")"
          pParmOutOfRange = True
        End If
      End If
      If Abs(Min + 999) > 1 Then
        If Val < Min Then
          retval = retval & " (below min value " & SignificantDigits(Min, 3) & ")"
          pParmOutOfRange = True
        End If
      End If
    Next
    If Me.Urban And Region.Region.UrbanNeedsRural Then
      retval = retval & vbCrLf & "   Rural Scenario = "
      If Me.RuralScenario Is Nothing Then
        retval = retval & "(Not specified)"
      Else
        retval = retval & Me.RuralScenario.Name
      End If
    End If
  Next
  If Me.RegCrippenBue > 0 Then
    retval = retval & vbCrLf & "Crippen & Bue Region " & Me.RegCrippenBue
  End If
  
  Select Case Me.Weight.WeightType
    Case 0: GoTo NoWeighting
    Case 1
      GagedYears = Me.Weight.GetGagedYears(pUserRegions(1).Region.DepVars(1).Name)
      retval = retval & vbCrLf & "  Weighted with " & GagedYears & " years of gaged data"
    Case 2
      GagedYears = "0"
      retval = retval & vbCrLf & "  Weighted as ungaged site"
      retval = retval & vbCrLf & "    Gaged area = " _
             & Me.Project.DB.Units("13").Convert(Me.Weight.AreaGaged, _
                                                 Me.Project.State.Metric, _
                                                 Me.Project.Metric)
  End Select
  For Each Param In pUserRegions(1).Region.DepVars
    retval = retval & vbCrLf & "   Interval " & StrPad(Param.Name, 3) _
          & " Gaged value = " & Me.Weight.GetGagedValue(Param.Name)
  Next
'  WeightText = Me.Weight.XML
'  WeightText = Mid(WeightText, InStr(WeightText, ">") + 1) 'Skip <weight type="1"> or <weight type="2">
'  WeightText = ReplaceString(WeightText, "<weightinterval year=""", " Interval ")
'  WeightText = ReplaceString(WeightText, """ gagedvalue=""", " Gaged value = ")
'  WeightText = ReplaceString(WeightText, """ gagedyears=""" & GagedYears & """/>", "")
'  WeightText = ReplaceString(WeightText, "</weight>", "")
'  WeightText = ReplaceString(WeightText, ">", "")
'  retval = retval & WeightText
NoWeighting:
  Summary = retval
End Property

Public Property Get UserRegions() As FastCollection
  If pUserRegions Is Nothing Then Set pUserRegions = New FastCollection
  Set UserRegions = pUserRegions
End Property

Public Property Get RuralScenario() As nssScenario
  Set RuralScenario = pRuralScenario
End Property
Public Property Set RuralScenario(newValue As nssScenario)
  Set pRuralScenario = newValue
  If pArea = 0 Then pArea = pRuralScenario.GetArea(pProject.State.Metric)
End Property

Public Sub Edit()
  Load frmEditScenario
  Set frmEditScenario.Scenario = Me
  frmEditScenario.Show
End Sub

Public Sub EditWeight()
  Load frmEditWeight
  Set frmEditWeight.Scenario = Me
  frmEditWeight.Show
End Sub

Public Function EquivalentYears() As Double()
  Dim RegionIndex As Long
  Dim ReturnIndex As Long
  Dim ComponentIndex As Long
  Dim AccumFlow As Double
  Dim FirstReturns As FastCollection
  Dim CurRegion As userRegion
  Dim CurReturn As nssDepVar
  Dim CurComponent As nssComponent
  Dim retval() As Double
  Dim curval As Double
  Dim ExpVal As Double
  Dim dArea2overG As Double
  Dim msg As String
  
  If Me.ROI Then
    EquivalentYears = pSaveEquivYears 'this was computed in call to ComputeROIDischarge
  Else
  
  'Commented out just in case we were saving old values too long
  '  If UBound(pSaveEquivYears) > 0 Then
  '    EquivalentYears = pSaveEquivYears
  '    Exit Function
  '  End If
    
    If pUserRegions.Count = 0 Then
      ReDim retval(0)
      EquivalentYears = retval
      Exit Function
    End If
    
  'temporarily comment out to make weighting method work,
  'but need something like this to not have Eq Yrs displayed on main form after weighting
  '  If pWeight.WeightType > 0 Then
  '    ReDim retval(1 To pUserRegions(1).Region.DepVars.Count)
  '    EquivalentYears = retval
  '    Exit Function
  '  End If
    
    msg = MismatchedRegions
    If Len(msg) > 0 Then
      ssMessageBox msg, vbOKOnly, "Cannot Compute Equivalent Years"
    Else
      Set FirstReturns = pUserRegions(1).Region.DepVars
      
      ReDim retval(1 To FirstReturns.Count)
      
      For ReturnIndex = 1 To FirstReturns.Count
        retval(ReturnIndex) = 0
        For RegionIndex = 1 To pUserRegions.Count
          Set CurRegion = pUserRegions(RegionIndex)
          Set CurReturn = CurRegion.Region.DepVars(ReturnIndex)
          retval(ReturnIndex) = retval(ReturnIndex) + CurReturn.EquivYears * CurRegion.UserParms(1).GetValue(Me.Project.Metric)
        Next
        retval(ReturnIndex) = retval(ReturnIndex) / Me.GetArea(Me.Project.Metric)
        If retval(ReturnIndex) > 0 Then
          If pWeight.WeightType = 1 Then
            retval(ReturnIndex) = retval(ReturnIndex) + pWeight.GetGagedYears(CurReturn.Name)
          ElseIf pWeight.WeightType = 2 Then
            dArea2overG = 2 * Abs(pWeight.AreaGaged - Me.GetArea(Me.Project.Metric)) / pWeight.AreaGaged
            retval(ReturnIndex) = dArea2overG * retval(ReturnIndex) + (1 - dArea2overG) * (Me.GetArea(Me.Project.Metric) / pWeight.AreaGaged) ^ CurRegion.Region.DepVars(ReturnIndex).ExpDA * pWeight.GetGagedYears(CurReturn.Name)
          End If
        End If
      Next
    End If
LastLine:
    EquivalentYears = retval 'Return empty array if there is a problem
  End If
End Function

Public Function StdError() As Double()
  Dim RegionIndex As Long
  Dim ReturnIndex As Long
  Dim ComponentIndex As Long
  Dim AccumFlow As Double
  Dim FirstReturns As FastCollection
  Dim CurRegion As userRegion
  Dim CurReturn As nssDepVar
  Dim CurComponent As nssComponent
  Dim retval() As Double
  Dim curval As Double
  Dim ExpVal As Double
  Dim msg As String
  
'Commented out just in case we were saving old values too long
'  If UBound(pSaveStdError) > 0 Then
'    StdError = pSaveStdError
'    Exit Function
'  End If
  
  If Me.ROI Then
    StdError = pSaveStdError 'this was computed in call to ComputeROIDischarge
  Else
    If pUserRegions.Count = 0 Then
      ReDim retval(0)
      StdError = retval
      Exit Function
    End If
    
    If pWeight.WeightType > 0 Then
      ReDim retval(1 To pUserRegions(1).Region.DepVars.Count)
      StdError = retval
      Exit Function
    End If
    
    msg = MismatchedRegions
    If Len(msg) > 0 Then
      ssMessageBox msg, vbOKOnly, "Cannot Compute Standard Error"
    Else
      Set FirstReturns = pUserRegions(1).Region.DepVars
      ReDim retval(1 To FirstReturns.Count)
      For ReturnIndex = 1 To FirstReturns.Count
        retval(ReturnIndex) = 0
        For RegionIndex = 1 To pUserRegions.Count
          Set CurRegion = pUserRegions(RegionIndex)
          Set CurReturn = CurRegion.Region.DepVars(ReturnIndex)
          retval(ReturnIndex) = retval(ReturnIndex) + CurReturn.StdErr * CurRegion.UserParms(1).GetValue(Me.Project.Metric)
        Next
        retval(ReturnIndex) = retval(ReturnIndex) / Me.GetArea(Me.Project.Metric) 'pUserRegions.Count
      Next
    End If
    StdError = retval
  End If
End Function

Public Function StdErrorPlus() As Double()
  StdErrorPlus = pSaveStdErrorPlus
End Function

Public Function PredictionError(ValidValues As Boolean) As Double()
  Dim RegionIndex As Long
  Dim ReturnIndex As Long
  Dim FirstReturns As FastCollection
  Dim CurRegion As userRegion
  Dim CurReturn As nssDepVar
  Dim retval() As Double
  
  ValidValues = False
  If Me.ROI Or pUserRegions.Count = 0 Then
    'only ROI-calculated standard error needed, no Prediction error (or no regions)
    ReDim retval(0)
  Else
    Set FirstReturns = pUserRegions(1).Region.DepVars
    ReDim retval(1 To FirstReturns.Count)
    'make sure we're not weighted or having mismatched regions
    If pWeight.WeightType = 0 And Len(MismatchedRegions) = 0 Then
      For ReturnIndex = 1 To FirstReturns.Count
        retval(ReturnIndex) = 0
        For RegionIndex = 1 To pUserRegions.Count
          Set CurRegion = pUserRegions(RegionIndex)
          Set CurReturn = CurRegion.Region.DepVars(ReturnIndex)
          retval(ReturnIndex) = retval(ReturnIndex) + CurReturn.PreErr * CurRegion.UserParms(1).GetValue(Me.Project.Metric)
        Next
        retval(ReturnIndex) = retval(ReturnIndex) / Me.GetArea(Me.Project.Metric) 'pUserRegions.Count
        If retval(ReturnIndex) > 0 Then ValidValues = True
      Next
    End If
  End If
  PredictionError = retval

End Function

Public Function EstimateError(ValidValues As Boolean) As Double()
  Dim RegionIndex As Long
  Dim ReturnIndex As Long
  Dim FirstReturns As FastCollection
  Dim CurRegion As userRegion
  Dim CurReturn As nssDepVar
  Dim retval() As Double

  ValidValues = False
  If Me.ROI Or pUserRegions.Count = 0 Then
    'only ROI-calculated standard error needed, no Estimate error (or no regions)
    ReDim retval(0)
  Else
    Set FirstReturns = pUserRegions(1).Region.DepVars
    ReDim retval(1 To FirstReturns.Count)
    'make sure we're not weighted or having mismatched regions
    If pWeight.WeightType = 0 And Len(MismatchedRegions) = 0 Then
      For ReturnIndex = 1 To FirstReturns.Count
        retval(ReturnIndex) = 0
        For RegionIndex = 1 To pUserRegions.Count
          Set CurRegion = pUserRegions(RegionIndex)
          Set CurReturn = CurRegion.Region.DepVars(ReturnIndex)
          retval(ReturnIndex) = retval(ReturnIndex) + CurReturn.EstErr * CurRegion.UserParms(1).GetValue(Me.Project.Metric)
        Next
        retval(ReturnIndex) = retval(ReturnIndex) / Me.GetArea(Me.Project.Metric) 'pUserRegions.Count
        If retval(ReturnIndex) > 0 Then ValidValues = True
      Next
    End If
  End If
  EstimateError = retval

End Function

Public Function PredInts() As Double()

  Dim RegionIndex As Long
  Dim ReturnIndex As Long
  Dim FirstReturns As FastCollection
  Dim CurRegion As userRegion
  Dim CurStat As nssDepVar
  Dim CurReturn As nssDepVar
  Dim CurComponent As nssComponent
  Dim i&, j&, MDim&, S!, T!
  Dim Prod() As Single, Prod2(1, 1) As Single
  Dim lv As Double
  Dim lParm As Single
  Dim X() As Single 'X vector
  Dim Xt() As Single 'transpose of X vector
  Dim Uv As Variant
  Dim U() As Single
  Dim retval() As Double
  Dim msg As String
  Dim StateMetric As Boolean
  Dim lCompName As String

  If pUserRegions.Count = 0 Then
    ReDim retval(0)
    PredInts = retval
    Exit Function
  End If

  If Me.ROI Then
    PredInts = pSavePredInts
  Else
    msg = MismatchedRegions
    If Len(msg) > 0 Then
      ssMessageBox msg, vbOKOnly, "Cannot Compute Prediction Intervals"
    Else
      If pUserRegions(1).Region.Name = pProject.NationalUrban.Name Then
        StateMetric = pProject.NationalUrban.State.Metric
      Else
        StateMetric = pProject.State.Metric
      End If
      Set FirstReturns = pUserRegions(1).Region.DepVars
      ReDim retval(2, 1 To FirstReturns.Count)
      For ReturnIndex = 1 To FirstReturns.Count
        Set CurReturn = FirstReturns(ReturnIndex)
        retval(1, ReturnIndex) = 0
        retval(2, ReturnIndex) = 0
        For RegionIndex = 1 To pUserRegions.Count
          Set CurRegion = pUserRegions(RegionIndex)
          Set CurStat = CurRegion.Region.DepVars(ReturnIndex)
          MDim = CurStat.Components.Count + 1
          ReDim X(MDim, 1)
          ReDim Xt(1, MDim)
          X(1, 1) = 1
          Xt(1, 1) = X(1, 1)
          For i = 2 To CurStat.Components.Count + 1
            Set CurComponent = CurStat.Components(i - 1)
            If i = 2 Then 'use area from whole scenario, not just the part in this region
              lParm = Me.GetArea(StateMetric)
            Else
              If Not CurComponent.Parameter Is Nothing Then
                lCompName = CurComponent.Parameter.Name
              ElseIf Not CurComponent.ExponentParameter Is Nothing Then
                lCompName = CurComponent.ExponentParameter.Name
              Else
                MsgBox "PROBLEM:  Can't find component name for computing PIs"
              End If
              j = i - 1
              While j <= CurRegion.UserParms.Count
                If lCompName = CurRegion.UserParms(j).Parameter.Name Then
                  lParm = CurRegion.UserParms(j).GetValue(Me.Project.Metric)
                  j = CurRegion.UserParms.Count
                End If
                j = j + 1
              Wend
            End If
            If CurComponent.ExpID = 0 Then 'base term, take log if possible
              lv = CurComponent.BaseCoeff * (CurComponent.BaseMod + lParm)
              If lv > 0 Then
                lv = Log10(lv)
              Else
                lv = 0
              End If
            Else 'exponent term, don't take log
              lv = CurComponent.ExpMod + lParm
            End If
            X(i, 1) = lv
'            If lv > 0 Then
'              X(i, 1) = Log10(lv)
'            Else
'              X(i, 1) = 0
'            End If
            Xt(1, i) = X(i, 1)
          Next i
          Uv = CurStat.PopulateMatrix
          ReDim U(MDim, MDim)
          For i = 1 To MDim
            For j = 1 To MDim
              U(i, j) = Uv(i, j)
            Next j
          Next i
          ReDim Prod(MDim, 1)
          Call mltply(Prod, U, X, MDim, MDim, 1, 1, 1, 1)
          Call mltply(Prod2, Xt, Prod, 1, MDim, 1, 1, 1, 1)
          S = Sqr(CurStat.Variance + Prod2(1, 1))
          T = 10 ^ (CurStat.tDist * S)
          retval(1, ReturnIndex) = retval(1, ReturnIndex) + ((1 / T) * (CurRegion.Discharges(ReturnIndex) / CurReturn.BCF)) * CurRegion.UserParms(1).GetValue(Me.Project.Metric)
          retval(2, ReturnIndex) = retval(2, ReturnIndex) + (T * (CurRegion.Discharges(ReturnIndex) / CurReturn.BCF)) * CurRegion.UserParms(1).GetValue(Me.Project.Metric)
        Next
        retval(1, ReturnIndex) = retval(1, ReturnIndex) / Me.GetArea(Me.Project.Metric) 'pUserRegions.Count
        retval(2, ReturnIndex) = retval(2, ReturnIndex) / Me.GetArea(Me.Project.Metric) 'pUserRegions.Count
      Next
    End If
    PredInts = retval
  End If
End Function

'Returns an error message if any region does not match the first one
'Returns empty string if all match
Private Function MismatchedRegions() As String
  Dim CurRegion As userRegion
  Dim RegionIndex As Long
  Dim ReturnIndex As Long
  Dim FirstReturns As FastCollection
  Set FirstReturns = pUserRegions(1).Region.DepVars
  For RegionIndex = 2 To pUserRegions.Count
    Set CurRegion = pUserRegions(RegionIndex)
    If CurRegion.Region.DepVars.Count <> FirstReturns.Count Then
      MismatchedRegions = "Regions '" & CurRegion.Region.Name & "' and '" & pUserRegions(1).Region.Name & "'" & vbCrLf _
                        & "have a different set of return periods, so they cannot be combined."
      Exit Function '!!!!!
    End If
    For ReturnIndex = 1 To CurRegion.Region.DepVars.Count
      If CurRegion.Region.DepVars(ReturnIndex).Name <> FirstReturns(ReturnIndex).Name Then
        MismatchedRegions = "Regions '" & CurRegion.Region.Name & "' and '" & pUserRegions(1).Region.Name & "'" & vbCrLf _
                          & "have a different set of return periods, so they cannot be combined." & vbCrLf _
                          & CurRegion.Region.DepVars(ReturnIndex).Name & " <> " & FirstReturns(ReturnIndex).Name
        Exit Function '!!!!!
      End If
    Next
  Next
  
  MismatchedRegions = ""

End Function

Public Function Intervals() As Single()
  Dim retval() As Single
  Dim FirstReturns As FastCollection
  Dim lReturn As String
  Dim i As Long
  Dim msg As String
  
  If pUserRegions.Count = 0 Then
    ReDim retval(0)
  Else
    msg = MismatchedRegions
    If Len(msg) > 0 Then
      ssMessageBox msg, vbOKOnly, "Cannot Compute Intervals"
    Else
      Set FirstReturns = pUserRegions(1).Region.DepVars
      ReDim retval(1 To FirstReturns.Count)
      For i = 1 To FirstReturns.Count
        If Left(FirstReturns(i).Name, 2) = "PK" Then
          lReturn = Mid(FirstReturns(i).Name, 3)
        Else
          lReturn = FirstReturns(i).Name
        End If
        If IsNumeric(lReturn) Then
          retval(i) = CSng(lReturn)
        End If
      Next
    End If
  End If
  Intervals = retval
End Function

Public Function Discharge(IntervalIndex As Long) As Double
  Static d() As Double
  Static GotDischarges As Integer
'  If IntervalIndex = 1 Then d = Discharges
  If GotDischarges = 0 Then
    d = Discharges
    GotDischarges = 1
  End If
  If IntervalIndex < LBound(d) Or IntervalIndex > UBound(d) Then
    Discharge = 0
  Else
    Discharge = d(IntervalIndex)
  End If
End Function

'Returns results in project units, ready to display
Public Function Discharges() As Double()
  Static lasterr As Single
  Dim RegionIndex As Long
  Dim ReturnIndex As Long
  Dim ComponentIndex As Long
  Dim AccumFlow As Double
  Dim FirstReturns As FastCollection
  Dim CurRegion As userRegion
  Dim CurReturn As nssDepVar
  Dim CurComponent As nssComponent
  Dim IntervalDischarge() As Double
  Dim RegionVal As Double
  Dim curval As Double
  Dim ExpVal As Double
  Dim StepName As String
  Dim WarnedAboutMissingComponents As Boolean
  Dim Equation As String
  Dim msg As String
  Dim StateMetric As Boolean
  
  On Error GoTo ErrHand
  
  If ROI Then
    IntervalDischarge = ComputeROIdischarge(Me, pSaveEquivYears, pSaveStdError, pSaveStdErrorPlus, pSavePredInts)
  ElseIf RevEqtn Then
    IntervalDischarge = ComputeRevEqtn(Me, pSavePredInts)
  Else
  
    If pUserRegions.Count = 0 Then
      ReDim IntervalDischarge(0)
      Discharges = IntervalDischarge
      Exit Function
    End If
    
    msg = MismatchedRegions
    If Len(msg) > 0 Then
      ssMessageBox msg, vbOKOnly, "Cannot Compute Discharges"
      GoTo LastLine
    End If
    
    Set FirstReturns = pUserRegions(1).Region.DepVars
    
    ReDim IntervalDischarge(1 To FirstReturns.Count)
    
    If pUserRegions(1).Region.Name = pProject.NationalUrban.Name Then
      StateMetric = pProject.NationalUrban.State.Metric
    Else
      StateMetric = pProject.State.Metric
    End If
    
    StepName = "For ReturnIndex = 1 To FirstReturns.Count"
    
    For ReturnIndex = 1 To FirstReturns.Count
      StepName = "IntervalDischarge(ReturnIndex) = 0"
      IntervalDischarge(ReturnIndex) = 0
      StepName = "For RegionIndex = 1 To pUserRegions.Count"
      For RegionIndex = 1 To pUserRegions.Count
        StepName = "Set CurRegion = pUserRegions(RegionIndex)"
        Set CurRegion = pUserRegions(RegionIndex)
        If ReturnIndex = 1 Then 'reset collection of discharges for this region
          CurRegion.Discharges.Clear
        End If
        StepName = "Set CurReturn = CurRegion.Region.DepVars(ReturnIndex)"
        Set CurReturn = CurRegion.Region.DepVars(ReturnIndex)
        RegionVal = CurReturn.Constant
        StepName = "For ComponentIndex = 1 To CurReturn.Components.Count"
        'Some 500-year returns are supposed to be missing components, so we don't want to warn about them
        If CurReturn.Components.Count = 0 And Not WarnedAboutMissingComponents And InStr(CurReturn.Name, "500") = 0 Then
          StepName = "Error: Components for return period " & CurReturn.Name & " of " & CurRegion.Region.Name & " are missing from the database"
          pErrorDescription = pErrorDescription & StepName & vbCrLf
          ssMessageBox StepName, vbOKOnly, "NSS Discharges"
          WarnedAboutMissingComponents = True
        End If
        For ComponentIndex = 1 To CurReturn.Components.Count
          StepName = "Set CurComponent = CurReturn.Components(ComponentIndex)"
          Set CurComponent = CurReturn.Components(ComponentIndex)
          StepName = "Select Case CurComponent.ParmID (" & CurComponent.ParmID & ")"
          Select Case CurComponent.ParmID
            Case -2: curval = findRuralValue(FirstReturns(ReturnIndex).Name, StateMetric) 'find computed rural discharge
                     If curval = -999 Then 'Could not find a rural discharge, so we can't compute this urban one
                       IntervalDischarge(ReturnIndex) = -1
                       GoTo NextInterval
                     End If
            Case -1: curval = Me.GetArea(StateMetric) 'Drainage area of whole basin, not just the part in this region
            Case 0:  curval = 0
            Case Is > 0
                   If ComponentIndex = 1 And InStr(LCase(CurComponent.Parameter.NSSName), "area") > 0 Then
                     curval = Me.GetArea(StateMetric) 'Drainage area of whole basin, not just the part in this region
                   Else
                     curval = CurRegion.UserParms(CurComponent.Parameter.NSSName).GetValue(StateMetric)
                   End If
            Case Else:
              ssMessageBox "Could not use component " & ComponentIndex & " (ParmID=" & CurComponent.ParmID & ")" & vbCr _
                   & "of region " & CurRegion.Region.Name & ", return period " & CurReturn.Name, vbOKOnly, "Compute Discharge Problem"
              GoTo LastLine
          End Select
          StepName = "If CurComponent.ExpID = 0 Then [If " & CurComponent.ExpID & " = 0 Then"
          If CurComponent.ExpID = 0 Then
            ExpVal = CurComponent.BaseExp
            Equation = ExpVal
          Else
            Equation = CurComponent.ExpMod & " + " & CurRegion.ParmIDVal(Abs(CurComponent.ExpID), StateMetric)
            ExpVal = CurComponent.ExpMod + CurRegion.ParmIDVal(Abs(CurComponent.ExpID), StateMetric)
            If CurComponent.ExpID < 0 Then
              StepName = "ExpVal = Log10(ExpVal) [ExpVal = Log10(" & ExpVal & ")]"
              Equation = "Log10(" & Equation & ")"
              ExpVal = Log10(ExpVal)
            Else
              Equation = "(" & Equation & ")"
            End If
            StepName = "ExpVal = ExpVal ^ CurComponent.ExpExp [ExpVal = " & ExpVal & " ^ " & CurComponent.ExpExp & "]"
            Equation = Equation & " ^ " & CurComponent.ExpExp
            ExpVal = ExpVal ^ CurComponent.ExpExp
            Equation = CurComponent.BaseExp & " * (" & Equation & ")"
            ExpVal = ExpVal * CurComponent.BaseExp
          End If
          Equation = "((" & curval & " + " & CurComponent.BaseMod & ") * " & CurComponent.BaseCoeff & ") ^ " & Equation
          curval = curval + CurComponent.BaseMod
          StepName = "curval = (curval * CurComponent.BaseCoeff) ^ ExpVal [curval = (" & curval & " * " & CurComponent.BaseCoeff & ") ^ " & ExpVal & "]"
          curval = (curval * CurComponent.BaseCoeff) ^ ExpVal
          RegionVal = RegionVal * curval
        Next
        StepName = "IntervalDischarge(ReturnIndex) = ..."
        IntervalDischarge(ReturnIndex) = IntervalDischarge(ReturnIndex) + RegionVal * CurRegion.UserParms(1).GetValue(StateMetric) / Me.GetArea(StateMetric)
        If StateMetric <> Me.Project.Metric Then
          If Me.Project.Metric Then
            RegionVal = RegionVal * 0.028317
          Else
            RegionVal = RegionVal / 0.028317
          End If
        End If
        CurRegion.Discharges.Add RegionVal
      Next
      StepName = "If StateMetric <> Me.Project.Metric"
      If StateMetric <> Me.Project.Metric Then
        If Me.Project.Metric Then
          IntervalDischarge(ReturnIndex) = IntervalDischarge(ReturnIndex) * 0.028317
        Else
          IntervalDischarge(ReturnIndex) = IntervalDischarge(ReturnIndex) / 0.028317
        End If
      End If
NextInterval:
    Next
    If InStr(FirstReturns(FirstReturns.Count).Name, "500") > 0 And _
       IntervalDischarge(FirstReturns.Count) < 0.001 Then
      'not able to compute 500 year flood, extrapolate it
      StepName = "nss500"
      IntervalDischarge(FirstReturns.Count) = nss500(IntervalDischarge, Me.Intervals)
      pExtrapolated = True
    End If
  End If
  Discharges = IntervalDischarge
  Exit Function

ErrHand:
    msg = "Error: " & Err.Description & vbCrLf & vbCrLf _
           & StepName & vbCrLf _
           & "ReturnIndex = " & ReturnIndex & vbCrLf _
           & "RegionIndex = " & RegionIndex & vbCrLf _
           & "ComponentIndex = " & ComponentIndex
    If Not CurComponent Is Nothing Then
      With CurComponent
        msg = msg & vbCrLf _
            & " Parameter Name = " & .Parameter.NSSName & vbCrLf _
            & " BaseMod = " & .BaseMod & vbCrLf _
            & " BaseCoeff = " & .BaseCoeff & vbCrLf _
            & " BaseExp = " & .BaseExp & vbCrLf _
            & " ExpMod = " & .ExpMod & vbCrLf _
            & " ExpID = " & .ExpID & vbCrLf _
            & " ExpExp = " & .ExpExp & vbCrLf _
            & "Equation: " & Equation
      End With
    End If
    pErrorDescription = pErrorDescription & msg
  'If Timer - lasterr > 1 Then
    ssMessageBox msg, vbExclamation, "Error computing discharges"
  '  lasterr = Timer
  'End If

LastLine:
  Discharges = IntervalDischarge 'Return empty array if there is a problem
End Function

Private Function findRuralValue(ReturnPeriod As String, WantMetric As Boolean)
  Dim ReturnIndex As Long
  Dim RuralReturns() As Single
  Dim lReturn As Single
  
  findRuralValue = -999
  If Me.RuralScenario Is Nothing Then
    pErrorDescription = pErrorDescription & "No rural flow value available for urban computation" & vbCrLf ', vbExclamation, "Error computing discharges"
  Else
    If Left(ReturnPeriod, 2) = "PK" Then
      lReturn = CSng(Mid(ReturnPeriod, 3))
    Else
      lReturn = CSng(ReturnPeriod)
    End If
    RuralReturns = Me.RuralScenario.Intervals
    For ReturnIndex = LBound(RuralReturns) To UBound(RuralReturns)
      If Abs(RuralReturns(ReturnIndex) - lReturn) < 0.001 Then
        findRuralValue = pProject.DB.FlowUnits.Convert(Me.RuralScenario.Discharge(ReturnIndex), pProject.Metric, WantMetric)
        Exit For
      End If
    Next
  End If

End Function

'Returns results in project units, ready to display
Public Function WeightedDischarges() As Double()
  Dim ReturnIndex As Long
  Dim IntervalDischarges() As Double
  Dim ThisDischarge As Double
  Dim Years() As Double
  Dim ThisInterval As nssDepVar
  IntervalDischarges = Me.Discharges
  Years = Me.EquivalentYears
  If Not Me.ROI Then
    For ReturnIndex = 1 To UBound(IntervalDischarges)
    
'      If Me.Project.State.Metric = Me.Project.Metric Then
        ThisDischarge = IntervalDischarges(ReturnIndex)
'      Else 'Convert back to state units from project units before weighting
'        If Me.Project.Metric Then
'          ThisDischarge = IntervalDischarges(ReturnIndex) / 0.028317
'        Else
'          ThisDischarge = IntervalDischarges(ReturnIndex) * 0.028317
'        End If
'      End If
      Set ThisInterval = pUserRegions(1).Region.DepVars(ReturnIndex)
      IntervalDischarges(ReturnIndex) = _
        pWeight.WeightedEstimate(ThisInterval.Name, _
                                 ThisDischarge, _
                                 Years(ReturnIndex), _
                                 Me.GetArea(pUserRegions(1).Region.State.Metric), _
                                 ThisInterval.ExpDA)
      
'      If Me.Project.State.Metric <> Me.Project.Metric Then 'Convert from state units to project units
'        If Me.Project.Metric Then
'          IntervalDischarges(ReturnIndex) = IntervalDischarges(ReturnIndex) * 0.028317
'        Else
'          IntervalDischarges(ReturnIndex) = IntervalDischarges(ReturnIndex) / 0.028317
'        End If
'      End If
    Next
  End If
  WeightedDischarges = IntervalDischarges
End Function

Public Property Get Matrix() As Variant
  If IsEmpty(pMatrix) Then 'SetMatrix
    'read from binary file
    GetMatrixBinary 2, pMatrix
  End If
  Matrix = pMatrix
End Property
Private Sub SetMatrix()
  Dim sql$
  Dim i&, j&
  Dim myRec As Recordset
  
  sql = "SELECT * FROM StateMatrices" & _
        " WHERE StID='" & Me.Project.State.Code & _
        "' ORDER BY RowID, ColID;"
  Set myRec = Me.Project.DB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    .MoveLast
    ReDim pMatrix(!RowID, !ColID)
    .MoveFirst
    While Not .EOF
      pMatrix(!RowID, !ColID) = !Value
      .MoveNext
    Wend
  End With
  For i = 1 To UBound(pMatrix, 1)
    For j = 1 To i
      pMatrix(j, i) = pMatrix(i, j)
    Next j
  Next i
End Sub

Public Property Get RHO() As Variant
  If IsEmpty(pRHO) Then 'SetRHO
    'read from binary file
    GetMatrixBinary 1, pRHO
  End If
  RHO = pRHO
End Property
Public Property Let RHO(ByVal newValue As Variant)
  pRHO = newValue
End Property
Private Sub SetRHO()
  Dim sql$
  Dim myRec As Recordset
  Dim i&, j&
  
  sql = "SELECT * FROM RHO " & _
        "WHERE StID='" & Me.Project.State.Code & _
        "' ORDER BY RowID, ColID;"
  Set myRec = Me.Project.DB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    .MoveLast
    ReDim pRHO(!RowID, !ColID)
    .MoveFirst
    While Not .EOF
      pRHO(!RowID, !ColID) = !Value
      .MoveNext
    Wend
  End With
  For i = 1 To UBound(pRHO, 1)
    For j = 1 To i
      pRHO(j, i) = pRHO(i, j)
    Next j
  Next i
End Sub

Public Property Get CF() As Variant
  If IsEmpty(pCF) Then SetCF
  CF = pCF
End Property
Public Property Let CF(ByVal newValue As Variant)
  pCF = newValue
End Property
Private Sub SetCF()
  Dim myRec As Recordset
  Dim i&
  
  Set myRec = Me.Project.DB.DB.OpenRecordset("ClimateFactor", dbOpenDynaset)
  With myRec
    .MoveLast
    ReDim pCF(1 To .RecordCount, 1 To 5)
    .MoveFirst
    While Not .EOF
      pCF(.AbsolutePosition + 1, 1) = !X
      pCF(.AbsolutePosition + 1, 2) = !Y
      pCF(.AbsolutePosition + 1, 3) = !C2
      pCF(.AbsolutePosition + 1, 4) = !C25
      pCF(.AbsolutePosition + 1, 5) = !C100
      .MoveNext
    Wend
  End With
End Sub

Public Function MaxFloodEnvelope(WantMetric As Boolean) As Double
  If pArea > 0 And Me.RegCrippenBue >= 1 And Me.RegCrippenBue <= 17 Then
    MaxFloodEnvelope = pProject.DB.Units("13").Convert(CrippenBueMaxFloodEnvelope(Me.GetArea(False), Me.RegCrippenBue), False, WantMetric)
  End If
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
Private Sub Class_Initialize()
  pName = "Unnamed Scenario"
  pUrban = False
  Set pWeight = New nssWeight
  ReDim pSaveEquivYears(0)
  ReDim pSaveStdError(0)
  pExtrapolated = False
End Sub

Public Function EstimateString() As String
  Dim Flows() As Double
  Dim Years() As Double
  Dim StdError() As Double
  Dim PredError() As Double
  Dim EstError() As Double
  Dim StdErrorPlus() As Double
  Dim Index As Integer
  Dim IntervalText() As String
  Dim FlowText() As String
  Dim StdErrText() As String
  Dim PredInts() As Double
  Dim PredIntsText() As String, PredIntsText2() As String
  Dim YearsText() As String
  Dim nIntervals As Integer
  Dim txt As String, lab1 As String, lab2 As String
  Dim lLen As Integer
  Dim labUnits As String
  Dim UsePredErr As Boolean, UseEstErr As Boolean
  Dim useMetric As Boolean, YearsOn As Boolean
  Dim AllNumeric As Boolean
  Dim lDepVarName As String
  Dim lUnitCode As Integer
  Dim lUnitStr As String
  
  useMetric = Me.Project.Metric
  
  txt = MismatchedRegions
  
  'the following code will become simpler when database mods are made
  'to reference the StatLabel from the DepVars table
'  'for starters, assume flow
'  If useMetric Then
'    lUnitStr = "    m3/s "
'  Else
'    lUnitStr = "    cfs  "
'  End If
'  If Me.UserRegions(1).Region.DepVars(1).Units.ID <> 13 Then
    lUnitStr = Me.UserRegions(1).Region.DepVars(1).Units.Abbrev(useMetric)
    lLen = Len(lUnitStr)
    If lLen > 9 Then 'too long, need to truncate
      lUnitStr = Left(lUnitStr, 9)
    ElseIf lLen < 9 Then
      lUnitStr = Space(9 - lLen) & lUnitStr & " "
    End If
'  End If
'  'for starters, assume cfs
'  If useMetric Then
'    lUnitStr = "    m3/s "
'  Else
'    lUnitStr = "    cfs  "
'  End If
'  lDepVarName = Me.UserRegions(1).Region.DepVars(1).Name
'  If Me.Project.DB.Parameters.KeyExists(lDepVarName) Then
'    lUnitCode = Me.Project.DB.Parameters(lDepVarName).ConvFlag
'    If lUnitCode <> 13 Then 'something besides flow (cfs or m3/s)
'      If Me.Project.DB.Units.KeyExists(CStr(lUnitCode)) Then
'        If useMetric Then
'          lUnitStr = Me.Project.DB.Units(CStr(lUnitCode)).MetricLabel
'        Else
'          lUnitStr = Me.Project.DB.Units(CStr(lUnitCode)).EnglishLabel
'        End If
'      End If
'      lLen = Len(lUnitStr)
'      If lLen > 9 Then 'too long, need to truncate
'        lUnitStr = Left(lUnitStr, 9)
'      ElseIf lLen < 9 Then
'        lUnitStr = Space(8 - lLen) & lUnitStr & " "
'      End If
'    End If
'  End If
  
  AllNumeric = True
  If Len(txt) > 0 Then
    EstimateString = txt
  ElseIf Me.Urban Then
    Flows = Me.Discharges
    StdError = Me.StdError
'    txt = "Recurrence     Peak,  Standard  " & vbCrLf _
'        & "Interval, yrs   cfs   Error, %  "
    txt = "              Value,  Standard  " & vbCrLf _
        & "Statistic  " & lUnitStr & "  Error, %  "
    If useMetric Then txt = ReplaceString(txt, " cfs", "m3/s")
    nIntervals = UBound(Flows)
    ReDim IntervalText(nIntervals)
    ReDim FlowText(nIntervals)
    ReDim StdErrText(nIntervals)
    For Index = 1 To nIntervals
      If Flows(Index) > -1 Then 'check for no urban estimate due to no rural at this interval
        IntervalText(Index) = Me.UserRegions(1).Region.DepVars(Index).Name
        If IsNumeric(IntervalText(Index)) Then
          IntervalText(Index) = SignificantDigits(Me.UserRegions(1).Region.DepVars(Index).Name, 3)
        ElseIf Len(IntervalText(Index)) < 7 Then
          IntervalText(Index) = IntervalText(Index) & Space(7 - Len(IntervalText(Index)))
        End If
        If IsNumeric(IntervalText(Index)) Then
          IntervalText(Index) = SignificantDigits(Me.UserRegions(1).Region.DepVars(Index).Name, 3)
        Else
          AllNumeric = False
        End If
        FlowText(Index) = SignificantDigits(Flows(Index), 3)
        If StdError(Index) > 0 Then
          StdErrText(Index) = SignificantDigits(StdError(Index), 2)
        End If
      End If
    Next
    If AllNumeric Then DecimalAlign IntervalText, True, True, 7
    DecimalAlign FlowText, True, True, 11
    DecimalAlign StdErrText, True, True, 8
    For Index = 1 To nIntervals
      If Len(Trim(IntervalText(Index))) > 0 Then
        txt = txt & vbCrLf & IntervalText(Index) _
                    & "  " & FlowText(Index) _
                           & StdErrText(Index)
      End If
    Next
    
  Else
    If Me.ROI Then
'      txt = "Recurrence     Peak,   Pred. Intervals   Standard Error    Equivalent" & vbCrLf _
'          & "Interval, yrs   cfs       Low     High     (-%)    (+%)    Years"
      txt = "              Value,   Pred. Intervals   Standard Error    Equivalent" & vbCrLf _
          & "Statistic  " & lUnitStr & "      Low     High     (-%)    (+%)    Years"
      Flows = Me.Discharges
      StdErrorPlus = Me.StdErrorPlus
      StdError = Me.StdError
      Years = Me.EquivalentYears
      PredInts = Me.PredInts
    Else
      Years = Me.EquivalentYears
      StdError = Me.StdError
      PredError = Me.PredictionError(UsePredErr)
      EstError = Me.EstimateError(UseEstErr)
      If UsePredErr Then 'fill Standard Error values with more desirable Prediction Errors
        For Index = 1 To UBound(StdError)
          StdError(Index) = PredError(Index)
        Next Index
      ElseIf UseEstErr Then 'fill Standard Error values with more desirable Estimation Errors
        For Index = 1 To UBound(StdError)
          StdError(Index) = EstError(Index)
        Next Index
      End If
      
      'use generic labels to handle different kinds of stats
      lab1 = "             Value,"
'      lab2 = "Statistic " & lUnitStr
      'use the following 3 lines when DepVars Units in place
      lab2 = "Statistic"
      labUnits = Me.UserRegions(1).Region.DepVars(1).Units.Abbrev(useMetric)
      lab2 = lab2 & Space((14 - Len(labUnits)) / 2) & labUnits

      If Me.LowFlow Then
        Flows = Me.Discharges
      Else
        Flows = Me.WeightedDischarges
      End If
      
      If Me.UsePredInts And Me.Weight.WeightType = 0 And Not Me.ParmOutOfRange Then
        PredInts = Me.PredInts
        txt = lab1 & "   Pred. Intervals   Standard  Equivalent" & vbCrLf & _
              lab2 & "      Low     High   Error, %  Years"
      ElseIf Not Me.ParmOutOfRange Then 'no pred intervals or don't show 'em for weighted estimates or parm value out or range
        txt = lab1 & "  Standard  Equivalent" & vbCrLf _
            & lab2 & "  Error, %  Years"
      Else
        txt = lab1 & vbCrLf & lab2
      End If
    End If
    
    'next line goes away when DepVars Units in place
    'If useMetric Then txt = ReplaceString(txt, " cfs", "m3/s")
    
    If Me.Weight.WeightType > 0 Then 'Don't show std err for weighted estimates
      txt = ReplaceString(txt, "Standard  ", "")
      txt = ReplaceString(txt, "Error, %  ", "")
    End If
    If Me.Weight.WeightType = 2 Then 'Don't show equiv years for weighted estimates for ungaged sites
      txt = ReplaceString(txt, "Equivalent", "")
      txt = ReplaceString(txt, "Years", "")
    End If
    If UsePredErr Then
      txt = ReplaceString(txt, "  Standard", "Prediction")
    ElseIf UseEstErr Then
      txt = ReplaceString(txt, "Standard", "Estimate")
    End If

    nIntervals = UBound(Flows)
    ReDim IntervalText(nIntervals)
    ReDim FlowText(nIntervals)
    ReDim StdErrText(nIntervals)
    If Me.UsePredInts And Me.Weight.WeightType = 0 Then
      ReDim PredIntsText(nIntervals)
      ReDim PredIntsText2(nIntervals)
    End If
    ReDim YearsText(nIntervals)
    YearsOn = False 'assume no equivalent years to display
    For Index = 1 To nIntervals
      
      IntervalText(Index) = Me.UserRegions(1).Region.DepVars(Index).Name
      If IsNumeric(IntervalText(Index)) Then
        IntervalText(Index) = SignificantDigits(Me.UserRegions(1).Region.DepVars(Index).Name, 3)
      ElseIf Len(IntervalText(Index)) < 7 Then
        IntervalText(Index) = IntervalText(Index) & Space(7 - Len(IntervalText(Index)))
      End If
      If Not IsNumeric(IntervalText(Index)) Then AllNumeric = False
              
      FlowText(Index) = SignificantDigits(Flows(Index), 3)
      
      If Not Me.ParmOutOfRange Then 'no Errors or PIs reported if parms out of range
      If Me.UsePredInts And Me.Weight.WeightType = 0 Then
        If PredInts(1, Index) > 0 Then
          PredIntsText(Index) = SignificantDigits(PredInts(1, Index), 3)
        End If
        If PredInts(2, Index) > 0 Then
          PredIntsText2(Index) = SignificantDigits(PredInts(2, Index), 3)
        End If
      End If
      If StdError(Index) <= 0 And Not Me.ROI Then
        StdErrText(Index) = ""
        'If Me.Weight.WeightType = 0 Then StdErrText = Space(8) Else StdErrText = ""
      Else
        If StdError(Index) < 5 And Not Me.ROI Then
          StdErrText(Index) = SignificantDigits(StdError(Index), 3) 'Log units to 3 sig digits
          txt = ReplaceString(txt, "Error, %", "Err, log")
        Else
          StdErrText(Index) = SignificantDigits(StdError(Index), 2) 'Percent to two sig digits
        End If
        'StdErrText = StrPad(StdErrText, 8)
      End If
      
      If Me.ROI Then
        If StdErrorPlus(Index) = 0 Then
          'StdErrText = StdErrText & Space(8)
        Else
          StdErrText(Index) = StrPad(StdErrText(Index), 6) & StrPad(SignificantDigits(StdErrorPlus(Index), 2), 8)
        End If
      End If
      
      If Years(Index) = 0 Then ' Or Me.Weight.WeightType > 1 Then
        YearsText(Index) = "" 'Space(10)
      Else
        YearsText(Index) = SignificantDigits(Years(Index), 2)
        YearsOn = True
      End If
      End If
    Next
    If Not YearsOn Then
      txt = ReplaceString(txt, "Equivalent", "")
      txt = ReplaceString(txt, "Years", "")
    End If
    If AllNumeric Then DecimalAlign IntervalText, True, True, 7
    If Me.LowFlow Then
      DecimalAlign FlowText, True, True, 9
    Else
      DecimalAlign FlowText, True, True, 11
    End If
    
    If pExtrapolated Then 'add "*" to extrapolated 500-year value
      lLen = Len(FlowText(nIntervals)) - 1
      FlowText(nIntervals) = Left(FlowText(nIntervals), lLen) & "*"
    End If
    
    DecimalAlign YearsText, True, True, 10
    If Me.UsePredInts And Me.Weight.WeightType = 0 Then
      DecimalAlign PredIntsText, True, True, 9
      DecimalAlign PredIntsText2, True, True, 9
    End If
    If Me.Weight.WeightType = 0 Then
      DecimalAlign StdErrText, True, True, 8
    End If
    For Index = 1 To nIntervals
      If Me.UsePredInts And Me.Weight.WeightType = 0 Then
        txt = txt & vbCrLf & IntervalText(Index) & "  " _
                           & FlowText(Index) _
                           & PredIntsText(Index) _
                           & PredIntsText2(Index) & Space(2) _
                           & StdErrText(Index)
      Else
        txt = txt & vbCrLf & IntervalText(Index) & "  " _
                           & FlowText(Index) _
                           & StdErrText(Index)
      End If
      If YearsOn Then txt = txt & YearsText(Index)
    Next
    If pExtrapolated Then txt = txt & vbCrLf & vbCrLf & "  *Extrapolated value"
    If Me.RegCrippenBue > 0 Then
      txt = txt & vbCrLf & "maximum: " _
                & SignificantDigits(Me.MaxFloodEnvelope(useMetric), 3) _
                & " (for C&B region " & Me.RegCrippenBue & ")"
    End If
  End If
  If Len(pErrorDescription) > 0 Then
    txt = txt & vbCrLf & Me.ErrorDescription
  End If
  EstimateString = txt
End Function

Private Sub GetMatrixBinary(MatType As Integer, MatVar As Variant)
  Dim i As Long, j As Long, Fun As Long, Fnv As Long, nv As Long
  Dim FileType As Integer '1 - RHO, 2 - REC
  Dim Filename As String
  Dim iVal As Integer
  Dim rval As Single
  Dim FTotVal As Single, TotVal As Single
  Dim ff As New ATCoFindFile

  If MatType = 1 Then 'RHO
'    Filename = GetSetting("StreamStatsDB", "Defaults", Me.Project.State.Abbrev & "_RHOImportFile")
    ff.SetDialogProperties "Please locate ROI RHO binary file", Me.Project.State.Abbrev & ".rho.bin"
    ff.SetRegistryInfo "StreamStatsDB", "Defaults", Me.Project.State.Abbrev & "RHOBinary"
    Filename = ff.GetName
  Else 'REC
'    Filename = GetSetting("StreamStatsDB", "Defaults", Me.Project.State.Abbrev & "_MConImportFile")
    ff.SetDialogProperties "Please locate ROI REC binary file", Me.Project.State.Abbrev & ".rec.bin"
    ff.SetRegistryInfo "StreamStatsDB", "Defaults", Me.Project.State.Abbrev & "RECBinary"
    Filename = ff.GetName
  End If
  Fun = FreeFile(0)
  Open Filename For Binary As #Fun
  Get #Fun, , FileType
  Get #Fun, , Fnv
  Get #Fun, , FTotVal
  'determine size of matrix array
  i = Fix(Sqr(CDbl(2 * Fnv)))
  ReDim MatVar(i, i)
  If FileType = 1 Then 'RHO values are reals
    For i = 1 To UBound(MatVar)
      For j = 1 To i
        Get #Fun, , rval
        MatVar(i, j) = rval
        nv = nv + 1
        TotVal = TotVal + rval
      Next j
    Next i
  Else 'REC values are integers
    For i = 1 To UBound(MatVar)
      For j = 1 To i
        Get #Fun, , iVal
        MatVar(i, j) = iVal
        nv = nv + 1
        TotVal = TotVal + iVal
      Next j
    Next i
  End If
  If Fnv <> nv Or Abs(FTotVal - TotVal) > 0.001 Then 'didn't read expected number or total of values
    ssMessageBox "PROBLEM: read different count or total of values than expected from binary matrix file!", vbExclamation
  Else 'fill out other half of matrix
    For i = 1 To UBound(MatVar)
      For j = 1 To i
        MatVar(j, i) = MatVar(i, j)
      Next j
    Next i
  End If
End Sub

Private Function ComputeRevEqtn(aScen As nssScenario, aPredInts() As Double) As Double()

End Function
