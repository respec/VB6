VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nssWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private pMaxIntervals As Long
Private pWeightType As Long '0=no weight, 1=Gaged, 2=Ungaged
Private pGagedValue() As Double
Private pGagedYears() As Double
Private pIntervals() As String
Private pIntervalsSet As Long 'Number of intervals with GagedYears and/or GagedValues
Private pAreaGaged As Double
'Private pAreaUnGaged As Double

Public Function Copy() As nssWeight
  Dim i As Long
  Dim retval As nssWeight
  Set retval = New nssWeight
  With retval
    .AreaGaged = Me.AreaGaged
    .WeightType = Me.WeightType
    For i = 1 To pIntervalsSet
      .SetGagedYears pIntervals(i), pGagedYears(i)
      .SetGagedValue pIntervals(i), pGagedValue(i)
    Next
  End With
  Set Copy = retval
End Function

'0=no weight, 1=Gaged, 2=Ungaged
Public Property Get WeightType() As Long
  WeightType = pWeightType
End Property
Public Property Let WeightType(newValue As Long)
  pWeightType = newValue
End Property

Public Property Get AreaGaged() As Double
  AreaGaged = pAreaGaged
End Property
Public Property Let AreaGaged(newValue As Double)
  pAreaGaged = newValue
End Property

'Public Property Get AreaUnGaged() As Double
'  AreaUnGaged = pAreaUnGaged
'End Property
'Public Property Let AreaUnGaged(newValue As Double)
'  pAreaUnGaged = newValue
'End Property

Private Sub InitArrays()
  ReDim pGagedValue(pMaxIntervals)
  ReDim pGagedYears(pMaxIntervals)
  ReDim pIntervals(pMaxIntervals)
End Sub
Private Sub EnlargeArrays()
  pMaxIntervals = pMaxIntervals * 2
  ReDim Preserve pGagedValue(pMaxIntervals)
  ReDim Preserve pGagedYears(pMaxIntervals)
  ReDim Preserve pIntervals(pMaxIntervals)
End Sub

'Interval is the label of the interval, i.e. "10" for the 10-year interval
'Area should be in the original state units rather than the project units since we store AreaGaged that way
Public Function WeightedEstimate(ByVal Interval As String, ByVal RegressionFlow As Double, RegressionYears As Double, Area As Double, AreaExponent As Double) As Double
  Dim i&
  Dim Qtgw As Double
  Dim dArea2overG As Double
  
  On Error GoTo ErrWeight
  Select Case pWeightType
    Case 0: WeightedEstimate = RegressionFlow
    Case 1, 2
      GoSub FindInterval
      If Abs(pGagedValue(i)) < 0.001 Then
        WeightedEstimate = RegressionFlow
      ElseIf pWeightType = 1 Then
        If Abs(pGagedYears(i)) < 0.001 Then
          WeightedEstimate = RegressionFlow
        ElseIf RegressionYears > 0 Then
          'subtract pGagedYears from RegressionYears since RegressionYears includes pGagedYears
          WeightedEstimate = 10 ^ ((pGagedYears(i) * Log10(pGagedValue(i)) + (RegressionYears - pGagedYears(i)) * Log10(RegressionFlow)) / (pGagedYears(i) + (RegressionYears - pGagedYears(i))))
        Else 'no equivalent years for equation, just use entered value
          WeightedEstimate = pGagedValue(i)
        End If
      Else
        If Abs(pAreaGaged) < 0.001 Then
          WeightedEstimate = RegressionFlow
        Else
          dArea2overG = 2 * Abs(pAreaGaged - Area) / pAreaGaged
'         original documentation applied exponent to area ratio term
         WeightedEstimate = dArea2overG * RegressionFlow + (1 - dArea2overG) * (Area / pAreaGaged) ^ AreaExponent * pGagedValue(i)
'         temporary replacement will not apply an exponent to this term (Jan 2002)
'          WeightedEstimate = dArea2overG * RegressionFlow + (1 - dArea2overG) * (Area / pAreaGaged) * pGagedValue(i)
'!!!!     future version will look up exponent value from data base (will be a value of 1 for most states)
        End If
      End If
    Case Else
      ssMessageBox "Unknown Weight Type '" & pWeightType & "' in WeightedEstimate", vbOKOnly, "NSS Weight"
  End Select
  
  Exit Function

FindInterval:
  For i = 1 To pIntervalsSet
    If pIntervals(i) = Interval Then
      Return
    End If
  Next
  ssMessageBox "Unknown Interval '" & Interval & "' in WeightedEstimate", vbOKOnly, "NSS Weight"
  Exit Function

ErrWeight:
  ssMessageBox "Error '" & Err.Description & "' in WeightedEstimate", vbOKOnly, "NSS Weight"
End Function

Public Sub SetGagedValue(ByVal Interval As String, ByVal GagedValue As Double)
  Dim i&
  If pIntervalsSet = 0 Then InitArrays
  For i = 1 To pIntervalsSet
    If pIntervals(i) = Interval Then
      pGagedValue(i) = GagedValue
      Exit Sub
    End If
  Next
  'Didn't already find interval, so add it
  If i > pMaxIntervals Then EnlargeArrays
  pIntervals(i) = Interval
  pGagedValue(i) = GagedValue
  pIntervalsSet = i
End Sub
Public Function GetGagedValue(ByVal Interval As String) As Double
  Dim i&
  For i = 1 To pIntervalsSet
    If pIntervals(i) = Interval Then
      GetGagedValue = pGagedValue(i)
      Exit Function
    End If
  Next
End Function

Public Sub SetGagedYears(ByVal Interval As String, ByVal GagedYears As Double)
  Dim i&
  If pIntervalsSet = 0 Then InitArrays
  For i = 1 To pIntervalsSet
    If pIntervals(i) = Interval Then
      pGagedYears(i) = GagedYears
      Exit Sub
    End If
  Next
  'Didn't already find interval, so add it
  If i > pMaxIntervals Then EnlargeArrays
  pIntervals(i) = Interval
  pGagedYears(i) = GagedYears
  pIntervalsSet = i
End Sub
Public Function GetGagedYears(ByVal Interval As String) As Double
  Dim i&
  For i = 1 To pIntervalsSet
    If pIntervals(i) = Interval Then
      GetGagedYears = pGagedYears(i)
      Exit Function
    End If
  Next
End Function

Public Property Get XML() As String
  Dim retval As String
  Dim i As Long
  retval = "<weight type=""" & pWeightType & """"
  Select Case pWeightType
    Case 0: retval = retval & ">"  'No body is needed: 0 = no weighting
    Case 1: retval = retval & ">" & vbCrLf
          For i = 1 To pIntervalsSet
            retval = retval & "  <weightinterval year=""" & pIntervals(i) & """ gagedvalue=""" & pGagedValue(i) & """ gagedyears=""" & pGagedYears(i) & """/>" & vbCrLf
          Next i
    Case 2
          retval = retval & " areagaged=""" & pAreaGaged & """>" & vbCrLf
          For i = 1 To pIntervalsSet
            retval = retval & "  <weightinterval year=""" & pIntervals(i) & """ gagedvalue=""" & pGagedValue(i) & """ gagedyears=""" & pGagedYears(i) & """/>" & vbCrLf
          Next i
  End Select
  retval = retval & "</weight>" & vbCrLf
  XML = retval
End Property
Public Property Let XML(newValue As String)
  Dim WeightXML As clsXMLelement
  Dim vIntervalXML As Variant
  Dim year As String
  If InStr(LCase(newValue), "weight") > 0 Then
    Set WeightXML = New clsXMLelement
    WeightXML.SetString newValue
    pWeightType = WeightXML.AttributeValue("type", "0")
    If pWeightType > 0 Then
      If pWeightType = 2 Then pAreaGaged = WeightXML.AttributeValue("areagaged", "0")
      pMaxIntervals = WeightXML.SubElements.Count
      InitArrays
      For Each vIntervalXML In WeightXML.SubElements
        year = vIntervalXML.AttributeValue("year", "1")
        Me.SetGagedValue year, vIntervalXML.AttributeValue("gagedvalue", "0")
        Me.SetGagedYears year, vIntervalXML.AttributeValue("gagedyears", "0")
      Next
    End If
  End If
End Property

Private Sub Class_Initialize()
  pMaxIntervals = 10
End Sub
