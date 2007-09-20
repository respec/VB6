VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nssProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private pState As nssState
Private pDB As nssDatabase
Private pCurrentRuralScenario As Long
Private pCurrentUrbanScenario As Long
Private pRuralScenarios As FastCollection 'of nssScenario
Private pUrbanScenarios As FastCollection 'of nssScenario
Private pFileName As String
Private pName As String
Private pUserName As String
Private pMetric As Boolean
Private pChanged As Boolean
Private pNationalState As nssState
Private pNationalUrban As nssRegion

Public Event Edited()

Public Sub RaiseEdited()
  RaiseEvent Edited
End Sub

Public Function UserMessage(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "NSS") As VbMsgBoxResult
  UserMessage = ssMessageBox(Prompt, Buttons, Title)
End Function

Public Property Get NationalUrban() As nssRegion
  Set NationalUrban = pNationalUrban
End Property

Public Property Get State() As nssState
  Set State = pState
End Property
Public Property Set State(ByVal newValue As nssState)
  Set pState = newValue
End Property

Public Property Get Changed() As Boolean
  Changed = pChanged
End Property

Public Property Let HelpFile(newValue As String)
  App.HelpFile = newValue
End Property

Public Property Get Metric() As Boolean
  Metric = pMetric
End Property
Public Property Let Metric(newValue As Boolean)
  If newValue <> pMetric Then
    pMetric = newValue
    pChanged = True
  End If
End Property

Public Property Get Name() As String
  Name = pName
End Property
Public Property Let Name(newValue As String)
  If newValue <> pName Then
    pName = newValue
    pChanged = True
  End If
End Property

Public Property Get UserName() As String
  UserName = pUserName
End Property
Public Property Let UserName(newValue As String)
  If newValue <> pUserName Then
    pUserName = newValue
    pChanged = True
  End If
End Property

Public Property Get CurrentRuralScenario() As Long
  If pCurrentRuralScenario > pRuralScenarios.Count Then
     pCurrentRuralScenario = pRuralScenarios.Count
  End If
  CurrentRuralScenario = pCurrentRuralScenario
End Property
Public Property Let CurrentRuralScenario(newValue As Long)
  If newValue > 0 And newValue <= pRuralScenarios.Count And newValue <> pCurrentRuralScenario Then
    pCurrentRuralScenario = newValue
    'pChanged = True
    RaiseEvent Edited
  End If
End Property

Public Property Get CurrentUrbanScenario() As Long
  If pCurrentUrbanScenario > pUrbanScenarios.Count Then
     pCurrentUrbanScenario = pUrbanScenarios.Count
  End If
  CurrentUrbanScenario = pCurrentUrbanScenario
End Property
Public Property Let CurrentUrbanScenario(newValue As Long)
  If newValue > 0 And newValue <= pUrbanScenarios.Count And newValue <> pCurrentUrbanScenario Then
    pCurrentUrbanScenario = newValue
    'pChanged = True
    RaiseEvent Edited
  End If
End Property

Public Property Get XML() As String
  Dim retval As String
  Dim vScenario As Variant, vRegion As Variant, vParameter As Variant
  Dim Scenario As nssScenario
  
  retval = "<NSSproject"
  retval = retval & " name=""" & Me.Name & """"
  retval = retval & " username=""" & Me.UserName & """"
  retval = retval & " state=""" & Me.State.Code & """"
  retval = retval & " metric=""" & Me.Metric & """"
  retval = retval & " currentrural=""" & Me.CurrentRuralScenario & """"
  retval = retval & " currenturban=""" & Me.CurrentUrbanScenario & """"
  retval = retval & ">" & vbCrLf

  For Each vScenario In Me.RuralScenarios
    GoSub AddScenario
  Next
  For Each vScenario In Me.UrbanScenarios
    GoSub AddScenario
  Next

  retval = retval & "</NSSproject>"
  XML = retval

  Exit Sub

AddScenario:
  
  Set Scenario = vScenario
  retval = retval & "<NSSScenario"
  retval = retval & " name=""" & Scenario.Name & """"
  If Scenario.RegCrippenBue > 0 Then
    retval = retval & " crippenbue=""" & Scenario.RegCrippenBue & """"
  End If
  retval = retval & " urban=""" & Scenario.Urban & """"
  If Not Scenario.RuralScenario Is Nothing Then
    retval = retval & " ruralname=""" & Scenario.RuralScenario.Name & """"
  End If
  retval = retval & " area=""" & Scenario.GetArea(Me.State.Metric) & """"
  retval = retval & ">" & vbCrLf
  If Scenario.Weight.WeightType > 0 Then
    retval = retval & Scenario.Weight.XML
  End If
  
  For Each vRegion In Scenario.UserRegions
    retval = retval & "<NSSRegion"
    retval = retval & " name=""" & vRegion.Region.Name & """"
    For Each vParameter In vRegion.UserParms
      retval = retval & " " & vParameter.Parameter.NSSName & "=""" & vParameter.GetValue(Me.State.Metric) & """"
    Next
    retval = retval & "/>" & vbCrLf
  Next
  
  retval = retval & "</NSSScenario>" & vbCrLf
  Return
End Property
Public Property Let XML(newValue As String)
  Dim projectXML As clsXMLelement
  Dim vScenario As Variant, vRegion As Variant, vParameter As Variant
  Dim ScenarioXML As clsXMLelement
  Dim newScenario As nssScenario
  Dim newRegion As userRegion
  Dim RuralName As String
  Dim searchRuralScenarios As Variant
  Dim RegionName As String
  
  Set projectXML = New clsXMLelement
  projectXML.SetString newValue
  
  If projectXML.tag <> "nssproject" Then
    ssMessageBox "Expected XML NSSproject, but did not find it.", vbOKOnly, "Failed to read NSS project XML"
  Else
    Clear
    'pFileName = newValue
    pName = projectXML.AttributeValue("name")
    pUserName = projectXML.AttributeValue("username")
    If LCase(projectXML.AttributeValue("metric")) = "true" Then pMetric = True Else pMetric = False
    Set Me.State = pDB.States(projectXML.AttributeValue("state"))
    pCurrentRuralScenario = projectXML.AttributeValue("currentrural", 0)
    pCurrentUrbanScenario = projectXML.AttributeValue("currenturban", 0)
    For Each vScenario In projectXML.SubElements
      If vScenario.tag = "nssscenario" Then
        Set ScenarioXML = vScenario
        Set newScenario = New nssScenario
        Set newScenario.Project = Me
        newScenario.Name = ScenarioXML.AttributeValue("name", "Unnamed")
        newScenario.RegCrippenBue = ScenarioXML.AttributeValue("crippenbue", 0)
        newScenario.SetArea ScenarioXML.AttributeValue("area", "0"), Me.State.Metric
        If LCase(ScenarioXML.AttributeValue("urban")) = "true" Then
          newScenario.Urban = True
          pUrbanScenarios.Add newScenario, LCase(newScenario.Name)
          RuralName = ScenarioXML.AttributeValue("ruralname", "")
          If Len(RuralName) > 0 Then
            For Each searchRuralScenarios In pRuralScenarios
              If LCase(searchRuralScenarios.Name) = LCase(RuralName) Then
                Set newScenario.RuralScenario = searchRuralScenarios
              End If
            Next
            If newScenario.RuralScenario Is Nothing Then
              ssMessageBox "Could not find rural scenario '" & RuralName & "' for urban scenario '" & newScenario.Name & "'", vbExclamation, "Project Let XML"
            End If
          End If
        Else
          newScenario.Urban = False
          'On Error Resume Next
          pRuralScenarios.Add newScenario, LCase(newScenario.Name)
        End If
        For Each vRegion In ScenarioXML.SubElements
          If vRegion.tag = "nssregion" Then
            Set newRegion = New userRegion
            RegionName = vRegion.AttributeValue("name")
            If RegionName = "National Urban" Then
              Set newRegion.Region = pNationalState.Regions(1)
            Else
              Set newRegion.Region = Me.State.Regions(RegionName)
            End If
            If newRegion.Region.LowFlowRegnID >= 0 Then newScenario.LowFlow = True
            If newRegion.Region.PredInt Then newScenario.UsePredInts = True
            If newRegion.Region.ROIRegnID >= 0 Then
              newScenario.ROI = True
              newScenario.UsePredInts = True 'ROI always uses prediction intervals
            End If
            For Each vParameter In vRegion.AttributeNames
              If LCase(vParameter) <> "name" Then
                newRegion.UserParms(vParameter).SetValue vRegion.AttributeValue(vParameter, 0), Me.State.Metric
              End If
            Next
            newScenario.UserRegions.Add newRegion, newRegion.Region.Name
            Set newRegion = Nothing
          ElseIf vRegion.tag = "weight" Then
            newScenario.Weight.XML = vRegion.GetString
          End If
        Next
        Set newScenario = Nothing
      Else
        ssMessageBox "Non-scenario '" & vScenario.tag & "' ignored in project.", vbOKOnly, "Reading Project"
      End If
    Next
    pChanged = False
    RaiseEdited
  End If

  Exit Property

ReadProblem:
  ssMessageBox "Problem Reading Project XML" & vbCr & Err.Description
End Property

Public Property Get FileName() As String
  FileName = pFileName
End Property
Public Property Let FileName(newValue As String)
  pFileName = newValue
End Property

Public Property Get RuralScenarios() As FastCollection
  Set RuralScenarios = pRuralScenarios
End Property
Public Property Get UrbanScenarios() As FastCollection
  Set UrbanScenarios = pUrbanScenarios
End Property

Public Property Get DB() As nssDatabase
  Set DB = pDB
End Property

Public Sub LoadNSSdatabase(FileName As String)
  Set pDB = Nothing
  Set pDB = New nssDatabase
  pDB.FileName = FileName
  Set pNationalState = New nssState
  pNationalState.Name = "Nation"
  pNationalState.Code = "00"
  Set pNationalState.DB = pDB
  Set pNationalUrban = pNationalState.Regions(1)
  Clear
End Sub

Public Sub Clear()
  Set pRuralScenarios = New FastCollection
  Set pUrbanScenarios = New FastCollection
  pCurrentRuralScenario = 0
  pCurrentUrbanScenario = 0
  pChanged = False
End Sub

Public Function Report() As String
  Dim retval As String
  Dim vScenario As Variant, vRegion As Variant, vParameter As Variant
  Dim Scenario As nssScenario
  Dim lScenario As nssScenario
  Dim Flows() As Double
  Dim Years() As Double
  Dim StdError() As Double
  Dim FlowText As String
  Dim IntervalText As String
  Dim YearsText As String
  Dim StdErrText As String
  Dim Index As Long
  Dim Val As Double, Max As Double, Min As Double
  Dim S As String
  Const indent As String = "    "
  Dim i As Long
  Dim RegArea As Single
  
  On Error GoTo ErrHandler
  
  retval = ""
  retval = retval & "National Streamflow Statistics Program" & vbCrLf
  retval = retval & "Version 4.0" ' & App.Major & "." & App.Minor '& "." & App.Revision & vbCrLf
  retval = retval & vbCrLf
  retval = retval & "Based on Techniques and Methods Book 4-A6" & vbCrLf
  retval = retval & "Equations from database " & pDB.FileName & vbCrLf
  retval = retval & "Updated by " & pDB.LastEdited & vbCrLf
  retval = retval & "Equations for " & Me.State.Name & " developed using "
  If Me.State.Metric Then
    retval = retval & "Metric units" & vbCrLf
  Else
    retval = retval & "English units" & vbCrLf
  End If
  retval = retval & vbCrLf
  retval = retval & "Site: " & Me.Name & ", " & Me.State.Name & vbCrLf
  retval = retval & "User: " & Me.UserName & vbCrLf
  retval = retval & "Date: " & Format(Date, "Long Date") & " " & Format(Time, "hh:mm ampm") & vbCrLf
  retval = retval & vbCrLf
  
'  For Each vScenario In Me.RuralScenarios
'    GoSub AddScenario
'  Next
'  For Each vScenario In Me.UrbanScenarios
'    GoSub AddScenario
'  Next
'
'  retval = retval & "Flood Peak Discharges, in cubic "
'  If Me.Metric Then retval = retval & "meters" Else retval = retval & "feet"
'  retval = retval & " per second" & vbCrLf
'  retval = retval & vbCrLf
'
'  retval = retval & "                     Recurrence     Peak,  Standard  Equivalent" & vbCrLf
'  retval = retval & "Estimate             Interval, yrs   cfs   Error, %  Years     " & vbCrLf
'  retval = retval & "____________________ _____________  _____  ________  __________" & vbCrLf
'
'  If Me.Metric Then retval = ReplaceString(retval, " cfs", "m3/s")

  For Each vScenario In Me.RuralScenarios
    GoSub AddFlows
  Next
  For Each vScenario In Me.UrbanScenarios
    GoSub AddFlows
  Next

  Report = retval
  Exit Function

'AddScenario:
'
'  Set Scenario = vScenario
'  If Scenario.Urban Then retval = retval & "Urban Estimate: " Else retval = retval & "Rural Estimate: "
'  retval = retval & ReplaceString(Scenario.Summary, vbLf, vbLf & "   ") & vbCrLf
'  retval = retval & Scenario.ErrorDescription & vbCrLf
'  Return

AddFlows:
  Set Scenario = vScenario
  retval = retval & vbCrLf
  If Scenario.Urban Then retval = retval & "Urban Estimate: " Else retval = retval & "Rural Estimate: "
  retval = retval & ReplaceString(Scenario.Summary, vbLf, vbLf & "   ") & vbCrLf
  retval = retval & Scenario.ErrorDescription
  retval = retval & vbCrLf & vbCrLf & "Results for:  " & Scenario.Name & vbCrLf & vbCrLf
  If Scenario.UserRegions.Count > 1 Then 'show results for each region in scenario
    For Each vRegion In Scenario.UserRegions
      Set lScenario = Scenario.Copy
      With lScenario
        .Name = .Name & " - " & vRegion.Region.Name
        .UserRegions.Clear
        RegArea = vRegion.UserParms(1).GetValue(Me.Metric) 'save original region area
        vRegion.UserParms(1).SetValue .GetArea(Me.Metric), Me.Metric
        .UserRegions.Add vRegion
        'add region name to scenario name
        S = indent & .Name & vbCrLf & .EstimateString
        vRegion.UserParms(1).SetValue RegArea, Me.Metric 'restore original region area
        'insert region results with scenario results
        While Len(S) > 0
          i = InStr(S, vbCrLf)
          If i > 0 Then 'take up to CR/LF
            retval = retval & Left(S, i - 1) & vbCrLf
            S = Mid(S, i + 2)
          Else 'take all
            retval = retval & S
            S = ""
          End If
          If Len(S) > 0 Then retval = retval & indent 'more lines to insert
        Wend
        retval = retval & vbCrLf & vbCrLf
      End With
    Next
    'include label for combined results
    retval = retval & Scenario.Name & " - Area-Averaged Estimates" & vbCrLf
  End If
  retval = retval & Scenario.EstimateString & vbCrLf & vbCrLf & vbCrLf
'  Flows = Scenario.WeightedDischarges
'  Years = Scenario.EquivalentYears
'  StdError = Scenario.StdError
'
'  retval = retval & vbCrLf
'  retval = retval & StrPad(Scenario.Name, 20, " ", False) & " "
'  For Index = LBound(Flows) To UBound(Flows)
''    If Index = LBound(Flows) Then
''      retval = retval & StrPad(Scenario.Name, 20, " ", False) & " "
''    Else
''      retval = retval & Space(21)
''    End If
'    If Flows(Index) > -1 Then
'      IntervalText = Scenario.UserRegions(1).Region.DepVars(Index).Name
'      If IsNumeric(IntervalText) Then IntervalText = SignificantDigits(Scenario.UserRegions(1).Region.DepVars(Index).Name, 2)
'      retval = retval & StrPad(IntervalText, 7) & "  "
'
'      FlowText = SignificantDigits(Flows(Index), 3)
'      retval = retval & StrPad(FlowText, 11)
'
'      If StdError(Index) = 0 Then
'        StdErrText = Space(8)
'      Else
'        StdErrText = StrPad(SignificantDigits(StdError(Index), 2), 8)
'        If StdError(Index) > 0 And StdError(Index) < 5 Then
'          retval = ReplaceString(retval, "Error, %", "Err, log")
'        End If
'      End If
'      retval = retval & StdErrText
'
'      If Years(Index) = 0 Then
'        YearsText = Space(10)
'      Else
'        YearsText = StrPad(SignificantDigits(Years(Index), 2), 10)
'      End If
'      retval = retval & YearsText & vbCrLf
'      retval = retval & Space(21)
'    End If
'  Next
'  If Scenario.RegCrippenBue > 0 Then
'    retval = retval & "maximum: " _
'                    & SignificantDigits(Scenario.MaxFloodEnvelope(Me.Metric), 3) _
'                    & " (for C&B region " & Scenario.RegCrippenBue & ")" & vbCrLf
'  End If
  retval = retval & Scenario.ErrorDescription
  Return
  
ErrHandler:
  retval = retval & "Error: " & Err.Description
  Resume Next
End Function

