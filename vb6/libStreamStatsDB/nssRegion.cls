VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nssRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private pName As String
Private pID As Long
Private pDB As nssDatabase
Private pUrban As Boolean, pUrbanNeedsRural As Boolean, pIsNew As Boolean
Private pROIRegnID As String
Private pLowFlowRegnID As String
Private pPredInt As Boolean
Private pParameters As FastCollection   'of ssStatistic
Private pROIParameters As FastCollection   'of ssStatistic
Private pDepVars As FastCollection      'of nssDepVar
Private pROIStations As FastCollection  'of ssStation
Private pState As nssState

Public Property Get Name() As String
  Name = pName
End Property
Public Property Let Name(ByVal newValue As String)
  pName = newValue
End Property

Public Property Get ID() As Long
  ID = pID
End Property
Public Property Let ID(ByVal newValue As Long)
  pID = newValue
End Property

Public Property Get Urban() As Boolean
  Urban = pUrban
End Property
Public Property Let Urban(ByVal newValue As Boolean)
  pUrban = newValue
End Property

Public Property Get IsNew() As Boolean
  IsNew = pIsNew
End Property
Public Property Let IsNew(ByVal newValue As Boolean)
  pIsNew = newValue
End Property

Public Property Get UrbanNeedsRural() As Boolean
  UrbanNeedsRural = pUrbanNeedsRural
End Property
Public Property Let UrbanNeedsRural(ByVal newValue As Boolean)
  pUrbanNeedsRural = newValue
End Property

Public Property Get ROIRegnID() As String
  ROIRegnID = pROIRegnID
End Property
Public Property Let ROIRegnID(ByVal newValue As String)
  pROIRegnID = newValue
End Property

Public Property Get LowFlowRegnID() As String
  LowFlowRegnID = pLowFlowRegnID
End Property
Public Property Let LowFlowRegnID(ByVal newValue As String)
  pLowFlowRegnID = newValue
End Property

Public Property Get ROI() As Boolean
  If pROIRegnID > 0 Then ROI = True
End Property
'Public Property Let ROI(ByVal newValue As Boolean)
'  pROI = newValue
'End Property

Public Property Get PredInt() As Boolean
  PredInt = pPredInt
End Property
Public Property Let PredInt(ByVal newValue As Boolean)
  pPredInt = newValue
End Property

Public Property Get DB() As nssDatabase
  Set DB = pDB
End Property
Public Property Set DB(newValue As nssDatabase)
  Set pDB = newValue
End Property

Public Property Get State() As nssState
  Set State = pState
End Property
Public Property Set State(ByVal newValue As nssState)
  Set pState = newValue
End Property

Public Property Set Parameters(ByRef newValue As FastCollection)
  Set pParameters = newValue
End Property
Public Property Get Parameters() As FastCollection
  ' ID must be set for this region before this will work
  If pParameters Is Nothing Then PopulateParameters
  Set Parameters = pParameters
End Property
Public Sub PopulateParameters()
  Dim newParm As nssParameter
  Dim parmRec As Recordset
  Dim sql$

  If Not pParameters Is Nothing Then
    pParameters.Clear
    Set pParameters = Nothing
  End If
  Set pParameters = New FastCollection

  If pDB Is Nothing Then
    ssMessageBox "Cannot read parameters" & vbCr & "DB not set."
  Else
    sql = "SELECT Parameters.ParmID, Parameters.Min, Parameters.Max, Parameters.ConvFlag, " & _
             "Parameters.Parameter, STATLABEL.StatLabel, STATLABEL.StatisticLabel, " & _
             "STATLABEL.StatisticLabelCode, STATTYPE.StatisticTypeCode, STATTYPE.StatisticType " & _
          "FROM STATTYPE INNER JOIN (STATLABEL INNER JOIN [Parameters] " & _
          "ON STATLABEL.StatisticLabelCode = Parameters.StatisticLabelCode) " & _
          "ON STATTYPE.StatisticTypeCode = STATLABEL.StatisticTypeCode " & _
          "WHERE Parameters.RegionID=" & pID & " " & _
          "ORDER BY Parameters.ParmID;"
    Set parmRec = pDB.DB.OpenRecordset(sql)
    While Not parmRec.EOF
      Set newParm = New nssParameter
      With newParm
        Set .Region = Me
        .ID = parmRec.Fields("ParmID")
        If Not IsNull(parmRec.Fields("StatisticLabelCode")) Then .LabelCode = parmRec.Fields("StatisticLabelCode")
        If Not IsNull(parmRec.Fields("StatisticLabel")) Then .Name = parmRec.Fields("StatisticLabel")
        If Not IsNull(parmRec.Fields("StatLabel")) Then .Abbrev = parmRec.Fields("StatLabel")
        If Not IsNull(parmRec.Fields("Parameter")) Then .NSSName = parmRec.Fields("Parameter")
        If Not IsNull(parmRec.Fields("ConvFlag")) Then Set .Units = pDB.Units(CStr(parmRec.Fields("ConvFlag")))
        If Not IsNull(parmRec.Fields("Min")) Then .SetMin parmRec.Fields("Min"), pState.Metric
        If Not IsNull(parmRec.Fields("Max")) Then .SetMax parmRec.Fields("Max"), pState.Metric
        If Not IsNull(parmRec.Fields("StatisticTypeCode")) Then .StatTypeCode = parmRec.Fields("StatisticTypeCode")
        If Not IsNull(parmRec.Fields("StatisticType")) Then .StatType = parmRec.Fields("StatisticType")
      End With
      pParameters.Add newParm, CStr(newParm.ID)
      Set newParm = Nothing
      parmRec.MoveNext
    Wend
    parmRec.Close
  End If
End Sub

Public Sub ClearROIUserparms()
  Dim sql As String
  Dim myRec As Recordset
  
  sql = "SELECT ROIUserParms.* From [ROIUserParms] " & _
        "WHERE RegionID=" & Me.ID & ";"
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    If Not .EOF Then
      .MoveLast
      While Not .RecordCount = 0
        .MoveFirst
        .Delete
      Wend
    End If
  End With

  sql = "SELECT Parameters.* From [Parameters] " & _
        "WHERE RegionID=" & Me.ID & ";"
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    If Not .EOF Then
      .MoveLast
      While Not .RecordCount = 0
        .MoveFirst
        .Delete
      Wend
    End If
  End With

  sql = "SELECT DepVars.* From [DepVars] " & _
        "WHERE RegionID=" & Me.ID & ";"
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    If Not .EOF Then
      .MoveLast
      While Not .RecordCount = 0
        .MoveFirst
        .Delete
      Wend
    End If
  End With

End Sub

Public Property Set ROIParameters(ByRef newValue As FastCollection)
  Set pROIParameters = newValue
End Property
Public Property Get ROIParameters() As FastCollection
  ' ID must be set for this region before this will work
  If pROIParameters Is Nothing Then PopulateROIParameters
  Set ROIParameters = pROIParameters
End Property
Public Sub PopulateROIParameters()
  Dim newParm As nssParameter
  Dim parmRec As Recordset
  Dim sql$

  If Not pROIParameters Is Nothing Then
    pROIParameters.Clear
    Set pROIParameters = Nothing
  End If
  Set pROIParameters = New FastCollection

  If pDB Is Nothing Then
    ssMessageBox "Cannot read parameters" & vbCr & "DB not set."
  Else
    sql = "SELECT STATLABEL.StatisticLabelCode, STATLABEL.StatLabel, STATLABEL.StatisticLabel, " & _
        "STATTYPE.StatisticTypeCode, STATTYPE.StatisticType, " & _
        "ROIUserParms.PosNegCorrelation, ROIUserParms.SimVar, ROIUserParms.RegVar, " & _
        "Parameters.ParmID, Parameters.Min, Parameters.Max, Parameters.ConvFlag " & _
        "FROM (ROIUserParms INNER JOIN " & _
        "(STATTYPE INNER JOIN STATLABEL ON STATTYPE.StatisticTypeCode = STATLABEL.StatisticTypeCode) " & _
        "ON ROIUserParms.ParmID = STATLABEL.StatisticLabelCode) " & _
        "INNER JOIN [Parameters] ON STATLABEL.StatisticLabelCode = Parameters.StatisticLabelCode " & _
        "WHERE Parameters.RegionID=" & pID & " AND ROIUserParms.RegionID=" & pID & ";"
    Set parmRec = pDB.DB.OpenRecordset(sql)
    While Not parmRec.EOF
      Set newParm = New nssParameter
      With newParm
        Set .Region = Me
        .ID = parmRec("ParmID")
        If Not IsNull(parmRec("StatisticLabelCode")) Then .LabelCode = parmRec("StatisticLabelCode")
        If Not IsNull(parmRec("StatisticLabel")) Then
          .Name = parmRec("StatisticLabel")
          .NSSName = .Name
        End If
        If Not IsNull(parmRec("StatLabel")) Then .Abbrev = parmRec("StatLabel")
        If Not IsNull(parmRec("ConvFlag")) Then Set .Units = pDB.Units(CStr(parmRec("ConvFlag")))
        If Not IsNull(parmRec("Min")) Then .SetMin parmRec("Min"), pState.Metric
        If Not IsNull(parmRec("Max")) Then .SetMax parmRec("Max"), pState.Metric
        If Not IsNull(parmRec("StatisticTypeCode")) Then .StatTypeCode = parmRec("StatisticTypeCode")
        If Not IsNull(parmRec("StatisticType")) Then .StatType = parmRec("StatisticType")
        If Not IsNull(parmRec("PosNegCorrelation")) Then .CorrelationType = CSng(parmRec("PosNegCorrelation"))
        If Not IsNull(parmRec("SimVar")) Then .SimulationVar = parmRec("SimVar")
        If Not IsNull(parmRec("RegVar")) Then .RegressionVar = parmRec("RegVar")
      End With
      pROIParameters.Add newParm, CStr(newParm.ID)
      Set newParm = Nothing
      parmRec.MoveNext
    Wend
    parmRec.Close
  End If
End Sub

Public Property Set DepVars(ByRef newValue As FastCollection)
  Set pDepVars = newValue
End Property
Public Property Get DepVars() As FastCollection
  ' ID must be set for this region before this will work
  If pDepVars Is Nothing Then PopulateDepVars
  Set DepVars = pDepVars
End Property
Public Sub PopulateDepVars()
  Dim newDepVar As nssDepVar
  Dim depVarRec As Recordset
  Dim sql$
  Dim AllNumeric As Boolean
  Dim DepVarID As Integer, i As Integer, ipos As Integer
  Dim vComponent As Variant
  Dim vDepVar As Variant
  Dim lName As String
  Dim lCurName As String
  
  AllNumeric = True 'assume all depvars for region are numbers and may be sorted
  Set pDepVars = Nothing
  Set pDepVars = New FastCollection
  If pDB Is Nothing Then
    ssMessageBox "Cannot read DepVars" & vbCr & "DB not set."
  Else
    sql = "SELECT DepVars.* FROM [DepVars] WHERE " & _
          "DepVars.RegionID=" & pID & _
          " ORDER BY Len(Name), Name;"
    'below will be the new query to pull Units from StatLabel table
'    sql = "SELECT DepVars.*, StatLabel.Units " & _
'          "FROM Regions INNER JOIN DepVars ON StatLabel.StatisticLableCode = DepVars.StatisticLabelCode " & _
'          "WHERE DepVars.RegionID=" & pID & _
'          " ORDER BY Len(DepVars.Name), DepVars.Name;"
    Set depVarRec = pDB.DB.OpenRecordset(sql)
    While Not depVarRec.EOF
      DepVarID = depVarRec.Fields("DepVarID")
      Set newDepVar = New nssDepVar
      With newDepVar
        Set .DB = pDB
        Set .Region = Me
        .ID = DepVarID
        If Not IsNull(depVarRec.Fields("Name")) Then .Name = depVarRec.Fields("Name")
        If Not IsNull(depVarRec.Fields("StdErr")) Then .StdErr = depVarRec.Fields("StdErr")
        If Not IsNull(depVarRec.Fields("EstErr")) Then .EstErr = depVarRec.Fields("EstErr")
        If Not IsNull(depVarRec.Fields("PreErr")) Then .PreErr = depVarRec.Fields("PreErr")
        If Not IsNull(depVarRec.Fields("EquivYears")) Then .EquivYears = depVarRec.Fields("EquivYears")
        If Not IsNull(depVarRec.Fields("Constant")) Then .Constant = depVarRec.Fields("Constant")
        If Not IsNull(depVarRec.Fields("BCF")) Then .BCF = depVarRec.Fields("BCF")
        If Not IsNull(depVarRec.Fields("t")) Then .tDist = depVarRec.Fields("t")
        If Not IsNull(depVarRec.Fields("Variance")) Then .Variance = depVarRec.Fields("Variance")
        If Not IsNull(depVarRec.Fields("ExpDA")) Then .ExpDA = depVarRec.Fields("ExpDA")
        'add the following line when database updated to link DepVars and StatLabel tables
'        If Not IsNull(depVarRec.Fields("Units")) Then Set .Units = pDB.Units(CStr(depVarRec.Fields("Units")))
      End With
      newDepVar.Components.Clear
      Set newDepVar.Components = Nothing
      If Left(newDepVar.Name, 2) = "PK" Then
        lName = Mid(newDepVar.Name, 3)
      Else
        lName = newDepVar.Name
      End If
      If Not IsNumeric(lName) Then AllNumeric = False
      If AllNumeric Then 'sort depvars for region
        ipos = 0
        i = 1
        While ipos = 0 And i <= pDepVars.Count
          If Left(pDepVars.ItemByIndex(i).Name, 2) = "PK" Then
            lCurName = Mid(pDepVars.ItemByIndex(i).Name, 3)
          Else
            lCurName = pDepVars.ItemByIndex(i).Name
          End If
          If CSng(lName) > CSng(lCurName) Then
            i = i + 1
          Else
            ipos = i
          End If
        Wend
      End If
      If ipos > 0 Then 'insert in proper sorted position
        pDepVars.Add newDepVar, CStr(DepVarID), ipos
      Else 'just put at end of collection
        pDepVars.Add newDepVar, CStr(DepVarID)
      End If
      Set newDepVar = Nothing
      depVarRec.MoveNext
    Wend
    depVarRec.Close
    For Each vDepVar In pDepVars
      Set newDepVar = vDepVar
      For Each vComponent In newDepVar.Components
        If vComponent.ParmID > 0 Then
          'XXX - the parameters in vComponent are from region indexed one below current region
          Set vComponent.Parameter = Me.Parameters(CStr(vComponent.ParmID))
        End If
        If vComponent.ExpID > 0 Then
          Set vComponent.ExponentParameter = Me.Parameters(CStr(vComponent.ExpID))
        End If
      Next
    Next
  End If
End Sub

Public Property Set ROIStations(ByRef newValue As FastCollection)
  'Set pROIStations = newValue
  ssMessageBox "This property is no longer valid - now in nssState class."
End Property
Public Property Get ROIStations() As FastCollection
  'this property is no longer functional as PopulateROIStations is now done in nssState.cls
  'If pROIStations Is Nothing Then PopulateROIStations
  ssMessageBox "This property is no longer valid - now in nssState class."
  'Set ROIStations = pROIStations
End Property

Public Sub PopulateROIStations()
  'this sub is no longer functional as PopulateROIStations is now done in nssState.cls
End Sub

Public Function Add(IsReturn As Boolean, ByVal RegName As String, ByVal Urban As Boolean, _
    ByVal RuralInput As Long, ByVal PredInt As Long, ByVal ROIRegnID As Long, Optional NoWarnFlg As Boolean) As Boolean
  Dim myRec As Recordset
  Dim ID&, lowFlowID&
  Dim sql$
  
  RegName = NoSpaces(RegName)
  Set myRec = Me.DB.DB.OpenRecordset("Regions", dbOpenDynaset)
  With myRec
    'check to see if region with this name already exists
    If .RecordCount > 0 Then
      .FindFirst "Name='" & RegName & "' AND StateCode='" & DB.State.Code & "'"
    End If
    If .NoMatch Or .RecordCount = 0 Then
      'find an unused ID
      .MoveLast
      ID = !RegionID
      .FindFirst "RegionID=" & ID
      While Not .NoMatch
        ID = !RegionID + 1
        .FindFirst "RegionID=" & ID
      Wend
    Else
      If Not NoWarnFlg Then
        ssMessageBox "There is already a region in " & DB.State.Name & _
               " with the name " & RegName & "." & vbCrLf & "No new region was created.", vbCritical
      End If
      Exit Function
    End If
    .Close
  End With
  'quicker to open new recordset than to search all regions for LowFlow ID
  sql = "SELECT * FROM Regions " & _
        "WHERE StateCode='" & DB.State.Code & "';"
  Set myRec = Me.DB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    If IsReturn Then
      lowFlowID = -1
    Else 'find a LowFlow ID
      If .RecordCount > 0 Then
        .MoveFirst
        lowFlowID = !LowFlow
        .MoveNext
        While Not .EOF
          If !LowFlow > lowFlowID Then lowFlowID = !LowFlow
          .MoveNext
        Wend
      End If
      If lowFlowID > 0 Then lowFlowID = lowFlowID + 1 Else lowFlowID = 1
    End If
    .AddNew
    !StateCode = DB.State.Code
    !RegionID = ID
    !RegionName = RegName
    !Urban = Urban
    !UrbanNeedsRural = -RuralInput
    !ROI = ROIRegnID
    !LowFlow = lowFlowID
    !PredInt = -PredInt
    .Update
    Add = True
  End With
End Function

Public Function Delete() As Boolean
  Dim myRec As Recordset
  Dim sql$
  
  sql = "SELECT * FROM Regions WHERE StateCode='" & DB.State.Code & "'"
  Set myRec = DB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    .FindFirst "RegionID=" & Me.ID
    If Not .NoMatch Then
      .Delete
      Delete = True
    Else
      Delete = False
    End If
  End With
End Function

Public Sub Edit(ByVal NewName As String, ByVal Urban As Boolean, _
                ByVal RuralInput As Integer, ByVal PredInt As Integer)
  Dim myRec As Recordset
  Dim sql$
  
  NewName = NoSpaces(NewName)
  sql = "SELECT * FROM Regions WHERE StateCode='" & Me.State.Code & "'"
  Set myRec = DB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    .MoveLast
    .FindFirst "RegionID=" & Me.ID
    If Not .NoMatch Then
      .Edit
      !RegionName = NewName
      !Urban = Urban
      !UrbanNeedsRural = -RuralInput
      !PredInt = -PredInt
      .Update
    Else
      ssMessageBox "There is no region in " & DB.State.Name & _
             " with the name " & Me.Name & "." & vbCrLf & "Region not found."
    End If
  End With
End Sub

Private Function GetStationType(TypeCode As String) As ssStationType
  Dim myRec As Recordset
  
  Set GetStationType = New ssStationType
  Set myRec = pDB.DB.OpenRecordset("STATIONTYPE", dbOpenSnapshot)
  With myRec
    .FindFirst "StationTypeCode='" & TypeCode & "'"
    If Not .NoMatch Then
      Set GetStationType.DB = pDB
      GetStationType.Code = .Fields("StationTypeCode")
      GetStationType.Name = .Fields("StationType")
    End If
  End With
End Function

Private Function NoSpaces(ObjName As String) As String
  Dim str$
  
  str = StrRetRem(ObjName)
  NoSpaces = str
  While Len(Trim(ObjName)) > 0
    str = StrRetRem(ObjName)
    NoSpaces = NoSpaces & "_" & str
  Wend
End Function

Private Sub Class_Initialize()
  pName = "Not Assigned"
End Sub

