VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nssDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const DatabasePassword = "H2O2U"

Private pStates As FastCollection 'of nssState
Private pState As nssState
Private pnssDB As Database
Private pFileName As String
Private pUnits As FastCollection 'of nssUnits
Private pAreaUnits As nssUnits
Private pFlowUnits As nssUnits
'Private pROIStats As FastCollection 'of ROI ssStatistics kept by this station
Private pStationTypes As FastCollection 'of strings
Private pStatisticTypes As FastCollection 'of ssStatType
Private pParameters As FastCollection 'of nssParameter
Private pLFStats As FastCollection 'of nssDepVar
Private pReturns As FastCollection 'of nssDepVar
Private pStatType As ssStatType
Private pSelStats As FastCollection 'of ssStatLabel
Private pSources As FastCollection 'of ssSource

Public Function AreaUnits() As nssUnits
  If pAreaUnits Is Nothing Then
    Set pAreaUnits = Units("1")
  End If
  Set AreaUnits = pAreaUnits
End Function

Public Function FlowUnits() As nssUnits
  If pFlowUnits Is Nothing Then
    Set pFlowUnits = Units("13")
  End If
  Set FlowUnits = pFlowUnits
End Function

Public Property Get DB() As Database
  Set DB = pnssDB
End Property

Public Property Get LastEdited() As String
  Dim myRec As Recordset
  Dim sql As String
  sql = "SELECT TransactionLog.* FROM [TransactionLog] ORDER BY TransID;"
  Set myRec = pnssDB.OpenRecordset(sql, dbOpenSnapshot)
  If Not myRec.EOF Then
    myRec.MoveLast
    LastEdited = myRec("UserName") _
         & " " & myRec("Date_Time") _
         & " " & myRec("Explanation")
  Else
    LastEdited = "(No entries in Transaction Log)"
  End If
End Property

Public Property Get filename() As String
  filename = pFileName
End Property
Public Property Let filename(ByVal newValue As String)
  pFileName = newValue
  Set pnssDB = OpenDatabase(pFileName, False, False, "MS Access; pwd=" & DatabasePassword)
End Property

Public Property Get State() As nssState
  Set State = pState
End Property
Public Property Set State(ByVal newValue As nssState)
  Set pState = newValue
End Property

Public Property Get States() As FastCollection
  ' Filename must be set before this will work
  If pStates Is Nothing Then PopulateStates
  Set States = pStates
End Property
Public Property Set States(ByVal newValue As FastCollection)
  Set pStates = newValue
End Property

Private Sub PopulateStates()
  Dim newState As nssState
  Dim lROIData As nssROI
  Dim lCode As String
  Dim myRec As Recordset
  Dim sql$
  
  Set pStates = New FastCollection

  'builds fastCollection of state objects
  sql = "SELECT STATES.* FROM [STATES] " & _
        "WHERE StateCode <> '00' " & _
        "ORDER BY Len(StateCode), StateCode;"
  Set myRec = pnssDB.OpenRecordset(sql, dbOpenForwardOnly)
  While Not myRec.EOF
    Set newState = New nssState
    Set lROIData = New nssROI
    With newState
      Set .DB = Me
      .Name = myRec.Fields("State")
      .Abbrev = myRec.Fields("St")
      .code = myRec.Fields("StateCode")
      .Metric = myRec.Fields("Metric")
      Set lROIData.DB = Me
      lROIData.StateCode = .code
      lROIData.ClimateFactor = myRec.Fields("ROIClimateFactor")
      lROIData.Distance = myRec.Fields("ROIDistance")
      lROIData.Regress = myRec.Fields("ROIBackStepRegression")
      lROIData.UseRegions = myRec.Fields("ROIRegions")
      lROIData.SimStations = myRec.Fields("ROISimStations")
      lCode = Right(.code, 2)
      If States.KeyExists(lCode) Then
        'just add other ROI data from this 2nd occurrence
        If CInt(.code) > 10000 Then
          Set States(lCode).ROILowData = lROIData
        Else
          Set States(lCode).ROIPeakData = lROIData
        End If
      Else 'add state (and this ROI data) to collection
        If CInt(.code) > 10000 Then 'low flow ROI data
          .code = lCode
          Set .ROILowData = lROIData
        Else
          Set .ROIPeakData = lROIData
        End If
        States.Add newState, lCode
      End If
    End With
    Set newState = Nothing
    myRec.MoveNext
  Wend
  myRec.Close
End Sub

Public Function Units() As FastCollection
  If pUnits Is Nothing Then PopulateUnits
  Set Units = pUnits
End Function
Private Sub PopulateUnits()
  Dim myRec As Recordset
  Dim sql$
  Dim newUnits As nssUnits
  
  Set pUnits = New FastCollection
  
  sql = "SELECT * FROM [Units] ORDER BY UnitID;"
  Set myRec = pnssDB.OpenRecordset(sql, dbOpenForwardOnly)
  While Not myRec.EOF
    Set newUnits = New nssUnits
    With newUnits
      .Id = myRec.Fields("UnitID")
      .MetricLabel = myRec.Fields("Metric")
      .MetricAbbrev = myRec.Fields("MetricAbbrev")
      .EnglishLabel = myRec.Fields("English")
      .EnglishAbbrev = myRec.Fields("EnglishAbbrev")
      .Factor = myRec.Fields("EnglishMetricFactor")
    End With
    pUnits.Add newUnits, CStr(newUnits.Id)
    Set newUnits = Nothing
    myRec.MoveNext
  Wend
  myRec.Close
End Sub

Public Property Set SelStats(ByRef newValue As FastCollection)
  Set pSelStats = newValue
End Property
Public Property Get SelStats() As FastCollection
  If pSelStats Is Nothing Then Set pSelStats = New FastCollection
  Set SelStats = pSelStats
End Property

Public Property Set StationTypes(ByRef newValue As FastCollection)
  Set pStationTypes = newValue
End Property
Public Property Get StationTypes() As FastCollection
  If pStationTypes Is Nothing Then PopulateStationTypes
  Set StationTypes = pStationTypes
End Property
Private Sub PopulateStationTypes()
  Dim myRec As Recordset
  Dim newType As ssStationType
  
  On Error GoTo x
  
  Set pStationTypes = New FastCollection
  'builds collection of StationTypes
  Set myRec = pnssDB.OpenRecordset("STATIONTYPE", dbOpenForwardOnly)
  On Error GoTo 0
  While Not myRec.EOF
    Set newType = New ssStationType
    With newType
      Set .DB = Me
      .code = myRec("StationTypeCode")
      .Name = myRec("StationType")
    End With
    StationTypes.Add newType, newType.code
    myRec.MoveNext
  Wend
  myRec.Close
  Exit Sub
x:
End Sub

Public Property Set StatType(ByRef newValue As ssStatType)
  Set pStatType = newValue
End Property
Public Property Get StatType() As ssStatType
  Set StatType = pStatType
End Property

Public Property Set StatisticTypes(ByRef newValue As FastCollection)
  Set pStatisticTypes = newValue
End Property
Public Property Get StatisticTypes() As FastCollection
  ' Filename must be set before this will work
  If pStatisticTypes Is Nothing Then PopulateStatisticTypes
  Set StatisticTypes = pStatisticTypes
End Property
Private Sub PopulateStatisticTypes()
  Dim myRec As Recordset
  Dim sql$
  Dim newStatType As ssStatType
  
  If Not pStatisticTypes Is Nothing Then
    pStatisticTypes.Clear
  End If
  Set pStatisticTypes = New FastCollection

  'builds collection of StatisticTypes
  sql = "SELECT * FROM STATTYPE ORDER BY StatisticTypeID;"
  Set myRec = pnssDB.OpenRecordset(sql, dbOpenForwardOnly)
  While Not myRec.EOF
    Set newStatType = New ssStatType
    With newStatType
      Set .DB = Me
      .Id = myRec("StatisticTypeID")
      .code = myRec("StatisticTypeCode")
      .Name = myRec("StatisticType")
    End With
    StatisticTypes.Add newStatType, newStatType.code
    myRec.MoveNext
  Wend
  myRec.Close
End Sub

Public Property Set Parameters(ByRef newValue As FastCollection)
  Set pParameters = newValue
End Property
Public Property Get Parameters() As FastCollection
  ' Filename must be set before this will work
  If pParameters Is Nothing Then PopulateParameters
  Set Parameters = pParameters
End Property
Private Sub PopulateParameters()
  Dim myRec As Recordset
  Dim sql$
  Dim newParm As nssParameter
  
  If Not pParameters Is Nothing Then
    pParameters.Clear
  End If
  Set pParameters = New FastCollection

  'builds collection of Parameters
  sql = "SELECT [STATLABEL].[StatLabel], [STATLABEL].[StatisticLabel], [UNITS].[UnitID]" & _
        "FROM STATLABEL INNER JOIN UNITS ON [STATLABEL].[UnitID]=[UNITS].[UnitID] " & _
        "ORDER BY [STATLABEL].[StatisticTypeCode], [STATLABEL].[StatisticLabelID];"
  Set myRec = pnssDB.OpenRecordset(sql, dbOpenForwardOnly)
  While Not myRec.EOF
    Set newParm = New nssParameter
    With newParm
      .Name = myRec("StatisticLabel")
      .Abbrev = myRec("StatLabel")
      .ConvFlag = myRec("UnitID")
    End With
    pParameters.Add newParm, newParm.Abbrev
    myRec.MoveNext
  Wend
  myRec.Close
End Sub

Public Property Set LFStats(ByRef newValue As FastCollection)
  Set pLFStats = newValue
End Property
Public Property Get LFStats() As FastCollection
  ' Filename must be set before this will work
  If pLFStats Is Nothing Then PopulateDepVars "LF"
  Set LFStats = pLFStats
End Property

Public Property Set Returns(ByRef newValue As FastCollection)
  Set pReturns = newValue
End Property
Public Property Get Returns() As FastCollection
  ' Filename must be set before this will work
  If pReturns Is Nothing Then PopulateDepVars "PF"
  Set Returns = pReturns
End Property
Private Sub PopulateDepVars(FlowType As String)
  Dim myRec As Recordset
  Dim sql$
  Dim newDepVar As nssDepVar
  
  On Error GoTo x
  
  If FlowType = "LF" Then
    If Not pLFStats Is Nothing Then
      pLFStats.Clear
      Set pLFStats = Nothing
    End If
    Set pLFStats = New FastCollection
    'sql = "SELECT DISTINCT Name FROM DepVars WHERE ID>10000 ORDER BY NAME;"
    sql = "SELECT DISTINCT StatLabel.StatLabel FROM StatLabel " & _
          "INNER JOIN DepVars ON StatLabel.StatisticLabelID = DepVars.StatisticLabelID " & _
          "WHERE DepVarID>10000 ORDER BY StatLabel;"
  ElseIf FlowType = "PF" Then
    If Not pReturns Is Nothing Then
      pReturns.Clear
      Set pReturns = Nothing
    End If
    Set pReturns = New FastCollection
    'sql = "SELECT DISTINCT Name FROM DepVars WHERE ID<10000 ORDER BY NAME;"
    sql = "SELECT DISTINCT StatLabel.StatLabel FROM StatLabel " & _
          "INNER JOIN DepVars ON StatLabel.StatisticLabelID = DepVars.StatisticLabelID " & _
          "WHERE DepVarID<10000 ORDER BY StatLabel;"
  End If
  
  'builds list of available Statistics/Returns
  Set myRec = pnssDB.OpenRecordset(sql, dbOpenForwardOnly)
  On Error GoTo 0
  While Not myRec.EOF
    Set newDepVar = New nssDepVar
    newDepVar.Name = myRec("StatLabel")
    If FlowType = "PF" Then
      pReturns.Add newDepVar, newDepVar.Name
    ElseIf FlowType = "LF" Then
      pLFStats.Add newDepVar, newDepVar.Name
    End If
    myRec.MoveNext
  Wend
  myRec.Close
  Exit Sub
x:
End Sub
Public Property Set Sources(ByRef newValue As FastCollection)
  Set pSources = newValue
End Property
Public Property Get Sources() As FastCollection
  ' Filename must be set before this will work
  If pSources Is Nothing Then PopulateSources
  Set Sources = pSources
End Property
Private Sub PopulateSources()
  Dim myRec As Recordset
  Dim newSource As ssSource
  Dim sql$
  
  Set pSources = New FastCollection

  'builds collection of Sources
  sql = "SELECT * FROM DATASOURCE" ' ORDER BY Citation"
  Set myRec = pnssDB.OpenRecordset(sql, dbOpenForwardOnly)
  With myRec
    While Not myRec.EOF
      Set newSource = New ssSource
      With newSource
        Set .DB = Me
        .Id = myRec("DataSourceID")
        .Name = myRec("Citation")
        If Not IsNull(myRec("CitationURL")) Then .URL = myRec("CitationURL")
      End With
      Sources.Add newSource, newSource.Id
      .MoveNext
    Wend
    .Close
  End With
End Sub

Public Function RecordUserInfo(ByVal UserName As String, ByVal LogOnName As String, _
    ByVal Organization As String, ByVal DateTime As String, _
    ByVal Explanation As String, ByVal Where As String) As Long
  Dim myRec As Recordset
  Set myRec = pnssDB.OpenRecordset("TransactionLog", dbOpenDynaset)
  With myRec
    .AddNew
    !UserName = UserName
    !LogOnName = LogOnName
    !Organization = Organization
    !date_time = DateTime
    !Explanation = Explanation
    !Region = Where
    RecordUserInfo = !TransID
    .Update
  End With
End Function

Public Function RecordChanges(ByVal Id As Long, ByVal TableName As String, _
    ByVal fldID As Long, ByVal RecID As String, _
    ByVal oldval As String, ByVal newval As String)
  Dim myRec As Recordset
  Dim fldName$
  
  Set myRec = pnssDB.OpenRecordset(TableName, dbOpenSnapshot)
  fldName = myRec.Fields(fldID).Name
  myRec.Close
  Set myRec = pnssDB.OpenRecordset("DetailedLog", dbOpenDynaset)
  With myRec
    .AddNew
    !TransID = Id
    !table = TableName
    !Field = fldName
    !Record = RecID
    !oldval = oldval
    !newval = newval
    .Update
  End With
  myRec.Close
End Function

Private Sub Class_terminate()
  On Error Resume Next
  Set pStates = Nothing
  Set pUnits = Nothing
  pnssDB.Close
End Sub
