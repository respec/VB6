VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nssROI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private pFlowType As String               '"P" or "L" (peak or low)
Private pStateCode As String
Private pDB As nssDatabase
Private pClimateFactor As Boolean
Private pDistance As Boolean
Private pRegress As Boolean
Private pUseRegions As Boolean
Private pSimStations As Integer       'number of similar stations to use
'Private pParms As FastCollection      'of nssparameter
Private pStations As FastCollection   'of ssStation
Private pFlowStats As FastCollection  'of ssStatistic

Public Property Get FlowType() As String
  FlowType = pFlowType
End Property
Public Property Let FlowType(ByVal newValue As String)
  pFlowType = newValue
End Property

Public Property Get StateCode() As String
  StateCode = pStateCode
End Property
Public Property Let StateCode(ByVal newValue As String)
  pStateCode = newValue
End Property

Public Property Get DB() As nssDatabase
  Set DB = pDB
End Property
Public Property Set DB(newValue As nssDatabase)
  Set pDB = newValue
End Property

Public Property Get ClimateFactor() As Boolean
  ClimateFactor = pClimateFactor
End Property
Public Property Let ClimateFactor(ByVal newValue As Boolean)
  pClimateFactor = newValue
End Property

Public Property Get Distance() As Boolean
  Distance = pDistance
End Property
Public Property Let Distance(ByVal newValue As Boolean)
  pDistance = newValue
End Property

Public Property Get Regress() As Boolean
  Regress = pRegress
End Property
Public Property Let Regress(ByVal newValue As Boolean)
  pRegress = newValue
End Property

Public Property Get UseRegions() As Boolean
  UseRegions = pUseRegions
End Property
Public Property Let UseRegions(ByVal newValue As Boolean)
  pUseRegions = newValue
End Property

Public Property Get SimStations() As Integer
  SimStations = pSimStations
End Property
Public Property Let SimStations(ByVal newValue As Integer)
  pSimStations = newValue
End Property

'Public Property Set Parms(ByRef newValue As FastCollection)
'  Set pParms = newValue
'End Property
'Public Property Get Parms() As FastCollection
'  If pParms Is Nothing Then PopulateParms
'  Set Parms = pParms
'End Property
'Private Sub PopulateParms()
'  Dim newParm As nssParameter
'  Dim parmRec As Recordset
'  Dim sql$
'  Dim i&
'
'  Set pParms = New FastCollection
'  If pDB Is Nothing Then
'    ssMessageBox "Cannot read ROIParms" & vbCr & "DB not set."
'  Else
'    sql = "SELECT STATLABEL.StatisticLabelID, STATLABEL.StatLabel, " & _
'        "STATLABEL.StatisticLabel, STATTYPE.StatisticTypeCode, STATTYPE.StatisticType, " & _
'        "ROIUserParms.PosNegCorrelation, ROIUserParms.SimVar, ROIUserParms.RegVar " & _
'        "FROM ROIUserParms " & _
'        "INNER JOIN (STATTYPE INNER JOIN STATLABEL ON STATTYPE.StatisticTypeCode = STATLABEL.StatisticTypeCode) " & _
'        "ON ROIUserParms.ParmID = STATLABEL.StatisticLabelID " & _
'        "Where ROIUserParms.StateID='" & pStateCode & "' " & _
'        "ORDER BY STATLABEL.StatisticLabelID;"
'    Set parmRec = pDB.DB.OpenRecordset(sql)
'    While Not parmRec.EOF
'      Set newParm = New nssParameter
'      With newParm
'        If Not IsNull(parmRec("StatisticLabelID")) Then .LabelCode = parmRec("StatisticLabelID")
'        If Not IsNull(parmRec("StatLabel")) Then .Abbrev = parmRec("StatLabel")
'        If Not IsNull(parmRec("StatisticLabel")) Then .Name = parmRec("StatisticLabel")
'        If Not IsNull(parmRec("StatisticTypeCode")) Then .StatTypeCode = parmRec("StatisticTypeCode")
'        If Not IsNull(parmRec("StatisticType")) Then .StatType = parmRec("StatisticType")
'        If Not IsNull(parmRec("PosNegCorrelation")) Then .CorrelationType = CSng(parmRec("PosNegCorrelation"))
'        If Not IsNull(parmRec("SimVar")) Then .SimulationVar = parmRec("SimVar")
'        If Not IsNull(parmRec("RegVar")) Then .RegressionVar = parmRec("RegVar")
'      End With
'      pParms.Add newParm, CStr(newParm.LabelCode)
'      Set newParm = Nothing
'      parmRec.MoveNext
'    Wend
'    parmRec.Close
'  End If
'End Sub

Public Property Set Stations(ByRef newValue As FastCollection)
  Set pStations = newValue
End Property
Public Property Get Stations() As FastCollection
  If pStations Is Nothing Then PopulateStations
  Set Stations = pStations
End Property
Public Sub PopulateStations()
  Dim newStation As ssStation
  Dim stationRec As Recordset
  Dim sql$
  
  Set pStations = New FastCollection
  If pDB Is Nothing Then
    ssMessageBox "Cannot read Stations" & vbCr & "DB not set."
  Else
'    sql = "SELECT STATION.* FROM STATION " & _
'        "WHERE STATION.StateCode='" & Me.code & _
'        "' ORDER BY STATION.ROI;"
    
    sql = "SELECT STATION.* " & _
          "FROM [StationState] INNER JOIN STATION ON [StationState].[StaID]=[STATION].[StaID] " & _
          "WHERE [StationState].[StateCode]='" & StateCode & "' " & _
          "AND [StationState].[ROI]>0 ORDER BY [StationState].[ROI];"
    Set stationRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
    While Not stationRec.EOF
      Set newStation = New ssStation
      With newStation
        Set .DB = pDB
        'Set .State = Me
        .ID = stationRec("StaID")
        If Not IsNull(stationRec("StationName")) Then .Name = stationRec("StationName")
        If Not IsNull(stationRec("StationName")) Then .Label = stationRec("StationName")
        If Not IsNull(stationRec("StationTypeCode")) Then Set .StationType = GetStationType(stationRec("StationTypeCode"))
        If Not IsNull(stationRec("IsRegulated")) Then .IsRegulated = stationRec("IsRegulated")
        If Not IsNull(stationRec("PeriodOfRecord")) Then .Period = stationRec("PeriodOfRecord")
        If Not IsNull(stationRec("StationRemarks")) Then .Remarks = stationRec("StationRemarks")
        If Not IsNull(stationRec("Latitude")) Then .Latitude = stationRec("Latitude")
        If Not IsNull(stationRec("Longitude")) Then .Longitude = stationRec("Longitude")
        If Not IsNull(stationRec("HUC")) Then .HUCCode = stationRec("HUC")
        If Not IsNull(stationRec("StatebasinCode")) Then .StatebasinCode = (stationRec("StatebasinCode"))
        If Not IsNull(stationRec("CountyCode")) Then .CountyCode = stationRec("CountyCode")
        If Not IsNull(stationRec("MCDCode")) Then .MCDCode = stationRec("MCDCode")
        If Not IsNull(stationRec("Directions")) Then .Directions = stationRec("Directions")
        'If Not IsNull(stationRec("HasData")) Then .HasData = stationRec("HasData")
        'Set .State = Me
        .IsNew = False
      End With
      pStations.Add newStation, newStation.ID
      Set newStation = Nothing
      stationRec.MoveNext
    Wend
    stationRec.Close
  End If
End Sub

Public Property Set FlowStats(ByVal newValue As FastCollection)
  Set pFlowStats = newValue
End Property
Public Property Get FlowStats() As FastCollection
  If pFlowStats Is Nothing Then PopulateFlowStats
  Set FlowStats = pFlowStats
End Property
Private Sub PopulateFlowStats()
  Dim newParm As ssStatLabel
  Dim myRec As Recordset, statRec As Recordset
  Dim sql$, str$, thisRetPd$
  
  Set pFlowStats = New FastCollection
  If pDB Is Nothing Then
    ssMessageBox "Cannot read ROIPeakFlows" & vbCr & "DB not set."
  Else
    'Read StatisticLabelCodes from STATES table
    sql = "SELECT ROIPeakFlows FROM [STATES] " & _
          "WHERE StateCode='" & StateCode & "';"
    Set myRec = pDB.DB.OpenRecordset(sql)
    If Not IsNull(myRec.Fields("ROIPeakFlows")) Then _
        str = myRec.Fields("ROIPeakFlows")
    'Match StatisticLabelCodes with statistics
    Set statRec = pDB.DB.OpenRecordset("STATLABEL")
    Set Me.DB.StatType = Me.DB.StatisticTypes("PFS")
    While Not Len(Trim(str)) = 0
      thisRetPd = StrSplit(str, ",", "")
      Set newParm = Me.DB.StatType.StatLabels(thisRetPd)
      pFlowStats.Add newParm, newParm.Code
      Set newParm = Nothing
      statRec.MoveNext
    Wend
    statRec.Close
    myRec.Close
  End If
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

