VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ssStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pDB As nssDatabase
Private pID As String
Private pName As String
Private pLabel As String
Private pStationType As ssStationType
Private pIsRegulated As String
Private pPeriod As String
Private pRemarks As String
Private pLatitude As Single
Private pLongitude As Single
Private pHUCCode As String
Private pStatebasinCode As String
Private pCountyCode As String
Private pMCDCode As String
Private pStatebasin As ssStateBasin
Private pCounty As ssCounty
Private pMCD As ssMCD
Private pHUC As ssHUC
Private pRegion As nssRegion
Private pDirections As String
Private pIsROI As Boolean
Private pROIIndex As Long
Private pROIRegionID As Long
Private pIsNew As Boolean
'Private pHasData As Boolean
Private pStateCode As String
Private pDistrictCode As String
Private pState As nssState
Private pStatistics As FastCollection 'of ssStatistics kept by this station
Private pSelStats As FastCollection 'of ssStatistics selected for this station

Public Property Get DB() As nssDatabase
  Set DB = pDB
End Property
Public Property Set DB(newValue As nssDatabase)
  Set pDB = newValue
End Property

Public Property Get ID() As String
  ID = pID
End Property
Public Property Let ID(ByVal newValue As String)
  pID = newValue
End Property

Public Property Get Name() As String
  Name = pName
End Property
Public Property Let Name(ByVal newValue As String)
  pName = newValue
End Property

Public Property Get Label() As String
  Label = pLabel
End Property
Public Property Let Label(ByVal newValue As String)
  pLabel = newValue
End Property

Public Property Get StationType() As ssStationType
  Set StationType = pStationType
End Property
Public Property Set StationType(ByVal newValue As ssStationType)
  Set pStationType = newValue
End Property

Public Property Get IsRegulated() As String
  IsRegulated = pIsRegulated
End Property
Public Property Let IsRegulated(ByVal newValue As String)
  pIsRegulated = newValue
End Property

Public Property Get Period() As String
  Period = pPeriod
End Property
Public Property Let Period(ByVal newValue As String)
  pPeriod = newValue
End Property

Public Property Get Remarks() As String
  Remarks = pRemarks
End Property
Public Property Let Remarks(ByVal newValue As String)
  pRemarks = newValue
End Property

Public Property Get Latitude() As Single
  Latitude = pLatitude
End Property
Public Property Let Latitude(ByVal newValue As Single)
  pLatitude = newValue
End Property

Public Property Get Longitude() As Single
  Longitude = pLongitude
End Property
Public Property Let Longitude(ByVal newValue As Single)
  pLongitude = newValue
End Property

Public Property Get HUCCode() As String
  HUCCode = pHUCCode
End Property
Public Property Let HUCCode(ByVal newValue As String)
  pHUCCode = newValue
End Property

Public Property Get StatebasinCode() As String
  StatebasinCode = pStatebasinCode
End Property
Public Property Let StatebasinCode(ByVal newValue As String)
  pStatebasinCode = newValue
End Property

Public Property Get CountyCode() As String
  CountyCode = pCountyCode
End Property
Public Property Let CountyCode(ByVal newValue As String)
  pCountyCode = newValue
End Property

Public Property Get MCDCode() As String
  MCDCode = pMCDCode
End Property
Public Property Let MCDCode(ByVal newValue As String)
  pMCDCode = newValue
End Property

Public Property Get Statebasin() As ssStateBasin
  Set Statebasin = pStatebasin
End Property
Public Property Set Statebasin(ByVal newValue As ssStateBasin)
  Set pStatebasin = newValue
End Property

Public Property Get County() As ssCounty
  Set County = pCounty
End Property
Public Property Set County(ByVal newValue As ssCounty)
  Set pCounty = newValue
End Property

Public Property Get MCD() As ssMCD
  Set MCD = pMCD
End Property
Public Property Set MCD(ByVal newValue As ssMCD)
  Set pMCD = newValue
End Property

Public Property Get HUC() As ssHUC
  Set HUC = pHUC
End Property
Public Property Set HUC(ByVal newValue As ssHUC)
  Set pHUC = newValue
End Property

Public Property Get Region() As nssRegion
  Set Region = pRegion
End Property
Public Property Set Region(ByVal newValue As nssRegion)
  Set pRegion = newValue
End Property

Public Property Get State() As nssState
  Set State = pState
End Property
Public Property Set State(ByVal newValue As nssState)
  Set pState = newValue
End Property

Public Property Get Directions() As String
  Directions = pDirections
End Property
Public Property Let Directions(ByVal newValue As String)
  pDirections = newValue
End Property

Public Property Get IsROI() As Boolean
  IsROI = pIsROI
End Property
Public Property Let IsROI(ByVal newValue As Boolean)
  pIsROI = newValue
End Property

Public Property Get ROIIndex() As Long
  ROIIndex = pROIIndex
End Property
Public Property Let ROIIndex(ByVal newValue As Long)
  pROIIndex = newValue
End Property

Public Property Get ROIRegionID() As Long
  ROIRegionID = pROIRegionID
End Property
Public Property Let ROIRegionID(ByVal newValue As Long)
  pROIRegionID = newValue
End Property

Public Property Get IsNew() As Boolean
  IsNew = pIsNew
End Property
Public Property Let IsNew(ByVal newValue As Boolean)
  pIsNew = newValue
End Property

'Public Property Get HasData() As Boolean
'  HasData = pHasData
'End Property
'Public Property Let HasData(ByVal newValue As Boolean)
'  pHasData = newValue
'End Property

Public Property Get StateCode() As String
  StateCode = pStateCode
End Property
Public Property Let StateCode(ByVal newValue As String)
  pStateCode = newValue
End Property

Public Property Get DistrictCode() As String
  DistrictCode = pDistrictCode
End Property
Public Property Let DistrictCode(ByVal newValue As String)
  pDistrictCode = newValue
End Property

Public Sub Add(ByRef NewValues() As String, ByVal StaIndex As Long, _
               Optional ImportFlg As Long, Optional ROIRegionID As Long = 0)
  Dim myRec As Recordset
  Dim fld&, ROIIndex As Long
  Dim sql$
  
  On Error Resume Next
  
  'instead of negative SatIndex kluge for indicating ROI station, just use IsROI property
  'prh, 2/2009
  'If StaIndex < 0 Then 'actually passing ROI station index
  If Me.IsROI Then
    'ROIIndex = Abs(StaIndex)
    'no need for negative StaIndex
    ROIIndex = StaIndex
    StaIndex = 1
  Else
    ROIIndex = 0
  End If
  If Len(NewValues(2, StaIndex, 1)) = 7 Then 'update 7-digit code to standard of 8 w/preceeding 0
    NewValues(2, StaIndex, 1) = "0" & NewValues(2, StaIndex, 1)
  End If
'  sql = "SELECT * FROM STATION WHERE StateCode='" & Me.State.code & "'"

  sql = "SELECT STATION.* " & _
        "FROM [StationState] INNER JOIN STATION ON [StationState].[StaID]=[STATION].[StaID] " & _
        "WHERE [StationState].[StateCode]='" & Me.State.Code & "';"
  
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    If .RecordCount > 0 Then 'check to see if station with this ID or name already exists
      .FindFirst "StaID='" & NewValues(2, StaIndex, 1) & "'"
      If Not .NoMatch Then  'Station with this code already exists for this state
        If ImportFlg <> 1 Then
          ssMessageBox "There is already a station in " & pDB.State.Name & " with the code " & _
              NewValues(2, StaIndex, 1) & "." & vbCrLf & "No new station was created.", _
               vbExclamation, "Code already exists"
        End If
        Me.Name = "bad"
        GoTo AddStateStationReference
      End If
      .FindFirst "StationName='" & NewValues(2, StaIndex, 2) & "'"
      If Not .NoMatch Then  'Station with this name already exists for this state
        If ImportFlg <> 1 Then
          ssMessageBox "There is already a station in " & pDB.State.Name & " with the name " & _
              NewValues(2, StaIndex, 2) & "." & vbCrLf & "No new station was created.", _
              vbExclamation, "Name already exists"
        End If
        Me.Name = "bad"
        GoTo AddStateStationReference
      End If
    End If
    .AddNew
    For fld = 1 To UBound(NewValues, 3)  'currently set to 15
      If Trim(NewValues(2, StaIndex, fld)) <> "" Then
        Select Case fld
          Case 1: .Fields("StaID") = Trim(NewValues(2, StaIndex, fld))
          Case 2: .Fields("StationName") = Trim(NewValues(2, StaIndex, fld))
          Case 3:  NewValues(2, StaIndex, fld) = GetStationTypeCode(NewValues(2, StaIndex, fld))
                  .Fields("StationTypeCode") = Trim(NewValues(2, StaIndex, fld))
          Case 4: .Fields("IsRegulated") = Trim(NewValues(2, StaIndex, fld))
          Case 5: .Fields("PeriodOfRecord") = Trim(NewValues(2, StaIndex, fld))
          Case 6: .Fields("Directions") = Trim(NewValues(2, StaIndex, fld))
          Case 7: .Fields("StationRemarks") = Trim(NewValues(2, StaIndex, fld))
          Case 8: .Fields("Latitude") = Trim(NewValues(2, StaIndex, fld))
          Case 9: .Fields("Longitude") = Trim(NewValues(2, StaIndex, fld))
          Case 10: .Fields("DistrictCode") = GetStateCode(Trim(NewValues(2, StaIndex, fld)))
          Case 11: .Fields("StateCode") = GetStateCode(Trim(NewValues(2, StaIndex, fld)))
          Case 12: .Fields("CountyCode") = Trim(NewValues(2, StaIndex, fld))
          Case 13: .Fields("MCDCode") = Trim(NewValues(2, StaIndex, fld))
          Case 14: .Fields("HUC") = Trim(NewValues(2, StaIndex, fld))
          Case 15: .Fields("StateBasinCode") = Trim(NewValues(2, StaIndex, fld))
        End Select
      End If
    Next fld
'    If ImportFlg = 1 Then
'      .Fields("HasData") = 1
'    Else
'      .Fields("HasData") = 0
'    End If
'this is now stored on StationState table
'    .Fields("StateCode") = Me.State.code
    .Update
  End With

AddStateStationReference:
  sql = "SELECT * FROM [StationState] WHERE StateCode='" & Me.State.Code & "'"
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    If .RecordCount > 0 Then 'check to see if station with this ID or name already exists
      If ROIRegionID <> 0 Then
        .FindFirst "StaID='" & NewValues(2, StaIndex, 1) & "' AND ROIRegionID=0"
      Else
        .FindFirst "StaID='" & NewValues(2, StaIndex, 1) & "'"
      End If
    End If
    If .NoMatch Or .RecordCount <= 0 Then 'add this station for this state
      .AddNew
      .Fields("StaID") = Trim(NewValues(2, StaIndex, 1))
      .Fields("StateCode") = Me.State.Code
      If Me.IsROI Then 'assign station index for ordering
        .Fields("ROI") = ROIIndex
      Else
        .Fields("ROI") = Me.IsROI
      End If
      .Fields("ROIRegionID") = ROIRegionID 'link this station's record to a specific ROI region
      .Update
    ElseIf Not .NoMatch And ROIIndex > 0 Then 'importing ROI, update ROI field
      .Edit
      .Fields("ROI") = ROIIndex
      .Fields("ROIRegionID") = ROIRegionID 'link this station's record to a specific ROI region
      .Update
    End If
  End With
End Sub

Public Sub Delete()
  Dim myRec As Recordset
  Dim sql$

'  sql = "SELECT * FROM STATION WHERE StateCode='" & pDB.State.code & "'"
  sql = "SELECT STATION.* " & _
        "FROM [StationState] INNER JOIN STATION ON [StationState].[StaID]=[STATION].[StaID] " & _
        "WHERE [StationState].[StateCode]='" & pDB.State.Code & "';"
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    .FindFirst "StaID='" & Me.ID & "'"
    If Not .NoMatch Then
      .Delete
    End If
  End With
' now delete from State Station table
  sql = "SELECT * FROM [StationState] WHERE StateCode='" & pDB.State.Code & "'"
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    .FindFirst "StaID='" & Me.ID & "'"
    If Not .NoMatch Then
      .Delete
    End If
  End With
  
End Sub

Public Sub Edit(ByRef NewValues() As String, StaIndex As Long)
  Dim myRec As Recordset
  Dim sql$
  Dim fld&
  
'  sql = "SELECT * FROM STATION WHERE StateCode='" & pDB.State.code & "'"
  sql = "SELECT STATION.* " & _
        "FROM [StationState] INNER JOIN STATION ON [StationState].[StaID]=[STATION].[StaID] " & _
        "WHERE [StationState].[StateCode]='" & pDB.State.Code & "';"
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    .FindFirst "StaID='" & Me.ID & "'"
    If Not .NoMatch Then
      .Edit
      For fld = 2 To UBound(NewValues, 3)
        If fld = 3 Then NewValues(2, StaIndex, fld) = _
            GetStationTypeCode(NewValues(2, StaIndex, fld))
        If (fld = 15 And NewValues(2, StaIndex, fld) = "") Then
          .Fields(fld - 1) = Null
        ElseIf Len(NewValues(2, StaIndex, fld)) > 0 Then 'only update if array element contains something
'          If fld > 13 Then 'State and District codes fall after HasData field on Station table
          If fld = 10 Or fld = 11 Then 'State or District codes
            .Fields(fld - 1) = GetStateCode(NewValues(2, StaIndex, fld))
          Else
            .Fields(fld - 1) = NewValues(2, StaIndex, fld)
          End If
        End If
      Next fld
      .Update
    Else
      ssMessageBox "There is no station in " & pDB.State.Name & _
             " with the name " & Me.Name & "." & vbCrLf & "Station not found."
    End If
    .Close
  End With
End Sub

Public Property Set SelStats(ByRef newValue As FastCollection)
  Set pSelStats = newValue
End Property
Public Property Get SelStats() As FastCollection
  If pSelStats Is Nothing Then Set pSelStats = New FastCollection
  Set SelStats = pSelStats
End Property

Public Property Set Statistics(ByRef newValue As FastCollection)
  Set pStatistics = newValue
End Property
Public Property Get Statistics() As FastCollection
  ' Code must be set for this state before this will work
  If pStatistics Is Nothing Then PopulateStatistics
  Set Statistics = pStatistics
End Property

Public Sub PopulateStatistics()
  Dim newStatistic As ssStatistic
  Dim statRec As Recordset
  Dim sql$
  
  Set pStatistics = New FastCollection
  If pDB Is Nothing Then
    ssMessageBox "Cannot read Statistics" & vbCr & "DB not set."
  Else
    'sql = "SELECT STATISTIC.StatisticValue, STATISTIC.Date, STATLABEL.StatLabel, " &
    sql = "SELECT STATISTIC.StatisticValue, STATISTIC.YearsRec, STATLABEL.StatLabel, " & _
        "STATLABEL.StatisticLabelID, STATLABEL.StatisticLabel, STATLABEL.Unitid, " & _
        "STATTYPE.StatisticTypeCode, STATTYPE.StatisticTypeID, STATTYPE.StatisticType, " & _
        "DATASOURCE.DataSourceID, DATASOURCE.Citation, DATASOURCE.CitationURL " & _
        "FROM DATASOURCE RIGHT JOIN (STATTYPE INNER JOIN (STATLABEL INNER JOIN STATISTIC ON " & _
        "STATLABEL.StatisticLabelID = STATISTIC.StatisticLabelID) ON " & _
        "STATTYPE.StatisticTypeCode = STATLABEL.StatisticTypeCode) ON " & _
        "DATASOURCE.DataSourceID = STATISTIC.DataSourceID " & _
        "WHERE STATISTIC.StaID='" & pID & "' " & _
        "ORDER BY STATTYPE.StatisticTypeID, STATLABEL.StatisticLabelID;"
    Set statRec = pDB.DB.OpenRecordset(sql)
    While Not statRec.EOF
      Set newStatistic = New ssStatistic
      With newStatistic
        Set .DB = pDB
        Set .Station = Me
        .Code = statRec.Fields("StatisticLabelID")
        .Abbrev = statRec.Fields("StatLabel")
        .Name = statRec.Fields("StatisticLabel")
        If Not IsNull(statRec.Fields("StatisticValue")) Then _
            .Value = statRec.Fields("StatisticValue")
        If Not IsNull(statRec("UnitID")) Then _
            Set .Units = pDB.Units(CStr(statRec("UnitID")))
'        If Not IsNull(statRec.Fields("Date")) Then _
'            .RecDate = statRec.Fields("Date")
        If Not IsNull(statRec.Fields("YearsRec")) Then _
          .YearsRec = statRec.Fields("YearsRec")
        .StatTypeID = statRec.Fields("StatisticTypeID")
        .StatTypeCode = statRec.Fields("StatisticTypeCode")
        .StatType = statRec.Fields("StatisticType")
        If Not IsNull(statRec.Fields("DataSourceID")) Then _
            .SourceID = statRec.Fields("DataSourceID")
        If Not IsNull(statRec.Fields("Citation")) Then _
            .Source = statRec.Fields("Citation")
        If Not IsNull(statRec.Fields("CitationURL")) Then _
            .SourceURL = statRec.Fields("CitationURL")
      End With
      pStatistics.Add newStatistic, CStr(newStatistic.Code)
      Set newStatistic = Nothing
      statRec.MoveNext
    Wend
    statRec.Close
  End If
End Sub

Private Function GetStationTypeCode(TypeName As String) As String
  Dim myRec As Recordset
  
  Set myRec = pDB.DB.OpenRecordset("STATIONTYPE", dbOpenSnapshot)
  With myRec
    .FindFirst "StationType='" & TypeName & "'"
    If .NoMatch Then 'try matching StationTypeCode field
      .FindFirst "StationTypeCode='" & TypeName & "'"
    End If
    If Not .NoMatch Then
      GetStationTypeCode = .Fields("StationTypeCode")
    End If
  End With
End Function

Private Function GetStateCode(StateAbbrev As String) As String
  Dim myRec As Recordset
  
  Set myRec = pDB.DB.OpenRecordset("STATES", dbOpenSnapshot)
  With myRec
    .FindFirst "St='" & StateAbbrev & "'"
    If Not .NoMatch Then
      GetStateCode = .Fields("StateCode")
    End If
  End With
End Function

Private Sub Class_Initialize()
  IsROI = False
End Sub
