VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ssStatistic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pDB As nssDatabase
Private pStation As ssStation
Private pCode As Long
Private pAbbrev As String
Private pName As String
Private pROI As Boolean
Private pValue As String
'Private pUnits As String
Private pUnits As nssUnits
'Private pRecDate As String
Private pStatTypeID As Long
Private pStatTypeCode As String
Private pStatType As String
Private pSourceID As Long
Private pSource As String
Private pSourceURL As String
Private pYearsRec As String

Public Property Get DB() As nssDatabase
  Set DB = pDB
End Property
Public Property Set DB(newValue As nssDatabase)
  Set pDB = newValue
End Property

Public Property Get Station() As ssStation
  Set Station = pStation
End Property
Public Property Set Station(newValue As ssStation)
  Set pStation = newValue
End Property

Public Property Get Code() As Long
  Code = pCode
End Property
Public Property Let Code(ByVal newValue As Long)
  pCode = newValue
End Property

Public Property Get Abbrev() As String
  Abbrev = pAbbrev
End Property
Public Property Let Abbrev(ByVal newValue As String)
  pAbbrev = newValue
End Property

Public Property Get Name() As String
  Name = pName
End Property
Public Property Let Name(ByVal newValue As String)
  pName = newValue
End Property

Public Property Get ROI() As Boolean
  ROI = pROI
End Property
Public Property Let ROI(ByVal newValue As Boolean)
  pROI = newValue
End Property

Public Property Get Value() As String
  Value = pValue
End Property
Public Property Let Value(ByVal newValue As String)
  pValue = newValue
End Property

Public Property Get Units() As nssUnits
  Set Units = pUnits
End Property
Public Property Set Units(ByVal newValue As nssUnits)
  Set pUnits = newValue
End Property

'Public Property Get Units() As String
'  Units = pUnits
'End Property
'Public Property Let Units(ByVal newValue As String)
'  pUnits = newValue
'End Property

'Public Property Get RecDate() As String
'  RecDate = pRecDate
'End Property
'Public Property Let RecDate(ByVal newValue As String)
'  pRecDate = newValue
'End Property
'
Public Property Get StatTypeID() As Long
  StatTypeID = pStatTypeID
End Property
Public Property Let StatTypeID(ByVal newValue As Long)
  pStatTypeID = newValue
End Property

Public Property Get StatTypeCode() As String
  StatTypeCode = pStatTypeCode
End Property
Public Property Let StatTypeCode(ByVal newValue As String)
  pStatTypeCode = newValue
End Property

Public Property Get StatType() As String
  StatType = pStatType
End Property
Public Property Let StatType(ByVal newValue As String)
  pStatType = newValue
End Property

Public Property Get SourceID() As Long
  SourceID = pSourceID
End Property
Public Property Let SourceID(ByVal newValue As Long)
  pSourceID = newValue
End Property

Public Property Get Source() As String
  Source = pSource
End Property
Public Property Let Source(ByVal newValue As String)
  pSource = newValue
End Property

Public Property Get SourceURL() As String
  SourceURL = pSourceURL
End Property
Public Property Let SourceURL(ByVal newValue As String)
  pSourceURL = newValue
End Property

Public Property Get YearsRec() As String
  YearsRec = pYearsRec
End Property
Public Property Let YearsRec(ByVal newValue As String)
  pYearsRec = newValue
End Property

Public Sub Add(ByRef NewValues() As String, StatIndex As Long, Optional NoWarnFlg As Long)
  Dim myRec As Recordset, stationRec As Recordset
  Dim sql$
  Dim response&
  
  On Error GoTo 0
  
  If NewValues(2, StatIndex, 2) = "" Then
    NewValues(2, StatIndex, 2) = GetLabelID(NewValues(2, StatIndex, 3))
  End If
  
  sql = "SELECT * FROM STATISTIC " & _
        "WHERE StaID='" & Me.Station.ID & "';"
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    If .RecordCount > 0 Then
      .FindFirst "StatisticLabelCode =" & NewValues(2, StatIndex, 2)
    End If
    If .NoMatch Or .RecordCount = 0 Then
      .AddNew
      !StaID = Me.Station.ID
      !DataSourceID = GetSourceID(NewValues(2, StatIndex, 7))
      !StatisticLabelCode = NewValues(2, StatIndex, 2)
      !StatisticValue = NewValues(2, StatIndex, 4)
      '!Date = NewValues(2, StatIndex, 6)
      !YearsRec = NewValues(2, StatIndex, 6)
      If !StatisticLabelCode = 0 Then
        NewValues(2, StatIndex, 2) = "bad"
      Else
        .Update
      End If
      'If adding first datum for station, change "HasData" field of STATION table to true
      If myRec.RecordCount > 1 Then
        .Close
      Else
        .Close
'        sql = "SELECT * FROM STATION " & _
'              "WHERE StaID='" & Me.Station.ID & "' " & _
'              "AND StateCode='" & Me.Station.State.code & "';"
        sql = "SELECT STATION.* " & _
              "FROM [StationState] INNER JOIN STATION ON [StationState].[StaID]=[STATION].[StaID] " & _
              "WHERE STATION.StaID='" & Me.Station.ID & "' " & _
              "AND [StationState].[StateCode]='" & Me.Station.State.Code & "';"
        Set stationRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
        With stationRec
          If Not .EOF Then
            .Edit
            !HasData = True
            .Update
          End If
          .Close
        End With
      End If
    Else
      .Close
      If NoWarnFlg <> 1 Then
        ssMessageBox "The statistic '" & NewValues(2, StatIndex, 2) & _
            "' already exists for station " & Me.Station.ID & "." & vbCrLf & _
            "The existing value will not be overwritten by the import value.", _
            vbCritical, "Statistic Already Exists"
      End If
    End If
  End With
End Sub

Public Sub Delete()
  Dim myRec As Recordset
  Dim sql$
  
  sql = "SELECT STATISTIC.StaID, STATLABEL.StatLabel FROM STATLABEL INNER JOIN " & _
      "STATISTIC ON STATLABEL.StatisticLabelCode = STATISTIC.StatisticLabelCode " & _
      "WHERE STATISTIC.StaID='" & Me.Station.ID & _
      "' AND STATLABEL.StatisticLabelCode=" & Me.Code
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    If .RecordCount > 0 Then
      .Delete
    End If
  End With
  
  'If deleting all datum for station, change "HasData" field of STATION table to true
  If myRec.RecordCount > 0 Then
    myRec.Close
    Exit Sub
  Else
    myRec.Close
'    sql = "SELECT * FROM STATION " & _
'          "WHERE StaID='" & Me.Station.ID & "' " & _
'          "AND StateCode='" & Me.Station.State.code & "';"
    sql = "SELECT STATION.* " & _
          "FROM [StationState] INNER JOIN STATION ON [StationState].[StaID]=[STATION].[StaID] " & _
          "WHERE STATION.StaID='" & Me.Station.ID & "' " & _
          "AND [StationState].[StateCode]='" & Me.Station.State.Code & "';"
    Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
    With myRec
      .Edit
      !HasData = False
      .Update
    End With
  End If
End Sub

Public Sub Edit(ByRef NewValues() As String, StatIndex As Long)
  Dim myRec As Recordset
  Dim sql$
  Dim flds&
  
  On Error GoTo 0

  sql = "SELECT STATISTIC.* FROM STATLABEL INNER JOIN STATISTIC ON " & _
      "STATLABEL.StatisticLabelCode = STATISTIC.StatisticLabelCode " & _
      "WHERE STATLABEL.StatLabel='" & Me.Abbrev & _
      "' AND STATISTIC.StaID='" & Me.Station.ID & "';"
  'using dbInconsistent to get recordset not read-only,
  'OK for this case since we are not updating StaID, to which the dbIconsistent option applies
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset, dbInconsistent)
  With myRec
    If .RecordCount > 0 Then
      .Edit
      !DataSourceID = GetSourceID(NewValues(2, StatIndex, 7))
      !StatisticLabelCode = NewValues(2, StatIndex, 2)
      !StatisticValue = NewValues(2, StatIndex, 4)
      '!Date = NewValues(2, StatIndex, 6)
      !YearsRec = NewValues(2, StatIndex, 6)
      .Update
    Else
      ssMessageBox "There is no " & Me.Name & " Statistic for " & Me.Station.Name & _
             "." & vbCrLf & "Statistic not found."
    End If
  End With
End Sub

Public Function GetSourceID(Citation As String) As String
  Dim myRec As Recordset
  Dim sql As String
  
  sql = "SELECT DATASOURCE.* FROM DATASOURCE WHERE DATASOURCE.Citation='" & Citation & "';"
  'using dbInconsistent to get recordset not read-only,
  'OK for this case since we are not updating StaID, to which the dbIconsistent option applies
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
'  Set myRec = pDB.DB.OpenRecordset("DATASOURCE", dbOpenSnapshot)
  With myRec
    .FindFirst "Citation='" & Citation & "'"
    If Not .NoMatch Then GetSourceID = .Fields("DataSourceID")
  End With
End Function

Public Function GetLabelID(StatLabel As String) As Long
  Dim myRec As Recordset
  
  Set myRec = pDB.DB.OpenRecordset("STATLABEL", dbOpenSnapshot)
  With myRec
    .FindFirst "StatLabel='" & StatLabel & "'"
    If .NoMatch Then .FindFirst "StatisticLabel='" & StatLabel & "'"
    If Not .NoMatch Then GetLabelID = .Fields("StatisticLabelCode")
  End With
End Function

