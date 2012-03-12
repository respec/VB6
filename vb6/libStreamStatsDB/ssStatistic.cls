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
Private pIsPreferred As Boolean
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

'More additional attributes
Private pStdError As Double
Private pVariance As Double
Private pLowerCI As Double
Private pUpperCI As Double
Private pStatStartDate As String
Private pStatEndDate As String
Private pStatisticRemarks As String
Private pStatistic_md As String

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

Public Property Get code() As Long
  code = pCode
End Property
Public Property Let code(ByVal newValue As Long)
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

Public Property Get IsPreferred() As Boolean
  IsPreferred = pIsPreferred
End Property
Public Property Let IsPreferred(ByVal newValue As Boolean)
  pIsPreferred = newValue
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

Public Property Get StdError() As Double
  StdError = pStdError
End Property
Public Property Let StdError(ByVal newValue As Double)
  pStdError = newValue
End Property

Public Property Get Variance() As Double
  Variance = pVariance
End Property
Public Property Let Variance(ByVal newValue As Double)
  pVariance = newValue
End Property

Public Property Get LowerCI() As Double
  LowerCI = pLowerCI
End Property
Public Property Let LowerCI(ByVal newValue As Double)
  pLowerCI = newValue
End Property

Public Property Get UpperCI() As Double
  UpperCI = pUpperCI
End Property
Public Property Let UpperCI(ByVal newValue As Double)
  pUpperCI = newValue
End Property

Public Property Get StatStartDate() As String
  StatStartDate = pStatStartDate
End Property
Public Property Let StatStartDate(ByVal newValue As String)
  pStatStartDate = newValue
End Property

Public Property Get StatEndDate() As String
  StatEndDate = pStatEndDate
End Property
Public Property Let StatEndDate(ByVal newValue As String)
  pStatEndDate = newValue
End Property

Public Property Get StatisticRemarks() As String
  StatisticRemarks = pStatisticRemarks
End Property
Public Property Let StatisticRemarks(ByVal newValue As String)
  pStatisticRemarks = newValue
End Property

Public Property Get Statistic_md() As String
  Statistic_md = pStatistic_md
End Property
Public Property Let Statistic_md(ByVal newValue As String)
  pStatistic_md = newValue
End Property

Public Sub Add(ByRef NewValues() As String, StatIndex As Long, Optional NoWarnFlg As Long)
  Dim myRec As Recordset, stationRec As Recordset, searchMatchRec As Recordset
  Dim sql$
  Dim response&
  Dim lStatisticLabelID$
  Dim lAddNew As Boolean
  Dim myMsgBox As ATCoMessage
  
  On Error GoTo 0
  
'  If NewValues(2, StatIndex, 2) = "" Then
'    NewValues(2, StatIndex, 2) = GetLabelID(NewValues(2, StatIndex, 2), pDB)
'  End If
  
  If IsNumeric(NewValues(2, StatIndex, 2)) Then
    lStatisticLabelID = NewValues(2, StatIndex, 2)
  ElseIf IsNumeric(NewValues(2, StatIndex, 3)) Then
    lStatisticLabelID = NewValues(2, StatIndex, 3)
  Else
    lStatisticLabelID = GetLabelID(NewValues(2, StatIndex, 2), pDB)
  End If
  
  sql = "SELECT * FROM STATISTIC " & _
        "WHERE StaID='" & Me.Station.Id & "';"
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  lAddNew = True
  With myRec
    If .RecordCount > 0 Then
      .FindFirst "StatisticLabelID =" & lStatisticLabelID
    End If
    'If .NoMatch Or .RecordCount = 0 Then
      .AddNew
      !staID = Me.Station.Id
      !DataSourceID = GetSourceID(NewValues(2, StatIndex, 8))
      !StatisticLabelID = lStatisticLabelID
      !StatisticValue = NewValues(2, StatIndex, 4)
      If LCase$(NewValues(2, StatIndex, 5)) = "yes" Or LCase$(NewValues(2, StatIndex, 5)) = "y" Then
        !IsPreferred = "Y"
      Else
        !IsPreferred = "N"
      End If
      '!Date = NewValues(2, StatIndex, 6)
      !YearsRec = NewValues(2, StatIndex, 7)
      
      If IsNumeric(NewValues(2, StatIndex, 10)) Then
        !StdError = NewValues(2, StatIndex, 10)
      End If
      If IsNumeric(NewValues(2, StatIndex, 11)) Then
        !Variance = NewValues(2, StatIndex, 11)
      End If
      If IsNumeric(NewValues(2, StatIndex, 12)) Then
        !LowerCI = NewValues(2, StatIndex, 12)
      End If
      If IsNumeric(NewValues(2, StatIndex, 13)) Then
        !UpperCI = NewValues(2, StatIndex, 13)
      End If
      If NewValues(2, StatIndex, 14) <> "" Then
        !StatStartDate = DateValue(NewValues(2, StatIndex, 14))
      End If
      If NewValues(2, StatIndex, 15) <> "" Then
        !StatEndDate = DateValue(NewValues(2, StatIndex, 15))
      End If
      !StatisticRemarks = NewValues(2, StatIndex, 16)
      !Statistic_md = DateTime.Now
      
      If !StatisticLabelID = 0 Then
        NewValues(2, StatIndex, 2) = "bad"
        lAddNew = False
      Else
        '.Update
      End If
'      'If adding first datum for station, change "HasData" field of STATION table to true
'      If myRec.RecordCount > 1 Then
'        .Close
'      Else
'        .Close
''        sql = "SELECT * FROM STATION " & _
''              "WHERE StaID='" & Me.Station.ID & "' " & _
''              "AND StateCode='" & Me.Station.State.code & "';"
'        sql = "SELECT STATION.* " & _
'              "FROM [StationState] INNER JOIN STATION ON [StationState].[StaID]=[STATION].[StaID] " & _
'              "WHERE STATION.StaID='" & Me.Station.ID & "' " & _
'              "AND [StationState].[StateCode]='" & Me.Station.State.Code & "';"
'        Set stationRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
'        With stationRec
'          If Not .EOF Then
'            .Edit
'            !HasData = True
'            .Update
'          End If
'          .Close
'        End With
'      End If

'    Else
'      .Close
'      If NoWarnFlg <> 1 Then
'        ssMessageBox "The statistic '" & NewValues(2, StatIndex, 2) & _
'            "' already exists for station " & Me.Station.Id & "." & vbCrLf & _
'            "The existing value will not be overwritten by the import value.", _
'            vbCritical, "Statistic Already Exists"
'      End If
'    End If
  End With
  
  'Check to overwrite the one record that has matching StaID, StatisticLabelID, and DataSourceID
  Dim lAskUser As Integer
    sql = "SELECT StaID, StatisticLabelID, DataSourceID FROM Statistic " & _
        "WHERE StaId='" & Me.Station.Id & "' " & _
        "AND StatisticLabelID =" & lStatisticLabelID & " " & _
        "AND DataSourceID ='" & myRec.Fields("DataSourceID") & "';"
    Set searchMatchRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
    With searchMatchRec
      If .RecordCount > 0 Then
        .MoveFirst
        Set myMsgBox = New ATCoMessage
        Do While Not .EOF
          If Me.Station.UpdateAction = 0 Then
            Me.Station.UpdateAction = myMsgBox.Show("For station " & Me.Station.Name & ", statistic " & CStr(lStatisticLabelID) & _
                                " already exist." & _
                                "What do you want to do?", "Excel Import - Data Exists", _
                                "&Replace", "Replace &All", "+&Keep", "K&eep All", "-&Cancel")
          ElseIf Me.Station.UpdateAction = 1 Or Me.Station.UpdateAction = 3 Then
            lAskUser = myMsgBox.Show("For station " & Me.Station.Name & ", statistic " & lStatisticLabelID & _
                                " already exist." & _
                                "What do you want to do?", "Excel Import - Data Exists", _
                                "&Replace", "Replace &All", "+&Keep", "K&eep All", "-&Cancel")
            If lAskUser = 2 Or lAskUser = 4 Then Me.Station.UpdateAction = lAskUser
          End If
          If Me.Station.UpdateAction = 1 Or Me.Station.UpdateAction = 2 Then
            If lAddNew Then
              .Delete 'delete the old one
            End If
            'lAddNew = True
          ElseIf Me.Station.UpdateAction = 3 Or Me.Station.UpdateAction = 4 Then
            lAddNew = False
          Else
            lAddNew = False
          End If
          .MoveNext
        Loop
      End If
      .Close
    End With
  
  If lAddNew Then
    'Check if another copy of this statistic exists and if it is also labeled as "IsPreferred"
    sql = "SELECT IsPreferred FROM Statistic " & _
        "WHERE StaId='" & Me.Station.Id & "' " & _
        "AND StatisticLabelID =" & lStatisticLabelID & " " & _
        "AND IsPreferred ='Y';"
    Set searchMatchRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
    With searchMatchRec
      If .RecordCount > 0 Then
        If myRec.Fields("IsPreferred") = "Y" Then
          .MoveFirst
          Do While Not .EOF
            .Edit
            .Fields("IsPreferred") = "N"
            .Update
            .MoveNext
          Loop
        End If
      End If
      .Close
    End With
    
    'Update it finally
    myRec.Update
  End If
  myRec.Close
  Set myMsgBox = Nothing
End Sub

Public Sub Delete()
  Dim myRec As Recordset
  Dim sql$, lSourceIDText$
  
  If Me.SourceID = 0 Then
    lSourceIDText = "00"
  Else
    lSourceIDText = CStr(Me.SourceID)
  End If
  
  sql = "SELECT STATISTIC.StaID, STATLABEL.StatLabel FROM STATLABEL INNER JOIN " & _
      "STATISTIC ON STATISTIC.StatisticLabelID = STATLABEL.StatisticLabelID " & _
      "WHERE STATISTIC.StaID='" & Me.Station.Id & "' " & _
      "AND STATLABEL.StatisticLabelID=" & Me.code & " " & _
      "AND Statistic.StatisticValue ='" & Me.Value & "' " & _
      "AND Statistic.DataSourceID ='" & lSourceIDText & "'"
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    If .RecordCount > 0 Then
      .Delete
    End If
  End With
  myRec.Close
'  'If deleting all datum for station, change "HasData" field of STATION table to true
'  If myRec.RecordCount > 0 Then
'    myRec.Close
'    Exit Sub
'  Else
'    myRec.Close
''    sql = "SELECT * FROM STATION " & _
''          "WHERE StaID='" & Me.Station.ID & "' " & _
''          "AND StateCode='" & Me.Station.State.code & "';"
'    sql = "SELECT STATION.* " & _
'          "FROM [StationState] INNER JOIN STATION ON [StationState].[StaID]=[STATION].[StaID] " & _
'          "WHERE STATION.StaID='" & Me.Station.ID & "' " & _
'          "AND [StationState].[StateCode]='" & Me.Station.State.Code & "';"
'    Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
'    With myRec
'      .Edit
'      !HasData = False
'      .Update
'    End With
'  End If
End Sub

Public Sub Edit(ByRef NewValues() As String, StatIndex As Long)
  Dim myRec As Recordset
  Dim sql$
  Dim flds&
  Dim lMsgErr$
  Dim lSourceIDText$
  
  On Error GoTo 0

  lMsgErr = ""
  
  If Me.SourceID = 0 Then
    lSourceIDText = "00"
  Else
    lSourceIDText = CStr(Me.SourceID)
  End If
  sql = "SELECT STATISTIC.* FROM STATLABEL INNER JOIN STATISTIC ON " & _
      "STATLABEL.StatisticLabelID = STATISTIC.StatisticLabelID " & _
      "WHERE STATLABEL.StatLabel='" & Me.Abbrev & _
      "' AND STATISTIC.StaID='" & Me.Station.Id & "' " & _
      "AND Statistic.DataSourceID='" & lSourceIDText & "';"
  'using dbInconsistent to get recordset not read-only,
  'OK for this case since we are not updating StaID, to which the dbIconsistent option applies
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset, dbInconsistent)
  With myRec
    If .RecordCount > 0 Then
      .Edit
      !DataSourceID = GetSourceID(NewValues(2, StatIndex, 8))
      !StatisticLabelID = NewValues(2, StatIndex, 2)
      !StatisticValue = NewValues(2, StatIndex, 4)
      If LCase(NewValues(2, StatIndex, 5)) = "yes" Or LCase(NewValues(2, StatIndex, 5)) = "y" Then
        !IsPreferred = "Y"
      ElseIf LCase(NewValues(2, StatIndex, 5)) = "no" Or LCase(NewValues(2, StatIndex, 5)) = "n" Then
        !IsPreferred = "N"
      Else
        !IsPreferred = "N"
      End If
      '!Date = NewValues(2, StatIndex, 6)
      !YearsRec = NewValues(2, StatIndex, 7)
      If IsNumeric(NewValues(2, StatIndex, 7)) Then
        !StdError = NewValues(2, StatIndex, 10)
      Else
        'lMsgErr = "StdErr is NAN." & vbCrLf
      End If
      If IsNumeric(NewValues(2, StatIndex, 11)) Then
        !Variance = NewValues(2, StatIndex, 11)
      Else
        'lMsgErr = lMsgErr & "Variance is NAN." & vbCrLf
      End If
      If IsNumeric(NewValues(2, StatIndex, 12)) Then
        !LowerCI = NewValues(2, StatIndex, 12)
      Else
        'lMsgErr = lMsgErr & "LowerCI is NAN." & vbCrLf
      End If
      If IsNumeric(NewValues(2, StatIndex, 13)) Then
        !UpperCI = NewValues(2, StatIndex, 13)
      Else
        'lMsgErr = lMsgErr & "UpperCI is NAN." & vbCrLf
      End If
      If NewValues(2, StatIndex, 14) <> "" Then
        !StatStartDate = DateValue(NewValues(2, StatIndex, 14))
      End If
      If NewValues(2, StatIndex, 15) <> "" Then
        !StatEndDate = DateValue(NewValues(2, StatIndex, 15))
      End If
      !StatisticRemarks = NewValues(2, StatIndex, 16)
      !Statistic_md = DateTime.Now
      If Len(lMsgErr) > 0 Then
        'Do not update as there is error
        .Close
        'Msgbox("This statistic has the following problems, update failed:" & vbcrlf & lMsg, vbInformation, "Not updated")
      Else
        .Update
      End If
    Else
      ssMessageBox "There is no " & Me.Name & " Statistic for " & Me.Station.Name & _
             "." & vbCrLf & "Statistic not found."
    End If
    .Close
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

