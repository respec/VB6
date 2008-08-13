VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nssParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private pDB As nssDatabase
Private pID As Long
Private pLabelCode As Long
Private pAbbrev As String
Private pName As String
Private pNSSName As String
Private pRegion As nssRegion
Private pMin As Double
Private pMax As Double
Private pConvFlag As Long
Private pUnits As nssUnits
Private pIsNew As Boolean
Private pStatTypeCode As String
Private pStatType As String
Private pCorrelationType As Single '0=none, <0-neg only, >0-pos only, >1000-pos or neg (value divided by 1000)
Private pSimulationVar As Boolean
Private pRegressionVar As Boolean

Public Property Get DB() As nssDatabase
  Set DB = pDB
End Property
Public Property Set DB(newValue As nssDatabase)
  Set pDB = newValue
End Property

Public Property Get Name() As String
  Name = pName
End Property
Public Property Let Name(ByVal newValue As String)
  pName = newValue
End Property

Public Property Get NSSName() As String
  NSSName = pNSSName
End Property
Public Property Let NSSName(ByVal newValue As String)
  pNSSName = newValue
End Property

Public Property Get ID() As Long
  ID = pID
End Property
Public Property Let ID(ByVal newValue As Long)
  pID = newValue
End Property

Public Property Get LabelCode() As Long
  LabelCode = pLabelCode
End Property
Public Property Let LabelCode(ByVal newValue As Long)
  pLabelCode = newValue
End Property

Public Property Get Abbrev() As String
  Abbrev = pAbbrev
End Property
Public Property Let Abbrev(ByVal newValue As String)
  pAbbrev = newValue
End Property

Public Function GetMin(WantMetric As Boolean) As Double
  GetMin = pUnits.Convert(pMin, pRegion.State.Metric, WantMetric)
End Function
Public Sub SetMin(ByVal newValue As Double, IsMetric As Boolean)
  pMin = pUnits.Convert(newValue, IsMetric, pRegion.State.Metric)
End Sub

Public Function GetMax(WantMetric As Boolean) As Double
  GetMax = pUnits.Convert(pMax, pRegion.State.Metric, WantMetric)
End Function
Public Sub SetMax(ByVal newValue As Double, IsMetric As Boolean)
  pMax = pUnits.Convert(newValue, IsMetric, pRegion.State.Metric)
End Sub

Public Property Get ConvFlag() As Long
  ConvFlag = pConvFlag
End Property
Public Property Let ConvFlag(ByVal newValue As Long)
  pConvFlag = newValue
End Property

Public Property Get Units() As nssUnits
  Set Units = pUnits
End Property
Public Property Set Units(ByVal newValue As nssUnits)
  Set pUnits = newValue
End Property

Public Property Get IsNew() As Boolean
  IsNew = pIsNew
End Property
Public Property Let IsNew(ByVal newValue As Boolean)
  pIsNew = newValue
End Property

Public Property Get Region() As nssRegion
  Set Region = pRegion
End Property
Public Property Set Region(newValue As nssRegion)
  Set pRegion = newValue
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

Public Property Get CorrelationType() As Single
  CorrelationType = pCorrelationType
End Property
Public Property Let CorrelationType(ByVal newValue As Single)
  pCorrelationType = newValue
End Property

Public Property Get SimulationVar() As Boolean
  SimulationVar = pSimulationVar
End Property
Public Property Let SimulationVar(ByVal newValue As Boolean)
  pSimulationVar = newValue
End Property

Public Property Get RegressionVar() As Boolean
  RegressionVar = pRegressionVar
End Property
Public Property Let RegressionVar(ByVal newValue As Boolean)
  pRegressionVar = newValue
End Property

Public Function Add(ByVal Region As nssRegion, ByVal StatLabel As String, _
    ByVal Min As String, ByVal Max As String, ByVal ConvFlag As String) As Boolean
  Dim myRec As Recordset, myRec2 As Recordset
  Dim ID As Long
  Dim statLabelCode As String, sql As String
  
  statLabelCode = GetLabelID(StatLabel, Region.DB)
  If statLabelCode > 0 Then
    Set myRec = Region.DB.DB.OpenRecordset("Parameters", dbOpenDynaset)
    With myRec
      If .RecordCount > 0 Then
        'check to see if parameter with this name already exists for this region
        .FindFirst "StatisticLabelID=" & statLabelCode & _
                   " AND RegionID=" & Region.ID
      End If
      If .NoMatch Or .RecordCount = 0 Then  'OK to add
        'find an unused ID
        If .RecordCount > 0 Then
          .MoveLast
          ID = !ParmID
        Else
          ID = 1
        End If
        .FindFirst "ParmID=" & ID
        While Not .NoMatch
          ID = !ParmID + 1
          .FindFirst "ParmID=" & ID
        Wend
        Me.ID = ID
        .AddNew
        !RegionID = Region.ID
        !ParmID = ID
        !StatisticLabelID = statLabelCode
        'get long label name from STATLABEL table and assign to Parameter field
        Set myRec2 = Region.DB.DB.OpenRecordset("STATLABEL", dbOpenDynaset)
        myRec2.FindFirst "StatisticLabelID=" & statLabelCode
        !Parameter = myRec2!StatisticLabel
        If IsNumeric(Min) Then !Min = Min
        If IsNumeric(Max) Then !Max = Max
        If IsNumeric(ConvFlag) Then !UnitID = ConvFlag
        .Update
        Add = True
      Else
        ssMessageBox "There is already a parameter in " & Region.Name & " with the name " _
               & StatLabel & "." & vbCrLf & "No new parameter was created."
      End If
    End With
  Else 'stat not found on STATLABEL table
    ssMessageBox "Parameter " & StatLabel & " not found on the " & _
           "STATLABEL table." & vbCrLf & "No new parameter was created." & _
           "Use the StreamStatDB Statistic Management tab to add new Statistics.", vbExclamation
  End If
End Function

Public Function AddROIUserParm(ByVal Region As nssRegion, ByVal StatLabel As String, _
    ByVal CorrelationType As String, ByVal SimVar As Boolean, ByVal RegVar As Boolean) As Boolean
  Dim myRec As Recordset
  Dim ID As Long
  Dim statLabelCode As String, sql As String
  
  statLabelCode = GetLabelID(StatLabel, Region.DB)
  sql = "SELECT ROIUserParms.* " & _
        "From ROIUserParms " & _
        "WHERE RegionID=" & Me.Region.ID & ";"
  Set myRec = Region.DB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    If .RecordCount > 0 Then
      .FindFirst "ParmID=" & statLabelCode
    End If
    If .NoMatch Or .RecordCount = 0 Then
      .AddNew
      !RegionID = Me.Region.ID
      !ParmID = statLabelCode
      If IsNumeric(CorrelationType) Then !PosNegCorrelation = CorrelationType
      !SimVar = SimVar
      !RegVar = RegVar
      .Update
      AddROIUserParm = True
    End If
    .Close
  End With
End Function

Public Function Delete() As Boolean
  Dim myRec As Recordset
  Dim sql$
  
  sql = "SELECT * FROM Parameters WHERE RegionID=" & Me.Region.ID
  Set myRec = Me.Region.DB.DB.OpenRecordset("Parameters", dbOpenDynaset)
  With myRec
    .FindFirst "ID=" & Me.ID
    If Not .NoMatch Then
      .Delete
      Delete = True
    Else
      Delete = False
    End If
  End With
End Function

Public Sub Edit(ByVal StatLabel As String, ByVal Min As String, _
    ByVal Max As String, ByVal ConvFlag As Integer)
  Dim myRec As Recordset
  
  Set myRec = Me.Region.DB.DB.OpenRecordset("Parameters", dbOpenDynaset)
  With myRec
    If .RecordCount = 0 Then Exit Sub
    .MoveLast
    .FindFirst "ParmID=" & Me.ID
    If Not .NoMatch Then
      .Edit
      !StatisticLabelID = GetLabelID(StatLabel, Me.Region.DB)
      If IsNumeric(Min) Then !Min = Min
      If IsNumeric(Max) Then !Max = Max
      If IsNumeric(ConvFlag) Then !UnitID = ConvFlag
      .Update
    End If
  End With
End Sub

Private Sub Class_Initialize()
  pAbbrev = "XX"
  pName = "Not Assigned"
End Sub

