VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ssStatType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private pDB As nssDatabase
Private pStatLabels As FastCollection ' of all available ssStatistics
Private pID As Long
Private pCode As String
Private pName As String

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

Public Property Get Code() As String
  Code = pCode
End Property
Public Property Let Code(ByVal newValue As String)
  pCode = newValue
End Property

Public Property Get Name() As String
  Name = pName
End Property
Public Property Let Name(ByVal newValue As String)
  pName = newValue
End Property

Public Property Get StatLabels() As FastCollection
  ' Code must be set for this state before this will work
  If pStatLabels Is Nothing Then PopulateStatLabels
  Set StatLabels = pStatLabels
End Property
Private Sub PopulateStatLabels()
  Dim newStatLabel As ssStatLabel
  Dim statRec As Recordset
  Dim sql$
  
  Set pStatLabels = New FastCollection
  If pDB Is Nothing Then
    ssMessageBox "Cannot read StatLabels" & vbCr & "DB not set."
  Else
'    sql = "SELECT STATLABEL.StatLabel, STATLABEL.StatisticLabel, STATLABEL.Units, " & _
'        "STATLABEL.StatisticLabelCode FROM STATTYPE INNER JOIN STATLABEL ON " & _
'        "STATTYPE.StatisticTypeCode = STATLABEL.StatisticTypeCode " & _
'        "WHERE STATTYPE.StatisticTypeCode='" & Me.Code & _
'        "' AND StatisticLabelCode > 0 ORDER BY StatisticLabelCode;"
    sql = "SELECT STATLABEL.* FROM STATTYPE INNER JOIN STATLABEL ON " & _
        "STATTYPE.StatisticTypeCode = STATLABEL.StatisticTypeCode " & _
        "WHERE STATTYPE.StatisticTypeCode='" & Me.Code & _
        "' AND StatisticLabelCode > 0 ORDER BY StatisticLabelCode;"
    Set statRec = pDB.DB.OpenRecordset(sql)
    While Not statRec.EOF
      Set newStatLabel = New ssStatLabel
      With newStatLabel
        Set .DB = pDB
        .ID = statRec.Fields("StatisticLabelCode")
        .TypeCode = Me.Code
        .TypeName = Me.Name
        .Code = statRec.Fields("StatLabel")
        .Name = statRec.Fields("StatisticLabel")
        .Units = statRec.Fields("UnitID")
        .IsNew = False
        .Definition = statRec.Fields("Definition")
      End With
      pStatLabels.Add newStatLabel, newStatLabel.Code
      Set newStatLabel = Nothing
      statRec.MoveNext
    Wend
    statRec.Close
  End If
End Sub

Public Sub Add(Code As String, Name As String)
  Dim myRec As Recordset
  Dim flds&
  Dim sql$
  
  sql = "SELECT * FROM STATTYPE ORDER BY StatisticTypeID"
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    If .RecordCount > 0 Then 'check to see if Stat Type with this Code or Name already exists
      .FindFirst "StatisticTypeCode='" & Code & "'"
      If Not .NoMatch Then  'Statistic with this code already exists for this state
        ssMessageBox "There is already a Statistic with the code " & Code & "." & _
            vbCrLf & "No new Statistic was created.", _
            vbCritical, "Code already exists"
        Exit Sub
      End If
      .FindFirst "StatisticType='" & Name & "'"
      If Not .NoMatch Then  'Statistic with this name already exists for this state
        ssMessageBox "There is already a Statistic with the name " & Name & "." & _
            vbCrLf & "No new Statistic was created.", _
            vbCritical, "Name already exists"
        Exit Sub
      End If
    End If
    .MoveLast
    flds = !StatisticTypeID + 1
    .AddNew
    !StatisticTypeID = flds
    !StatisticTypeCode = Code
    !StatisticType = Name
    .Update
  End With
End Sub

Public Sub Edit(Code As String, Name As String)
  Dim myRec As Recordset
  
  Set myRec = pDB.DB.OpenRecordset("STATTYPE", dbOpenDynaset)
  With myRec
    If .RecordCount = 0 Then
      Exit Sub
    Else
      .FindFirst "StatisticTypeID=" & Me.ID
      .Edit
      !StatisticTypeCode = Code
      !StatisticType = Name
      .Update
    End If
  End With
End Sub

Public Sub Delete()
  Dim myRec As Recordset
  
  Set myRec = pDB.DB.OpenRecordset("STATTYPE", dbOpenDynaset)
  With myRec
    .FindFirst "ID=" & Me.ID
    If Not .NoMatch Then
      .Delete
    End If
  End With
End Sub
