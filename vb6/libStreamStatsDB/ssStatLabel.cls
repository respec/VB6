VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ssStatLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pDB As nssDatabase
Private pID As Long
Private pCode As String
Private pTypeCode As String
Private pTypeName As String
Private pName As String
Private pUnits As String
Private pIsNew As Boolean
Private pDefinition As String
Private pAlias As String

Public Property Get DB() As nssDatabase
  Set DB = pDB
End Property
Public Property Set DB(newValue As nssDatabase)
  Set pDB = newValue
End Property

Public Property Get Id() As Long
  Id = pID
End Property
Public Property Let Id(ByVal newValue As Long)
  pID = newValue
End Property

Public Property Get TypeCode() As String
  TypeCode = pTypeCode
End Property
Public Property Let TypeCode(ByVal newValue As String)
  pTypeCode = newValue
End Property

Public Property Get TypeName() As String
  TypeName = pTypeName
End Property
Public Property Let TypeName(ByVal newValue As String)
  pTypeName = newValue
End Property

Public Property Get code() As String
  code = pCode
End Property
Public Property Let code(ByVal newValue As String)
  pCode = newValue
End Property

Public Property Get Name() As String
  Name = pName
End Property
Public Property Let Name(ByVal newValue As String)
  pName = newValue
End Property

Public Property Get Units() As String
  Units = pUnits
End Property
Public Property Let Units(ByVal newValue As String)
  pUnits = newValue
End Property

Public Property Get IsNew() As Boolean
  IsNew = pIsNew
End Property
Public Property Let IsNew(ByVal newValue As Boolean)
  pIsNew = newValue
End Property

Public Property Get Definition() As String
  Definition = pDefinition
End Property
Public Property Let Definition(ByVal newValue As String)
  pDefinition = newValue
End Property

Public Property Get Alias() As String
  Alias = pAlias
End Property
Public Property Let Alias(ByVal newValue As String)
  pAlias = newValue
End Property

Public Sub Add(ByRef NewValues() As String, StatIndex As Long)
  'NewValues(dim 1) - 0, 1, or 2, with 2 being the modified value
  'NewValues(dim 2) - index in array of stations (always? 1 in StreamStatsDB)
  'NewValues(dim 3) - 1 = Stat Type
  '                   2 = Code (one word name)
  '                   3 = Full Name (must include underscores between words)
  '                   4 = Units
  '                   5 = Full Definition (no underscores needed)
  '                   6 = Alias (alternative to Code)
  Dim myRec As Recordset
  Dim fld&
  Dim sql$
  
  sql = "SELECT STATLABEL.* FROM STATLABEL ORDER BY StatisticLabelID;"
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    If .RecordCount > 0 Then 'check to see if Stat with this Abbrev or Name already exists
      .FindFirst "StatLabel='" & NewValues(2, StatIndex, 2) & "'"
      If Not .NoMatch Then  'Stat with this Abbrev already exists
        ssMessageBox "There is already a Statistic with the abbreviation " & _
            NewValues(2, StatIndex, 2) & "." & vbCrLf & "No new Statistic was created.", _
            vbCritical, "Abbreviation already exists"
        Exit Sub
      End If
      .FindFirst "StatisticLabel='" & NewValues(2, StatIndex, 3) & "'"
      If Not .NoMatch Then  'Stat with this Name already exists
        ssMessageBox "There is already a Statistic with the name " & _
            NewValues(2, StatIndex, 3) & "." & vbCrLf & _
            "No new Statistic was created.", _
            vbCritical, "Name already exists"
        Exit Sub
      End If
    End If
    .MoveLast
    fld = !StatisticLabelID + 1
    .AddNew
    .Fields("StatisticLabelID") = fld
    .Fields("StatisticTypeID") = GetStatTypeID(Trim(NewValues(2, StatIndex, 1)))
    .Fields("StatisticTypeCode") = GetStatTypeCode(Trim(NewValues(2, StatIndex, 1)))
    .Fields("StatLabel") = Trim(NewValues(2, StatIndex, 2))
    'be sure to use underscores for blanks in Full Name
    .Fields("StatisticLabel") = Trim(ReplaceString(NewValues(2, StatIndex, 3), " ", "_"))
    'convert unit string to numeric
    .Fields("UnitID") = GetUnitCode(Trim(NewValues(2, StatIndex, 4)))
    If UBound(NewValues, 3) < 5 Then 'just use StatisticLabel field, no underscores
      .Fields("Definition") = Trim(ReplaceString(NewValues(2, StatIndex, 3), "_", " "))
    Else
      .Fields("Definition") = Trim(NewValues(2, StatIndex, 5))
    End If
    If UBound(NewValues, 3) < 6 Then 'just use StatLabel field
      .Fields("Alias") = .Fields("StatLabel")
    Else
      .Fields("Alias") = Trim(NewValues(2, StatIndex, 6))
    End If
    .Update
  End With
End Sub

Public Sub Delete()
  Dim myRec As Recordset
  
  Set myRec = pDB.DB.OpenRecordset("STATLABEL", dbOpenDynaset)
  With myRec
    .FindFirst "StatisticLabelID=" & Me.Id
    If Not .NoMatch Then
      .Delete
    End If
  End With
End Sub

Public Sub Edit(ByRef NewValues() As String, StatIndex As Long)
  'NewValues(dim 1) - 0, 1, or 2, with 2 being the modified value
  'NewValues(dim 2) - index in array of stations (always? 1 in StreamStatsDB)
  'NewValues(dim 3) - 1 = Stat Type
  '                   2 = Code (one word name)
  '                   3 = Full Name (must include underscores between words)
  '                   4 = Units
  '                   5 = Full Definition (no underscores needed)
  '                   6 = Alias (alternative to Code)
  Dim myRec As Recordset
  Dim fld&
  
  Set myRec = pDB.DB.OpenRecordset("STATLABEL", dbOpenDynaset)
  With myRec
    .FindFirst "StatisticLabelID=" & Me.Id
    If Not .NoMatch Then
      .Edit
      .Fields("StatisticTypeID") = GetStatTypeID(Trim(NewValues(2, StatIndex, 1)))
      .Fields("StatisticTypeCode") = GetStatTypeCode(Trim(NewValues(2, StatIndex, 1)))
      .Fields("StatLabel") = Trim(NewValues(2, StatIndex, 2))
      'be sure to use underscores for blanks in Full Name
      .Fields("StatisticLabel") = Trim(ReplaceString(NewValues(2, StatIndex, 3), " ", "_"))
      'convert unit string to numeric
      .Fields("UnitID") = GetUnitCode(Trim(NewValues(2, StatIndex, 4)))
      If UBound(NewValues, 3) < 5 Then 'just use StatisticLabel field, no underscores
        .Fields("Definition") = Trim(ReplaceString(NewValues(2, StatIndex, 3), "_", " "))
      Else
        .Fields("Definition") = Trim(NewValues(2, StatIndex, 5))
      End If
      If UBound(NewValues, 3) < 6 Then 'just use StatLabel field
        .Fields("Alias") = .Fields("StatLabel")
      Else
        .Fields("Alias") = Trim(NewValues(2, StatIndex, 6))
      End If
      .Update
    Else
      ssMessageBox "This Statistic is not found in the database." & _
          vbCrLf & "Statistic not found."
    End If
  End With
End Sub

Private Function GetStatTypeCode(TypeName As String) As String
  Dim myRec As Recordset
  
  Set myRec = pDB.DB.OpenRecordset("STATTYPE", dbOpenSnapshot)
  With myRec
    .FindFirst "StatisticType='" & TypeName & "'"
    If Not .NoMatch Then GetStatTypeCode = .Fields("StatisticTypeCode")
  End With
End Function

Private Function GetStatTypeID(TypeName As String) As Integer
  Dim myRec As Recordset
  
  Set myRec = pDB.DB.OpenRecordset("STATTYPE", dbOpenSnapshot)
  With myRec
    .FindFirst "StatisticType='" & TypeName & "'"
    If Not .NoMatch Then GetStatTypeID = .Fields("StatisticTypeID")
  End With
End Function

Private Function GetUnitCode(UnitName As String) As Integer
  'uses English unit labels
  Dim myRec As Recordset
  
  Set myRec = pDB.DB.OpenRecordset("UNITS", dbOpenSnapshot)
  With myRec
    .FindFirst "English='" & UnitName & "'"
    If Not .NoMatch Then GetUnitCode = .Fields("UnitID")
  End With
End Function
