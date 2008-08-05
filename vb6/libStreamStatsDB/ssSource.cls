VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ssSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pDB As nssDatabase
Private pID As String
Private pName As String
Private pURL As String

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

Public Property Get URL() As String
  URL = pURL
End Property
Public Property Let URL(ByVal newValue As String)
  pURL = newValue
End Property

Public Sub Add(ByVal Citation As String, ByVal CitationURL As String)
  Dim myRec As Recordset
  Dim i$, sql$
  Dim j&
  
  sql = "SELECT * FROM DATASOURCE ORDER BY DataSourceID"
  Set myRec = pDB.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    If .RecordCount > 0 Then
      While Not .EOF
        If Citation = myRec("Citation") Then
          Citation = "not added"
          Exit Sub
        End If
        .MoveNext
      Wend
      i = CStr(.RecordCount)
      .FindLast "DataSourceID='" & i & "'"
      While Not .NoMatch
        j = j + 1
        i = CStr(.RecordCount + j)
        .FindLast "DataSourceID='" & i & "'"
      Wend
    Else
      i = 1
    End If
    .AddNew
    !DataSourceID = i
    !Citation = Citation
    !CitationURL = CitationURL
    .Update
  End With
End Sub

Public Sub Edit(ByVal Citation As String, ByVal CitationURL As String)
  Dim myRec As Recordset
  
  Set myRec = pDB.DB.OpenRecordset("DATASOURCE", dbOpenDynaset)
  With myRec
    .FindFirst "DataSourceID='" & Me.ID & "'"
    If Not .NoMatch Then
      .Edit
      !Citation = Citation
      !CitationURL = CitationURL
      .Update
    End If
  End With
End Sub

