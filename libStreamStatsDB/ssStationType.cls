VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ssStationType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pDB As nssDatabase
Private pCode As String
Private pName As String

Public Property Get DB() As nssDatabase
  Set DB = pDB
End Property
Public Property Set DB(newValue As nssDatabase)
  Set pDB = newValue
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

