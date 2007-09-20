VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nssComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private pDepVarID As Long
Private pParmID As Long
Private pExpID As Long
Private pDB As nssDatabase
Private pParameter As nssParameter
Private pExponentParameter As nssParameter
Private pDepVar As nssDepVar
Private pBaseMod As Double
Private pBaseCoeff As Double
Private pBaseExp As Double
Private pExpMod As Double
Private pExpExp As Double
Private pIsNew As Boolean

Public Property Get DepVarID() As Long
  DepVarID = pDepVarID
End Property
Public Property Let DepVarID(ByVal newValue As Long)
  pDepVarID = newValue
End Property

Public Property Get ParmID() As Long
  ParmID = pParmID
End Property
Public Property Let ParmID(ByVal newValue As Long)
  pParmID = newValue
End Property

Public Property Get ExpID() As Long
  ExpID = pExpID
End Property
Public Property Let ExpID(ByVal newValue As Long)
  pExpID = newValue
End Property

Public Property Set DB(newValue As nssDatabase)
  Set pDB = newValue
End Property

Public Property Get BaseMod() As Double
  BaseMod = pBaseMod
End Property
Public Property Let BaseMod(ByVal newValue As Double)
  pBaseMod = newValue
End Property

Public Property Get BaseCoeff() As Double
  BaseCoeff = pBaseCoeff
End Property
Public Property Let BaseCoeff(ByVal newValue As Double)
  pBaseCoeff = newValue
End Property

Public Property Get BaseExp() As Double
  BaseExp = pBaseExp
End Property
Public Property Let BaseExp(ByVal newValue As Double)
  pBaseExp = newValue
End Property

Public Property Get ExpMod() As Double
  ExpMod = pExpMod
End Property
Public Property Let ExpMod(ByVal newValue As Double)
  pExpMod = newValue
End Property

Public Property Get ExpExp() As Double
  ExpExp = pExpExp
End Property
Public Property Let ExpExp(ByVal newValue As Double)
  pExpExp = newValue
End Property

Public Property Get IsNew() As Boolean
  IsNew = pIsNew
End Property
Public Property Let IsNew(ByVal newValue As Boolean)
  pIsNew = newValue
End Property

Public Property Get Parameter() As nssParameter
  Set Parameter = pParameter
End Property
Public Property Set Parameter(newValue As nssParameter)
  Set pParameter = newValue
End Property

Public Property Get ExponentParameter() As nssParameter
  Set ExponentParameter = pExponentParameter
End Property
Public Property Set ExponentParameter(newValue As nssParameter)
  Set pExponentParameter = newValue
End Property

Public Property Get DepVar() As nssDepVar
  Set DepVar = pDepVar
End Property
Public Property Set DepVar(newValue As nssDepVar)
  Set pDepVar = newValue
End Property

Public Sub Add(ByVal Region As nssRegion, ByVal DepVarID As Long, ByVal ParmID As Long, _
    ByVal BaseMod As String, ByVal BaseCoeff As String, ByVal BaseExp As String, _
    ByVal ExpID As Long, ByVal ExpMod As String, ByVal ExpExp As String)
  Dim myRec As Recordset
  
  Set myRec = Region.DB.DB.OpenRecordset("Components", dbOpenDynaset)
  With myRec
'    'check to see if parameter with this name already exists for this region
'    .FindFirst "DepVarID=" & DepVarID & " AND ParmID=" & ParmID & " AND ExpIndex=" & ExpID
'    If Not .NoMatch Then  'Erase if a legacy is found
'      Me.Delete
'    End If
    .AddNew
    .Fields(0) = DepVarID
    !ParmID = ParmID
    If IsNumeric(BaseMod) Then !BaseMod = BaseMod
    If IsNumeric(BaseCoeff) Then !BaseCoeff = BaseCoeff
    If IsNumeric(BaseExp) Then !BaseExp = BaseExp
    If IsNumeric(ExpID) Then !ExpIndex = ExpID
    If IsNumeric(ExpMod) Then !ExpMod = ExpMod
    If IsNumeric(ExpExp) Then !ExpExp = ExpExp
    .Update
  End With
End Sub

Public Sub Delete()
  Dim myRec As Recordset
  
  Set myRec = pDB.DB.OpenRecordset("Components", dbOpenDynaset)
  With myRec
    .FindFirst "DepVarID=" & Me.DepVarID & " AND ParmID=" & Me.ParmID & _
              " AND ExpIndex=" & Me.ExpID
    If Not .NoMatch Then
      .Delete
    End If
  End With
End Sub


