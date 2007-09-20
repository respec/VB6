VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "userParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private pParm As nssParameter
Private pValue As Double

Public Property Get Parameter() As nssParameter
  Set Parameter = pParm
End Property
Public Property Set Parameter(ByVal newValue As nssParameter)
  Set pParm = newValue
End Property

Public Function GetValue(WantMetric As Boolean) As Double
  GetValue = pParm.Units.Convert(pValue, pParm.Region.State.Metric, WantMetric)
End Function
Public Sub SetValue(ByVal newValue As Double, IsMetric As Boolean)
  pValue = pParm.Units.Convert(newValue, IsMetric, pParm.Region.State.Metric)
End Sub

Private Sub Class_Initialize()
  pValue = -999
End Sub
