VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nssUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private pID As Long
Private pMetricLabel As String
Private pEnglishLabel As String
Private pMetricAbbrev As String
Private pEnglishAbbrev As String
Private pFactor As Double 'English * pFactor = metric

Public Property Get ID() As Long
  ID = pID
End Property
Public Property Let ID(ByVal newValue As Long)
  pID = newValue
End Property

Public Function Label(Metric As Boolean) As String
  If Metric Then
    Label = pMetricLabel
  Else
    Label = pEnglishLabel
  End If
End Function

Public Function Abbrev(Metric As Boolean) As String
  If Metric Then
    Abbrev = pMetricAbbrev
  Else
    Abbrev = pEnglishAbbrev
  End If
End Function

Public Property Get MetricLabel() As String
  MetricLabel = pMetricLabel
End Property
Public Property Let MetricLabel(ByVal newValue As String)
  If newValue = "-" Then
    pMetricLabel = ""
  Else
    pMetricLabel = newValue
  End If
End Property

Public Property Get EnglishLabel() As String
  EnglishLabel = pEnglishLabel
End Property
Public Property Let EnglishLabel(ByVal newValue As String)
  If newValue = "-" Then
    pEnglishLabel = ""
  Else
    pEnglishLabel = newValue
  End If
End Property

Public Property Get EnglishAbbrev() As String
  EnglishAbbrev = pEnglishAbbrev
End Property
Public Property Let EnglishAbbrev(ByVal newValue As String)
  If newValue = "-" Then
    pEnglishAbbrev = ""
  Else
    pEnglishAbbrev = newValue
  End If
End Property

Public Property Get MetricAbbrev() As String
  MetricAbbrev = pMetricAbbrev
End Property
Public Property Let MetricAbbrev(ByVal newValue As String)
  If newValue = "-" Then
    pMetricAbbrev = ""
  Else
    pMetricAbbrev = newValue
  End If
End Property

Public Property Get Factor() As Double
  Factor = pFactor
End Property
Public Property Let Factor(ByVal newValue As Double)
  pFactor = newValue
End Property

Public Function Convert(Value As Double, IsMetric As Boolean, WantMetric As Boolean) As Double
  If IsMetric = WantMetric Or Abs(Value + 999) < 0.1 Then
    Convert = Value
  ElseIf WantMetric Then
    If pMetricLabel = "deg C" Then
      Convert = (Value - 32) * pFactor
    Else
      Convert = Value * pFactor
    End If
  Else
    If pMetricLabel = "deg C" Then
      Convert = Value / pFactor + 32
    Else
      Convert = Value / pFactor
    End If
  End If
End Function

Private Sub Class_Initialize()
  pID = -1
  pMetricLabel = ""
  pEnglishLabel = ""
  pMetricAbbrev = ""
  pEnglishAbbrev = ""
  pFactor = 1
End Sub
