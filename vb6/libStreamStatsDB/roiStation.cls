VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "roiStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private pDB As nssDatabase
Private pID As Long
Private pLat As Double
Private pLng As Double
Private pSD As Double
Private pCF As Double
Private pDA As Double
Private pPeak2 As Double
Private pPeak5 As Double
Private pPeak10 As Double
Private pPeak25 As Double
Private pPeak50 As Double
Private pPeak100 As Double
Private pPeak200 As Double
Private pPeak500 As Double
Private pEqYears As Double
Private pSlope As Double
Private pShape As Double
Private pUnits As nssUnits

Public Property Set DB(newValue As nssDatabase)
  Set pDB = newValue
End Property

Public Property Get ID() As Long
  ID = pID
End Property
Public Property Let ID(ByVal newValue As Long)
  pID = newValue
End Property

Public Property Get Lat() As Double
  Lat = pLat
End Property

Public Property Get Lng() As Double
  Lng = pLng
End Property

Public Property Get SD() As Double
  SD = pSD
End Property

Public Property Get CF() As Double
  CF = pCF
End Property

Public Property Get DA() As Double
  DA = pDA
End Property

Public Property Get EqYears() As Double
  EqYears = pEqYears
End Property

Public Property Get Slope() As Double
  Slope = pSlope
End Property

Public Property Get Shape() As Double
  Shape = pShape
End Property

Public Property Get Peak2() As Double
  Peak2 = pPeak2
End Property

Public Property Get Peak5() As Double
  Peak5 = pPeak5
End Property

Public Property Get Peak10() As Double
  Peak10 = pPeak10
End Property

Public Property Get Peak25() As Double
  Peak25 = pPeak25
End Property

Public Property Get Peak50() As Double
  Peak50 = pPeak50
End Property

Public Property Get Peak100() As Double
  Peak100 = pPeak100
End Property

Public Property Get Peak200() As Double
  Peak200 = pPeak200
End Property

Public Property Get Peak500() As Double
  Peak500 = pPeak500
End Property

Public Property Get Units() As nssUnits
  Set Units = pUnits
End Property
Public Property Set Units(ByVal newValue As nssUnits)
  Set pUnits = newValue
End Property

Private Sub Class_Initialize()
  pID = -999
End Sub


