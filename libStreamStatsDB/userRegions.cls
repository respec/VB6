VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "userRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Private pRegion As nssRegion
Private pUserParms As FastCollection 'of userParameter
Private pDischarges As FastCollection 'of individual region's discharges estimates

Public Property Get Region() As nssRegion
  Set Region = pRegion
End Property
Public Property Set Region(ByVal newValue As nssRegion)
  Set pRegion = newValue
  PopulateUserParms
End Property

Public Property Get UserParms() As FastCollection
  ' Region must be set before this will work
  Set UserParms = pUserParms
End Property
Public Property Set UserParms(newValue As FastCollection)
  Set pUserParms = Nothing
  Set pUserParms = newValue
End Property

Public Property Get Discharges() As FastCollection
  Set Discharges = pDischarges
End Property
Public Property Set Discharges(newValue As FastCollection)
  Set pDischarges = newValue
End Property

'Returns the userParameter value of a parameter given its ID
Public Function ParmIDVal(ParmID As Long, Metric As Boolean) As Double
  Dim vParm As Variant
  For Each vParm In UserParms
    If vParm.Parameter.ID = ParmID Then
      ParmIDVal = vParm.GetValue(Metric)
      Exit Function
    End If
  Next
End Function

Public Function CopyUserParms() As FastCollection
  Dim oldParm As Variant
  Dim newParm As userParameter
  Dim retval As FastCollection
  Set retval = New FastCollection
  For Each oldParm In UserParms
    Set newParm = New userParameter
    Set newParm.Parameter = oldParm.Parameter
    newParm.SetValue oldParm.GetValue(pRegion.State.Metric), pRegion.State.Metric
    retval.Add newParm, oldParm.Parameter.NSSName
    Set newParm = Nothing
  Next
  Set CopyUserParms = retval
  Set retval = Nothing
End Function

Private Sub PopulateUserParms()
  Dim newParm As userParameter
  Dim baseParm As Variant 'nssParameter
  
  Set pUserParms = Nothing
  Set pUserParms = New FastCollection
  For Each baseParm In pRegion.Parameters
    Set newParm = New userParameter
    Set newParm.Parameter = baseParm
    pUserParms.Add newParm, baseParm.NSSName 'baseParm.Abbrev
    Set newParm = Nothing
  Next
End Sub

Private Sub Class_Initialize()
  Set pUserParms = New FastCollection
  Set pDischarges = New FastCollection
End Sub
