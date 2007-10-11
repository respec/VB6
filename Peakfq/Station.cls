VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pfqStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pID As String
Private pName As String
Private pActive As Boolean
Private pBegYear As Long
Private pEndYear As Long
Private pSkewOpt As Long  '-1 - Station, 0 - Weighted, 1 - Generalized
Private pUrbanRegPeaks As Boolean
Private pHistoricPeriod As Single
Private pGenSkew As Single
Private pHighSysPeak As Single
Private pHighOutlier As Single
Private pLowHistPeak As Single
Private pLowOutlier As Single
Private pGageBaseDischarge As Single
Private pSESkew As Single
Private pLat As Single
Private pLng As Single
Private pPlotName As String
Private pPlotMade As Boolean
'the following are for storing comments for various specification records
Private pComment As String
Private pCGenSkew As String
Private pCSESkew As String
Private pCBegYear As String
Private pCEndYear As String
Private pCHistoric As String
Private pCSkewOpt As String
Private pCUrban As String
Private pCLowOutlier As String
Private pCHighOutlier As String
Private pCGageBase As String
Private pCLat As String
Private pCLong As String
Private pCPlotName As String

Private SOText(-1 To 1) As String

Public Property Get id() As String
  id = pID
End Property
Public Property Let id(ByVal newValue As String)
  pID = newValue
End Property

Public Property Get Name() As String
  Name = pName
End Property
Public Property Let Name(ByVal newValue As String)
  pName = newValue
End Property

Public Property Get Active() As Boolean
  Active = pActive
End Property
Public Property Let Active(ByVal newValue As Boolean)
  pActive = newValue
End Property

Public Property Get BegYear() As Long
  BegYear = pBegYear
End Property
Public Property Let BegYear(ByVal newValue As Long)
  pBegYear = newValue
End Property

Public Property Get EndYear() As Long
  EndYear = pEndYear
End Property
Public Property Let EndYear(ByVal newValue As Long)
  pEndYear = newValue
End Property

Public Property Get SkewOpt() As Long
  SkewOpt = pSkewOpt
End Property
Public Property Let SkewOpt(ByVal newValue As Long)
  pSkewOpt = newValue
End Property

Public Property Get UrbanRegPeaks() As Boolean
  UrbanRegPeaks = pUrbanRegPeaks
End Property
Public Property Let UrbanRegPeaks(ByVal newValue As Boolean)
  pUrbanRegPeaks = newValue
End Property

Public Property Get HistoricPeriod() As Single
  HistoricPeriod = pHistoricPeriod
End Property
Public Property Let HistoricPeriod(ByVal newValue As Single)
  pHistoricPeriod = newValue
End Property

Public Property Get GenSkew() As Single
  GenSkew = pGenSkew
End Property
Public Property Let GenSkew(ByVal newValue As Single)
  pGenSkew = newValue
End Property

Public Property Get HighSysPeak() As Single
  HighSysPeak = pHighSysPeak
End Property
Public Property Let HighSysPeak(ByVal newValue As Single)
  pHighSysPeak = newValue
End Property

Public Property Get HighOutlier() As Single
  HighOutlier = pHighOutlier
End Property
Public Property Let HighOutlier(ByVal newValue As Single)
  pHighOutlier = newValue
End Property

Public Property Get LowHistPeak() As Single
  LowHistPeak = pLowHistPeak
End Property
Public Property Let LowHistPeak(ByVal newValue As Single)
  pLowHistPeak = newValue
End Property

Public Property Get LowOutlier() As Single
  LowOutlier = pLowOutlier
End Property
Public Property Let LowOutlier(ByVal newValue As Single)
  pLowOutlier = newValue
End Property

Public Property Get GageBaseDischarge() As Single
  GageBaseDischarge = pGageBaseDischarge
End Property
Public Property Let GageBaseDischarge(ByVal newValue As Single)
  pGageBaseDischarge = newValue
End Property

Public Property Get SESkew() As Single
  SESkew = pSESkew
End Property
Public Property Let SESkew(ByVal newValue As Single)
  pSESkew = newValue
End Property

Public Property Get Lat() As Single
  Lat = pLat
End Property
Public Property Let Lat(ByVal newValue As Single)
  pLat = newValue
End Property

Public Property Get Lng() As Single
  Lng = pLng
End Property
Public Property Let Lng(ByVal newValue As Single)
  pLng = newValue
End Property

Public Property Get PlotName() As String
  PlotName = pPlotName
End Property
Public Property Let PlotName(ByVal newValue As String)
  pPlotName = newValue
End Property

Public Property Get PlotMade() As Boolean
  PlotMade = pPlotMade
End Property
Public Property Let PlotMade(ByVal newValue As Boolean)
  pPlotMade = newValue
End Property

Public Property Get Comment() As String
  Comment = pComment
End Property
Public Property Let Comment(ByVal newValue As String)
  pComment = newValue
End Property

Public Property Get CGenSkew() As String
  CGenSkew = pCGenSkew
End Property
Public Property Let CGenSkew(ByVal newValue As String)
  pCGenSkew = newValue
End Property

Public Property Get CSESkew() As String
  CSESkew = pCSESkew
End Property
Public Property Let CSESkew(ByVal newValue As String)
  pCSESkew = newValue
End Property

Public Property Get CBegYear() As String
  CBegYear = pCBegYear
End Property
Public Property Let CBegYear(ByVal newValue As String)
  pCBegYear = newValue
End Property

Public Property Get CEndYear() As String
  CEndYear = pCEndYear
End Property
Public Property Let CEndYear(ByVal newValue As String)
  pCEndYear = newValue
End Property

Public Property Get CHistoric() As String
  CHistoric = pCHistoric
End Property
Public Property Let CHistoric(ByVal newValue As String)
  pCHistoric = newValue
End Property

Public Property Get CSkewOpt() As String
  CSkewOpt = pCSkewOpt
End Property
Public Property Let CSkewOpt(ByVal newValue As String)
  pCSkewOpt = newValue
End Property

Public Property Get CUrban() As String
  CUrban = pCUrban
End Property
Public Property Let CUrban(ByVal newValue As String)
  pCUrban = newValue
End Property

Public Property Get CLowOutlier() As String
  CLowOutlier = pCLowOutlier
End Property
Public Property Let CLowOutlier(ByVal newValue As String)
  pCLowOutlier = newValue
End Property

Public Property Get CHighOutlier() As String
  CHighOutlier = pCHighOutlier
End Property
Public Property Let CHighOutlier(ByVal newValue As String)
  pCHighOutlier = newValue
End Property

Public Property Get CGageBase() As String
  CGageBase = pCGageBase
End Property
Public Property Let CGageBase(ByVal newValue As String)
  pCGageBase = newValue
End Property

Public Property Get CLat() As String
  CLat = pCLat
End Property
Public Property Let CLat(ByVal newValue As String)
  pCLat = newValue
End Property

Public Property Get CLong() As String
  CLong = pCLong
End Property
Public Property Let CLong(ByVal newValue As String)
  pCLong = newValue
End Property

Public Property Get CPlotName() As String
  CPlotName = pCPlotName
End Property
Public Property Let CPlotName(ByVal newValue As String)
  pCPlotName = newValue
End Property

Public Function WriteSpecsVerbose() As String

  Dim s As String
  Const pad As String = "     "

  If Len(pComment) > 0 Then
    s = pComment & vbCrLf & "Station " & pID & vbCrLf
  Else
    s = "Station " & pID & vbCrLf
  End If
  If Len(pCBegYear) > 0 Then s = s & pad & pCBegYear & vbCrLf
  If pBegYear > 0 Then s = s & pad & "BegYear " & CStr(pBegYear) & vbCrLf
  If Len(pCEndYear) > 0 Then s = s & pad & pCEndYear & vbCrLf
  If pEndYear > 0 Then s = s & pad & "EndYear " & CStr(pEndYear) & vbCrLf
  If Len(pCHistoric) > 0 Then s = s & pad & pCHistoric & vbCrLf
  If pHistoricPeriod > 0 Then s = s & pad & "HistPeriod " & CStr(pHistoricPeriod) & vbCrLf
  If Len(pCSkewOpt) > 0 Then s = s & pad & pCSkewOpt & vbCrLf
  s = s & pad & "SkewOpt " & SOText(pSkewOpt) & vbCrLf
  If Len(pCGenSkew) > 0 Then s = s & pad & pCGenSkew & vbCrLf
  s = s & pad & "GenSkew " & pGenSkew & vbCrLf
  If Len(pCSESkew) > 0 Then s = s & pad & pCSESkew & vbCrLf
  s = s & pad & "SkewSE " & pSESkew & vbCrLf
  If Len(pCUrban) > 0 Then s = s & pad & pCUrban & vbCrLf
  If pUrbanRegPeaks Then s = s & pad & "Urb/Reg YES" & vbCrLf
  If Len(pCLowOutlier) > 0 Then s = s & pad & pCLowOutlier & vbCrLf
  If pLowOutlier > 0 Then s = s & pad & "LoThresh " & CStr(pLowOutlier) & vbCrLf
  If Len(pCHighOutlier) > 0 Then s = s & pad & pCHighOutlier & vbCrLf
  If pHighOutlier > 0 Then s = s & pad & "HiThresh " & CStr(pHighOutlier) & vbCrLf
  If Len(pCGageBase) > 0 Then s = s & pad & pCGageBase & vbCrLf
  If pGageBaseDischarge > 0 Then s = s & pad & "GageBase " & CStr(pGageBaseDischarge) & vbCrLf
  If Len(pCLat) > 0 Then s = s & pad & pCLat & vbCrLf
  If pLat > 0 Then s = s & pad & "Latitude " & CStr(pLat) & vbCrLf
  If Len(pCLong) > 0 Then s = s & pad & pCLong & vbCrLf
  If pLng > 0 Then s = s & pad & "Longitude " & CStr(pLng) & vbCrLf
  If Len(pCLowOutlier) > 0 Then s = s & pad & pCLowOutlier & vbCrLf
  If Len(pCPlotName) > 0 Then s = s & pad & pCPlotName & vbCrLf
  If Len(pPlotName) > 0 Then s = s & pad & "PlotName " & pPlotName & vbCrLf
  WriteSpecsVerbose = s

End Function

Public Function WriteSpecsNonDefault(defsta As pfqStation) As String

  Dim s As String
  Const pad As String = "     "

  '!!! KLUGE WARNING: Station comments get lost in active project,
  'due to batch program not preserving comments, so comments are
  'extracted from default project
  If Len(defsta.Comment) > 0 Then
    s = defsta.Comment & vbCrLf & "Station " & pID & vbCrLf
  Else
    s = "Station " & pID & vbCrLf
  End If
  If Len(defsta.CBegYear) > 0 Then s = s & pad & defsta.CBegYear & vbCrLf
  If pBegYear <> defsta.BegYear Then s = s & pad & "BegYear " & CStr(pBegYear) & vbCrLf
  If Len(defsta.CEndYear) > 0 Then s = s & pad & defsta.CEndYear & vbCrLf
  If pEndYear <> defsta.EndYear Then s = s & pad & "EndYear " & CStr(pEndYear) & vbCrLf
  If Len(defsta.CHistoric) > 0 Then s = s & pad & defsta.CHistoric & vbCrLf
  If pHistoricPeriod <> defsta.HistoricPeriod Then s = s & pad & "HistPeriod " & CStr(pHistoricPeriod) & vbCrLf
  If Len(defsta.CSkewOpt) > 0 Then s = s & pad & defsta.CSkewOpt & vbCrLf
  If pSkewOpt <> defsta.SkewOpt Then s = s & pad & "SkewOpt " & SOText(pSkewOpt) & vbCrLf
  If Len(defsta.CGenSkew) > 0 Then s = s & pad & defsta.CGenSkew & vbCrLf
  If pGenSkew <> defsta.GenSkew Then s = s & pad & "GenSkew " & pGenSkew & vbCrLf
  If Len(defsta.CSESkew) > 0 Then s = s & pad & defsta.CSESkew & vbCrLf
  If pSESkew <> defsta.SESkew Then s = s & pad & "SkewSE " & pSESkew & vbCrLf
  If Len(defsta.CUrban) > 0 Then s = s & pad & defsta.CUrban & vbCrLf
  If pUrbanRegPeaks <> defsta.UrbanRegPeaks Then s = s & pad & "Urb/Reg YES" & vbCrLf
  If Len(defsta.CLowOutlier) > 0 Then s = s & pad & defsta.CLowOutlier & vbCrLf
  If pLowOutlier <> defsta.LowOutlier Then s = s & pad & "LoThresh " & CStr(pLowOutlier) & vbCrLf
  If Len(defsta.CHighOutlier) > 0 Then s = s & pad & defsta.CHighOutlier & vbCrLf
  If pHighOutlier <> defsta.HighOutlier Then s = s & pad & "HiThresh " & CStr(pHighOutlier) & vbCrLf
  If Len(defsta.CGageBase) > 0 Then s = s & pad & defsta.CGageBase & vbCrLf
  If pGageBaseDischarge <> defsta.GageBaseDischarge Then s = s & pad & "GageBase " & CStr(pGageBaseDischarge) & vbCrLf
  If Len(defsta.CLat) > 0 Then s = s & pad & defsta.CLat & vbCrLf
  If pLat <> defsta.Lat Then s = s & pad & "Latitude " & CStr(pLat) & vbCrLf
  If Len(defsta.CLong) > 0 Then s = s & pad & defsta.CLong & vbCrLf
  If pLng <> defsta.Lng Then s = s & pad & "Longitude " & CStr(pLng) & vbCrLf
  If Len(defsta.CPlotName) > 0 Then s = s & pad & defsta.CPlotName & vbCrLf
  If pPlotName <> defsta.PlotName Then s = s & pad & "PlotName " & pPlotName & vbCrLf
  WriteSpecsNonDefault = s

End Function

Private Sub Class_Initialize()

  pActive = True     'init all stations to be analyzed
  pSkewOpt = 0       'Weighted skew option
  pUrbanRegPeaks = False
  pBegYear = 0
  pEndYear = 0
  pHistoricPeriod = 0#
  pGenSkew = -0.5
  pHighOutlier = 0#
  pLowOutlier = 0#
  pGageBaseDischarge = 0#
  pSESkew = 0.55
  pLat = 0#
  pLng = 0#
  SOText(-1) = "Station"
  SOText(0) = "Weighted"
  SOText(1) = "Generalized"

End Sub
