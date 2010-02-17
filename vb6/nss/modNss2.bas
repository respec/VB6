Attribute VB_Name = "modNss2"
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Global Project As nssProject
Global Const HYDRO_SIZE = 45  '# elements in dimensionless hydrograph
Global disch_ratio(0 To HYDRO_SIZE - 1) As Single

Global DefaultSaveFile As String

Public Sub Main()
  Dim progress As String
  Dim dbPath As String
  Dim ff As New ATCoFindFile

  'App.HelpFile = App.path & "\nss.chm"
  ff.SetDialogProperties "Please locate NSS help file 'nss.chm'", App.path & "\nss.chm"
  ff.SetRegistryInfo "NSS", "files", "nss.chm"
  App.HelpFile = ff.GetName
  
FindDB:
  On Error GoTo NoDB
  'Open Stream Stats Database
  'dbPath = GetSetting("StreamStatsDB", "Defaults", "nssDatabase", App.path & "\StreamStatsDB.mdb")
  ff.SetDialogProperties "Please locate NSS or StreamStats database version 5" ', "NSSv4.mdb"
  ff.SetRegistryInfo "StreamStatsDB", "Defaults", "nssDatabaseV4"
  dbPath = ff.GetName
  
  If Not FileExists(dbPath) Then GoTo NoDB

  On Error GoTo ShowProgress
  
  progress = "dbPath = " & dbPath
  DefaultSaveFile = App.path & "\current.nss"
  progress = progress & vbCr & "Creating Project"
  Set Project = New nssProject
  Project.HelpFile = App.HelpFile
  progress = progress & vbCr & "Loading NSS database " & dbPath
  
  On Error GoTo NoDB
  Project.LoadNSSdatabase dbPath
  
  Project.FileName = DefaultSaveFile
  If Len(Dir(Project.FileName)) > 0 Then
    progress = progress & vbCr & "Loading " & Project.FileName
    Project.XML = WholeFileString(Project.FileName)
  Else
    progress = progress & vbCr & "Not Loading " & Project.FileName & " (not found)"
    Project.XML = "<NSSproject name=""Unnamed"" username="""" state=""01"" metric=""False"" currentrural=""0"" currenturban=""0""></NSSproject>"
  End If
  
  progress = progress & vbCr & "Showing frmStart"
  On Error GoTo ShowProgress
  frmStart.Show
  
  Exit Sub

NoDB:
  If MsgBox("Could not open database or project" & vbCr & vbCr _
        & progress & vbCr & vbCr _
        & Err.Description & vbCr & vbCr _
        & "Search for current database?", vbOKCancel, "NSS Database Problem") = vbOK Then
    SaveSetting "StreamStatsDB", "Defaults", "nssDatabaseV4", "NSSv4.mdb"
    GoTo FindDB
  Else
    End
  End If

ShowProgress:
  MsgBox progress & vbCr & Err.Description, vbExclamation, "Error starting NSS"
End Sub

Public Sub InitDischRatio()
'   dimensionless hydrograph data
    disch_ratio(0) = 0.12
    disch_ratio(1) = 0.16
    disch_ratio(2) = 0.21
    disch_ratio(3) = 0.26
    disch_ratio(4) = 0.33
    disch_ratio(5) = 0.4
    disch_ratio(6) = 0.49
    disch_ratio(7) = 0.58
    disch_ratio(8) = 0.67
    disch_ratio(9) = 0.76
    disch_ratio(10) = 0.84
    disch_ratio(11) = 0.9
    disch_ratio(12) = 0.95
    disch_ratio(13) = 0.98
    disch_ratio(14) = 1#
    disch_ratio(15) = 0.99
    disch_ratio(16) = 0.96
    disch_ratio(17) = 0.92
    disch_ratio(18) = 0.86
    disch_ratio(19) = 0.8
    disch_ratio(20) = 0.74
    disch_ratio(21) = 0.68
    disch_ratio(22) = 0.62
    disch_ratio(23) = 0.56
    disch_ratio(24) = 0.51
    disch_ratio(25) = 0.47
    disch_ratio(26) = 0.43
    disch_ratio(27) = 0.39
    disch_ratio(28) = 0.36
    disch_ratio(29) = 0.33
    disch_ratio(30) = 0.3
    disch_ratio(31) = 0.28
    disch_ratio(32) = 0.26
    disch_ratio(33) = 0.24
    disch_ratio(34) = 0.22
    disch_ratio(35) = 0.2
    disch_ratio(36) = 0.19
    disch_ratio(37) = 0.17
    disch_ratio(38) = 0.16
    disch_ratio(39) = 0.15
    disch_ratio(40) = 0.14
    disch_ratio(41) = 0.13
    disch_ratio(42) = 0.12
    disch_ratio(43) = 0.11
    disch_ratio(44) = 0.1
End Sub

Public Sub AllIntervals(nAllInt As Long, allint() As Single)
  Dim IntervalIndex As Long
  Dim IntervalShiftIndex As Long
  Dim IntervalValue As Single
  Dim Scenario As Variant 'nssScenario
  Dim vReturn As Variant
  
  ReDim allint(50)
  
  nAllInt = 0
  
  For Each Scenario In Project.RuralScenarios
    If Not Scenario.lowflow Then GoSub InsertSort
  Next
  For Each Scenario In Project.UrbanScenarios
    If Not Scenario.lowflow Then GoSub InsertSort
  Next
      
  If nAllInt > 0 Then
    ReDim Preserve allint(nAllInt - 1)
  Else
    ReDim allint(0)
  End If
  
  Exit Sub

InsertSort:
  For Each vReturn In Scenario.UserRegions(1).region.DepVars
    If Left(vReturn.Name, 2) = "PK" Then
      IntervalValue = CSng(Mid(vReturn.Name, 3))
    Else
      IntervalValue = CSng(vReturn.Name)
    End If
    IntervalIndex = 0
    While IntervalIndex < nAllInt And allint(IntervalIndex) <= IntervalValue
      If allint(IntervalIndex) = IntervalValue Then Return 'Already have this interval
      IntervalIndex = IntervalIndex + 1
    Wend
    For IntervalShiftIndex = IntervalIndex To nAllInt - 1
      allint(IntervalShiftIndex + 1) = allint(IntervalShiftIndex)
    Next
    allint(IntervalIndex) = IntervalValue
    nAllInt = nAllInt + 1
  Next
  Return

End Sub

'Public Function gausex(exprob!) As Single
'    'GAUSSIAN PROBABILITY FUNCTIONS   W.KIRBY  JUNE 71
'       'GAUSEX=VALUE EXCEEDED WITH PROB EXPROB
'
'    'GAUSCF MODIFIED 740906 WK -- REPLACED ERF FCN REF BY RATIONAL APPRX N
'    'ALSO REMOVED DOUBLE PRECISION FROM GAUSEX AND GAUSAB.
'    '76-05-04 WK -- TRAP UNDERFLOWS IN EXP IN GUASCF AND DY.
'
'    'rev 8/96 by PRH for VB
'    Const c0! = 2.515517
'    Const c1! = 0.802853
'    Const c2! = 0.010328
'    Const d1! = 1.432788
'    Const d2! = 0.189269
'    Const d3! = 0.001308
'    Dim pr!, rtmp!, p!, t!, numerat!, Denom!
'
'    p = exprob
'    If p >= 1# Then
'      'set to minimum
'      rtmp = -10#
'    ElseIf p <= 0# Then
'      'set at maximum
'      rtmp = 10#
'    Else
'      'compute value
'      pr = p
'      If p > 0.5 Then pr = 1# - pr
'      t = (-2# * Log(pr)) ^ 0.5
'      numerat = (c0 + t * (c1 + t * c2))
'      Denom = (1# + t * (d1 + t * (d2 + t * d3)))
'      rtmp = t - numerat / Denom
'      If p > 0.5 Then rtmp = -rtmp
'    End If
'    gausex = rtmp
'End Function
'
