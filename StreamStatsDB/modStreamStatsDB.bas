Attribute VB_Name = "modStreamStatsDB"
Option Explicit
Global myMsgBox As ATCoMessage
Global IPC As ATCoIPC
Global ExePath As String
Global SSDB As nssDatabase
'StationFields contains items for the Station table
'It is made up of 'StatisticLabelCode' values 1 - 14 and 24 ('DISTRICT') on the StatLabel table
Global StationFields(1 To 15) As String
Global StatFields(1 To 6) As String
Global DataFields(1 To 8) As String
Global ROIImportRegnIDs() As Long
Global ROIImportRegnNames() As String
Global UserInfoOK As Boolean
Global TransID As String
Global IsBasin As Boolean, IsCounty As Boolean, IsMCD As Boolean, _
       IsHUC As Boolean, IsState As Boolean, ImportedNewData As Boolean

Private Sub Main()
  Dim StepName As String
  Dim RunningVB As Boolean
  Dim ExeName As String 'name of executable
  Dim s As String * 80
  Dim hdle&, binpos&, i&
  Dim filename As String

  On Error GoTo MiscError
  
  StepName = "GetModuleHandle"
  hdle = GetModuleHandle("StreamStatsDB")
  StepName = "GetModuleFileName"
  i = GetModuleFileName(hdle, s, 80)
  StepName = "ExeName = UCase"
  ExeName = UCase(Left(s, InStr(s, Chr(0)) - 1))
  If InStr(ExeName, "VB6.EXE") Then
    RunningVB = True
    ExeName = UCase("c:\VBExperimental\StreamStatsDB\data")
  Else
    RunningVB = False
  End If
  ' reset ExePath for particular machine
  binpos = InStr(ExeName, "\BIN")
  If binpos < 1 Then binpos = InStrRev(ExeName, "\")
  If binpos < 1 Then
    ExePath = CurDir
  Else
    ExePath = Left(ExeName, binpos)
  End If
  If Right(ExePath, 1) <> "\" Then ExePath = ExePath & "\"
  
  'Open Stream Stats Database
  
  StepName = "Set SSDB = New nssDatabase"
  Set SSDB = New nssDatabase
  
  StepName = "GetDatabaseFilename"
  filename = GetDatabaseFilename
  
  StepName = "SSDB.Filename = " & filename
  SSDB.filename = filename
  
  StepName = "Set IPC = New ATCoIPC"
  Set IPC = New ATCoIPC
  
  StepName = "frmStreamStatsDB.Show"
  frmStreamStatsDB.Show
  Exit Sub

MiscError:
  Select Case MsgBox(Err.Description & vbCr & "At: " & StepName, vbAbortRetryIgnore, "StreamStatsDB Main")
    Case vbAbort: End
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
  End Select

CancelError:
  On Error Resume Next
  Unload frmCDLG
End Sub

Public Function GetDatabaseFilename(Optional aVerify As Boolean = False) As String
  Dim ff As ATCoFindFile
  Set ff = New ATCoFindFile
  ff.SetRegistryInfo "StreamStatsDB", "files", "StreamStatsDB.mdb"
  ff.SetDialogProperties "Please locate StreamStats database", ExePath & "StreamStatsDB.mdb"
  GetDatabaseFilename = ff.GetName(aVerify)
  
'  Dim fileTitle As String
'  fileTitle = GetSetting("StreamStatsDB", "Defaults", "nffDatabase")
'  If Len(fileTitle) = 0 Then GoTo NoFile:
'  If Len(Dir(fileTitle)) = 0 Then
'NoFile:
'    fileTitle = "StreamStatsDB.mdb"
'    StepName = "With frmCDLG.CDLG"
'    With frmCDLG.CDLG
'      .DialogTitle = "Select the Stream Stats database"
'      .Filename = fileTitle
'      .Filter = "(*.mdb)|*.mdb"
'      .filterIndex = 1
'      On Error GoTo CancelError
'      .CancelError = True
'
'      StepName = "ShowOpen"
'      .ShowOpen
'      On Error GoTo MiscError
'      fileTitle = .Filename
'    End With
'    StepName = "Unload frmCDLG"
'    Unload frmCDLG
'
'    StepName = "SaveSetting"
'    SaveSetting "StreamStatsDB", "Defaults", "nffDatabase", fileTitle
'  End If

End Function

Public Function GetStatTypeCode(StatName As String)
  Dim i&
  For i = 1 To SSDB.StatisticTypes.Count
    If StatName = SSDB.StatisticTypes(i).Name Then
      GetStatTypeCode = SSDB.StatisticTypes(i).code
      Exit Function
    End If
  Next i
End Function

Public Function Decimal2DMS(DecimalDeg As String) As String
  Dim i&
  Dim str$, ConvertStr$
  
  i = InStr(1, DecimalDeg, ".")
  If i > 0 Then
    str = CStr(Left(DecimalDeg, i - 1))
    While Len(str) < 2
      str = "0" & str
    Wend
    ConvertStr = str
    DecimalDeg = Mid(DecimalDeg, i) * 60
    i = InStr(1, DecimalDeg, ".")
    If i > 0 Then
      str = CStr(Left(DecimalDeg, i - 1))
      While Len(str) < 2
        str = "0" & str
      Wend
      ConvertStr = ConvertStr & str
      DecimalDeg = Round(Mid(DecimalDeg, i) * 60, 0)
      While Len(DecimalDeg) < 2
        DecimalDeg = "0" & DecimalDeg
      Wend
      ConvertStr = ConvertStr & DecimalDeg
    Else
      While Len(DecimalDeg) < 2
        DecimalDeg = "0" & DecimalDeg
      Wend
      ConvertStr = ConvertStr & DecimalDeg & "00"
    End If
  Else
    While Len(DecimalDeg) < 2
      DecimalDeg = "0" & DecimalDeg
    Wend
    ConvertStr = DecimalDeg & "0000"
  End If

  Decimal2DMS = ConvertStr

End Function

Public Function DMS2Decimal(DMS As String) As Single
  Dim lenValue&
  Dim degStr$, minStr$, secStr$
  Dim Degrees As Single
  Dim minsec As Single
  
  On Error GoTo BadLatLong
  lenValue = Len(DMS)
  degStr = Mid(DMS, 1, lenValue - 4)
  minStr = Mid(DMS, lenValue - 3, 2)
  secStr = Mid(DMS, lenValue - 1, 2)
  
  Degrees = CSng(degStr)
  minsec = CSng(secStr)
  minsec = minsec / 60 'Convert seconds to minutes
  minsec = minsec + CSng(minStr) 'Add minutes
  minsec = minsec / 60  'Convert to degrees
  If Degrees > 0 Then
    Degrees = Degrees + minsec
  Else
    Degrees = Degrees - minsec
  End If

BadLatLong:
  
  DMS2Decimal = Degrees

End Function

Public Function GetLabelID(StatLabel As String, DB As nssDatabase) As Long
  Dim myRec As Recordset

  Set myRec = DB.DB.OpenRecordset("STATLABEL", dbOpenSnapshot)
  With myRec
    .FindFirst "StatLabel='" & StatLabel & "'"
    If .NoMatch Then .FindFirst "StatisticLabel='" & StatLabel & "'"
    If Not .NoMatch Then GetLabelID = .Fields("StatisticLabelID")
  End With
End Function
