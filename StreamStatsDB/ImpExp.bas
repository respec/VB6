Attribute VB_Name = "Import"
Option Explicit

Sub NWISImport(ImpFileName As String)
  Dim XLRange As Excel.Range, XLApp As Excel.Application, _
      XLBook As Excel.Workbook
  Dim bumSheets$, stationID$, bumFields$, filename$, stateFIPS$
  Dim staCnt&, header&, lastRow&, firstCol&, fldCnt&, h&, i&, response&
  Dim value As String, staName As String
  Dim stationValues() As String, dataValues() As String, impStates() As String
  Dim OutFile As Integer
  Dim myStation As ssStation
  Dim myStatistic As ssStatistic
  
  On Error GoTo errTrap
  
  IPC.SendMonitorMessage "(OPEN StreamStatsDB)"
  IPC.SendMonitorMessage "(BUTTOFF DETAILS)"
  IPC.SendMonitorMessage "(BUTTON CANCEL)"
  IPC.SendMonitorMessage "(BUTTON PAUSE)"
  IPC.SendMonitorMessage "(MSG1 Importing NWIS file)"
  IPC.SendMonitorMessage "(PROGRESS 0)"
  
  Set XLApp = New Excel.Application
  Set XLBook = Excel.Workbooks.Open(ImpFileName, 0, True)
  
  For h = 1 To ActiveWorkbook.Sheets.Count
    With ActiveWorkbook
      Worksheets(h).Activate
    End With
    With ActiveWorkbook.ActiveSheet
      'Check for the key word "site_no"
      Set XLRange = Range(Cells(1, 1), Cells(20, 10))
      Set XLRange = XLRange.Find("site_no", , , , xlByRows, xlNext, True)
      If XLRange Is Nothing Then
        bumSheets = bumSheets & " " & .Name
        GoTo badSheet
      End If
      'Establish boundaries of import file
      header = XLRange.row
      firstCol = XLRange.Column
      Set XLRange = Cells(XLRange.row, XLRange.Column)
      Set XLRange = XLRange.End(xlDown)
      lastRow = XLRange.row
      'Loop thru stations in spreadsheet
      ReDim impStates(0)
      impStates(0) = .Cells(header + 1, firstCol + 6)
      If Len(impStates(0)) = 1 Then impStates(0) = "0" & impStates(0)
      For staCnt = header + 1 To lastRow
        IPC.SendMonitorMessage "(PROGRESS " & _
            (staCnt - header) * 100 / (lastRow - header) & ")"
        Select Case NextPipeCharacter(IPC.hPipeReadFromProcess("Status"))
          Case "P"
            While NextPipeCharacter(IPC.hPipeReadFromProcess("Status")) <> "R"
              DoEvents
            Wend
          Case "C"
              MsgBox "The NWIS import operation has been interrupted." & _
                  vbCrLf & vbCrLf & "Data has been imported up to station " & _
                  stationID & ".", , "Import interrupted"
            Err.Raise 999
        End Select
        ReDim stationValues(2 To 2, 1 To 1, 1 To UBound(StationFields))
        ReDim dataValues(2 To 2, 1 To 4, 1 To UBound(DataFields))
        'Loop thru column fields in spreadsheet
        For fldCnt = firstCol To firstCol + 10
          value = .Cells(staCnt, fldCnt)
          Select Case fldCnt - firstCol + 1
            Case 1:
              If Len(value) < 8 Or Len(value) > 15 Then
                MsgBox "The site_no (Station ID) on row " & staCnt & _
                    " must be 8-15 digits long." & vbCrLf & _
                    "No data will be imported for this station."
                Exit For
              End If
              stationValues(2, 1, 1) = value
              stationID = value
            Case 2:
              stationValues(2, 1, 2) = value
              staName = value
            Case 3:
              If value <> "" Then
                If value < 17.5 Or (value > 360 And value < 173000) Then
                  bumFields = vbCrLf & stationID & " on row " & staCnt & _
                      " indicates a latitude south of the Virgin Islands."
                ElseIf (value > 72 And value <= 360) Or value > 720000 Then
                  bumFields = vbCrLf & stationID & " on row " & staCnt & _
                      " indicates a latitude north of Alaska."
                End If
                'Convert degrees, minutes, seconds to decimal degrees if necessary
                If value > 360 Then
                  stationValues(2, 1, 7) = DMS2Decimal(value)
                Else
                  stationValues(2, 1, 7) = value
                End If
              End If
            Case 4:
              If value <> "" Then
                If Abs(value) < 64 Or (Abs(value) > 360 And Abs(value) < 640000) Then
                  bumFields = vbCrLf & stationID & " on row " & staCnt & _
                      " indicates a longitude east of the Virgin Islands."
                ElseIf (Abs(value) > 172 And Abs(value) <= 360) Or Abs(value) > 1720000 Then
                  bumFields = vbCrLf & stationID & " on row " & staCnt & _
                      " indicates a longitude west of Alaska."
                End If
                'Convert degrees, minutes, seconds to decimal degrees if necessary
                'always store longitude as negative (prh from kries, 5/2005)
                If Abs(value) > 360 Then
                  stationValues(2, 1, 8) = -Abs(DMS2Decimal(value))
                Else
                  stationValues(2, 1, 8) = -Abs(value)
                End If
              End If
            Case 5:
              dataValues(2, 1, 3) = "DATUM"
              dataValues(2, 1, 4) = value
              dataValues(2, 1, 7) = "Imported from NWIS file"
              dataValues(2, 1, 8) = "http://waterdata.usgs.gov/nwis/si"
            Case 7:  'state code
              dataValues(2, 2, 3) = "DISTRICT"
              If Len(value) = 1 Then value = "0" & value
              dataValues(2, 2, 4) = value
              dataValues(2, 2, 7) = "Imported from NWIS file"
              dataValues(2, 1, 8) = "http://waterdata.usgs.gov/nwis/si"
              stateFIPS = value
              If Len(stateFIPS) = 1 Then stateFIPS = "0" & stateFIPS
              'Keep collection of state codes for all states in import file
              For i = 0 To UBound(impStates)
                If impStates(i) = stateFIPS Then Exit For
              Next i
              If i > UBound(impStates) Then
                ReDim Preserve impStates(i)
                impStates(i) = stateFIPS
              End If
            Case 8:  'county code
              While Len(value) < 3
                value = "0" & value
              Wend
              stationValues(2, 1, 11) = value
            Case 9:  'HUC code
              stationValues(2, 1, 9) = value
            Case 10:
              If Not IsNumeric(value) Then
                If value <> "" Then bumFields = bumFields & vbCrLf & stationID & _
                    " on row " & staCnt & " has a non-numeric value for the drainage area."
              ElseIf value < 0 Then
                bumFields = vbCrLf & stationID & " on row " & staCnt & _
                    " has a negative value for the drainage area."
              End If
              dataValues(2, 3, 3) = "AREA"
              dataValues(2, 3, 4) = value
              dataValues(2, 3, 7) = "Imported from NWIS file"
              dataValues(2, 1, 8) = "http://waterdata.usgs.gov/nwis/si"
            Case 11:
              If Not IsNumeric(value) Then
                If value <> "" Then bumFields = bumFields & vbCrLf & stationID & _
                    " on row " & staCnt & " has a non-numeric value for the " & _
                    "contributing drainage area."
              ElseIf value < 0 Then
                bumFields = vbCrLf & stationID & " on row " & staCnt & _
                    " has a negative value for the contributing drainage area."
              End If
              dataValues(2, 4, 3) = "CONTDA"
              dataValues(2, 4, 4) = value
              dataValues(2, 4, 7) = "Imported from NWIS file"
              dataValues(2, 1, 8) = "http://waterdata.usgs.gov/nwis/si"
          End Select
        Next fldCnt
        If SSDB.States.IndexFromKey(stateFIPS) > 0 Then
          Set SSDB.state = SSDB.States(stateFIPS)
        Else
          MsgBox "State FIPS code " & stateFIPS & " for station " & stationID & " not recognized.  Station not imported."
          GoTo nextSta
        End If
        If SSDB.state.Stations.IndexFromKey(stationID) > 0 Then
          'MsgBox "There is already data stored for station " & _
              stationID & "-" & staName & vbCrLf & _
              "No data will be imported for this station.", , _
              "Station already in database"
          GoTo nextSta
        End If
        Set myStation = New ssStation
        Set myStation.Db = SSDB
        Set myStation.state = SSDB.States(stateFIPS)
        myStation.id = stationID
        myStation.Add stationValues(), 1, 1
        If myStation.Name <> "bad" Then
          Set myStatistic = New ssStatistic
          Set myStatistic.Db = SSDB
          Set myStatistic.station = myStation
          For i = 1 To 4
            If dataValues(2, i, 4) <> "" Then myStatistic.Add dataValues(), i
          Next i
        End If
nextSta:
      Next staCnt
badSheet:
    End With
  Next h
errTrap:
  'set state collection of stations to nothing for all import states
  For i = 0 To UBound(impStates)
    If Not SSDB.States(impStates(i)).Stations Is Nothing Then
      SSDB.States(impStates(i)).Stations.Clear
      Set SSDB.States(impStates(i)).Stations = Nothing
    End If
  Next i
  If Err.Number = 999 Then
    IPC.SendMonitorMessage "(MSG1 User Canceled Import)"
  ElseIf Err.Number > 0 Then
    MsgBox "Problem importing:  " & Err.Description, , "NWIS Import Problem"
  End If
  IPC.SendMonitorMessage "(CLOSE)"
  If Len(Trim(bumSheets)) > 0 Then
    myMsgBox.Show "The keyword 'site_no' was not found in the spreadsheet(s)" & _
        bumSheets & "." & vbCrLf & vbCrLf & "'site_no' must appear in the top " & _
        "left corner of the data block." & vbCrLf & "No data was imported from " & _
        "this spreadsheet(s).", "Keyword not found", "+-&OK"
  End If
  'Close EXCEL
  XLBook.Close False
  Set XLBook = Nothing
  XLApp.Quit
  Set XLApp = Nothing
  'Print text file with suspected data errors
  If Len(bumFields) > 0 Then
    OutFile = FreeFile
    Open "NWIS-Import.txt" For Output As OutFile
    Print #OutFile, bumFields
    Close OutFile
    myMsgBox.Show "The import file " & ImpFileName & " contains bad values." _
        & vbCrLf & "Check the '" & CurDir & "\NWIS-Import.txt' file " & _
        "to see which fields these were.", _
        "Bad Data Value(s)", "+-&OK"
  End If
End Sub

Sub BCFImport(ImpFileName As String)
  Dim staCnt&, attCnt&, inFile&, OutFile&, i&, staIndex&, attIndex&, response&
  Dim textLine$, str$, badFields$, badStats$, OverWriteInfo$
  Dim stationValues() As String, dataValues() As String
  Dim myStation As ssStation
  Dim myStatistic As ssStatistic
  
  On Error GoTo errTrap
  
  IPC.SendMonitorMessage "(OPEN StreamStatsDB)"
  IPC.SendMonitorMessage "(BUTTOFF DETAILS)"
  IPC.SendMonitorMessage "(BUTTON CANCEL)"
  IPC.SendMonitorMessage "(BUTTON PAUSE)"
  IPC.SendMonitorMessage "(MSG1 Importing BCF file)"
  IPC.SendMonitorMessage "(PROGRESS 0)"
  
  inFile = FreeFile
  Open ImpFileName For Input As inFile
  
  Line Input #inFile, textLine
  Do Until EOF(inFile)
    'Progress bar
    If SSDB.state.Stations.Count > 0 Then
      IPC.SendMonitorMessage "(PROGRESS " & _
          staCnt * 100 / SSDB.state.Stations.Count & ")"
    Else 'use an estimate based on number of counties
      IPC.SendMonitorMessage "(PROGRESS " & _
          staCnt * 100 / (SSDB.state.Counties.Count * 40) & ")"
    End If
    'Cancel/Pause
    Select Case NextPipeCharacter(IPC.hPipeReadFromProcess("Status"))
      Case "P"
        While NextPipeCharacter(IPC.hPipeReadFromProcess("Status")) <> "R"
          DoEvents
        Wend
      Case "C"
        MsgBox "The BCF import operation has been interrupted." & _
            vbCrLf & vbCrLf & "Data has been imported for stations in " & _
            SSDB.state.Name & " up to station " & stationValues(2, 1, 1) & ".", , _
            "Import interrupted"
        Err.Raise 999
    End Select
    If Left(textLine, 1) = 1 Then  'Station data card
      staCnt = staCnt + 1
      'Read in station ID and Name
      ReDim stationValues(2 To 2, 1 To 1, 1 To UBound(StationFields))
      ReDim dataValues(2 To 2, 1 To 166, 1 To UBound(DataFields))
      stationValues(2, 1, 1) = Trim(Mid(textLine, 2, 15)) 'StationID
      stationValues(2, 1, 2) = Trim(Mid(textLine, 21))    'Station Name
      'Read in "District_Code"
      attCnt = 1
      dataValues(2, 1, 2) = "24"
      If Mid(textLine, 17, 1) = " " Then  'add leading 0 to code
        dataValues(2, 1, 4) = "0" & Mid(textLine, 18, 1)
      Else
        dataValues(2, 1, 4) = Mid(textLine, 17, 2)
      End If
      dataValues(2, 1, 7) = "Imported from Basin Characteristics file"
      Line Input #inFile, textLine
      'Loop thru Statistic data cards
      While Left(textLine, 1) = 2 And Not EOF(inFile)
        textLine = Mid(textLine, 17)
        While Len(Trim(textLine)) > 2
          attCnt = attCnt + 1
          str = Left(textLine, 3)
          'update from BCF ID to StreamStatsDB ID
          dataValues(2, attCnt, 2) = ConvertBCFfield(Trim(str))
          If dataValues(2, attCnt, 2) = -999 Then  'can not identify BCF stat ID - ignore this stat
            attCnt = attCnt - 1
            badFields = badFields & vbCrLf & "Numeric identifier " & str & _
                        " for station " & stationValues(2, 1, 1) & "."
          Else  'have match for BCF stat ID
            dataValues(2, attCnt, 4) = Trim(Mid(textLine, 4, 7))
            dataValues(2, attCnt, 7) = "Imported from Basin Characteristics file"
          End If
          textLine = Mid(textLine, 11)
        Wend
        Line Input #inFile, textLine
      Wend
    End If
    'Station and its statistics have been read in
    'Check whether station/statistics exist - add if not, write to text file if so
    staIndex = SSDB.state.Stations.IndexFromKey(stationValues(2, 1, 1))
    If staIndex = -1 Then 'Station does not exist - add it and its statistics
      Set myStation = New ssStation
      Set myStation.Db = SSDB
      Set myStation.state = SSDB.state
      myStation.Add stationValues(), 1, 1
      myStation.id = stationValues(2, 1, 1)
      For i = 1 To attCnt
        'Retrieve the StatLabel based upon the BCF stat code
        Set myStatistic = New ssStatistic
        With myStatistic
          Set .Db = SSDB
          Set .station = myStation
          .Add dataValues(), i, 1
          If dataValues(2, i, 2) = "bad" Then
            badFields = badFields & vbCrLf & "Numeric identifier " & str & _
                " for station " & myStation.id & "."
          End If
        End With
      Next i
    ElseIf staIndex > 0 Then 'Station exists - check whether its stats also exist
      Set myStation = SSDB.state.Stations(staIndex)
      For i = 1 To attCnt
        attIndex = myStation.Statistics.IndexFromKey(dataValues(2, i, 2))
        If attIndex = -1 Then 'Statistic does not exist
          Set myStatistic = New ssStatistic
          With myStatistic
            Set .Db = SSDB
            Set .station = myStation
            .Add dataValues(), i, 1
          End With
        ElseIf attIndex > 0 Then 'Statistic already exists - write to text file
          If response <> 2 And response <> 4 Then 'user hasn't chosen to replace or keep all
            response = myMsgBox.Show("For station " & myStation.Name & " the statistic " & myStation.Statistics(attIndex).Name & _
                                     " already exists." & vbCrLf & "Existing value: " & myStation.Statistics(attIndex).value & _
                                     "  New value: " & dataValues(2, i, 4) & vbCrLf & "What do you want to do?", "BCF Import - Data Exists", _
                                     "&Replace", "Replace &All", "+&Keep", "K&eep All", "-&Cancel")
          End If
          If response = 5 Then Err.Raise 999
          If response > 2 Then 'keeping existing, note it to file
            badStats = badStats & vbCrLf & vbTab & _
                       Left(myStation.id & "       ", 15) & vbTab & _
                       StrPad(dataValues(2, i, 2), 7) & vbTab & _
                       StrPad(myStation.Statistics(attIndex).value, 10) & _
                       vbTab & StrPad(dataValues(2, i, 4), 8)
          Else 'overwrite value
            OverWriteInfo = OverWriteInfo & vbCrLf & vbTab & _
                            Left(myStation.id & "       ", 15) & vbTab & _
                            StrPad(dataValues(2, i, 2), 7) & vbTab & _
                            StrPad(myStation.Statistics(attIndex).value, 10) & _
                            vbTab & StrPad(dataValues(2, i, 4), 8)
            myStation.Statistics(attIndex).Edit dataValues(), i
          End If
        End If
      Next i
    End If
    'check for empty lines/end of file
    While textLine = "" And Not EOF(inFile)
      Line Input #inFile, textLine
    Wend
  Loop
errTrap:
  If Err.Number = 999 Then
    IPC.SendMonitorMessage "(MSG1 User Canceled Import)"
  ElseIf Err.Number > 0 Then
    MsgBox "Problem importing:  " & Err.Description, , "BCF Import Problem"
  End If
  IPC.SendMonitorMessage "(CLOSE)"
  Close inFile
  If Len(Trim(OverWriteInfo)) > 0 Or Len(Trim(badStats)) > 0 Or Len(Trim(badFields)) > 0 Then
    OutFile = FreeFile
    Open FilenameNoExt(ImpFileName) & "_Import.txt" For Output As OutFile
    Print #OutFile, "This file contains summary information regarding the"
    Print #OutFile, "import of the Basin Characteristics file " & ImpFileName
    Print #OutFile, "into the StreamStats database for " & SSDB.state.Name
    If Len(Trim(OverWriteInfo)) > 0 Then
      Print #OutFile, ""
      Print #OutFile, ""
      Print #OutFile, "The following records show statistic values that were"
      Print #OutFile, "updated from the import:"
      Print #OutFile, ""
      Print #OutFile, vbTab & "               " & vbTab & "       " & vbTab & "Previous" & vbTab & " Updated"
      Print #OutFile, vbTab & "Station ID     " & vbTab & "Stat ID" & vbTab & "   Value" & vbTab & "   Value"
      Print #OutFile, OverWriteInfo
    End If
    If Len(Trim(badStats)) > 0 Then
      Print #OutFile, ""
      Print #OutFile, ""
      Print #OutFile, ""
      Print #OutFile, "The following records show statistic values that were"
      Print #OutFile, "not updated from the import (i.e. the new value was ignored):"
      Print #OutFile, ""
      Print #OutFile, vbTab & "               " & vbTab & "       " & vbTab & "Existing" & vbTab & " Ignored"
      Print #OutFile, vbTab & "Station ID     " & vbTab & "Stat ID" & vbTab & "   Value" & vbTab & "   Value"
      Print #OutFile, badStats
    End If
    If Len(Trim(badFields)) > 0 Then
      Print #OutFile, ""
      Print #OutFile, ""
      Print #OutFile, ""
      Print #OutFile, "The following records list statistics that were not imported"
      Print #OutFile, "because they contained unrecognized numeric identifiers:"
      Print #OutFile, badFields
    End If
    MsgBox "Summary information about this import was saved in the file " & _
           FilenameNoExt(ImpFileName) & "_Import.txt"
    Close OutFile
  End If
End Sub

Public Function NextPipeCharacter(PipeHandle As Long) As String
  Dim res&, lread&, lavail&, lmessage&
  Dim inbuf As Byte
    
  DoEvents
  res = PeekNamedPipe(PipeHandle, ByVal 0&, 0, lread, lavail, lmessage)
  If res <> 0 And lavail > 0 Then
    lavail = 1 'Only get one character
    res = ReadFile(PipeHandle, inbuf, lavail, lread, 0)
    NextPipeCharacter = Chr(inbuf)
  End If
End Function

Private Function ConvertBCFfield(BCFfield As String, Optional Flag As Long) As Variant
  'If Len(Trim(BCFfield)) = 0 Then Exit Function
  BCFfield = CLng(BCFfield)
  Select Case BCFfield
    Case 1: ConvertBCFfield = 73
    Case 2: ConvertBCFfield = 70
    Case 3: ConvertBCFfield = 166
    Case 4: ConvertBCFfield = 104
    Case 5: ConvertBCFfield = 89
    Case 6: ConvertBCFfield = 88
    Case 7: ConvertBCFfield = 171
    Case 8: ConvertBCFfield = 103
    Case 9: ConvertBCFfield = 77
    Case 10: ConvertBCFfield = 78
    Case 11: ConvertBCFfield = 79
    Case 12: ConvertBCFfield = 155
    Case 13: ConvertBCFfield = 149
    Case 14: ConvertBCFfield = 142
    Case 15: ConvertBCFfield = 143
    Case 16: ConvertBCFfield = 164
    Case 17: ConvertBCFfield = 86
    Case 18: ConvertBCFfield = 64
    Case 19: ConvertBCFfield = 84
    Case 20: ConvertBCFfield = 87
    Case 21: ConvertBCFfield = 169
    Case 22: ConvertBCFfield = 7
    Case 23: ConvertBCFfield = 8
    Case 24: ConvertBCFfield = 24
    Case 32: ConvertBCFfield = 96
    Case 33: ConvertBCFfield = 29
    Case 34: ConvertBCFfield = 30
    Case 35: ConvertBCFfield = 31
    Case 36: ConvertBCFfield = 32
    Case 37: ConvertBCFfield = 33
    Case 41: ConvertBCFfield = 125
    Case 42: ConvertBCFfield = 127
    Case 43: ConvertBCFfield = 129
    Case 44: ConvertBCFfield = 107
    Case 45: ConvertBCFfield = 110
    Case 46: ConvertBCFfield = 113
    Case 47: ConvertBCFfield = 99
    Case 48: ConvertBCFfield = 116
    Case 49: ConvertBCFfield = 118
    Case 50: ConvertBCFfield = 120
    Case 51: ConvertBCFfield = 101
    Case 52: ConvertBCFfield = 123
    Case 53: ConvertBCFfield = 97
    Case 54: ConvertBCFfield = 91
    Case 55: ConvertBCFfield = 38
    Case 56: ConvertBCFfield = 36
    Case 57: ConvertBCFfield = 26
    Case 58: ConvertBCFfield = 34
    Case 59: ConvertBCFfield = 27
    Case 60: ConvertBCFfield = 109
    Case 61: ConvertBCFfield = 108
    Case 63: ConvertBCFfield = 122
    Case 64: ConvertBCFfield = 121
    Case 65: ConvertBCFfield = 92
    Case 70: ConvertBCFfield = 94
    Case 71: ConvertBCFfield = 95
    Case 72: ConvertBCFfield = 112
    Case 75: ConvertBCFfield = 172
    Case 76: ConvertBCFfield = 174
    Case 77: ConvertBCFfield = 176
    Case 78: ConvertBCFfield = 178
    Case 79: ConvertBCFfield = 180
    Case 80: ConvertBCFfield = 182
    Case 81: ConvertBCFfield = 184
    Case 82: ConvertBCFfield = 186
    Case 83: ConvertBCFfield = 226
    Case 84: ConvertBCFfield = 227
    Case 85: ConvertBCFfield = 228
    Case 86: ConvertBCFfield = 428
    Case 87: ConvertBCFfield = 1178
    Case 88: ConvertBCFfield = 473
    Case 89: ConvertBCFfield = 477
    Case 90: ConvertBCFfield = 481
    Case 91: ConvertBCFfield = 437
    Case 92: ConvertBCFfield = 441
    Case 93: ConvertBCFfield = 445
    Case 94: ConvertBCFfield = 449
    Case 95: ConvertBCFfield = 453
    Case 96: ConvertBCFfield = 457
    Case 97: ConvertBCFfield = 461
    Case 98: ConvertBCFfield = 465
    Case 99: ConvertBCFfield = 469
    Case 100: ConvertBCFfield = 475
    Case 101: ConvertBCFfield = 479
    Case 102: ConvertBCFfield = 483
    Case 103: ConvertBCFfield = 439
    Case 104: ConvertBCFfield = 443
    Case 105: ConvertBCFfield = 447
    Case 106: ConvertBCFfield = 451
    Case 107: ConvertBCFfield = 455
    Case 108: ConvertBCFfield = 459
    Case 109: ConvertBCFfield = 463
    Case 110: ConvertBCFfield = 467
    Case 111: ConvertBCFfield = 471
    Case 112: ConvertBCFfield = 325
    Case 113: ConvertBCFfield = 329
    Case 114: ConvertBCFfield = 331
    Case 115: ConvertBCFfield = 333
    Case 116: ConvertBCFfield = 337
    Case 117: ConvertBCFfield = 339
    Case 118: ConvertBCFfield = 341
    Case 119: ConvertBCFfield = 343
    Case 120: ConvertBCFfield = 345
    Case 121: ConvertBCFfield = 347
    Case 122: ConvertBCFfield = 349
    Case 123: ConvertBCFfield = 353
    Case 124: ConvertBCFfield = 355
    Case 125: ConvertBCFfield = 357
    Case 126: ConvertBCFfield = 361
    Case 127: ConvertBCFfield = 363
    Case 128: ConvertBCFfield = 365
    Case 129: ConvertBCFfield = 369
    Case 130: ConvertBCFfield = 371
    Case 131: ConvertBCFfield = 247
    Case 132: ConvertBCFfield = 291
    Case 133: ConvertBCFfield = 309
    Case 134: ConvertBCFfield = 313
    Case 135: ConvertBCFfield = 319
    Case 136: ConvertBCFfield = 235
    Case 137: ConvertBCFfield = 239
    Case 138: ConvertBCFfield = 241
    Case 139: ConvertBCFfield = 237
    Case 140: ConvertBCFfield = 243
    Case 141: ConvertBCFfield = 245
    Case 142: ConvertBCFfield = 253
    Case 143: ConvertBCFfield = 255
    Case 144: ConvertBCFfield = 257
    Case 145: ConvertBCFfield = 259
    Case 146: ConvertBCFfield = 261
    Case 147: ConvertBCFfield = 263
    Case 148: ConvertBCFfield = 265
    Case 149: ConvertBCFfield = 271
    Case 150: ConvertBCFfield = 273
    Case 151: ConvertBCFfield = 275
    Case 152: ConvertBCFfield = 277
    Case 153: ConvertBCFfield = 279
    Case 154: ConvertBCFfield = 281
    Case 155: ConvertBCFfield = 283
    Case 156: ConvertBCFfield = 289
    Case 157: ConvertBCFfield = 293
    Case 158: ConvertBCFfield = 295
    Case 159: ConvertBCFfield = 297
    Case 160: ConvertBCFfield = 299
    Case 161: ConvertBCFfield = 301
    Case 162: ConvertBCFfield = 307
    Case 163: ConvertBCFfield = 311
    Case 164: ConvertBCFfield = 315
    Case 165: ConvertBCFfield = 317
    Case 167: ConvertBCFfield = 37
    Case 168: ConvertBCFfield = 414
    Case 169: ConvertBCFfield = 1177
    Case 171: ConvertBCFfield = 420
    Case 172: ConvertBCFfield = 416
    Case 173: ConvertBCFfield = 410
    Case 174: ConvertBCFfield = 408
    Case 175: ConvertBCFfield = 400
    Case 176: ConvertBCFfield = 390
    Case 177: ConvertBCFfield = 384
    Case 178: ConvertBCFfield = 188
    Case 179: ConvertBCFfield = 231
    Case 180: ConvertBCFfield = 229
    Case 181: ConvertBCFfield = 230
    Case 196: ConvertBCFfield = 233
    Case 197: ConvertBCFfield = 234
    Case 198: ConvertBCFfield = 436
    Case 199: ConvertBCFfield = 373
    Case Else: ConvertBCFfield = -999
  End Select
End Function

'
'Private Function ConvertBCFfield(BCFfield As String, Optional Flag As Long) As Variant
'  'If Len(Trim(BCFfield)) = 0 Then Exit Function
'  BCFfield = CLng(BCFfield)
'  Select Case BCFfield
'    Case 1: If Flag = 1 Then ConvertBCFfield = "AREA" Else ConvertBCFfield = 73
'    Case 2: If Flag = 1 Then ConvertBCFfield = "CONTDA" Else ConvertBCFfield = 70
'    Case 3: If Flag = 1 Then ConvertBCFfield = "CSL1085" Else ConvertBCFfield = 166
'    Case 4: If Flag = 1 Then ConvertBCFfield = "BSLOPGM" Else ConvertBCFfield = 104
'    Case 5: If Flag = 1 Then ConvertBCFfield = "LENGTH" Else ConvertBCFfield = 89
'    Case 6: If Flag = 1 Then ConvertBCFfield = "CLENBLUE" Else ConvertBCFfield = 88
'    Case 7: If Flag = 1 Then ConvertBCFfield = "VALLEN" Else ConvertBCFfield = 171
'    Case 8: If Flag = 1 Then ConvertBCFfield = "ELEV" Else ConvertBCFfield = 103
'    Case 9: If Flag = 1 Then ConvertBCFfield = "ELV10,85" Else ConvertBCFfield = 77
'    Case 10: If Flag = 1 Then ConvertBCFfield = "EL5000" Else ConvertBCFfield = 78
'    Case 11: If Flag = 1 Then ConvertBCFfield = "EL6000" Else ConvertBCFfield = 79
'    Case 12: If Flag = 1 Then ConvertBCFfield = "STORAGE" Else ConvertBCFfield = 155
'    Case 13: If Flag = 1 Then ConvertBCFfield = "LAKEAREA" Else ConvertBCFfield = 149
'    Case 14: If Flag = 1 Then ConvertBCFfield = "FOREST" Else ConvertBCFfield = 142
'    Case 15: If Flag = 1 Then ConvertBCFfield = "GLACIER" Else ConvertBCFfield = 143
'    Case 16: If Flag = 1 Then ConvertBCFfield = "SOIL INF" Else ConvertBCFfield = 164
'    Case 17: If Flag = 1 Then ConvertBCFfield = "LOESSDEP" Else ConvertBCFfield = 86
'    Case 18: If Flag = 1 Then ConvertBCFfield = "AZIMUTH" Else ConvertBCFfield = 64
'    Case 19: If Flag = 1 Then ConvertBCFfield = "LAT" Else ConvertBCFfield = 84
'    Case 20: If Flag = 1 Then ConvertBCFfield = "LONG" Else ConvertBCFfield = 87
'    Case 21: If Flag = 1 Then ConvertBCFfield = "TIMETOPK" Else ConvertBCFfield = 169
'    Case 22: If Flag = 1 Then ConvertBCFfield = "LAT GAGE" Else ConvertBCFfield = 7
'    Case 23: If Flag = 1 Then ConvertBCFfield = "LNG GAGE" Else ConvertBCFfield = 8
'    Case 24: If Flag = 1 Then ConvertBCFfield = "DISTRICT" Else ConvertBCFfield = 24
'    Case 32: If Flag = 1 Then ConvertBCFfield = "PRECIP" Else ConvertBCFfield = 96
'    Case 33: If Flag = 1 Then ConvertBCFfield = "I24,2" Else ConvertBCFfield = 29
'    Case 34: If Flag = 1 Then ConvertBCFfield = "I24,10" Else ConvertBCFfield = 30
'    Case 35: If Flag = 1 Then ConvertBCFfield = "I24,25" Else ConvertBCFfield = 31
'    Case 36: If Flag = 1 Then ConvertBCFfield = "I24,50" Else ConvertBCFfield = 32
'    Case 37: If Flag = 1 Then ConvertBCFfield = "I24,100" Else ConvertBCFfield = 33
'    Case 41: If Flag = 1 Then ConvertBCFfield = "PRC10" Else ConvertBCFfield = 125
'    Case 42: If Flag = 1 Then ConvertBCFfield = "PRC11" Else ConvertBCFfield = 127
'    Case 43: If Flag = 1 Then ConvertBCFfield = "PRC12" Else ConvertBCFfield = 129
'    Case 44: If Flag = 1 Then ConvertBCFfield = "PRC1" Else ConvertBCFfield = 107
'    Case 45: If Flag = 1 Then ConvertBCFfield = "PRC2" Else ConvertBCFfield = 110
'    Case 46: If Flag = 1 Then ConvertBCFfield = "PRC3" Else ConvertBCFfield = 113
'    Case 47: If Flag = 1 Then ConvertBCFfield = "PRC4" Else ConvertBCFfield = 99
'    Case 48: If Flag = 1 Then ConvertBCFfield = "PRC5" Else ConvertBCFfield = 116
'    Case 49: If Flag = 1 Then ConvertBCFfield = "PRC6" Else ConvertBCFfield = 118
'    Case 50: If Flag = 1 Then ConvertBCFfield = "PRC7" Else ConvertBCFfield = 120
'    Case 51: If Flag = 1 Then ConvertBCFfield = "PRC8" Else ConvertBCFfield = 101
'    Case 52: If Flag = 1 Then ConvertBCFfield = "PRC9" Else ConvertBCFfield = 123
'    Case 53: If Flag = 1 Then ConvertBCFfield = "SNOFALL" Else ConvertBCFfield = 97
'    Case 54: If Flag = 1 Then ConvertBCFfield = "SNOMAR" Else ConvertBCFfield = 91
'    Case 55: If Flag = 1 Then ConvertBCFfield = "SNOAPR" Else ConvertBCFfield = 38
'    Case 56: If Flag = 1 Then ConvertBCFfield = "SN2" Else ConvertBCFfield = 36
'    Case 57: If Flag = 1 Then ConvertBCFfield = "SN10" Else ConvertBCFfield = 26
'    Case 58: If Flag = 1 Then ConvertBCFfield = "SN25" Else ConvertBCFfield = 34
'    Case 59: If Flag = 1 Then ConvertBCFfield = "SN100" Else ConvertBCFfield = 27
'    Case 60: If Flag = 1 Then ConvertBCFfield = "JANMIN" Else ConvertBCFfield = 109
'    Case 61: If Flag = 1 Then ConvertBCFfield = "JANAV" Else ConvertBCFfield = 108
'    Case 63: If Flag = 1 Then ConvertBCFfield = "JULYMAX" Else ConvertBCFfield = 122
'    Case 64: If Flag = 1 Then ConvertBCFfield = "JULYAV" Else ConvertBCFfield = 121
'    Case 65: If Flag = 1 Then ConvertBCFfield = "WE MAR2" Else ConvertBCFfield = 92
'    Case 70: If Flag = 1 Then ConvertBCFfield = "EVAP" Else ConvertBCFfield = 94
'    Case 71: If Flag = 1 Then ConvertBCFfield = "EVAPAN" Else ConvertBCFfield = 95
'    Case 72: If Flag = 1 Then ConvertBCFfield = "FROST" Else ConvertBCFfield = 112
'    Case 75: If Flag = 1 Then ConvertBCFfield = "P1,25" Else ConvertBCFfield = 172
'    Case 76: If Flag = 1 Then ConvertBCFfield = "P2" Else ConvertBCFfield = 174
'    Case 77: If Flag = 1 Then ConvertBCFfield = "P5" Else ConvertBCFfield = 176
'    Case 78: If Flag = 1 Then ConvertBCFfield = "P10" Else ConvertBCFfield = 178
'    Case 79: If Flag = 1 Then ConvertBCFfield = "P25" Else ConvertBCFfield = 180
'    Case 80: If Flag = 1 Then ConvertBCFfield = "P50" Else ConvertBCFfield = 182
'    Case 81: If Flag = 1 Then ConvertBCFfield = "P100" Else ConvertBCFfield = 184
'    Case 82: If Flag = 1 Then ConvertBCFfield = "P200" Else ConvertBCFfield = 186
'    Case 83: If Flag = 1 Then ConvertBCFfield = "MEANPK" Else ConvertBCFfield = 226
'    Case 84: If Flag = 1 Then ConvertBCFfield = "SDPK" Else ConvertBCFfield = 227
'    Case 85: If Flag = 1 Then ConvertBCFfield = "SKEWPK" Else ConvertBCFfield = 228
'    Case 86: If Flag = 1 Then ConvertBCFfield = "QA" Else ConvertBCFfield = 428
'    Case 87: If Flag = 1 Then ConvertBCFfield = "SDQA" Else ConvertBCFfield = 1178
'    Case 88: If Flag = 1 Then ConvertBCFfield = "Q10" Else ConvertBCFfield = 473
'    Case 89: If Flag = 1 Then ConvertBCFfield = "Q11" Else ConvertBCFfield = 477
'    Case 90: If Flag = 1 Then ConvertBCFfield = "Q12" Else ConvertBCFfield = 481
'    Case 91: If Flag = 1 Then ConvertBCFfield = "Q1" Else ConvertBCFfield = 437
'    Case 92: If Flag = 1 Then ConvertBCFfield = "Q2" Else ConvertBCFfield = 441
'    Case 93: If Flag = 1 Then ConvertBCFfield = "Q3" Else ConvertBCFfield = 445
'    Case 94: If Flag = 1 Then ConvertBCFfield = "Q4" Else ConvertBCFfield = 449
'    Case 95: If Flag = 1 Then ConvertBCFfield = "Q5" Else ConvertBCFfield = 453
'    Case 96: If Flag = 1 Then ConvertBCFfield = "Q6" Else ConvertBCFfield = 457
'    Case 97: If Flag = 1 Then ConvertBCFfield = "Q7" Else ConvertBCFfield = 461
'    Case 98: If Flag = 1 Then ConvertBCFfield = "Q8" Else ConvertBCFfield = 465
'    Case 99: If Flag = 1 Then ConvertBCFfield = "Q9" Else ConvertBCFfield = 469
'    Case 100: If Flag = 1 Then ConvertBCFfield = "SDQ10" Else ConvertBCFfield = 475
'    Case 101: If Flag = 1 Then ConvertBCFfield = "SDQ11" Else ConvertBCFfield = 479
'    Case 102: If Flag = 1 Then ConvertBCFfield = "SDQ12" Else ConvertBCFfield = 483
'    Case 103: If Flag = 1 Then ConvertBCFfield = "SDQ1" Else ConvertBCFfield = 439
'    Case 104: If Flag = 1 Then ConvertBCFfield = "SDQ2" Else ConvertBCFfield = 443
'    Case 105: If Flag = 1 Then ConvertBCFfield = "SDQ3" Else ConvertBCFfield = 447
'    Case 106: If Flag = 1 Then ConvertBCFfield = "SDQ4" Else ConvertBCFfield = 451
'    Case 107: If Flag = 1 Then ConvertBCFfield = "SDQ5" Else ConvertBCFfield = 455
'    Case 108: If Flag = 1 Then ConvertBCFfield = "SDQ6" Else ConvertBCFfield = 459
'    Case 109: If Flag = 1 Then ConvertBCFfield = "SDQ7" Else ConvertBCFfield = 463
'    Case 110: If Flag = 1 Then ConvertBCFfield = "SDQ8" Else ConvertBCFfield = 467
'    Case 111: If Flag = 1 Then ConvertBCFfield = "SDQ9" Else ConvertBCFfield = 471
'    Case 112: If Flag = 1 Then ConvertBCFfield = "M1,2" Else ConvertBCFfield = 325
'    Case 113: If Flag = 1 Then ConvertBCFfield = "M1,10" Else ConvertBCFfield = 329
'    Case 114: If Flag = 1 Then ConvertBCFfield = "M1,20" Else ConvertBCFfield = 331
'    Case 115: If Flag = 1 Then ConvertBCFfield = "M3,2" Else ConvertBCFfield = 333
'    Case 116: If Flag = 1 Then ConvertBCFfield = "M3,10" Else ConvertBCFfield = 337
'    Case 117: If Flag = 1 Then ConvertBCFfield = "M3,20" Else ConvertBCFfield = 339
'    Case 118: If Flag = 1 Then ConvertBCFfield = "M7,2" Else ConvertBCFfield = 341
'    Case 119: If Flag = 1 Then ConvertBCFfield = "M7,5" Else ConvertBCFfield = 343
'    Case 120: If Flag = 1 Then ConvertBCFfield = "M7,10" Else ConvertBCFfield = 345
'    Case 121: If Flag = 1 Then ConvertBCFfield = "M7,20" Else ConvertBCFfield = 347
'    Case 122: If Flag = 1 Then ConvertBCFfield = "M14,2" Else ConvertBCFfield = 349
'    Case 123: If Flag = 1 Then ConvertBCFfield = "M14,10" Else ConvertBCFfield = 353
'    Case 124: If Flag = 1 Then ConvertBCFfield = "M14,20" Else ConvertBCFfield = 355
'    Case 125: If Flag = 1 Then ConvertBCFfield = "M30,2" Else ConvertBCFfield = 357
'    Case 126: If Flag = 1 Then ConvertBCFfield = "M30,10" Else ConvertBCFfield = 361
'    Case 127: If Flag = 1 Then ConvertBCFfield = "M30,20" Else ConvertBCFfield = 363
'    Case 128: If Flag = 1 Then ConvertBCFfield = "M90,2" Else ConvertBCFfield = 365
'    Case 129: If Flag = 1 Then ConvertBCFfield = "M90,10" Else ConvertBCFfield = 369
'    Case 130: If Flag = 1 Then ConvertBCFfield = "M90,20" Else ConvertBCFfield = 371
'    Case 131: If Flag = 1 Then ConvertBCFfield = "V1,100" Else ConvertBCFfield = 247
'    Case 132: If Flag = 1 Then ConvertBCFfield = "V15,5" Else ConvertBCFfield = 291
'    Case 133: If Flag = 1 Then ConvertBCFfield = "V30,5" Else ConvertBCFfield = 309
'    Case 134: If Flag = 1 Then ConvertBCFfield = "V30,20" Else ConvertBCFfield = 313
'    Case 135: If Flag = 1 Then ConvertBCFfield = "V30,100" Else ConvertBCFfield = 319
'    Case 136: If Flag = 1 Then ConvertBCFfield = "V1,2" Else ConvertBCFfield = 235
'    Case 137: If Flag = 1 Then ConvertBCFfield = "V1,5" Else ConvertBCFfield = 239
'    Case 138: If Flag = 1 Then ConvertBCFfield = "V1,10" Else ConvertBCFfield = 241
'    Case 139: If Flag = 1 Then ConvertBCFfield = "V1,20" Else ConvertBCFfield = 237
'    Case 140: If Flag = 1 Then ConvertBCFfield = "V1,25" Else ConvertBCFfield = 243
'    Case 141: If Flag = 1 Then ConvertBCFfield = "V1,50" Else ConvertBCFfield = 245
'    Case 142: If Flag = 1 Then ConvertBCFfield = "V3,2" Else ConvertBCFfield = 253
'    Case 143: If Flag = 1 Then ConvertBCFfield = "V3,5" Else ConvertBCFfield = 255
'    Case 144: If Flag = 1 Then ConvertBCFfield = "V3,10" Else ConvertBCFfield = 257
'    Case 145: If Flag = 1 Then ConvertBCFfield = "V3,20" Else ConvertBCFfield = 259
'    Case 146: If Flag = 1 Then ConvertBCFfield = "V3,25" Else ConvertBCFfield = 261
'    Case 147: If Flag = 1 Then ConvertBCFfield = "V3,50" Else ConvertBCFfield = 263
'    Case 148: If Flag = 1 Then ConvertBCFfield = "V3,100" Else ConvertBCFfield = 265
'    Case 149: If Flag = 1 Then ConvertBCFfield = "V7,2" Else ConvertBCFfield = 271
'    Case 150: If Flag = 1 Then ConvertBCFfield = "V7,5" Else ConvertBCFfield = 273
'    Case 151: If Flag = 1 Then ConvertBCFfield = "V7,10" Else ConvertBCFfield = 275
'    Case 152: If Flag = 1 Then ConvertBCFfield = "V7,20" Else ConvertBCFfield = 277
'    Case 153: If Flag = 1 Then ConvertBCFfield = "V7,25" Else ConvertBCFfield = 279
'    Case 154: If Flag = 1 Then ConvertBCFfield = "V7,50" Else ConvertBCFfield = 281
'    Case 155: If Flag = 1 Then ConvertBCFfield = "V7,100" Else ConvertBCFfield = 283
'    Case 156: If Flag = 1 Then ConvertBCFfield = "V15,2" Else ConvertBCFfield = 289
'    Case 157: If Flag = 1 Then ConvertBCFfield = "V15,10" Else ConvertBCFfield = 293
'    Case 158: If Flag = 1 Then ConvertBCFfield = "V15,20" Else ConvertBCFfield = 295
'    Case 159: If Flag = 1 Then ConvertBCFfield = "V15,25" Else ConvertBCFfield = 297
'    Case 160: If Flag = 1 Then ConvertBCFfield = "V15,50" Else ConvertBCFfield = 299
'    Case 161: If Flag = 1 Then ConvertBCFfield = "V15,100" Else ConvertBCFfield = 301
'    Case 162: If Flag = 1 Then ConvertBCFfield = "V30,2" Else ConvertBCFfield = 307
'    Case 163: If Flag = 1 Then ConvertBCFfield = "V30,10" Else ConvertBCFfield = 311
'    Case 164: If Flag = 1 Then ConvertBCFfield = "V30,25" Else ConvertBCFfield = 315
'    Case 165: If Flag = 1 Then ConvertBCFfield = "V30,50" Else ConvertBCFfield = 317
'    Case 167: If Flag = 1 Then ConvertBCFfield = "SN50" Else ConvertBCFfield = 37
'    Case 168: If Flag = 1 Then ConvertBCFfield = "D85" Else ConvertBCFfield = 414
'    Case 169: If Flag = 1 Then ConvertBCFfield = "DEPH25" Else ConvertBCFfield = 1177
'    Case 171: If Flag = 1 Then ConvertBCFfield = "D95" Else ConvertBCFfield = 420
'    Case 172: If Flag = 1 Then ConvertBCFfield = "D90" Else ConvertBCFfield = 416
'    Case 173: If Flag = 1 Then ConvertBCFfield = "D75" Else ConvertBCFfield = 410
'    Case 174: If Flag = 1 Then ConvertBCFfield = "D70" Else ConvertBCFfield = 408
'    Case 175: If Flag = 1 Then ConvertBCFfield = "D50" Else ConvertBCFfield = 400
'    Case 176: If Flag = 1 Then ConvertBCFfield = "D25" Else ConvertBCFfield = 390
'    Case 177: If Flag = 1 Then ConvertBCFfield = "D10" Else ConvertBCFfield = 384
'    Case 178: If Flag = 1 Then ConvertBCFfield = "P500" Else ConvertBCFfield = 188
'    Case 179: If Flag = 1 Then ConvertBCFfield = "WRC SKEW" Else ConvertBCFfield = 231
'    Case 180: If Flag = 1 Then ConvertBCFfield = "WRC MEAN" Else ConvertBCFfield = 229
'    Case 181: If Flag = 1 Then ConvertBCFfield = "WRC STD" Else ConvertBCFfield = 230
'    Case 196: If Flag = 1 Then ConvertBCFfield = "YRSPK" Else ConvertBCFfield = 233
'    Case 197: If Flag = 1 Then ConvertBCFfield = "YRSHISPK" Else ConvertBCFfield = 234
'    Case 198: If Flag = 1 Then ConvertBCFfield = "YRSDAY" Else ConvertBCFfield = 436
'    Case 199: If Flag = 1 Then ConvertBCFfield = "YRSLOW" Else ConvertBCFfield = 373
'    Case Else: If Flag = 1 Then ConvertBCFfield = "NoLabel" Else ConvertBCFfield = -999
'  End Select
'End Function

Sub XLSImport(ImpFileName As String, DataSource As String, SourceURL As String)
  Dim XLRange As Excel.Range, XLApp As Excel.Application, _
      XLBook As Excel.Workbook
  Dim bumSheets$, stationID$, filename$, stateFIPS$
  Dim staCnt&, header&, lastRow&, firstCol&, fldCnt&, h&, i&, response&, lastCol&
  Dim value As String, staName As String
  Dim badFields As String, OverWriteInfo As String, badStats As String
  Dim stationValues() As String, dataValues() As String, impStates() As String
  Dim OutFile As Integer
  Dim myStation As ssStation
  Dim myStatistic As ssStatistic
  Dim ImportCol() As Long, nAtts As Long, attCnt As Long, StaAttCnt As Long
  Dim staIndex As Long, attIndex As Long
  Dim Src As New ssSource
  Dim YrsRecCol As Long
  Dim YearsOfRecord As Double
  
  On Error GoTo errTrapXLS
  
  IPC.SendMonitorMessage "(OPEN StreamStatsDB)"
  IPC.SendMonitorMessage "(BUTTOFF DETAILS)"
  IPC.SendMonitorMessage "(BUTTON CANCEL)"
  IPC.SendMonitorMessage "(BUTTON PAUSE)"
  IPC.SendMonitorMessage "(MSG1 Importing Excel Spreadsheet file)"
  IPC.SendMonitorMessage "(PROGRESS 0)"

  'add user-entered Datasource
  Set Src.Db = SSDB
  Src.Add DataSource, SourceURL

  Set XLApp = New Excel.Application
  Set XLBook = Excel.Workbooks.Open(ImpFileName, 0, True)
  
  For h = 1 To ActiveWorkbook.Sheets.Count
    With ActiveWorkbook
      Worksheets(h).Activate
    End With
    With ActiveWorkbook.ActiveSheet
      'Check for the key word "site_no"
      Set XLRange = Range(Cells(1, 1), Cells(20, 10))
      Set XLRange = XLRange.Find("site_no", , , , xlByRows, xlNext, True)
      If XLRange Is Nothing Then 'try STAID
        Set XLRange = Range(Cells(1, 1), Cells(20, 10))
        Set XLRange = XLRange.Find("STAID", , , , xlByRows, xlNext, False)
        If XLRange Is Nothing Then
          bumSheets = bumSheets & " " & .Name
          GoTo badSheetXLS
        End If
      End If
      'Establish boundaries of import file
      header = XLRange.row
      firstCol = XLRange.Column
      Set XLRange = Cells(XLRange.row, XLRange.Column)
      Set XLRange = XLRange.End(xlDown)
      lastRow = XLRange.row
      Set XLRange = Cells(header, XLRange.Column)
      Set XLRange = XLRange.End(xlToRight)
      lastCol = XLRange.Column
      'check column headers for valid attributes
      nAtts = 0
      YrsRecCol = 0
      ReDim ImportCol(lastCol)
      Set myStatistic = New ssStatistic
      For fldCnt = firstCol + 1 To lastCol
        Set myStatistic.Db = SSDB
        ImportCol(fldCnt) = myStatistic.GetLabelID(Cells(header, fldCnt))
        If ImportCol(fldCnt) > 0 Then
          nAtts = nAtts + 1
        ElseIf UCase(Cells(header, fldCnt)) = "YEARS" Or _
               UCase(Cells(header, fldCnt)) = "YEARSREC" Then
          'apply years of record value from this column to all other stats
          YrsRecCol = fldCnt
        Else 'note fields that can't be imported
          badFields = badFields & Cells(header, fldCnt) & vbCrLf
        End If
      Next fldCnt
      'Loop thru stations in spreadsheet
      For staCnt = header + 1 To lastRow
        IPC.SendMonitorMessage "(PROGRESS " & _
            (staCnt - header) * 100 / (lastRow - header) & ")"
        Select Case NextPipeCharacter(IPC.hPipeReadFromProcess("Status"))
          Case "P"
            While NextPipeCharacter(IPC.hPipeReadFromProcess("Status")) <> "R"
              DoEvents
            Wend
          Case "C"
              MsgBox "The NWIS import operation has been interrupted." & _
                  vbCrLf & vbCrLf & "Data has been imported up to station " & _
                  stationID & ".", , "Import interrupted"
            Err.Raise 999
        End Select
        ReDim stationValues(2 To 2, 1 To 1, 1 To UBound(StationFields))
        ReDim dataValues(2 To 2, 1 To nAtts, 1 To UBound(DataFields))
        YearsOfRecord = 0
        If YrsRecCol > 0 Then 'read years of record for this station
          YearsOfRecord = Cells(staCnt, YrsRecCol)
        End If
        'Loop thru column fields in spreadsheet
        StaAttCnt = 0
        attCnt = 0
        For fldCnt = firstCol To lastCol
          value = Cells(staCnt, fldCnt)
          If fldCnt = firstCol Then 'station id field
            If Len(value) < 8 Or Len(value) > 15 Then
              MsgBox "The site_no (Station ID) on row " & staCnt & _
                  " must be 8-15 digits long." & vbCrLf & _
                  "No data will be imported for this station."
              Exit For
            End If
            stationValues(2, 1, 1) = value
            stationID = value
          ElseIf ImportCol(fldCnt) > 0 And Len(value) > 0 Then
            If ImportCol(fldCnt) <= 14 Or ImportCol(fldCnt) = 24 Then 'station data
              StaAttCnt = StaAttCnt + 1
              If ImportCol(fldCnt) = 24 Then 'map DistrictCode to 15th element in array
                stationValues(2, 1, 15) = value
              Else
                stationValues(2, 1, ImportCol(fldCnt)) = value
              End If
            Else 'statistic data
              attCnt = attCnt + 1
              dataValues(2, attCnt, 2) = CStr(ImportCol(fldCnt))
              dataValues(2, attCnt, 4) = value
              If YearsOfRecord > 0 Then
                dataValues(2, attCnt, 6) = YearsOfRecord
              End If
              dataValues(2, attCnt, 7) = DataSource
              dataValues(2, attCnt, 8) = SourceURL
            End If
          End If
        Next fldCnt
        'Check whether station/statistics exist - add if not, write to text file if so
        staIndex = SSDB.state.Stations.IndexFromKey(stationValues(2, 1, 1))
        If staIndex = -1 Then 'Station does not exist - add it and its statistics
          Set myStation = New ssStation
          Set myStation.Db = SSDB
          Set myStation.state = SSDB.state
          myStation.Add stationValues(), 1, 1
          myStation.id = stationValues(2, 1, 1)
          For i = 1 To attCnt
            'Retrieve the StatLabel based upon the BCF stat code
            Set myStatistic = New ssStatistic
            With myStatistic
              Set .Db = SSDB
              Set .station = myStation
              .Add dataValues(), i, 1
            End With
          Next i
        ElseIf staIndex > 0 Then 'Station exists - check whether its stats also exist
          Set myStation = SSDB.state.Stations(staIndex)
          If StaAttCnt > 0 Then 'station data to update
            myStation.Edit stationValues(), 1
          End If
          For i = 1 To attCnt
            attIndex = myStation.Statistics.IndexFromKey(dataValues(2, i, 2))
            If attIndex = -1 Then 'Statistic does not exist
              Set myStatistic = New ssStatistic
              With myStatistic
                Set .Db = SSDB
                Set .station = myStation
                .Add dataValues(), i, 1
              End With
            ElseIf attIndex > 0 Then 'Statistic already exists - write to text file
              If response <> 2 And response <> 4 Then 'user hasn't chosen to replace or keep all
                response = myMsgBox.Show("For station " & myStation.Name & " the statistic " & myStation.Statistics(attIndex).Name & _
                                         " already exists." & vbCrLf & "Existing value: " & myStation.Statistics(attIndex).value & _
                                         "  New value: " & dataValues(2, i, 4) & vbCrLf & "What do you want to do?", "BCF Import - Data Exists", _
                                         "&Replace", "Replace &All", "+&Keep", "K&eep All", "-&Cancel")
              End If
              If response = 5 Then Err.Raise 999
              If response > 2 Then 'keeping existing, note it to file
                badStats = badStats & vbCrLf & vbTab & _
                           Left(myStation.id & "       ", 15) & vbTab & _
                           StrPad(dataValues(2, i, 2), 7) & vbTab & _
                           StrPad(myStation.Statistics(attIndex).value, 10) & _
                           vbTab & StrPad(dataValues(2, i, 4), 8)
              Else 'overwrite value
                OverWriteInfo = OverWriteInfo & vbCrLf & vbTab & _
                                Left(myStation.id & "       ", 15) & vbTab & _
                                StrPad(dataValues(2, i, 2), 7) & vbTab & _
                                StrPad(myStation.Statistics(attIndex).value, 10) & _
                                vbTab & StrPad(dataValues(2, i, 4), 8)
                myStation.Statistics(attIndex).Edit dataValues(), i
              End If
            End If
          Next i
        End If
      Next staCnt
badSheetXLS:
    End With
  Next h
errTrapXLS:
  If Err.Number = 999 Then
    IPC.SendMonitorMessage "(MSG1 User Canceled Import)"
  ElseIf Err.Number > 0 Then
    MsgBox "Problem importing:  " & Err.Description, , "Excel Import Problem"
  End If
  IPC.SendMonitorMessage "(CLOSE)"
  If Len(Trim(bumSheets)) > 0 Then
    myMsgBox.Show "The keyword 'site_no' was not found in the spreadsheet(s)" & _
        bumSheets & "." & vbCrLf & vbCrLf & "'site_no' must appear in the top " & _
        "left corner of the data block." & vbCrLf & "No data was imported from " & _
        "this spreadsheet(s).", "Keyword not found", "+-&OK"
  End If
  'Close EXCEL
  XLBook.Close False
  Set XLBook = Nothing
  XLApp.Quit
  Set XLApp = Nothing

  If Len(Trim(OverWriteInfo)) > 0 Or Len(Trim(badStats)) > 0 Or Len(Trim(badFields)) > 0 Then
    OutFile = FreeFile
    Open FilenameNoExt(ImpFileName) & "_Import.txt" For Output As OutFile
    Print #OutFile, "This file contains summary information regarding the"
    Print #OutFile, "import of the Excel Spreadsheet file " & ImpFileName
    Print #OutFile, "into the StreamStats database for " & SSDB.state.Name
    If Len(Trim(badFields)) > 0 Then
      Print #OutFile, ""
      Print #OutFile, ""
      Print #OutFile, ""
      Print #OutFile, "The following columns in the spreadsheet were not imported"
      Print #OutFile, "because they contained unrecognized statistic names:"
      Print #OutFile, badFields
      Print #OutFile, "(To make a statistic available for import, it must be added"
      Print #OutFile, " to the list of available statistics through the Statistic"
      Print #OutFile, " Management tab of StreamStatsDB)"
    End If
    If Len(Trim(OverWriteInfo)) > 0 Then
      Print #OutFile, ""
      Print #OutFile, ""
      Print #OutFile, "The following records show statistic values that were"
      Print #OutFile, "updated from the import:"
      Print #OutFile, ""
      Print #OutFile, vbTab & "               " & vbTab & "       " & vbTab & "Previous" & vbTab & " Updated"
      Print #OutFile, vbTab & "Station ID     " & vbTab & "Stat ID" & vbTab & "   Value" & vbTab & "   Value"
      Print #OutFile, OverWriteInfo
    End If
    If Len(Trim(badStats)) > 0 Then
      Print #OutFile, ""
      Print #OutFile, ""
      Print #OutFile, ""
      Print #OutFile, "The following records show statistic values that were"
      Print #OutFile, "not updated from the import (i.e. the new value was ignored):"
      Print #OutFile, ""
      Print #OutFile, vbTab & "               " & vbTab & "       " & vbTab & "Existing" & vbTab & " Ignored"
      Print #OutFile, vbTab & "Station ID     " & vbTab & "Stat ID" & vbTab & "   Value" & vbTab & "   Value"
      Print #OutFile, badStats
    End If
    MsgBox "Summary information about this import was saved in the file " & _
           FilenameNoExt(ImpFileName) & "_Import.txt"
    Close OutFile
  End If
End Sub
