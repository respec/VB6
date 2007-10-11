VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pfqProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPFQExeFileName As String
Private pSpecFileName As String
Private pDataFileName As String
Private pDataType As Long '0 - ASCII(Watstore), 1 - WDM
Private pStations As FastCollection 'of class pfqStation
Private pOutFile As String
Private pAdditionalOutput As Long
Private pAddOutFileName As String
Private pIntermediateResults As Boolean
Private pLinePrinter As Boolean
Private pGraphic As Boolean
Private pGraphFormat As String
Private pPrintPlotPos As Boolean
Private pPlotPos As Single
Private pConfidenceLimits As Single
Private pInputDir As String
Private pOutputDir As String
Private pEMA As Boolean
'the following are for storing comments for various specification records
Private pCDataFile As String
Private pCOutFile As String
Private pCPlotStyle As String
Private pCPlotFormat As String
Private pCPrintPlotPos As String
Private pCPlotPos As String
Private pCAdditional As String
Private pCIntermediate As String
Private pCConfidenceLimits As String
Private pCEMA As String

Private FType(1) As String

Public Property Get PFQExeFileName() As String
  PFQExeFileName = pPFQExeFileName
End Property
Public Property Let PFQExeFileName(ByVal newValue As String)
  pPFQExeFileName = newValue
End Property

Public Property Get SpecFileName() As String
  SpecFileName = pSpecFileName
End Property
Public Property Let SpecFileName(ByVal newValue As String)
  Dim s As String, LastFile As String, CurFile As String
  Dim FileIsStatic As Boolean

  pSpecFileName = newValue
'  ReadSpecFile
'  If pStations.Count = 0 Then 'may be defaulted to do all stations
  s = WholeFileString(pSpecFileName)
  If UCase(Left(s, 7)) <> "VERBOSE" Then
    'update spec file to verbose mode that explicitly defines all station specs
    If InStr(Right(s, 2), vbLf) > 0 Then 'already have line feed
      s = s & "Update"
    Else 'include CR/LF
      s = s & vbCrLf & "Update"
    End If
    SaveFileString pSpecFileName, s
    ReadSpecFile 'populate what we can now (at least get output file)
    RunBatchModel
    
  End If
  ReadSpecFile
End Property

Public Property Get DataFileName() As String
  DataFileName = pDataFileName
End Property
Public Property Let DataFileName(ByVal newValue As String)
  pDataFileName = newValue
End Property

Public Property Get DataType() As Long
  DataType = pDataType
End Property
Public Property Let DataType(ByVal newValue As Long)
  pDataType = newValue
End Property

Public Property Set Stations(ByRef newValue As FastCollection)
  Set pStations = newValue
End Property
Public Property Get Stations() As FastCollection
  If pStations Is Nothing Then Set pStations = New FastCollection
  Set Stations = pStations
End Property

Public Property Get OutFile() As String
  OutFile = pOutFile
End Property
Public Property Let OutFile(ByVal newValue As String)
  pOutFile = newValue
End Property

Public Property Get AdditionalOutput() As Long
  AdditionalOutput = pAdditionalOutput
End Property
Public Property Let AdditionalOutput(ByVal newValue As Long)
  pAdditionalOutput = newValue
End Property

Public Property Get AddOutFileName() As String
  AddOutFileName = pAddOutFileName
End Property
Public Property Let AddOutFileName(ByVal newValue As String)
  pAddOutFileName = newValue
End Property

Public Property Get IntermediateResults() As Boolean
  IntermediateResults = pIntermediateResults
End Property
Public Property Let IntermediateResults(ByVal newValue As Boolean)
  pIntermediateResults = newValue
End Property

Public Property Get ConfidenceLimits() As Single
  ConfidenceLimits = pConfidenceLimits
End Property
Public Property Let ConfidenceLimits(ByVal newValue As Single)
  pConfidenceLimits = newValue
End Property

Public Property Get LinePrinter() As Boolean
  LinePrinter = pLinePrinter
End Property
Public Property Let LinePrinter(ByVal newValue As Boolean)
  pLinePrinter = newValue
End Property

Public Property Get Graphic() As Boolean
  Graphic = pGraphic
End Property
Public Property Let Graphic(ByVal newValue As Boolean)
  pGraphic = newValue
End Property

Public Property Get GraphFormat() As String
  GraphFormat = pGraphFormat
End Property
Public Property Let GraphFormat(ByVal newValue As String)
  pGraphFormat = newValue
End Property

Public Property Get PlotPos() As Single
  PlotPos = pPlotPos
End Property
Public Property Let PlotPos(ByVal newValue As Single)
  pPlotPos = newValue
End Property

Public Property Get PrintPlotPos() As Boolean
  PrintPlotPos = pPrintPlotPos
End Property
Public Property Let PrintPlotPos(ByVal newValue As Boolean)
  pPrintPlotPos = newValue
End Property

Public Property Get InputDir() As String
  InputDir = pInputDir
End Property
Public Property Let InputDir(ByVal newValue As String)
  pInputDir = newValue
End Property

Public Property Get OutputDir() As String
  OutputDir = pOutputDir
End Property
Public Property Let OutputDir(ByVal newValue As String)
  pOutputDir = newValue
End Property

Public Property Get EMA() As Boolean
  EMA = pEMA
End Property
Public Property Let EMA(ByVal newValue As Boolean)
  pEMA = newValue
End Property

Public Property Get CDataFile() As String
  CDataFile = pCDataFile
End Property
Public Property Let CDataFile(ByVal newValue As String)
  pCDataFile = newValue
End Property

Public Property Get COutFile() As String
  COutFile = pCOutFile
End Property
Public Property Let COutFile(ByVal newValue As String)
  pCOutFile = newValue
End Property

Public Property Get CPlotStyle() As String
  CPlotStyle = pCPlotStyle
End Property
Public Property Let CPlotStyle(ByVal newValue As String)
  pCPlotStyle = newValue
End Property

Public Property Get CPlotFormat() As String
  CPlotFormat = pCPlotFormat
End Property
Public Property Let CPlotFormat(ByVal newValue As String)
  pCPlotFormat = newValue
End Property

Public Property Get CPrintPlotPos() As String
  CPrintPlotPos = pCPrintPlotPos
End Property
Public Property Let CPrintPlotPos(ByVal newValue As String)
  pCPrintPlotPos = newValue
End Property

Public Property Get CPlotPos() As String
  CPlotPos = pCPlotPos
End Property
Public Property Let CPlotPos(ByVal newValue As String)
  pCPlotPos = newValue
End Property

Public Property Get CAdditional() As String
  CAdditional = pCAdditional
End Property
Public Property Let CAdditional(ByVal newValue As String)
  pCAdditional = newValue
End Property

Public Property Get CIntermediate() As String
  CIntermediate = pCIntermediate
End Property
Public Property Let CIntermediate(ByVal newValue As String)
  pCIntermediate = newValue
End Property

Public Property Get CConfidenceLimits() As String
  CConfidenceLimits = pCConfidenceLimits
End Property
Public Property Let CConfidenceLimits(ByVal newValue As String)
  pCConfidenceLimits = newValue
End Property

Public Property Get CEMA() As String
  CEMA = pCEMA
End Property
Public Property Let CEMA(ByVal newValue As String)
  pCEMA = newValue
End Property

Public Sub ReadSpecFile()

  Dim i As Integer
  Dim SpecFile As String
  Dim Rec As String, Kwd As String
  Dim lCom As String
  Dim CommentPending As Boolean
  Dim CurStation As pfqStation

  CommentPending = False
  Set pStations = New FastCollection
  SpecFile = WholeFileString(pSpecFileName)

  While Len(SpecFile) > 0
    Rec = StrSplit(SpecFile, vbCrLf, "")
    If Left(Rec, 1) = "'" Then 'process comment
      If CommentPending Then 'multiple line comment
        lCom = lCom & vbCrLf & Rec
      Else 'new comment
        lCom = Rec
        CommentPending = True
      End If
    Else 'process specification
      Kwd = UCase(StrRetRem(Rec))
      Select Case Kwd
        Case "I"
          Kwd = UCase(StrRetRem(Rec))
          For i = 0 To 1
            If Kwd = FType(i) Then pDataType = i
          Next i
          pDataFileName = Rec
          If CommentPending Then pCDataFile = lCom
        Case "O"
          Kwd = UCase(StrRetRem(Rec))
          Select Case Kwd
            Case "FILE"
              pOutFile = Rec
              If CommentPending Then pCOutFile = lCom
            Case "PLOT"
              Kwd = UCase(StrRetRem(Rec))
              If Kwd = "STYLE" Then
                If UCase(Rec) = "PRINTER" Or UCase(Rec) = "BOTH" Then
                  pLinePrinter = True
                End If
                If UCase(Rec) = "GRAPHICS" Or UCase(Rec) = "BOTH" Then
                  pGraphic = True
                End If
                If CommentPending Then pCPlotStyle = lCom
              ElseIf Kwd = "FORMAT" Then
                pGraphFormat = UCase(Rec)
                If CommentPending Then pCPlotFormat = lCom
              ElseIf Kwd = "PRINTPOS" Then
                If UCase(Rec) = "YES" Then
                  pPrintPlotPos = True
                Else
                  pPrintPlotPos = False
                End If
                If CommentPending Then pCPrintPlotPos = lCom
              ElseIf Kwd = "POSITION" Then
                pPlotPos = CSng(Rec)
                If CommentPending Then pCPlotPos = lCom
              End If
            Case "ADDITIONAL"
              Kwd = UCase(StrRetRem(Rec))
              If Kwd = "WDM" Then
                pAdditionalOutput = 1
              ElseIf Left(Kwd, 3) = "WAT" Then
                pAdditionalOutput = 2
              ElseIf Kwd = "BOTH" Then
                pAdditionalOutput = 3
              Else
                pAdditionalOutput = 0
              End If
              If pAdditionalOutput >= 2 Then
                'remaining text should be file name
                pAddOutFileName = Rec
              End If
              If CommentPending Then pCAdditional = lCom
            Case "DEBUG"
              If UCase(Rec) = "YES" Then
                pIntermediateResults = True
              Else
                pIntermediateResults = False
              End If
              If CommentPending Then pCIntermediate = lCom
            Case "CONFIDENCE"
              pConfidenceLimits = CSng(Rec)
              If CommentPending Then pCConfidenceLimits = lCom
            Case "EMA"
              If UCase(Rec) = "YES" Then
                pEMA = True
              Else
                pEMA = False
              End If
              If CommentPending Then pCEMA = lCom
          End Select
        Case "STATION"
          If Not CurStation Is Nothing Then
            'previous station info exists, add it to collection
            pStations.Add CurStation
          End If
          'build new station
          Set CurStation = New pfqStation
          CurStation.id = Rec
          If CommentPending Then CurStation.Comment = lCom
        Case "GENSKEW"
          CurStation.GenSkew = CSng(Rec)
          If CommentPending Then CurStation.CGenSkew = lCom
        Case "SKEWSE"
          CurStation.SESkew = CSng(Rec)
          If CommentPending Then CurStation.CSESkew = lCom
        Case "BEGYEAR"
          CurStation.BegYear = CLng(Rec)
          If CommentPending Then CurStation.CBegYear = lCom
        Case "ENDYEAR"
          CurStation.EndYear = CLng(Rec)
          If CommentPending Then CurStation.CEndYear = lCom
        Case "HISTPERIOD"
          CurStation.HistoricPeriod = CSng(Rec)
          If CommentPending Then CurStation.CHistoric = lCom
        Case "SKEWOPT"
          If UCase(Rec) = "STATION" Then
            CurStation.SkewOpt = -1
          ElseIf UCase(Rec) = "WEIGHTED" Then
            CurStation.SkewOpt = 0
          ElseIf UCase(Rec) = "GENERALIZED" Then
            CurStation.SkewOpt = 1
          End If
          If CommentPending Then CurStation.CSkewOpt = lCom
        Case "URB/REG"
          If UCase(Rec) = "YES" Then
            CurStation.UrbanRegPeaks = 1
          Else
            CurStation.UrbanRegPeaks = 0
          End If
          If CommentPending Then CurStation.CUrban = lCom
        Case "LOTHRESH"
          CurStation.LowOutlier = CSng(Rec)
          If CommentPending Then CurStation.CLowOutlier = lCom
        Case "HITHRESH"
          CurStation.HighOutlier = CSng(Rec)
          If CommentPending Then CurStation.CHighOutlier = lCom
        Case "GAGEBASE"
          CurStation.GageBaseDischarge = CSng(Rec)
          If CommentPending Then CurStation.CGageBase = lCom
        Case "LATITUDE"
          CurStation.Lat = CSng(Rec)
          If CommentPending Then CurStation.CLat = lCom
        Case "LONGITUDE"
          CurStation.Lng = CSng(Rec)
          If CommentPending Then CurStation.CLong = lCom
        Case "HISYS" 'no comment for HISYS as it is just info for the interface
          CurStation.HighSysPeak = CSng(Rec)
        Case "LOHIST" 'no comment for LOHIST as it is just info for the interface
          CurStation.LowHistPeak = CSng(Rec)
        Case "PLOTNAME"
          CurStation.PlotName = Rec
          If CommentPending Then CurStation.CPlotName = lCom
      End Select
      CommentPending = False 'assume any pending comment was stored with a specification
    End If
  Wend
  If Not CurStation Is Nothing Then
    'station info exists, add it to collection
    pStations.Add CurStation
    Set CurStation = Nothing
  End If
End Sub

Public Function SaveAsString(Optional DefPrj As pfqProject = Nothing) As String

  Dim i As Long
  Dim s As String
  Dim vSta As Variant
  Dim defsta As New pfqStation

  s = ""
  If DefPrj Is Nothing Then 'no default specs, write out verbose
    s = "Verbose" & vbCrLf
  End If
  If Len(pCDataFile) > 0 Then s = s & pCDataFile & vbCrLf
  s = s & "I " & FType(pDataType) & " " & pInputDir & "\" & FilenameNoPath(pDataFileName) & vbCrLf
  If Len(pCOutFile) > 0 Then s = s & pCOutFile & vbCrLf
  s = s & "O File " & pOutputDir & "\" & FilenameNoPath(pOutFile) & vbCrLf
  If Len(pCPlotStyle) > 0 Then s = s & pCPlotStyle & vbCrLf
  If pLinePrinter And pGraphic Then
    s = s & "O Plot Style Both" & vbCrLf
  ElseIf pLinePrinter Then
    s = s & "O Plot Style Printer" & vbCrLf
  ElseIf pGraphic Then
    s = s & "O Plot Style Graphics" & vbCrLf
  End If
  If Len(pCPlotFormat) > 0 Then s = s & pCPlotFormat & vbCrLf
  If pGraphic Then s = s & "O Plot Format " & pGraphFormat & vbCrLf
  If Len(pCPrintPlotPos) > 0 Then s = s & pCPrintPlotPos & vbCrLf
  If pPrintPlotPos Then 'default, don't print it
'    S = S & "O Plot PrintPos YES" & vbCrLf
  Else
    s = s & "O Plot PrintPos NO" & vbCrLf
  End If
  If Len(pCPlotPos) > 0 Then s = s & pCPlotPos & vbCrLf
  If Abs(pPlotPos) > 0.000001 Then 'not using default of 0, print it
    s = s & "O Plot Position " & CStr(pPlotPos) & vbCrLf
  End If
  If Len(pCAdditional) > 0 Then s = s & pCAdditional & vbCrLf
  If pAdditionalOutput = 1 Then
    s = s & "O Additional WDM" & vbCrLf
  ElseIf pAdditionalOutput = 2 Then
    s = s & "O Additional Watstore " & pAddOutFileName & vbCrLf
  ElseIf pAdditionalOutput = 3 Then
    s = s & "O Additional Both " & pAddOutFileName & vbCrLf
  End If
  If Len(pCIntermediate) > 0 Then s = s & pCIntermediate & vbCrLf
  If pIntermediateResults Then s = s & "O Debug YES" & vbCrLf
  If Len(pCConfidenceLimits) > 0 Then s = s & pCConfidenceLimits & vbCrLf
  If Abs(pConfidenceLimits - 0.95) > 0.000001 Then 'not using .95, print it
    s = s & "O Confidence " & CStr(pConfidenceLimits) & vbCrLf
  End If
  If Len(pCEMA) > 0 Then s = s & pCEMA & vbCrLf
  If pEMA Then
    s = s & "O EMA YES" & vbCrLf
  End If
  i = 0
  For Each vSta In pStations
    i = i + 1
    If vSta.Active Then 'write station specs to string
      If DefPrj Is Nothing Then 'write out all station specs
        s = s & vSta.WriteSpecsVerbose
      Else 'only write out non-default station specs
        Set defsta = DefPrj.Stations(i)
        s = s & vSta.WriteSpecsNonDefault(defsta)
        Set defsta = Nothing
      End If
    End If
  Next
  SaveAsString = s

End Function

Public Sub BuildNewSpecFile(Optional tmpSpecName As String = "PKFQWPSF.TMP")
  'BuildNewSpecFile is called when a PeakFQ data file is opened
  'and a new spec file is needed for it to run PeakFQ.
  'Contains simplest set of specs: Input file, Output file, "Update" flag
  'It is given a temporary name (tmpSpecName) that will not be saved.
  Dim s As String

  pSpecFileName = tmpSpecName
  If UCase(Right(pDataFileName, 3)) = "WDM" Then 'WDM data file
    s = "I WDM " & FilenameNoPath(pDataFileName) & vbCrLf
  Else 'assume asci text data file
    s = "I ASCI " & FilenameNoPath(pDataFileName) & vbCrLf
  End If
  pOutFile = FilenameNoExt(pDataFileName) & ".prt" '".out"
  s = s & "O File " & FilenameNoPath(pOutFile) & vbCrLf
  s = s & "Update"
  SaveFileString pSpecFileName, s

End Sub

Public Sub RunBatchModel()
  Dim s As String
  Dim i As Long, oldlen As Long, curlen As Long

  On Error Resume Next
  If Len(pPFQExeFileName) > 0 Then
    If CurDir <> PathNameOnly(pPFQExeFileName) Then
      'copy support files for batch executable
      FileCopy PathNameOnly(pPFQExeFileName) & "\pkfqms.wdm", "pkfqms.wdm"
      FileCopy PathNameOnly(pPFQExeFileName) & "\interact.ini", "interact.ini"
    End If
    
    If FileExists(PfqPrj.OutFile) Then 'delete old output file
      Kill PfqPrj.OutFile
    End If
    
    gIPC.SendMonitorMessage "(Caption PKFQWin Status)"
    gIPC.SendMonitorMessage "Starting " & pPFQExeFileName
    If Not gIPC.StartProcess("PeakFQ", """" & pPFQExeFileName & """ " & FilenameNoPath(pSpecFileName), 10000, 20000) Then
      gIPC.SendMonitorMessage "(Open)"
      gIPC.SendMonitorMessage "(MSG1 Unable to start PeakFQ batch program.)"
    Else
      If Not FileExists(PfqPrj.OutFile) Then
        gIPC.SendMonitorMessage "(Open)"
        gIPC.SendMonitorMessage "(MSG1 Problem running PeakFQ batch program.)"
      End If
    End If
        
'    i = Shell(pPFQExeFileName & " " & FilenameNoPath(pSpecFileName) & " >PeakFQ.run")
'    oldlen = -1
'    curlen = 0
'    If i > 0 Then
'      'this while loop should be replaced with StatusMonitor
'      While oldlen <> curlen
'        If FileExists(PfqPrj.OutFile) Then
'          oldlen = FileLen(PfqPrj.OutFile)
'        Else
'          oldlen = 0
'        End If
'        Sleep 2000
'        If FileExists(PfqPrj.OutFile) Then
'          curlen = FileLen(PfqPrj.OutFile)
'        Else 'problem if still no output file
'          curlen = 0
'        End If
'      Wend
'    End If
'    If curlen = 0 Then
'      MsgBox "Problem running PeakFQ batch program." & vbCrLf & _
'             "Check PeakFQ.RUN file for details", vbExclamation, "PKFQWin"
'    End If
    If CurDir <> PathNameOnly(pPFQExeFileName) Then
      Kill "pkfqms.wdm"
      Kill "interact.ini"
    End If

'    this code works fine for first run,
'    but not when VERBOSE spec file already exists
'    s = WholeFileString(pSpecFileName)
'    While UCase(Left(s, 7)) <> "VERBOSE"
'      'when specfile has been written in Verbose mode the batch run has ended
'      Sleep 2000
'      s = WholeFileString(pSpecFileName)
'    Wend
  End If

End Sub

Public Function Copy() As pfqProject
  Dim oldStation As Variant
  Dim newStation As pfqStation
  Dim retval As New pfqProject
  With retval
    .AdditionalOutput = pAdditionalOutput
    .AddOutFileName = pAddOutFileName
    .ConfidenceLimits = pConfidenceLimits
    .DataFileName = pDataFileName
    .DataType = pDataType
    .Graphic = pGraphic
    .GraphFormat = pGraphFormat
    .IntermediateResults = pIntermediateResults
    .LinePrinter = pLinePrinter
    .OutFile = pOutFile
    .PFQExeFileName = pPFQExeFileName
    .PlotPos = pPlotPos
    .PrintPlotPos = pPrintPlotPos
    .InputDir = pInputDir
    .OutputDir = pOutputDir
    .EMA = pEMA
    .CDataFile = pCDataFile
    .COutFile = pCOutFile
    .CPlotStyle = pCPlotStyle
    .CPlotFormat = pCPlotFormat
    .CPrintPlotPos = pCPrintPlotPos
    .CPlotPos = pCPlotPos
    .CAdditional = pCAdditional
    .CIntermediate = pCIntermediate
    .CConfidenceLimits = pCConfidenceLimits
    .CEMA = pCEMA
    
'    .SpecFileName = pSpecFileName
    Set .Stations = Nothing
    For Each oldStation In Stations
      Set newStation = New pfqStation
      With newStation
        .Active = oldStation.Active
        .BegYear = oldStation.BegYear
        .EndYear = oldStation.EndYear
        .GageBaseDischarge = oldStation.GageBaseDischarge
        .GenSkew = oldStation.GenSkew
        .HighSysPeak = oldStation.HighSysPeak
        .HighOutlier = oldStation.HighOutlier
        .HistoricPeriod = oldStation.HistoricPeriod
        .id = oldStation.id
        .Lat = oldStation.Lat
        .Lng = oldStation.Lng
        .LowHistPeak = oldStation.LowHistPeak
        .LowOutlier = oldStation.LowOutlier
        .Name = oldStation.Name
        .PlotMade = oldStation.PlotMade
        .SESkew = oldStation.SESkew
        .SkewOpt = oldStation.SkewOpt
        .UrbanRegPeaks = oldStation.UrbanRegPeaks
        .PlotName = oldStation.PlotName
        .Comment = oldStation.Comment
        .CGenSkew = oldStation.CGenSkew
        .CSESkew = oldStation.CSESkew
        .CBegYear = oldStation.CBegYear
        .CEndYear = oldStation.CEndYear
        .CHistoric = oldStation.CHistoric
        .CSkewOpt = oldStation.CSkewOpt
        .CUrban = oldStation.CUrban
        .CLowOutlier = oldStation.CLowOutlier
        .CHighOutlier = oldStation.CHighOutlier
        .CGageBase = oldStation.CGageBase
        .CLat = oldStation.CLat
        .CLong = oldStation.CLong
        .CPlotName = oldStation.CPlotName
      End With
      .Stations.Add newStation
      Set newStation = Nothing
    Next
  End With
  Set Copy = retval
  Set retval = Nothing
End Function

Public Function SaveDefaults(FileStr As String) As pfqProject
  'Reads an existing spec file (contained in FileStr),
  'any station specfications found are assumed to be non-default
  'and are set to -999 so they will be written out if file is saved
  Dim prj As New pfqProject
  Dim i As Long
  Dim Rec As String, Kwd As String
  Dim lCom As String
  Dim CommentPending As Boolean
'  Dim CurStation As pfqStation

  CommentPending = False
  i = 0
  Set prj = Me.Copy
  With prj
    While Len(FileStr) > 0
      Rec = StrSplit(FileStr, vbCrLf, "")
      If Left(Rec, 1) = "'" Then 'process comment
        If CommentPending Then 'multiple line comment
          lCom = lCom & vbCrLf & Rec
        Else 'new comment
          lCom = Rec
          CommentPending = True
        End If
      Else 'process specification
        Kwd = UCase(StrRetRem(Rec))
        Select Case Kwd
'        Case "I"
'          Kwd = UCase(StrRetRem(Rec))
'          For i = 0 To 1
'            If Kwd = FType(i) Then .DataType = i
'          Next i
'          .DataFileName = Rec
'        Case "O"
'          Kwd = UCase(StrRetRem(Rec))
'          Select Case Kwd
'            Case "FILE"
'              .OutFile = Rec
'            Case "PLOT"
'              Kwd = UCase(StrRetRem(Rec))
'              If Kwd = "STYLE" Then
'                If UCase(Rec) = "PRINTER" Or UCase(Rec) = "BOTH" Then
'                  .LinePrinter = True
'                End If
'                If UCase(Rec) = "GRAPHICS" Or UCase(Rec) = "BOTH" Then
'                  .Graphic = True
'                End If
'              ElseIf Kwd = "PRINTPOS" Then
'                If UCase(Rec) = "YES" Then
'                  .PrintPlotPos = True
'                Else
'                  .PrintPlotPos = False
'                End If
'              ElseIf Kwd = "POSITION" Then
'                .PlotPos = CSng(Rec)
'              End If
'            Case "ADDITIONAL"
'              Kwd = UCase(StrRetRem(Rec))
'              If Kwd = "WDM" Then
'                .AdditionalOutput = 1
'              ElseIf Left(Kwd, 3) = "WAT" Then
'                .AdditionalOutput = 2
'              ElseIf Kwd = "BOTH" Then
'                .AdditionalOutput = 3
'              Else
'                .AdditionalOutput = 0
'              End If
'              If pAdditionalOutput >= 2 Then
'                'remaining text should be file name
'                .AddOutFileName = Rec
'              End If
'            Case "DEBUG"
'              If UCase(Rec) = "YES" Then
'                .IntermediateResults = True
'              Else
'                .IntermediateResults = False
'              End If
'            Case "CONFIDENCE"
'              .ConfidenceLimits = CSng(Rec)
'          End Select
          Case "STATION"
'          If Not CurStation Is Nothing Then
'            'previous station info exists, add it to collection
'            .Stations.Add CurStation
'          End If
'          'build new station
'          Set CurStation = New pfqStation
'          CurStation.id = Rec
            i = i + 1
            If CommentPending Then .Stations(i).Comment = lCom
          Case "GENSKEW"
            .Stations(i).GenSkew = -999#
            If CommentPending Then .Stations(i).CGenSkew = lCom
          Case "SKEWSE"
            .Stations(i).SESkew = -999#
            If CommentPending Then .Stations(i).CSESkew = lCom
          Case "BEGYEAR"
            .Stations(i).BegYear = -999#
            If CommentPending Then .Stations(i).CBegYear = lCom
          Case "ENDYEAR"
            .Stations(i).EndYear = -999#
            If CommentPending Then .Stations(i).CEndYear = lCom
          Case "HISTPERIOD"
            .Stations(i).HistoricPeriod = -999#
            If CommentPending Then .Stations(i).CHistoric = lCom
          Case "SKEWOPT"
            .Stations(i).SkewOpt = -999#
            If CommentPending Then .Stations(i).CSkewOpt = lCom
          Case "URB/REG"
            'set default to opposite of spec file so this spec will be written out
            If Me.Stations(i).UrbanRegPeaks Then
              .Stations(i).UrbanRegPeaks = False
            Else
              .Stations(i).UrbanRegPeaks = True
            End If
            If CommentPending Then .Stations(i).CUrban = lCom
          Case "LOTHRESH"
            .Stations(i).LowOutlier = -999#
            If CommentPending Then .Stations(i).CLowOutlier = lCom
          Case "HITHRESH"
            .Stations(i).HighOutlier = -999#
            If CommentPending Then .Stations(i).CHighOutlier = lCom
          Case "GAGEBASE"
            .Stations(i).GageBaseDischarge = -999#
            If CommentPending Then .Stations(i).CGageBase = lCom
          Case "LATITUDE"
            .Stations(i).Lat = -999#
            If CommentPending Then .Stations(i).CLat = lCom
          Case "LONGITUDE"
            .Stations(i).Lng = -999#
            If CommentPending Then .Stations(i).CLong = lCom
          Case "PLOTNAME"
            .Stations(i).PlotName = "-999"
            If CommentPending Then .Stations(i).CPlotName = lCom
        End Select
        CommentPending = False 'assume any pending comment was stored with a specification
      End If
    Wend
'    If Not CurStation Is Nothing Then
'      'station info exists, add it to collection
'      .Stations.Add CurStation
'      Set CurStation = Nothing
'    End If
  End With

  Set SaveDefaults = prj
  
End Function


Private Sub Class_Initialize()
  pPFQExeFileName = ""
  pSpecFileName = ""
  pDataFileName = ""
  Set pStations = New FastCollection
  pAdditionalOutput = 0
  pIntermediateResults = False
  pLinePrinter = False
  pGraphic = False
  pGraphFormat = ""
  pPrintPlotPos = True
  pPlotPos = 0#
  pConfidenceLimits = 0.95
  FType(0) = "ASCI"
  FType(1) = "WDM"
  pOutputDir = ""
  pEMA = False
End Sub
