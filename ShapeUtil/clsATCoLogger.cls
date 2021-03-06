VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsATCoLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'##MODULE_REMARKS Copyright 2001-3AQUA TERRA Consultants - Royalty-free use permitted under open source license

'##MODULE_DESCRIPTION Class for logging messages to a variety of targets (File, Status Monitor, VB Debugger, Message Box)

Dim pFileName As String
Dim pLogFile As Long
Dim pLog2Debug As Boolean
Dim pDateTime As Boolean

Private Const cMaxLogFiles As Long = 9

Public Property Let Log2Debug(ByVal newValue As Boolean)
  pLog2Debug = newValue
End Property

Public Property Let DateTime(ByVal newValue As Boolean)
  pDateTime = newValue
End Property

Public Sub SetFileName(FileName As String, Optional OverWrite As Boolean = True)
  Dim LogNumber As Long
  Dim moveFromFilename As String
  Dim moveToFilename As String

  On Error GoTo ErrHand

  If FileName <> pFileName Then
    If Len(pFileName) > 0 Then Close pLogFile: pLogFile = 0
    pFileName = FileName
    
    If Len(pFileName) > 0 Then
      If Not OverWrite And FileExists(pFileName) Then 'don't overwrite existing file, rename
        LogNumber = 1
        While FileExists(FilenameNoExt(pFileName) & LogNumber & "." & FileExt(pFileName)) And _
              LogNumber <= cMaxLogFiles
          LogNumber = LogNumber + 1
        Wend
        If LogNumber > cMaxLogFiles Then 'delete oldest log and shift others down a notch
          moveFromFilename = FilenameNoExt(pFileName) & "1." & FileExt(pFileName)
          If FileExists(moveFromFilename) Then Kill moveFromFilename
          For LogNumber = 2 To cMaxLogFiles
            moveToFilename = moveFromFilename
            moveFromFilename = FilenameNoExt(pFileName) & LogNumber & "." & FileExt(pFileName)
            If FileExists(moveFromFilename) Then
              Name moveFromFilename As moveToFilename
            End If
          Next
          LogNumber = cMaxLogFiles
        End If
        pFileName = FilenameNoExt(pFileName) & LogNumber & "." & FileExt(pFileName)
      ElseIf FileExists(pFileName) Then
        Kill pFileName
      Else 'File doesn't exist, maybe directory doesn't either
        MkDirPath PathNameOnly(pFileName)
      End If
      pLogFile = FreeFile(0)
      Open pFileName For Output As pLogFile
    End If
  End If
  
  Exit Sub
  
ErrHand:
  Log "clsATCoLogger: Error in SetFileName: " & Err.Description
  Resume Next
End Sub

Public Function GetFileName() As String
  Close pLogFile
  GetFileName = pFileName
  Open pFileName For Append As pLogFile
End Function

Public Function LogMsg(Message As String, title As String) As Long
  MsgBox Message, , title
  LogMsg = 1
End Function

'sends 'message' to desired logging targets
'Specs allows for conditions to be specified as follows:
'  If Specs contains - "F", send to file
'                      "S", send to status reporter
'                      "D", send to debugger
'  If Specs is empty then message is sent to targets that are "turned on"
Public Sub Log(Message As String, Optional Specs As String = "")
  Static lastMessage As String
  Dim lmessage As String
  Dim lLog2File As Boolean, lLog2Status As Boolean
  Dim lLog2Debug As Boolean

  If pDateTime Then
    lmessage = Format(Now, "yyyy/mm/dd hh:mm:ss  ") & Message
  Else
    lmessage = Message
  End If

  If lmessage <> lastMessage Then
    'assume local spec flags are "on"
    lLog2File = True
    lLog2Status = True
    lLog2Debug = True
    If Len(Specs) > 0 Then 'process specs, set local flags
      If InStr(UCase(Specs), "F") = 0 Then lLog2File = False
      If InStr(UCase(Specs), "S") = 0 Then lLog2Status = False
      If InStr(UCase(Specs), "D") = 0 Then lLog2Debug = False
    End If
  
    If Len(pFileName) > 0 And lLog2File Then
  '    AppendFileString pFileName, lmessage & vbCrLf
      Print #pLogFile, lmessage
    End If
        
    If pLog2Debug And lLog2Debug Then
      Debug.Print lmessage
    End If
    lastMessage = lmessage
  End If
End Sub

Public Function CurrentLog() As String
  If Len(pFileName) > 0 Then
    If FileExists(pFileName) Then
      Close pLogFile
      CurrentLog = WholeFileString(pFileName)
      Open pFileName For Append As pLogFile
    End If
  End If
End Function

Private Sub Class_Initialize()
  pLogFile = 0
  pFileName = ""
End Sub

Private Sub Class_Terminate()
  If pLogFile > 0 Then 'need to close log file
    Close #pLogFile
  End If
End Sub


