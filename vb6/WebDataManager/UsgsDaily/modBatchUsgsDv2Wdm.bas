Attribute VB_Name = "modBatchUsgsDv2Wdm"
Option Explicit

Public Function ConvertUsgsDv2Wdm(myWDMFileName As String, _
                                myDownloadFiles As Collection) As ATCclsTserFile
  Dim myWDM As ATCclsTserFile
  Dim myWDMExt As clsTSerWDM
  Dim myMem As ATCclsTserFile
  Dim myRegistry As ATCoRegistry
  Dim myScriptLocation As String
  Dim myDownloadFileName As String
  Dim vDownloadFileName As Variant
  Dim p As String
  Dim s As String
  Dim myMsg As ATCoMessage
    
  Set myWDM = New clsTSerWDM
  Set myWDMExt = myWDM
  Set ConvertUsgsDv2Wdm = myWDM
  
  Set myRegistry = New ATCoRegistry
  Set myMsg = New ATCoMessage

  Dim AllAttributes As Collection

  myRegistry.AppName = "GenScn"
'  New version of clsTSerWDM no longer requires us to set MsgUnit from here
'  p = GetSetting("HASS_ENT", "files", "HSPFMsgWDM", "c:\vbexpermental\genscn\bin\hspfmsg.wdm")
'  myWDMExt.MsgUnit = F90_WDBOPN(CLng(0), p, Len(p))

  If FileExists(myWDMFileName) Then
    myWDM.Filename = myWDMFileName
  Else
    myWDM.SaveAs myWDMFileName
    If LCase(myWDM.Filename) <> LCase(myWDMFileName) Then
      MsgBox "Could not create WDM file " & myWDMFileName, vbOKOnly, "USGS Daily Values"
      
      Exit Function '!!!!!!!!!!!!!!!!!!!!
    
    End If
  End If
  
  myMsg.ShowNoWait "", "Writing to " & myWDMFileName, "Hide"

  'use a registry key for the WDMUtil bin folder (scripts just below)
  p = myRegistry.RegGetString(HKEY_LOCAL_MACHINE, "SOFTWARE\AQUA TERRA Consultants\WDMUtil\ExePath", "")
  If Len(p) = 0 Then p = "c:\vbexperimental\wdmutil"
  myScriptLocation = p & "\scripts\USGSdvBASINSdownload.ws"
  
  Set AllAttributes = myWDM.AvailableAttributes
  
  For Each vDownloadFileName In myDownloadFiles
    myDownloadFileName = vDownloadFileName
    'Set myMsg.Icon = frmShapeAttributes.Icon
    myMsg.ShowNoWait "Reading downloaded file: " & vDownloadFileName & vbCr _
                   & " (This may take some time)", _
                     "Writing to " & myWDMFileName, "Hide"
    Set myMem = Nothing
    Set myMem = New clsTSerMemory
    myMem.Filename = myDownloadFileName & "|" & myScriptLocation
    If myMem.DataCount > 0 Then
      myWDM.AddTimSer myMem.Data(1), TsIdRenum
      With myWDM.Data(myWDM.DataCount)
        Debug.Print myDownloadFileName, myWDM.DataCount, .header.ID, .Max
      End With
    End If
  Next
  Set myMem = Nothing
  myMsg.Hide
End Function

Sub WriteShapeFile(dataFile As ATCclsTserFile, ByVal baseFilename As String, XattribName As String, YattribName As String, VattribNames As Variant)
  Dim shpfile%, shxfile%
  Dim RecordNumber As Long, FieldNumber As Long, FieldNumberCheck As Long
  Dim foundDuplicate As Long
  Dim FileLength As Long, ShapeType As Long
  Dim ThisFieldLength As Long
  Dim X#, Y#, lowX#, lowY#, uppX#, uppY#
  Dim dbf As clsDBF
  Dim strVal As String
  Dim tryNumber As Long
  Dim attribNames() As String
  attribNames = VattribNames
  
  ShapeType = 1 'Point
  
  baseFilename = FilenameNoExt(baseFilename)
  If Len(FilenameOnly(baseFilename)) > 7 Then
    MsgBox "Warning: Base name of shape file exceeds seven characters." & vbCr _
         & "You may need to shorten it later for the shp, shx, and dbf shape files.", _
         vbOKOnly, "Shape file '" & FilenameOnly(baseFilename) & "'"
  End If
  
  'Shape data (.dbf)
  Set dbf = New clsDBF
  dbf.NumRecords = dataFile.DataCount
  dbf.NumFields = UBound(attribNames)
  
  For FieldNumber = 1 To dbf.NumFields
    dbf.FieldType(FieldNumber) = "N" 'Default to numeric fields, change to character later if needed
    dbf.FieldName(FieldNumber) = attribNames(FieldNumber)
    dbf.FieldLength(FieldNumber) = 1
    dbf.FieldDecimalCount(FieldNumber) = 0
    
    tryNumber = 1
    foundDuplicate = 0
CheckForDuplicateFieldname:
    strVal = UCase(dbf.FieldName(FieldNumber))
    For FieldNumberCheck = 1 To FieldNumber - 1
      If strVal = UCase(dbf.FieldName(FieldNumberCheck)) Then
        foundDuplicate = FieldNumberCheck
        tryNumber = tryNumber + 1
        dbf.FieldName(FieldNumber) = Left(attribNames(FieldNumber), 8) & tryNumber
        GoTo CheckForDuplicateFieldname
      End If
    Next
    If foundDuplicate > 0 Then
      MsgBox "Field #" & foundDuplicate & " '" & attribNames(foundDuplicate) & " and" & vbCr _
           & "Field #" & FieldNumber & " '" & attribNames(FieldNumber) & vbCr _
           & "both shortened to '" & dbf.FieldName(foundDuplicate) & "' in the DBF of the shape file." & vbCr _
           & "Field #" & FieldNumber & " has been renamed to '" & dbf.FieldName(FieldNumber) & "'", vbOKOnly, "WriteShapeFile"
    End If
  Next
  
  For RecordNumber = 1 To dbf.NumRecords
    For FieldNumber = 1 To dbf.NumFields
      strVal = dataFile.Data(RecordNumber).Attrib(attribNames(FieldNumber))
      ThisFieldLength = Len(strVal)
      If ThisFieldLength > dbf.FieldLength(FieldNumber) Then
        dbf.FieldLength(FieldNumber) = ThisFieldLength
      End If
      If ThisFieldLength > 0 Then
        If dbf.FieldType(FieldNumber) = "N" Then
          If ThisFieldLength > 1 And Left(strVal, 1) = "0" Then
            dbf.FieldType(FieldNumber) = "C" 'Leading zeroes usually mean this is an ID number of some sort rather than a numeric value
          ElseIf Not IsNumeric(strVal) Then
            dbf.FieldType(FieldNumber) = "C" 'If there is a non-numeric value in this field, make it type Character
          End If
        End If
      End If
    Next
  Next
  
'  For FieldNumber = 1 To dbf.NumFields
'    Debug.Print FieldNumber & ":" & dbf.FieldType(FieldNumber) & "," & dbf.FieldName(FieldNumber) & "," & dbf.FieldLength(FieldNumber) & " attribName = " & attribNames(FieldNumber)
'  Next
  
  dbf.InitData
  
  CreateNewShapeFile baseFilename, ShapeType
  'Point Shape Main file (.shp)
  shpfile = FreeFile(0)
  Open baseFilename & ".shp" For Binary Access Read Write As shpfile
  Call ReadShapeHeader(shpfile, FileLength, ShapeType, lowX, lowY, uppX, uppY)
  Seek #shpfile, 101
  
  'Point Shape Index file (.shx)
  shxfile = FreeFile(0)
  Open baseFilename & ".shx" For Binary Access Read Write As shxfile
  Seek #shxfile, 101
  
  RecordNumber = 0
  For RecordNumber = 1 To dbf.NumRecords
    With dataFile.Data(RecordNumber)
      dbf.CurrentRecord = RecordNumber
      For FieldNumber = 1 To dbf.NumFields
        'If dbf.FieldType(FieldNumber) = "C" Then
        dbf.Value(FieldNumber) = .Attrib(attribNames(FieldNumber))
        If Trim(dbf.Value(FieldNumber)) <> Trim(.Attrib(attribNames(FieldNumber))) Then Stop
      Next
      X = .AttribNumeric(XattribName)
      Y = .AttribNumeric(YattribName)
    End With
    Call WriteShapePointAll(shpfile, RecordNumber, X, Y)
    Call WriteShapePointIndex(shxfile, RecordNumber)
    
    FileLength = FileLength + 14
    If X > uppX Then uppX = X Else If X < lowX Then lowX = X
    If Y > uppY Then uppY = Y Else If Y < lowY Then lowY = Y
  Next
  dbf.WriteDBF baseFilename & ".dbf"
    
  'Debug.Print "Wrote " & dbf.NumRecords & " records to DBF and " & (FileLength - 50) / 14 & " records to SHP"
  
  Call WriteShapeHeader(shpfile, FileLength, ShapeType, lowX - 0.01, lowY - 0.01, uppX + 0.01, uppY + 0.01)
  Close shpfile
  
  FileLength = 50 + dbf.NumRecords * 4
  Call WriteShapeHeader(shxfile, FileLength, ShapeType, lowX - 0.01, lowY - 0.01, uppX + 0.01, uppY + 0.01)
  Close shxfile

End Sub


