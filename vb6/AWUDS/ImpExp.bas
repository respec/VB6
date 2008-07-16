Attribute VB_Name = "ImpExp"
Option Explicit
'Copyright 2003 by AQUA TERRA Consultants

' ##MODULE_NAME ImpExp
' ##MODULE_DATE December 8, 2003
' ##MODULE_AUTHOR Robert Dusenbury of AQUA TERRA CONSULTANTS
' ##MODULE_SUMMARY Imports data from and Exports data to Excel spreadsheets.
' ##MODULE_REMARKS <P>State data for counties, HUCs or aquifers can be exported to and _
          imported from specifically formatted Excel spreadsheets. This utility is useful _
          for editing large amounts of data at one time, as opposed to editing one _
          category at a time via the <EM>Interactive Data Input/Edit</EM> option on the _
          main interface.</P> _
          <P>The keyword 'Area' must appear in the upper left corner of the data table _
          with the unit area codes beneath in that column and the field abbreviations _
          (i.e., 'PS-GWPop') to the right in that row. Each cell containing a code or _
          abbreviation has a pop-up comment with the associated full name. Data _
          values are stored in the grid defined by the area codes and field names. When _
          data are exported, the fields are referenced against the <EM>Required</EM> field _
          in <EM>state</EM> database to determine which fields are required for that _
          particular state. All required fields are outlined in red in the resultant _
          spreadsheet.</P> _
          <P>When data are imported, QA checks are performed to ensure that the unit area _
          codes and field abbreviations listed on the spreadsheet actually exist. _
          Also, no negative values area allowed, except for HY-OfPow, which can be _
          negative. If any errors occur during import, they will be written to the _
          text file <EM>ImportError.txt</EM> in the <EM>AWUDSReports</EM> _
          directory.</P>
'
' <><><><><><>< Objects Used by Section ><><><><><><><>
'
' ##SUMMARY Box of fields in Excel spreadsheet
Dim xlRange As Excel.Range
'
' <><><><><><>< Variables Section ><><><><><><><>
'
' ##SUMMARY Name of Excel file to import
Dim ImportFile As String

Sub FillExportData(NewBook As Excel.Workbook, SelCatRec As Recordset, Locns() As String)
Attribute FillExportData.VB_Description = "Fills Excel workbook with export data."
' ##SUMMARY Fills Excel workbook with export data.
' ##PARAM NewBook M Excel workbook object.
' ##PARAM SelCatRec I Recordset of selected categories.
' ##PARAM Locns I 2-D Array of selected locations: 1st dim = code, 2nd dim = name.
' ##REMARKS Places one category per spreadsheet and puts a descriptive header at the top _
          of each page. Each cell containing a unit area code or field abbreviation is _
          assigned a pop-up comment with the associated full name. Data values are stored _
          in the grid defined by the area codes and field names. Cells with required _
          fields are outlined in red.
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim l As Long
  Dim numAreas As Long
  Dim opt As Long
  Dim ReqSt As Long
  Dim sql As String
  Dim table As String
  Dim fldTable As String
  Dim Border As String
  Dim fldNameRec As Recordset
  Dim dataRec As Recordset
  Dim pauseCancelMessage As String
  
  AtcoLaunch1.SendMonitorMessage "(MSG1 Creating Export File)"
  AtcoLaunch1.SendMonitorMessage "(PROGRESS 0)"
  AtcoLaunch1.SendMonitorMessage "(BUTTOFF DETAILS)"
  
  On Error GoTo ErrHndlr
  
  'Determine which group of special fields are required for this state
  sql = "SELECT Required FROM state WHERE state_cd='" & MyP.stateCode & "'"
  Set dataRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  ReqSt = dataRec("Required")
  dataRec.Close
  'Determine which data dictionary this data uses
  Set dataRec = MyP.stateDB.OpenRecordset("LastExport", dbOpenSnapshot)
  If dataRec("QualFlg") < 7 Then
    opt = dataRec("QualFlg")
  ElseIf dataRec("QualFlg") = 7 Then
    opt = 1
  ElseIf dataRec("QualFlg") = 8 Then
    opt = 5
  End If
  dataRec.Close
  fldTable = "Field" & opt
  'Determine which unit area table to use & create recordset with all distinct areas
  If Left(MyP.AreaTable, 2) = "co" Then
    table = "county"
  ElseIf Left(MyP.AreaTable, 3) = "huc" Then
    table = "huc"
  ElseIf Left(MyP.AreaTable, 2) = "Aq" Then
    table = "aquifer"
  End If
  'Create category recordset to retrieve name and description of selected
  'Category IDs, which are stored in SelCatRec recordset
  'Add sheets from previously existing file if necessary
  While Application.Worksheets.Count < SelCatRec.RecordCount
    Application.Worksheets.Add
  Wend
  'Delete sheets if necessary
  While Application.Worksheets.Count > SelCatRec.RecordCount
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(Application.Worksheets.Count).Delete
    Application.DisplayAlerts = True
  Wend
  
  'Loop thru each category to be exported and create a sheet for it
  SelCatRec.MoveFirst
  numAreas = UBound(Locns, 2)
  For i = 0 To SelCatRec.RecordCount - 1
    'Create a recordset with names of all data fields in current category
    sql = "SELECT DISTINCT " & fldTable & ".Name, " & fldTable & ".Description, " & fldTable & ".ID " & _
        "FROM ((" & CatTable & " INNER JOIN " & fldTable & " ON " & CatTable & ".ID = " & fldTable & ".CategoryID) " & _
        "LEFT JOIN LastExport ON " & fldTable & ".ID = LastExport.FieldID) " & _
        "INNER JOIN " & MyP.YearFields & " ON " & fldTable & ".ID = [" & MyP.YearFields & "].FieldID " & _
        "WHERE (" & CatTable & ".ID=" & SelCatRec("ID") & " And " & fldTable & ".Formula = " & "''" & ") " & _
        "ORDER BY " & fldTable & ".ID"
    Set fldNameRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    fldNameRec.MoveLast
    fldNameRec.MoveFirst
    'Create a recordset with data for all fields within current category
    sql = "SELECT LastExport.Location, " & fldTable & ".ID, LastExport.Value, " & _
        CatTable & ".Description, " & CatTable & ".ID " & _
        "FROM (" & CatTable & " INNER JOIN " & fldTable & " ON " & CatTable & ".ID = " & fldTable & ".CategoryID) " & _
        "INNER JOIN LastExport ON " & fldTable & ".ID = LastExport.FieldID " & _
        "WHERE (" & CatTable & ".ID)=" & SelCatRec("ID") & " And (" & fldTable & ".Formula = " & "''" & ") " & _
        "ORDER BY LastExport.Location, " & fldTable & ".ID;"
    Set dataRec = MyP.stateDB.OpenRecordset(sql, dbOpenForwardOnly)
    With ActiveWorkbook
      Set XLSheet = Worksheets(i + 1)
      XLSheet.Activate
    End With
    With XLSheet
      Range(Cells(5, 2), Cells(numAreas + 5, fldNameRec.RecordCount + 1)) _
           .NumberFormat = "#0.00########"
      'Write and format a header for current spreadsheet in the export file
      With Range(Cells(2, 1), Cells(2, fldNameRec.RecordCount + 1))
        .HorizontalAlignment = xlHAlignLeft
        .Font.Color = RGB(0, 55, 400)
        Cells(2, 1).value = UCase(SelCatRec("Description")) & " Export Data for " & _
                 MyP.Year1Opt & " in " & MyP.State
      End With
      'Write in keyword 'Area' and the field names for current spreadsheet
      Cells(4, 1).value = "Area"
      Cells(4, 1).BorderAround Weight:=xlThin
      For j = 2 To fldNameRec.RecordCount + 1
        Cells(4, j).value = fldNameRec("Name")
        Cells(4, j).AddComment (fldNameRec("Description"))
        Cells(4, j).BorderAround Weight:=xlThin
        If ((Mid(Trim(fldNameRec("Name")), 4, 1) = "F") Or _
            (Right(Trim(fldNameRec("Name")), 2) = "DB") Or _
            (Right(Trim(fldNameRec("Name")), 3) = "Fac")) Then
          Range(Cells(5, j), Cells(numAreas + 5, j)).NumberFormat = "####0"
        ElseIf (Right(Trim(fldNameRec("Name")), 3) = "Pop") Then
          Range(Cells(5, j), Cells(numAreas + 5, j)).NumberFormat = "#0.000#######"
        End If
        fldNameRec.MoveNext
      Next j
      'Fill in locations, data values and comments for each data field 'Name'
      For k = 5 To numAreas + 5
        fldNameRec.MoveFirst
        With Cells(k, 1)
          AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (i + _
              (k - 4) / (numAreas + 1)) * 100 / SelCatRec.RecordCount & " )"
          Cells(k, 1).value = "'" & Locns(0, k - 5)
          Cells(k, 1).AddComment Locns(1, k - 5) & " " & MyP.UnitArea
          Cells(k, 1).BorderAround Weight:=xlThin
          Select Case NextPipeCharacter(AtcoLaunch1.ComputeRead)
            Case "P"
              While NextPipeCharacter(AtcoLaunch1.ComputeRead) <> "R"
                DoEvents
              Wend
            Case "C"
              ImportDone = True
              MyMsgBox.Show "The export was stopped on the " & SelCatRec("Description") _
                  & vbCrLf & " category on " & MyP.UnitArea & " " & Locns(0, k - 5) & ".", _
                  "Import not successful", "+-&OK"
              Err.Raise 999
          End Select
          'Fill in data values
          For l = 2 To fldNameRec.RecordCount + 1
            Border = MyP.Required(fldNameRec("ID"), ReqSt)
            If Border = "red" Then
              Cells(k, l).BorderAround ColorIndex:=3, Weight:=xlThin
            ElseIf Border = "blue" Then
              Cells(k, l).BorderAround ColorIndex:=23, Weight:=xlThin
            Else
              Cells(k, l).BorderAround , Weight:=xlHairline
            End If
            If Not IsNull(dataRec("Value")) Then Cells(k, l).value = dataRec("Value")
            fldNameRec.MoveNext
            dataRec.MoveNext
          Next l
        End With
      Next k
      'Format sheet and set data validation
      .Name = SelCatRec("Name")
      With Range(Cells(5, 2), Cells(numAreas + 5, fldNameRec.RecordCount + 1)).Validation
        .Add Type:=xlValidateDecimal, _
            AlertStyle:=xlValidAlertStop, _
            operator:=xlBetween, Formula1:="0", Formula2:="99999.99"
        .ErrorMessage = "You must enter a number between 0 and 99999.99"
        .ErrorTitle = "Bad Data Value"
      End With
      'Check to ensure -999.99 < HY-OfPow < 99999.99
      If SelCatRec("Name") = "HY" Then
        If opt = 0 Then
          With Range(Cells(5, 6), Cells(numAreas + 5, 6)).Validation
            .Modify Type:=xlValidateDecimal, _
                AlertStyle:=xlValidAlertStop, _
                operator:=xlBetween, Formula1:="-999.99", Formula2:="99999.99"
            .ErrorMessage = "Values for HY-OfPow must be between -999.99 and 99999.99"
          End With
        Else
          With Range(Cells(5, 5), Cells(numAreas + 5, 5)).Validation
            .Modify Type:=xlValidateDecimal, _
                AlertStyle:=xlValidAlertStop, _
                operator:=xlBetween, Formula1:="-999.99", Formula2:="99999.99"
            .ErrorMessage = "Values for HY-OfPow must be between -999.99 and 99999.99"
          End With
        End If
      End If
      'Distinguish '# of Facilities' fields and set allowable range to integers
      fldNameRec.MoveFirst
      For l = 2 To fldNameRec.RecordCount + 1
        If InStr(1, LCase(Cells(4, l)), "fac") > 1 Then
          With Range(Cells(5, l), Cells(numAreas + 5, l)).Validation
            .Modify Type:=xlValidateWholeNumber, _
                AlertStyle:=xlValidAlertStop, _
                operator:=xlBetween, Formula1:="0", Formula2:="99999"
            .ErrorMessage = "The number of facilities must be whole number between 0 and 99999"
          End With
        End If
        fldNameRec.MoveNext
      Next l
      'Format the report
      Columns("A:Z").AutoFit
      Columns(1).ColumnWidth = 15
      With Cells(k + 2, 1).Characters
        .Text = "Notes: 1) When editing numeric codes under the 'Area' heading, add a single quote " _
            & "(') before the numeric sequence and include any leading zeros."
        .Font.size = 7
      End With
      With Cells(k + 3, 1).Characters
        .Text = "              2) Cells with red borders indicate required data elements. " & _
                "The completed compilation cannot contain null values for these data elements."
        .Font.size = 7
      End With
      With Cells(k + 4, 1).Characters
        .Text = "              3) Cells with blue borders indicate required data elements " & _
                "that allow null values for individual areas where no information is available."
        .Font.size = 7
      End With
    End With
    fldNameRec.Close
    dataRec.Close
    SelCatRec.MoveNext
  Next i
  Application.Worksheets(1).Select
  SelCatRec.Close
  
  Exit Sub
ErrHndlr:
End Sub

Public Function NextPipeCharacter(PipeHandle As Long) As String
Attribute NextPipeCharacter.VB_Description = "Checks pipe to see if user has clicked either Cancel, Pause, or Resume button on status monitor."
' ##SUMMARY Checks pipe to see if user has clicked either <EM>Cancel</EM>, _
          <EM>Pause</EM>, or <EM>Resume</EM> button on status monitor.
' ##PARAM PipeHandle Integer ID for message in pipe.
' ##RETURNS Message sent through pipe based on ReadIn from Status Monitor.
' ##REMARKS For Status Monitor messages we expect one character at a time: C for Cancel, _
          P for Pause, R for Resume.
  Dim res As Long      ' return code for message read from pipe
  Dim lread As Long    ' number of bytes actually read
  Dim lavail As Long   ' number of bytes available to read
  Dim lmessage As Long ' dummy
  Dim inbuf As Byte    ' buffer
    
  DoEvents
  res = PeekNamedPipe(PipeHandle, ByVal 0&, 0, lread, lavail, lmessage)
  If res <> 0 And lavail > 0 Then
    lavail = 1 'Only get one character
    res = ReadFile(PipeHandle, inbuf, lavail, lread, 0)
    NextPipeCharacter = Chr(inbuf)
  End If
End Function

Function CreateNewXLBook(BookName As String, SelCatRec As Recordset, Locns() As String) As Workbook
Attribute CreateNewXLBook.VB_Description = "Creates new workbook file and saves it to file."
' ##SUMMARY Creates new workbook file and saves it to file.
' ##PARAM SelCatRec I Recordset specifies IDs of selected categories and determines _
          number of worksheets in workbook.
' ##PARAM Locns I 2-D Array of selected locations: 1st dim = code, 2nd dim = name.<BR>
' ##PARAM BookName I Full path and filename of report.
' ##RETURNS Excel workbook object.
  Dim NewBook As Excel.Workbook 'new Excel workbook object
  
  'Open Status Monitor
  AtcoLaunch1.SendMonitorMessage "(OPEN AWUDS)"
  AtcoLaunch1.SendMonitorMessage "(BUTTON CANCEL)"
  AtcoLaunch1.SendMonitorMessage "(BUTTON PAUSE)"
  
  On Error GoTo x
  
  Set XLApp = New Excel.Application
  Set NewBook = Excel.Workbooks.Add
  Application.SheetsInNewWorkbook = SelCatRec.RecordCount
  NewBook.SaveAs BookName, XLFileFormatNum
  FillExportData NewBook, SelCatRec, Locns
  ImportDone = True
x:
  If Err.Number = 999 Then _
      AtcoLaunch1.SendMonitorMessage "(MSG1 User Canceled Export)"
  If ImportDone = False Then
    MyMsgBox.Show "The import was not successful." & vbCrLf & _
        "Make sure the destination file is not currently open.", _
        "Import not successful", "+-&OK"
  End If
  NewBook.Close True
  Set NewBook = Nothing
  XLApp.Quit
  AtcoLaunch1.SendMonitorMessage "(CLOSE)"

End Function

Sub XLImport(ImpFileName As String)
Attribute XLImport.VB_Description = "Imports data from properly formatted Excel workbook (as created by export function)."
' ##SUMMARY Imports data from properly formatted Excel workbook (as created by export _
          function).
' ##REMARKS Performs numerous QA checks before importing data to DB. Any records not _
          imported to State database are written to 'AWUDSReports\ImportError.txt' for _
          reference.
' ##PARAM ImpFileName I Full path and filename of Excel workbook to be imported.
  Dim sql As String            ' recordset query
  Dim locn As String           ' name of location read from spreadsheet
  Dim roundedVals As String    ' message documenting rounded import data values
  Dim numFormat As String      ' Excel number format of the import values
  Dim errMsg As String         ' message to user characterizing bad import value
  Dim rowCntr As Long          ' counter for loop thru rows in spreadsheet
  Dim colCntr As Long          ' counter for loop thru columns in spreadsheet
  Dim lastRow As Long          ' last row of data table in spreadsheet
  Dim firstCol As Long         ' first column of data table in spreadsheet
  Dim lastCol As Long          ' last column of data table in spreadsheet
  Dim header As Long           ' header row
  Dim numFlds As Long          ' number of data fields in spreadsheet
  Dim Length As Long           ' length of unit area code
  Dim locnIndex As Long        ' index for position in LocnArray
  Dim h As Long                ' counter for loop thru spreadsheets in workbook
  Dim i As Long                ' counter for loop thru fields
  Dim j As Long                ' index for data storage option read from QualFlg field in State DB
  Dim k As Long                ' loop counter
  Dim response As Long         ' index to store user response to message box
  Dim rangeTemp As Excel.Range ' object representing block of cells in spreadsheet
  Dim value As Variant         ' value read from cell in spreadsheet
  Dim existing As Variant      ' value already in DB for specific field and location
  Dim impFlds() As Long        ' array of field IDs being imported
  Dim tmpLocns() As Variant    ' array of locations being imported
  Dim fldRec As Recordset      ' recordset of all fields in current category
  Dim allDataRec As Recordset  ' recordset with all data in State DB for given unit area & year
  Dim areaRec As Recordset     ' recordset of all areas (i.e., counties, hucs, aquifers) in state
  Dim OutFile As Integer       ' next available file number, if creating "ImportError.txt" file.
  Dim ok As Boolean            ' true if import value is acceptable
  Dim tooBig As Boolean        ' true if sql query has too many characters
  
  On Error GoTo ErrTrap
  
  ImportFile = ImpFileName
  Set XLApp = New Excel.Application
  Set XLBook = Workbooks.Open(ImpFileName)

  AtcoLaunch1.SendMonitorMessage "(OPEN AWUDS)"
  AtcoLaunch1.SendMonitorMessage "(BUTTON CANCEL)"
  AtcoLaunch1.SendMonitorMessage "(BUTTON PAUSE)"
  AtcoLaunch1.SendMonitorMessage "(BUTTOFF DETAILS)"
  AtcoLaunch1.SendMonitorMessage "(MSG1 Importing:)"
  AtcoLaunch1.SendMonitorMessage "(MSG2 " & ImpFileName & ")"
  AtcoLaunch1.SendMonitorMessage "(PROGRESS 0)"
      
  Areas = ""
  errMsg = ""
  roundedVals = ""
  
  'Loop thru all sheets in workbook - should be 1 per category
  For h = 1 To XLBook.Sheets.Count
    With ActiveWorkbook
      Set XLSheet = Worksheets(h)
      XLSheet.Activate
    End With
    With XLSheet
      'Check for the key word "Area"
      Set xlRange = .UsedRange
      Set rangeTemp = xlRange.Find("Area", , , , xlByRows, xlNext, False)
      If rangeTemp Is Nothing Then
        MyMsgBox.Show "The key word 'Area' was not found in " & .Name & "." & vbCrLf _
            & vbCrLf & "'Area' must appear in the top left corner of the data block." _
            & vbCrLf & "See any AWUDS export file for an example." _
            & vbCrLf & vbCrLf & "NO DATA WAS IMPORTED FROM " & .Name & ".", _
            "Field name not found", "+-&OK"
        GoTo y
      End If
      'Determine the size and location of the data fields in XL
      ' and size the array for fields appropriately
      header = rangeTemp.Row
      firstCol = rangeTemp.Column
      lastCol = firstCol + xlRange.Columns.Count - 1
      lastRow = header + xlRange.Rows.Count - 1
      numFlds = lastCol - firstCol
      ReDim impFlds(numFlds - 1)
      'use length of unit area code to determine type of unit area
      Length = Len(Cells(header + 1, firstCol))
      If Length = 3 Then
        TableName = "County"
      ElseIf Length = 4 Or Length = 8 Then
        TableName = "HUC"
      ElseIf Length = 10 Then
        TableName = "Aquifer"
      Else
        MyMsgBox.Show "The unit areas on worksheet number " & h & _
            " appear to have " & Length & " digits." _
            & vbCrLf & "They must have either 3 digits for counties, " & _
            " 4 for HUC-4s, 8 for HUC-8s, or 10 for aquifers." & vbCrLf & _
            "Edit this sheet and try again.  Any other sheets will still be imported.", _
            "Unrecognized Unit Areas", "+-&OK"
        GoTo y
      End If
      
      If h = 1 Then 'first xl sheet.
        'Create recordset with all data in DB for given unit area & year.
        ' All imported data are added to this recordset.
        sql = "SELECT * FROM " & TableName & "Data " & _
              "WHERE Date=" & MyP.Year1Opt
        Set allDataRec = MyP.stateDB.OpenRecordset(sql, dbOpenDynaset)
        allDataRec.FindFirst "FieldID=195"
        If allDataRec.NoMatch Then IRinTwo = True Else IRinTwo = False
        'Create data array.
        If allDataRec.RecordCount > 0 Then
          allDataRec.MoveFirst
          'Check data storage type in database
          If allDataRec("QualFlg") < 7 Then
            j = allDataRec("QualFlg")
          ElseIf allDataRec("QualFlg") = 7 Then
            j = 1
          ElseIf allDataRec("QualFlg") = 8 Then
            j = 5
          End If
          FieldTable = "Field" & j
          Years = "(" & TableName & "Data.Date = " & MyP.Year1Opt & ")"
          If j = 0 Then
            MyP.YearFields = "1995Fields1"
            CatTable = "Category1"
          Else
            MyP.YearFields = "2000Fields" & j
            CatTable = "Category2"
          End If
          'Put info in 'Field_' table into set of arrays to be used for comparison
          PopulateFieldArrays FieldTable
          
          'Create recordset of all areas (i.e., counties, hucs, aquifers) in state
          MyP.Length = Len(Trim(Cells(header + 1, firstCol)))
          If MyP.Length = 3 Or MyP.Length = 10 Then 'importing county or aquifer
            sql = "SELECT * FROM " & TableName & _
                  " WHERE state_cd='" & MyP.stateCode & "'" & _
                  " AND ((" & TableName & ".begin<=" & MyP.Year1Opt & " OR IsNull(" & TableName & ".begin))" & _
                  " AND (" & TableName & ".end>=" & MyP.Year1Opt & " OR IsNull(" & TableName & ".end)))"
          Else
            sql = "SELECT * FROM " & TableName & _
                  " WHERE state_cd='" & MyP.stateCode & "'"
          End If
          sql = sql & " And Len(Trim(" & LCase(TableName) & "_cd))=" & MyP.Length & _
                  " ORDER BY " & LCase(TableName) & "_cd;"
          Set areaRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
          areaRec.MoveLast
          ReDim LocnArray(0, (areaRec.RecordCount - 1))
          areaRec.MoveFirst
          While Not areaRec.EOF
            LocnArray(0, areaRec.AbsolutePosition) = Trim(areaRec(LCase(TableName) & "_cd"))
            areaRec.MoveNext
          Wend
          
          frmAwuds2.CreateTable
          PopulateDataArray
        Else  'No data exists yet for this unit area & year.
          sql = "It was inferred from the Area entry '" & Cells(header + 1, firstCol) & _
                "' on row " & header + 1 & " of Sheet " & .Name & _
                " that you are attempting to import "
          Select Case Length
            Case 3: sql = sql & "County"
            Case 4: sql = sql & "HUC - 4"
            Case 8: sql = sql & "HUC - 8"
            Case 10: sql = sql & "Aquifer"
          End Select
          sql = sql & " data," & vbCrLf & "but there is no such " & TableName & _
              " data in the State database for " & MyP.State & ", " _
              & MyP.Year1Opt & "." & vbCrLf & vbCrLf & _
              "If you are in fact trying to import " & TableName & " data, " & _
              "a blank template must be created in the State database using the" & vbCrLf & _
              "'Create New Year of Data' utility before data can be imported " & _
              "to the database." & vbCrLf & _
              "If you did not intend to import " & TableName & " data, " & _
              "recheck the Area entries in the Import workbook."
          MyMsgBox.Show sql, "No " & TableName & " Data for This Year", "+-&OK"
          GoTo ErrTrap
        End If
      End If
      'Check to make sure irrigation options are consistent
      If IRinTwo And (.Name = "IR") Then 'do not include Total Irrigation fields
        MsgBox "The import file contains the category 'IR', but irrigation for " & TableName & " areas in " & _
              MyP.State & " for " & MyP.Year1Opt & " is divided into Golf and Crop." & vbCrLf & _
              "No data will be imported for 'IR'."
        GoTo y
      ElseIf Not IRinTwo And (.Name = "IC" Or .Name = "IG") Then
        MsgBox "The import file contains the category '" & .Name & "', but irrigation for " & TableName & _
              " areas in " & MyP.State & " for " & MyP.Year1Opt & " is kept as a total, not divided into Golf and Crop." & _
              vbCrLf & "No data will be imported for 'IC' or 'IG'."
        GoTo y
      End If
      'Create recordset with all fields in current category
      sql = "SELECT " & FieldTable & ".ID, " & FieldTable & ".Name, " & FieldTable & ".Formula FROM " & CatTable & _
            " INNER JOIN " & FieldTable & " ON " & CatTable & ".ID = " & FieldTable & ".CategoryID " & _
            "WHERE " & CatTable & ".Name='" & XLSheet.Name & _
            "' ORDER BY " & FieldTable & ".ID;"
      Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
      If fldRec.RecordCount = 0 Then 'no data on record for this "category"
          sql = "There are no fields for the category '" & .Name & "' for " & TableName & _
                " areas in " & MyP.State & " for " & MyP.Year1Opt & "." & vbCrLf & _
                "The name of the worksheet must be the 2-letter abbreviation for the data category."
        MsgBox sql
        GoTo y
      End If
      
      'Get FieldIDs from field names in header columns, check whether they exist,
      ' then create ImpFlds array containing FieldIDs
      i = 0   ' field counter
      For colCntr = firstCol + 1 To lastCol
        fldRec.FindFirst "Name='" & .Cells(header, colCntr) & "'"
        If Len(Trim(.Cells(header, colCntr))) = 0 Then
          sql = "No field name was entered for column " & colCntr & " on sheet " & _
              XLSheet.Name & "." & vbCrLf & "Since the column header is blank, " & _
              "no data will be imported for column " & colCntr & "."
        ElseIf fldRec.NoMatch Then
          If .Name = "PS" And (MyP.DataOpt = 3 Or MyP.DataOpt = 4) And InStr(1, .Cells(header, colCntr), "Pop") > 0 Then
            sql = "For " & MyP.Year1Opt & ", " & MyP.State & " stores Public Supply population data by state" _
                & " total, not by individual " & TableName & "." & vbCrLf & _
                "Therefore, no data will be imported for the " & .Cells(header, colCntr) & " field."
          ElseIf .Name = "PS" And (MyP.DataOpt = 5 Or MyP.DataOpt = 6) And InStr(1, .Cells(header, colCntr), "WPop") > 0 Then
            sql = "For " & MyP.Year1Opt & ", " & MyP.State & " stores Public Supply population data by " _
                & TableName & " total, not divided into SW and GW." & vbCrLf & _
                "Therefore, no data will be imported for the " & .Cells(header, colCntr) & " field."
          ElseIf .Name = "PS" And (MyP.DataOpt = 1 Or MyP.DataOpt = 2) And InStr(1, .Cells(header, colCntr), "tPop") > 0 Then
            sql = "For " & MyP.Year1Opt & ", " & MyP.State & " Public Supply population data is " _
                & "divided into SW and GW, not kept as a " & TableName & " total." & vbCrLf & _
                "Therefore, no data will be imported for the " & .Cells(header, colCntr) & " field."
          ElseIf .Name = "DO" And (MyP.DataOpt = 2 Or MyP.DataOpt = 4) And InStr(1, .Cells(header, colCntr), "WFr") > 0 Then
            sql = "For " & MyP.Year1Opt & ", " & MyP.State & " stores Domestic withdrawals by state" _
                & " total, not by individual " & TableName & "." & vbCrLf & _
                "Therefore, the " & .Cells(header, colCntr) & " field will not be imported."
          Else
            sql = "There is no field named '" & .Cells(header, colCntr) & "' on file for this area." _
                & vbCrLf & vbCrLf & "Check the spelling of this field on worksheet '" & _
                XLSheet.Name & "' to make sure it is correct." & vbCrLf & _
                "All acceptable field names are listed in the table '" & FieldTable & _
                "' in " & AwudsDataPath & "Categories.mdb." & vbCrLf & _
                "No data will be imported for this field."
          End If
        ElseIf Len(Trim(fldRec("Formula"))) > 0 Then
            sql = "The import file contains the calculated field '" & .Cells(header, colCntr) & _
                "' on Sheet " & .Name & "." & vbCrLf & _
                "Calculated fields are not stored in the database, but rather calculated at runtime." & _
                "Therefore, no data will be imported for this field."
        Else
          'Check to see if field is duplicated on this sheet
          For j = 0 To i - 1
            If impFlds(j) = fldRec("ID") Then
              errMsg = errMsg & "The import file contains duplicate fields for " & fldRec("Name") & " on sheet " & .Name & "." & vbCrLf
            End If
          Next j
          impFlds(i) = fldRec("ID")
          sql = ""
        End If
        If sql <> "" Then
          MyMsgBox.Show sql, "Bad or Missing Field Name", "+-&OK"
        End If
        i = i + 1
      Next colCntr
      'Run a check to see if there are any duplicate entries for areas
      ReDim tmpLocns(header + 1 To lastRow)
      For rowCntr = header + 1 To lastRow
        locn = Trim(Cells(rowCntr, firstCol))
        If Len(locn) = Length And (IsNumeric(locn) Or Length = 10) Then
          tmpLocns(rowCntr) = locn
          For j = header + 1 To rowCntr - 1
            If locn = tmpLocns(j) Then
              errMsg = errMsg & "The import file contains duplicate entries for the location '" & _
                       locn & "' on sheet '" & .Name & "'." & vbCrLf
            End If
          Next j
        End If
      Next rowCntr
      'Get unit areas and associated data values from XLSheet. Run checks on all values.
      For rowCntr = header + 1 To lastRow
        locn = Trim(Cells(rowCntr, firstCol))
        If Not (Len(locn) = Length And (IsNumeric(locn) Or Length = 10)) Then
          If LCase(Left(locn, 5)) = "notes" Then
            rowCntr = lastRow
          ElseIf locn <> "" Then
            MyMsgBox.Show "The worksheet '" & .Name & "' contains the entry " & _
                locn & " in row " & rowCntr & " of the Area column, " & _
                "which is not a valid location for a " & TableName & " in " & _
                MyP.State & "." & vbCrLf & vbCrLf & "No data in this row will be imported.", _
                "Code for " & MyP.UnitArea & " not found", "+-&OK"
          End If
          GoTo NextRow
        End If
        k = 0
        Select Case NextPipeCharacter(AtcoLaunch1.ComputeRead)
          Case "P"
            While NextPipeCharacter(AtcoLaunch1.ComputeRead) <> "R"
              DoEvents
            Wend
          Case "C"
            ImportDone = True
              MyMsgBox.Show "The import was stopped before any data was saved in the " _
                  & TableName & " database.", _
                  "Import cancelled", "+-&OK"
            Err.Raise 999
        End Select
        For colCntr = firstCol To lastCol
          If k = 0 Then 'first column with the unit areas
            AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (h * 100) / (XLBook.Sheets.Count * 2) + _
                (rowCntr - header + 1) / (lastRow - header + 1) & " )"
            'Check to see if unit area code on current row of XLSheet really exists.
            areaRec.FindFirst TableName & "_cd='" & locn & "'"
            locnIndex = areaRec.AbsolutePosition
            If areaRec.NoMatch Then  'area does not exist for that year
              MyMsgBox.Show "There is no " & TableName & " in " & MyP.State & _
                  " with the code " & locn & " for " & MyP.Year1Opt & "." & _
                  vbCrLf & vbCrLf & "Check the '" & LCase(TableName) & _
                  "_cd' column in the '" & LCase(TableName) & "' table" & vbCrLf & _
                  "in '" & AwudsDataPath & "General.mdb' to find the proper " & Length & "-digit code." _
                  & vbCrLf & "No data associated with this code will be imported.", _
                  "Code for " & MyP.UnitArea & " not found", "+-&OK"
              Exit For 'pop out of column loop and go to next row
            End If
          Else 'field contains data element for an existing area
            If impFlds(k - 1) > 0 Then  'we have a data field to match data element
              value = .Cells(rowCntr, colCntr).value
              If tooBig = True Then  'all areas referenced
                NoRec = True
              Else
                existing = EvalArray(impFlds(k - 1), locnIndex)
              End If
              'check to ensure that import value is a number within permissible range
              numFormat = .Cells(rowCntr, colCntr).NumberFormat
              If Not IsNumeric(value) Then
                ok = False
                errMsg = errMsg & locn & " has a non-numeric value (" & value & ") for the field " & _
                    .Cells(header, colCntr) & "." & vbCrLf
              ElseIf value <> Empty And _
                    (numFormat = "General" Or _
                     InStr(1, numFormat, ":") > 1 Or _
                     InStr(1, numFormat, "-") > 1 Or _
                     Left(numFormat, 1) = "$" Or _
                     InStr(1, numFormat, "/") > 1 Or _
                     InStr(1, numFormat, "%")) Then
                errMsg = errMsg & locn & " has improper formatting ('" & numFormat & "') for the field " & _
                    .Cells(header, colCntr) & "." & vbCrLf & _
                    "  Excel files created by exporting from AWUDS are created with the proper " & _
                    "formatting ('Number' with 0-3 decimal places, depending on the field)." & vbCrLf & _
                    "  Do not change the formatting in the spreadsheets from that of a 'Number' " & _
                    "with the appropriate precision" & vbCrLf
              ElseIf (value < 0 Or value > 99999.999) Then
                If impFlds(k - 1) = 218 And (value >= -999.99 And value <= 99999.999) Then
                  'HY-OfPow can be negative
                  ok = True
                  NoRec = False
                Else
                  ok = False
                  errMsg = errMsg & locn & " has a value (" & value & ") for the field " & _
                      .Cells(header, colCntr) & " that is outside of the permissible range"
                  If impFlds(k - 1) = 218 Then
                    errMsg = errMsg & " (-999.99 to 99999.99)." & vbCrLf
                  Else
                    errMsg = errMsg & " (0.00 to 99999.99)." & vbCrLf
                  End If
                End If
              Else
                ok = True
              End If
              'Check to ensure decimal place limit not exceeded.
              i = InStr(1, CStr(value), ".")
              If i > 0 Then
                i = Len(value) - i
                If LCase(Right(.Cells(header, colCntr), 6)) = "-facil" Or _
                   LCase(Right(.Cells(header, colCntr), 3)) = "fac" Then
                  errMsg = errMsg & .Cells(header, colCntr) & " must be an integer.  " & _
                      value & " was entered for " & locn & "." & vbCrLf
                  value = RoundDc(CDbl(value), 0, True, False)
                  ok = False
                ElseIf i > 2 Then
                  If InStr(1, LCase(.Cells(header, colCntr)), "pop") > 0 Then
                    If i > 3 Then roundedVals = roundedVals & "You entered a value with " & i & _
                        " decimal places for the field '" & .Cells(header, colCntr) & "' at " & locn & "." & _
                        "  Only 3 are allowed; the value was rounded off using the USGS Rounding function." & vbCrLf
                    value = RoundDc(CDbl(value), -3, True, False)
                  Else
                    roundedVals = roundedVals & "You entered a value with " & i & _
                        " decimal places for the field '" & .Cells(header, colCntr) & "' at " & locn & "." & _
                        "  Only 2 are allowed; the value was rounded off using the USGS Rounding function." & vbCrLf
                    value = RoundDc(CDbl(value), -2, True, False)
                  End If
                End If
              End If
              'Put updated values in a buffer
              If Len(Trim(errMsg)) > 0 Then response = 2
              With allDataRec
                If Not (response = 2 Or NoRec) Then
                  If CStr(existing) <> value And Not IsEmpty(existing) Then
                    If IsEmpty(value) Or value = Null Then sql = "null" Else sql = CStr(value)
                    fldRec.FindFirst "Name='" & XLSheet.Cells(header, colCntr) & "'"
                    response = MyMsgBox.Show("Data already exists for " & fldRec("Name") & _
                        " in " & Trim(areaRec(TableName & "_nm")) & " " & TableName & " (" & _
                        Trim(areaRec(TableName & "_cd")) & ") for " & MyP.Year1Opt & _
                        "." & vbCrLf & vbCrLf & "Do you want to overwrite this data?" _
                        & vbCrLf & "The existing value is " & existing & _
                        " and the import value is " & sql & ".", "Confirm Overwrite of Data", _
                        "+&Yes", "&Yes to All", "&No", "-&Cancel Import")
                  End If
                End If
                If response < 3 Then
                  If ok Then
                    If IsEmpty(value) Then
                      DataArray(areaRec.AbsolutePosition, impFlds(k - 1)) = -9999
                    Else
                      DataArray(areaRec.AbsolutePosition, impFlds(k - 1)) = value
                    End If
                  End If
                ElseIf response = 3 Then
                  response = 1
                ElseIf response = 4 Then
                  GoTo ErrTrap
                End If
              End With
            End If
          End If
          k = k + 1
        Next colCntr
NextRow:
      Next rowCntr
      areaRec.MoveFirst
      fldRec.Close
y:
    End With
  Next h
  If Len(errMsg) > 0 Then 'write error report
    OutFile = FreeFile
    Open ReportPath & "ImportError.txt" For Output As OutFile
    Print #OutFile, errMsg
    Close OutFile
    MyMsgBox.Show "The import file '" & ImpFileName & "' contains erroneous values." _
        & vbCrLf & "Check the '" & ReportPath & "ImportError.txt' file to see " & _
        "which fields are problematic, then edit the import file to correct those values." & _
        vbCrLf & vbCrLf & "NO DATA WAS IMPORTED from '" & ImpFileName & "'.", _
        "Bad Data Value(s)", "+-&OK"
    PopulateDataArray
  Else 'No errors; update records
    For h = 0 To UBound(LocnArray, 2)
      locn = LocnArray(0, h)
      sql = "SELECT * FROM " & TableName & "Data " & _
            "WHERE Location='" & locn & "' AND Date=" & MyP.Year1Opt
      Set allDataRec = MyP.stateDB.OpenRecordset(sql, dbOpenDynaset)
      With allDataRec
        Select Case NextPipeCharacter(AtcoLaunch1.ComputeRead)
          Case "P"
            While NextPipeCharacter(AtcoLaunch1.ComputeRead) <> "R"
              DoEvents
            Wend
          Case "C"
            ImportDone = True
              MyMsgBox.Show "The import was stopped while writing data to the " & TableName & " database." _
                  & vbCrLf & "The " & MyP.stateCode & " database was updated through " & locn & ".", _
                  "Import interrupted", "+-&OK"
            Err.Raise 999
        End Select
        For i = 1 To NFields
          value = EvalArray(i, h)
          .FindFirst "FieldID=" & i
          If Not .NoMatch Then
            .Edit
            If NoRec Then !value = Null Else !value = value
            .Update
          End If
        Next i
        AtcoLaunch1.SendMonitorMessage "(PROGRESS " & 50 + (h * 100) / (2 * UBound(LocnArray, 2)) & " )"
      End With
      allDataRec.Close
    Next h
  End If
  'Write report on rounded values, if any
  If Len(roundedVals) > 0 Then
    OutFile = FreeFile
    Open ReportPath & "ImportRounding.txt" For Output As OutFile
    Print #OutFile, roundedVals
    Close OutFile
    sql = "The import file '" & ImpFileName & "' contains certain values with too much precision." & vbCrLf
    If Len(errMsg) > 0 Then
      sql = sql & "These values will be rounded off to the proper decimal place."
    Else
      sql = sql & "These values were rounded off to the proper decimal place."
    End If
    MyMsgBox.Show sql & vbCrLf & "Check the '" & ReportPath & "ImportRounding.txt' file " & _
        "to see which fields were rounded off.", _
        "Bad Data Value(s)", "+-&OK"
  End If
ErrTrap:
  If Err.Number = 999 Then AtcoLaunch1.SendMonitorMessage _
      "(MSG1 User Canceled Import)"
  AtcoLaunch1.SendMonitorMessage "(CLOSE)"
  XLBook.Close (False)
  Set XLBook = Nothing
  XLApp.Quit
  Set XLApp = Nothing
End Sub

Sub SetNatExpHeaders(NewBook As Excel.Workbook, SelCatRec As Recordset, ExpType As String)
' ##SUMMARY Fills Excel workbook with headers for National data export.
' ##PARAM NewBook M Excel workbook object.
' ##PARAM SelCatRec I Recordset of selected categories.
' ##PARAM ExpType I Export area type - "co" or "aq"
' ##REMARKS Places one category per spreadsheet and puts a descriptive header at the top _
          of each page. Each cell containing a unit area code or field abbreviation is _
          assigned a pop-up comment with the associated full name.  All irrigation categories _
          are stored on one irrigation spreadsheet.
  Dim i As Long
  Dim j As Long
  Dim CatID As Long
  Dim fldOffset As Long
  Dim nFlds As Long
  Dim sCol As Long
  Dim sql As String
  Dim fldNameRec As Recordset
  Dim pauseCancelMessage As String
  
  AtcoLaunch1.SendMonitorMessage "(MSG1 Creating National Export File)"
  AtcoLaunch1.SendMonitorMessage "(PROGRESS 0)"
  AtcoLaunch1.SendMonitorMessage "(BUTTOFF DETAILS)"
  
  On Error GoTo ErrHndlr
  
  'Create category recordset to retrieve name and description of selected
  'Category IDs, which are stored in SelCatRec recordset
  'Add sheets from previously existing file if necessary
  While Application.Worksheets.Count < SelCatRec.RecordCount
    Application.Worksheets.Add
  Wend
  'Delete sheets if necessary
  While Application.Worksheets.Count > SelCatRec.RecordCount
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(Application.Worksheets.Count).Delete
    Application.DisplayAlerts = True
  Wend

  If ExpType = "aq" Then
    sCol = 4
  Else
    sCol = 2
  End If
  fldOffset = 0
  'Loop thru each category to be exported and create a sheet for it
  SelCatRec.MoveFirst
  For i = 0 To SelCatRec.RecordCount - 1
    'Create a recordset with names of all data fields in current category
    CatID = SelCatRec("ID")
MoreFlds:
    If ExpType = "aq" Then
      sql = "SELECT FieldA.ID, FieldA.Name, FieldA.CategoryID, FieldA.Description " & _
            "From FieldA " & _
            "WHERE ((FieldA.CategoryID)=" & CatID & " And Len(Trim([Formula]))=0)"
    Else
      sql = "SELECT AllFields.ID, AllFields.Name, AllFields.CategoryID, AllFields.Description " & _
            "From AllFields " & _
            "WHERE ((AllFields.CategoryID)=" & CatID & " And Len(Trim([Formula]))=0)"
    End If

'    sql = "SELECT DISTINCT " & FldTable & ".Name, " & FldTable & ".Description, " & FldTable & ".ID " & _
'        "FROM ((" & CatTable & " INNER JOIN " & FldTable & " ON " & CatTable & ".ID = " & FldTable & ".CategoryID) " & _
'        "LEFT JOIN LastExport ON " & FldTable & ".ID = LastExport.FieldID) " & _
'        "INNER JOIN " & MyP.YearFields & " ON " & FldTable & ".ID = [" & MyP.YearFields & "].FieldID " & _
'        "WHERE (" & CatTable & ".ID=" & SelCatRec("ID") & " And " & FldTable & ".Formula = " & "''" & ") " & _
'        "ORDER BY " & FldTable & ".ID"
    Set fldNameRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    fldNameRec.MoveLast
    fldNameRec.MoveFirst
    nFlds = fldNameRec.RecordCount
    With ActiveWorkbook
      Set XLSheet = Worksheets(i + 1)
      XLSheet.Activate
    End With
    With XLSheet
      Range(Cells(5, sCol + fldOffset), Cells(5, fldNameRec.RecordCount + fldOffset + sCol - 1)) _
          .NumberFormat = "####0.00"
      If CatID < 18 Or CatID > 19 Then
        'Write and format a header for current spreadsheet in the export file
        With Range(Cells(2, 1), Cells(2, fldNameRec.RecordCount + 1))
          .HorizontalAlignment = xlHAlignLeft
          .Font.Color = RGB(0, 55, 400)
          Cells(2, 1).value = UCase(SelCatRec("Description")) & " Export Data for " & _
                   MyP.Year1Opt & " in " & MyP.State
        End With
        'Write in keyword 'Area' and the field names for current spreadsheet
        Cells(4, 1).value = "Area"
        Cells(4, 1).BorderAround Weight:=xlThin
        If ExpType = "aq" Then 'add two extra fields for State and Aquifer names
          Cells(4, 2).value = "State"
          Cells(4, 3).value = "Aquifer"
          Range(Cells(4, 2), Cells(4, 3)).BorderAround Weight:=xlThin
        End If
        .Name = SelCatRec("Name") 'set sheet name
      End If
      'set field headers
      For j = sCol To nFlds + sCol - 1
        Cells(4, j + fldOffset).value = fldNameRec("Name")
        Cells(4, j + fldOffset).AddComment (fldNameRec("Description"))
        Cells(4, j + fldOffset).BorderAround Weight:=xlThin
        Cells(4, j + fldOffset).id = fldNameRec("ID")
        If ((Mid(Trim(fldNameRec("Name")), 4, 1) = "F") Or _
            (Right(Trim(fldNameRec("Name")), 2) = "DB") Or _
            (Right(Trim(fldNameRec("Name")), 3) = "Fac")) Then
          Range(Cells(5, j), Cells(5, j)).NumberFormat = "####0"
        ElseIf (Right(Trim(fldNameRec("Name")), 3) = "Pop") Then
          Range(Cells(5, j), Cells(5, j)).NumberFormat = "###0.000"
        End If
        fldNameRec.MoveNext
      Next j
    End With
    fldNameRec.Close
    If CatID = 17 Or CatID = 18 Then 'add irrig fields from other irrigation categories
      CatID = CatID + 1
      fldOffset = fldOffset + nFlds
      GoTo MoreFlds
    Else
      fldOffset = 0
      SelCatRec.MoveNext
    End If
  Next i

  Exit Sub
ErrHndlr:
End Sub

Sub FillNationalExport(NewBook As Excel.Workbook, SelCatRec As Recordset, Locns() As String, CurRow As Long, ExpType As String)
' ##SUMMARY Fills Excel workbook with National export data.
' ##PARAM NewBook M Excel workbook object.
' ##PARAM SelCatRec I Recordset of selected categories.
' ##PARAM Locns I 2-D Array of selected locations: 1st dim = code, 2nd dim = name.
' ##PARAM CurRow M Current row number on worksheets for exporting national data - begin at 5.
' ##REMARKS Worksheets for selected categories and data headers for National export are _
          built in SetNatExpHeaders. Data values are stored in the grid defined by the _
          area codes and field names. Cells with required fields are outlined in red.
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim l As Long
  Dim lCol As Long
  Dim sCol As Long
  Dim mxCol As Long
  Dim dataID As Long
  Dim fldID As Long
  Dim lCatID As Long
  Dim nFlds As Long
  Dim fldOffset As Long
  Dim numAreas As Long
  Dim opt As Long
  Dim ReqSt As Long
  Dim sql As String
  Dim fldTable As String
  Dim Border(50) As String
  Dim fldNameRec As Recordset
  Dim dataRec As Recordset
  Dim pauseCancelMessage As String
  
  AtcoLaunch1.SendMonitorMessage "(MSG2 Exporting " & MyP.State & ")"
  AtcoLaunch1.SendMonitorMessage "(PROGRESS 0)"
  AtcoLaunch1.SendMonitorMessage "(BUTTOFF DETAILS)"
  
  On Error GoTo ErrHndlr
  
  'Determine which group of special fields are required for this state
  sql = "SELECT Required FROM state WHERE state_cd='" & MyP.stateCode & "'"
  Set dataRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  ReqSt = dataRec("Required")
  dataRec.Close
  'Determine which data dictionary this data uses
  Set dataRec = MyP.stateDB.OpenRecordset("LastExport", dbOpenSnapshot)
  If dataRec("QualFlg") < 7 Then
    opt = dataRec("QualFlg")
  ElseIf dataRec("QualFlg") = 7 Then
    opt = 1
  ElseIf dataRec("QualFlg") = 8 Then
    opt = 5
  End If
  dataRec.Close
  fldTable = "Field" & opt

  If ExpType = "aq" Then
    sCol = 4
  Else
    sCol = 2
  End If
  'Loop thru each category to be exported and output its data
  SelCatRec.MoveFirst
  numAreas = UBound(Locns, 2)
  If ExpType = "co" Then CurRow = CurRow + 1 'leave blank row to fill in state totals
  For i = 0 To SelCatRec.RecordCount - 1
    mxCol = sCol - 1 'set max column number to last ID column
    lCatID = SelCatRec("ID")
MoreFlds:
    'Create a recordset with names of all data fields in current category
    sql = "SELECT DISTINCT " & fldTable & ".Name, " & fldTable & ".Description, " & fldTable & ".ID " & _
        "FROM ((" & CatTable & " INNER JOIN " & fldTable & " ON " & CatTable & ".ID = " & fldTable & ".CategoryID) " & _
        "LEFT JOIN LastExport ON " & fldTable & ".ID = LastExport.FieldID) " & _
        "INNER JOIN " & MyP.YearFields & " ON " & fldTable & ".ID = [" & MyP.YearFields & "].FieldID " & _
        "WHERE (" & CatTable & ".ID=" & lCatID & " And " & fldTable & ".Formula = " & "''" & ") " & _
        "ORDER BY " & fldTable & ".ID"
    Set fldNameRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    fldNameRec.MoveLast
    fldNameRec.MoveFirst
    nFlds = fldNameRec.RecordCount
    'Create a recordset with data for all fields within current category
    sql = "SELECT LastExport.Location, " & fldTable & ".ID, LastExport.Value, " & _
        CatTable & ".Description, " & CatTable & ".ID " & _
        "FROM (" & CatTable & " INNER JOIN " & fldTable & " ON " & CatTable & ".ID = " & fldTable & ".CategoryID) " & _
        "INNER JOIN LastExport ON " & fldTable & ".ID = LastExport.FieldID " & _
        "WHERE (" & CatTable & ".ID)=" & lCatID & " And (" & fldTable & ".Formula = " & "''" & ") " & _
        "ORDER BY LastExport.Location, " & fldTable & ".ID;"
    Set dataRec = MyP.stateDB.OpenRecordset(sql, dbOpenForwardOnly)
    If dataRec.RecordCount > 0 Then
      With ActiveWorkbook
        Set XLSheet = Worksheets(i + 1)
        XLSheet.Activate
      End With
      With XLSheet
        'Fill in locations, data values and comments for each data 'Name'
        For k = CurRow To numAreas + CurRow
          fldNameRec.MoveFirst
          With Cells(k, 1)
            AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (i + _
              (k - CurRow + 1) / (numAreas + 1)) * 100 / SelCatRec.RecordCount & " )"
            If Len(Cells(k, 1)) = 0 Then 'fill in location data
              If ExpType = "co" Then
                Cells(k, 1).value = "'" & MyP.stateCode & Locns(0, k - CurRow)
                Cells(k, 1).AddComment MyP.State & "-" & Locns(1, k - CurRow) & " " & MyP.UnitArea
                Cells(k, 1).BorderAround Weight:=xlThin
              Else
                Cells(k, 1).value = "'" & MyP.stateCode
                Cells(k, 1).AddComment MyP.State
                Cells(k, 2).value = "'" & MyP.State
                Cells(k, 3).value = "'" & Locns(0, k - CurRow)
                Cells(k, 3).AddComment Locns(1, k - CurRow) & " " & MyP.UnitArea
                Range(Cells(k, 1), Cells(k, 3)).BorderAround Weight:=xlThin
              End If
            End If
            Select Case NextPipeCharacter(AtcoLaunch1.ComputeRead)
              Case "P"
                While NextPipeCharacter(AtcoLaunch1.ComputeRead) <> "R"
                  DoEvents
                Wend
              Case "C"
                ImportDone = True
                MyMsgBox.Show "The export was stopped on the " & SelCatRec("Description") _
                    & vbCrLf & " category on " & MyP.UnitArea & " " & Locns(0, k - CurRow) & ".", _
                    "Import not successful", "+-&OK"
                Err.Raise 999
            End Select
            'Fill in data values
            lCol = sCol + fldOffset
            For l = 1 To nFlds
              dataID = fldNameRec("ID")
              If Len(Cells(4, lCol).id) > 0 Then 'valid field ID for this column
                fldID = CLng(Cells(4, lCol).id)
              Else 'column doesn't have a field ID, problem!
                fldID = dataID + 1
              End If
              While fldID < dataID 'move through header IDs to match current data ID
                lCol = lCol + 1
                If Len(Cells(4, lCol).id) > 0 Then
                  fldID = CLng(Cells(4, lCol).id)
                Else
                  fldID = dataID + 1
                End If
              Wend
              If fldID = dataID Then 'found matching IDs
                Border(lCol) = MyP.Required(dataID, ReqSt)
                If Border(lCol) = "red" Then
                  Cells(k, lCol).BorderAround ColorIndex:=3, Weight:=xlThin
                ElseIf Border(lCol) = "blue" Then
                  Cells(k, lCol).BorderAround ColorIndex:=23, Weight:=xlThin
                Else
                  Cells(k, lCol).BorderAround , Weight:=xlHairline
                End If
                If k = CurRow Then Range(Cells(k - 1, lCol), Cells(k + numAreas, lCol)).NumberFormat = Cells(5, lCol).NumberFormat
                If Not IsNull(dataRec("Value")) Then Cells(k, lCol).value = dataRec("Value")
                If lCol > mxCol Then mxCol = lCol
              End If
              fldNameRec.MoveNext
              dataRec.MoveNext
            Next l
          End With
        Next k
        If ExpType = "co" Then
          'fill state total row
          k = CurRow - 1
          If Len(Cells(k, 1)) = 0 Then
            Cells(k, 1).value = "'" & MyP.stateCode & "000"
            Cells(k, 1).AddComment MyP.State
            Cells(k, 1).BorderAround Weight:=xlThin
          End If
          For lCol = 2 + fldOffset To mxCol
            If Len(Cells(CurRow, lCol)) > 0 Then 'values to sum up
              Cells(k, lCol).value = Application.WorksheetFunction.Sum(Range(Cells(CurRow, lCol), Cells(CurRow + numAreas, lCol)))
              If Border(lCol) = "red" Then
                Cells(k, lCol).BorderAround ColorIndex:=3, Weight:=xlThin
              ElseIf Border(lCol) = "blue" Then
                Cells(k, lCol).BorderAround ColorIndex:=23, Weight:=xlThin
              Else
                Cells(k, lCol).BorderAround , Weight:=xlHairline
              End If
            End If
          Next lCol
        Else
          k = CurRow
        End If
        'Format sheet and set data validation
        Columns("A:AB").AutoFit
        Columns(1).ColumnWidth = 15
        With Range(Cells(k, sCol + fldOffset), Cells(CurRow + numAreas, mxCol)).Validation
          .Add Type:=xlValidateDecimal, _
              AlertStyle:=xlValidAlertStop, _
              operator:=xlBetween, Formula1:="0", Formula2:="99999.99"
          .ErrorMessage = "You must enter a number between 0 and 99999.99"
          .ErrorTitle = "Bad Data Value"
        End With
        'Check to ensure -999 < HY-OfPow < 99999.99
        If SelCatRec("Name") = "HY" Then
          If opt = 0 Then
            With Range(Cells(k, 6), Cells(CurRow + numAreas, 6)).Validation
              .Modify Type:=xlValidateDecimal, _
                  AlertStyle:=xlValidAlertStop, _
                  operator:=xlBetween, Formula1:="-999", Formula2:="99999.99"
              .ErrorMessage = "Values for HY-OfPow must be between -999 and 99999.99"
            End With
          Else
            With Range(Cells(k, 5), Cells(CurRow + numAreas, 5)).Validation
              .Modify Type:=xlValidateDecimal, _
                  AlertStyle:=xlValidAlertStop, _
                  operator:=xlBetween, Formula1:="-999", Formula2:="99999.99"
              .ErrorMessage = "Values for HY-OfPow must be between -999 and 99999.99"
            End With
          End If
        End If
      End With
    Else
      mxCol = mxCol + nFlds
    End If
    fldNameRec.Close
    dataRec.Close
    If lCatID = 17 Or lCatID = 18 Then 'export irrig fields from other irrigation categories
      lCatID = lCatID + 1
      fldOffset = mxCol - sCol + 1 'FldOffset + NFlds
      GoTo MoreFlds
    Else
      fldOffset = 0
      SelCatRec.MoveNext
    End If
  Next i
  Application.Worksheets(1).Select
  SelCatRec.Close
  CurRow = CurRow + numAreas + 1
  
  Exit Sub
ErrHndlr:
  If Err.Number = 999 Then
    ExpType = "done"
  Else
    MsgBox "The National Export report encountered a problem during production:" & vbCrLf & vbCrLf & _
      Err.Description, vbCritical, "ERROR"
  End If
End Sub

