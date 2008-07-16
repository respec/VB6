Attribute VB_Name = "modAwudsArray"
Option Explicit
'Copyright 2003 by AQUA TERRA Consultants

'##MODULE_NAME modAwudsArray
'##MODULE_DATE December 4, 2003
'##MODULE_AUTHOR Robert Dusenbury and Mark Gray of AQUA TERRA CONSULTANTS
'##MODULE_SUMMARY Contains set of global variable declarations and 3 routines associated _
          with data management.
'##MODULE_REMARKS 2 of the routines in this module populate parallel data arrays that _
          store 4 pieces of information for each data field. Sub _
          <EM>PopulateFieldArrays</EM> populates 3 arrays that store the field name, _
          field formula, and whether the field value has a formula (i.e., is a product of _
          other data fields). Sub <EM>PopulateDataArray</EM> populates a 2-dimensional _
          array with a full data set for each location. If the user has selected the _
          <EM>Compare Data for 2 Years</EM> report option, a twin set of these 4 arrays _
          is created for the 2nd year of data. _
          <P></P> _
          <P>The 3rd routine, <EM>EvalArray</EM>,&nbsp;is a function that retrieves a _
          data value given a field ID and location index. This function&nbsp;calls _
          itself&nbsp;reiteratively if the data field is a function of other data _
          fields.</P>
'
' <><><><><><>< Global Variables Section ><><><><><><><>
'
'##SUMMARY Global Variable Double -- 2-D array containing all data values for selected state: _
  1 = location index, 2 = field index
Global DataArray() As Double
'##SUMMARY Global Variable Double -- 2-D array containing 2nd year of data for selected state: _
  1 = location index, 2 = field index
Global DataArray2() As Double
'##SUMMARY Global Variable String -- Array of field names for each field in select data dictionary
Global FldName() As String
'##SUMMARY Global Variable String -- Array of field names for 2nd year
Global FldName2() As String
'##SUMMARY Global Variable String -- Array of field formulas for each field in select data dictionary
Global FldFormula() As String
'##SUMMARY Global Variable String -- Array of field formulas for 2nd year
Global FldFormula2() As String
'##SUMMARY Global Variable Boolean -- Array for each field in select data dictionary; True if value is product of other data fields
Global FldHasFormula() As Boolean
'##SUMMARY Global Variable Boolean -- Array for 2nd year of data; True if value is product of other data fields
Global FldHasFormula2() As Boolean
'##SUMMARY Global Variable String -- Text string that is written to the Excel output, at _
  the top of the report output, used to describe the contents of the report (one line header _
  title for the report).
Global RepTitle As String
'##SUMMARY Global Variable String -- Previous year selected on 3rd tab of main form
Global OldYr As String
'##SUMMARY Global Variable String -- Label addendum for <EM>Instructions</EM> box on main form
Global operation As String
'##SUMMARY Global Variable String -- Name of data table in state DB being used by current operation: _
  either 'LastEdit',  'LastImport',  'LastExport', or 'LastReport'.
Global File As String
'##SUMMARY Global Variable String -- Part of sql statement that determines which _
    categories are used , for example: _
    <br>Cats = "And ([Category1].ID=16" will fetch information for 'Irrigation, Total'.
Global Cats As String
'##SUMMARY Global Variable String -- Part of sql statement that determines _
    which areas are used, for example: _
    <br>Areas = " AND len(trim(CountyData.Location)) = 3" retrieves all _
    counties; the length of county codes = 3. _
    <br>Areas = " AND (CountyData.Location='001' )" retrieves a specific _
    county by county area code.
Global Areas As String
'##SUMMARY Global Variable String -- Name of '<EM>Fieldx</EM> table in the _
  Categories.mdb data dictionary where x is the value A, 0, 1, 2, 3, 4, 5, or 6 _
  See the Categories documentation for table descriptions.
Global FieldTable As String
'##SUMMARY Global Variable String -- Name of '<EM>Fieldx</EM> table in the _
  Categories.mdb data dictionary where x is the value A, 0, 1, 2, 3, 4, 5, or 6. _
  Used for 2nd year of data for 'Compare Data for 2 Years' report. _
  See the Categories documentation for table descriptions.
Global FieldTable2 As String
'##SUMMARY Global Variable String -- Name of selected unit area (i.e., county, huc, or aquifer)
Global TableName As String
'##SUMMARY Global Variable String -- Name of selected unit area (i.e., county, huc, or aquifer) for 2nd area
Global TableName2 As String
'##SUMMARY Global Variable String -- Name of data table in National DB used by the current operation: _
  either 'LastEdit', 'LastImport', 'LastExport', or 'LastReport'.
Global NewTable As String
'##SUMMARY Global Variable String -- Part of sql statement that determines which fields are excluded
Global OmitFlds As String
'##SUMMARY Global Variable String -- Part of sql statement that determines which years are used
Global Years As String
'##SUMMARY Global Variable String -- Name of Excel file used to provide header for report
Global HeaderFile As String
'##SUMMARY Global Variable String -- Name of categories table to be used in queries
Global CatTable As String
'##SUMMARY Global Variable Long -- Number of data fields in data dictionary
Global NFields As Long
'##SUMMARY Global Variable Long -- Number of data fields in data dictionary used by 2nd year of data
Global NFields2 As Long
'##SUMMARY Global Variable Long -- Number of rows in report header
Global HeaderRows As Long
'##SUMMARY Global Variable Long -- Option for labeling HUCs/counties/aquifers: 0 = name, 1 = code, 2 = both
Global AreaID As Long
'##SUMMARY Global Variable Long -- Numeric code defining which data fields are required for this state: Integer Code 0-7 for _
  state's particular data requirements: 0 = not required; 1 = Mining fields 146,147,149,150; _
  2 = Livestock fields 171,174; 3 = Aquaculture fields 183,186; 4 = both 1 & 2; 5 = both 1 & 3; _
  6 = both 2 & 3; 7 = All 1,2 & 3
Global ReqSt As Long
'##SUMMARY Global Variable Long -- In data flds w/ formulas, tracks whether other data flds used by formula are available
Global Z As Long
'##SUMMARY Global Variable String -- 2-D array with code, name, and 'code - name' in 1st dimension, location in 2nd. _
  For example, counties in Delaware: LocnArray(0, 1) = "001", LocnArray(1, 1) = "Kent", LocnArray(0, 1) = "Kent - 001"
Global LocnArray() As String
'##SUMMARY Global Variable String -- Array of values for data fields in category being edited by user on 5th tab
Global PreEditVal() As String
'##SUMMARY Global Variable Object -- Recordset of total population for user-selected areas; 1st field of data dictionary
Global TotalPopRec As Recordset
'##SUMMARY Global Variable Object -- Recordset of all data fields
Global AllFldRec As Recordset
'##SUMMARY Global Variable Boolean -- True if user-selections on 3rd tab of main form are OK
Global DataSelOK As Boolean
'##SUMMARY Global Variable Boolean -- True if user has verified he wishes to delete all data values in category
Global DeleteOK As Boolean
'##SUMMARY Global Variable Boolean -- True if aggregating HUC-8s to HUC-4s
Global AggregateHUCs As Boolean
'##SUMMARY Global Variable Boolean -- True if import of data from Excel spreadsheet is complete
Global ImportDone As Boolean
'##SUMMARY Global Variable Boolean -- True if user has verified he wishes to overwrite data values in category
Global EditOK As Boolean
'##SUMMARY Global Variable Boolean -- True if user has selected national database instead of individual state
Global NationalDB As Boolean
'##SUMMARY Global Variable Boolean -- True if data needed to execute formula is not available
Global NoRec As Boolean
'##SUMMARY Global Variable Boolean -- True if user has selected an new Excel workbook for import
Global NewImpFile As Boolean
'##SUMMARY Global Variable Boolean -- True if Irrigation category is divided into crops and golf
Global IRinTwo As Boolean
'##SUMMARY Global Variable Boolean -- True if Irrigation category is divided into crops and golf; _
    used for 2nd data set when doing Compare by Area or Compare 2 Years report.
Global IRinTwo2 As Boolean
'##SUMMARY Global Variable Boolean -- True if user has selected <EM>Compare State Totals for by Area</EM> operation
Global TwoAreas As Boolean
'##SUMMARY Global Variable Boolean -- True if user has selected <EM>Compare Data for 2 Years</EM> operation
Global TwoYears As Boolean
'##SUMMARY Global Variable Boolean -- Used to determine when to inform the user that there _
  is no data entered for that category. If an asterisk is preceeding a data category _
  description (* reclaimed wastewater, in Mgal/d) then the value of asterisk is set to _
  TRUE and the following text is displayed to the user: _
  <br>An * before the listing indicates there is no data for that category
Global Asterisk As Boolean
'##SUMMARY Variable Object -- Instance of an Excel application object
Global XLApp As Excel.Application
'##SUMMARY Variable Object -- Instance of an Excel workbook object used for report
Global XLBook As Workbook
'##SUMMARY Variable Object -- Instance of an Excel workbook object used to provide header for report
Global HeaderBook As Workbook
'##SUMMARY Variable Object -- Instance of an Excel worksheet object
Global XLSheet As Worksheet
'##SUMMARY Variable Object -- Instance of an Excel worksheet object used to provide header for report
Global HeaderSheet As Worksheet
'##SUMMARY Variable String -- File extension filter for Excel _
                              ("(*.xls)|*.xls" for 2003 and earlier, _
                               "(*.xlsx)|*.xlsx" for 2007)
Global XLFileFilter As String
'##SUMMARY Variable String -- File extension for Excel _
                              (".xls" for 2003 and earlier, ".xlsx" for 2007)
Global XLFileExt As String
'##SUMMARY Variable Integer -- File format number for Excel _
                               (-4143 for 97-2003, 51 for 2007)
Global XLFileFormatNum As Integer

Sub PopulateFieldArrays(TableName As String)
Attribute PopulateFieldArrays.VB_Description = "Fills in array values for name and formula of water-use data fields."
' ##SUMMARY Fills in array values for name and formula of water-use data fields.
' ##REMARKS Fills in twin sets of arrays if "Compare Data for 2 Years" report selected.
' ##PARAM TableName I Name of selected unit area (i.e., county, huc, or aquifer).
  Dim fldRec As Recordset 'Data then Field recordset.
  Dim id As Long          'field ID read from data dictionary
  
  If MyP.UserOpt = 10 Then  'Comparing 2 years. Create second set of arrays.
    If TableName = "Aquifer" Then
      'Aquifers have only 1 data storage option
      FieldTable = "FieldA"
      FieldTable2 = "FieldA"
    Else
      'ID data storage option for 1st and 2nd years of data, respectively
      FieldTable = "Field" & MyP.DataOpt
      FieldTable2 = "Field" & MyP.DataOpt2
    End If
    'Open recordset with all fields used by 2nd data dictionary option
    Set fldRec = MyP.stateDB.OpenRecordset(FieldTable2)
    fldRec.MoveLast
    'Record ID of final field in 2nd data dictionary and dimension arrays accordingly
    NFields2 = fldRec("ID") 'Have to look at last ID since there are gaps of unused IDs
    ReDim FldName2(1 To NFields2)
    ReDim FldFormula2(1 To NFields2)
    ReDim FldHasFormula2(1 To NFields2)
    fldRec.MoveFirst
    id = 0
    'Fill 3 parallel arrays with data field information
    While id < NFields2
      id = fldRec("ID")
      If id > 0 And id <= NFields2 Then
        FldName2(id) = fldRec("Name")
        FldFormula2(id) = Trim(fldRec("Formula"))
        If Len(FldFormula2(id)) > 0 Then FldHasFormula2(id) = True
      End If
      fldRec.MoveNext
    Wend
  End If

  'Open recordset with all fields
  Set fldRec = MyP.stateDB.OpenRecordset(FieldTable)
  fldRec.MoveLast
  'Record ID of final field and dimension arrays accordingly
  NFields = fldRec("ID") 'Have to look at last ID since there are gaps of unused IDs
  ReDim FldName(1 To NFields)
  ReDim FldFormula(1 To NFields)
  ReDim FldHasFormula(1 To NFields)
  fldRec.MoveFirst
  id = 0
  'Fill 3 parallel arrays with data field information
  While id < NFields
    id = fldRec("ID")
    If id > 0 And id <= NFields Then
      FldName(id) = fldRec("Name")
      FldFormula(id) = Trim(fldRec("Formula"))
      If Len(FldFormula(id)) > 0 Then FldHasFormula(id) = True
    End If
    fldRec.MoveNext
  Wend
End Sub

Sub PopulateDataArray()
Attribute PopulateDataArray.VB_Description = "Fills values from database into 2-D data array dimensioned by location index and water-use data field index."
' ##SUMMARY Fills values from database into 2-D data array dimensioned by location index _
          and water-use data field index.
' ##REMARKS Fills in twin data array if "Compare Data for 2 Years" report selected.
  Dim dataRec As Recordset 'contains records of all user-selected data elements
  Dim locID As String      'location ID
  Dim sql As String        'database query
  Dim fldIndex As Long     'field index for 2nd dimension of DataArray
  Dim locIndex As Long     'location index for 1st dimension of DataArray
  Dim byState As Boolean   'national database by state
  
  'If national DB and not exporting, determine if data grouped by state or other, natural boundary
  If (NationalDB And MyP.UserOpt <> 11) And Not (TableName = "HUC8" Or TableName = "Aquifer") Then byState = True
  
  'Dimension and fill data array with 'missing' values
  ReDim DataArray(0 To UBound(LocnArray, 2), 1 To NFields)
  For locIndex = 0 To UBound(LocnArray, 2)
    For fldIndex = 1 To NFields
      DataArray(locIndex, fldIndex) = -9999
    Next
  Next
  
  'Select all current-year records from table being used for report
  sql = "SELECT * FROM " & File & _
        " WHERE Date=" & MyP.Year1Opt
  Set dataRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  'Write in all data to data array
  With dataRec
    While Not .EOF
      locID = !Location
      locIndex = 0
      'Match location from DB with location in 1st dimension of data array
      If byState Then
        While LocnArray(0, locIndex) <> Left(locID, 2)
          locIndex = locIndex + 1
        Wend
      Else
        While LocnArray(0, locIndex) <> locID
          locIndex = locIndex + 1
        Wend
      End If
      'Fill data value
      If Not IsNull(!value) Then DataArray(locIndex, !FieldID) = !value
      .MoveNext
    Wend
    .Close
  End With
  
  If MyP.UserOpt = 10 Then  'Write only second year of data to second data array
    ReDim DataArray2(0 To UBound(LocnArray, 2), 1 To NFields2)
    For locIndex = 0 To UBound(LocnArray, 2)
      For fldIndex = 1 To NFields2
        DataArray2(locIndex, fldIndex) = -9999
      Next
    Next
    
    sql = "SELECT * FROM " & File & _
          " WHERE Date=" & MyP.Year2Opt
    Set dataRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    With dataRec
      While Not .EOF
        If MyP.Length = 2 Then 'National DB by state
          locID = Left(!Location, 2)
        Else
          locID = !Location
        End If
        locIndex = 0
        While LocnArray(0, locIndex) <> locID
          locIndex = locIndex + 1
        Wend
        If Not IsNull(!value) Then DataArray2(locIndex, !FieldID) = !value
        .MoveNext
      Wend
    End With
    dataRec.Close
  End If
  
End Sub

Function EvalArray(ByVal FieldID As Long, ByVal locIndex As Long, Optional Reit As Boolean) As Double
Attribute EvalArray.VB_Description = "Retrieves value of specified data element from array of water-use data previously populated from database."
' ##SUMMARY Retrieves value of specified data element from array of water-use data _
          previously populated from database.
' ##PARAM FieldID I Field ID of datum to be retrieved; 2nd dimension of data array.
' ##PARAM locIndex I Location index of datum to be retrieved;&nbsp;1st dimension of data _
          array.
' ##PARAM Reit I True if function called reiteratively.
' ##RETURNS Value of specified water-use datum.
' ##REMARKS Data element specified by&nbsp;location index&nbsp;and field ID. Function may _
          call itself reiteratively if data element is a function of other data fields.
  Dim operator As String    'mathematical operator
  Dim valSoFar As Double    'value of data element, may progress if call is iterative
  Dim thisVal As Double     'value read from DataArray, which was filled from database
  Dim fldNum As Long        'field ID read from formula, if element function of others
  Dim formula As String     'formula read from data dictionary
  Dim hasFormula As Boolean 'true if data element is required for this state

  On Error GoTo x
  
  'If data field uses formula, Z tracks whether other data flds used by formula are available
  If Not Reit Then 'not a reiterative call
    'Initialize parameters
    Z = 0
    NoRec = True
  End If
  'Check to see if field has a formula, ergo is a function of other fields
  If TwoYears Then
    hasFormula = FldHasFormula2(FieldID)
  Else
    hasFormula = FldHasFormula(FieldID)
  End If
  If Not hasFormula Then 'read value from data array
    If TwoYears Then
      thisVal = DataArray2(locIndex, FieldID)
    Else
      thisVal = DataArray(locIndex, FieldID)
    End If
    If thisVal > -9999 Then
      EvalArray = thisVal
      Z = Z + 1
    End If
  Else 'process formula iteratively
    If TwoYears Then
      formula = FldFormula2(FieldID)
    Else
      formula = FldFormula(FieldID)
    End If
    'initalize parameters before loop
    valSoFar = 0
    operator = "+"
    While Len(formula) > 0
      'Read leading field number from formula
      fldNum = StrFirstInt(formula)
      If fldNum = 1000 Then '1000 used as numerical value, not field ID
        thisVal = 1000#
      Else 'make reiterative call to this function
        thisVal = EvalArray(fldNum, locIndex, True)
      End If
      'Determine operator following field ID and perform operation
      Select Case operator
        Case "+": If Not NoRec Then valSoFar = valSoFar + thisVal
        Case "-": If Not NoRec Then valSoFar = valSoFar - thisVal
        Case "*": If NoRec Then valSoFar = 0 Else valSoFar = valSoFar * thisVal
        Case "/": If thisVal = 0 Or NoRec Then _
            valSoFar = 0 Else valSoFar = valSoFar / thisVal
      End Select
      If Len(formula) > 0 Then 'strip leading operator from formula
        operator = Left(formula, 1)
        formula = Mid(formula, 2)
      End If
    Wend
    EvalArray = valSoFar
  End If
  If Z > 0 Then NoRec = False
  Exit Function
x:
  NoRec = True
  
End Function
