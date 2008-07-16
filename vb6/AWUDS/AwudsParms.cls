VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AwudsParms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Copyright 2000 by AQUA TERRA Consultants
  '##MODULE_NAME AwudsParms
  '##MODULE_DATE December 12, 2001
  '##MODULE_AUTHOR Jack Kittle and Robert Dusenbury, AQUA TERRA CONSULTANTS
  '##MODULE_SUMMARY Opens the State database, General.mdb, and Categories.mdb _
    and maintains 15 attributes describing these databases as well as _
    data selections made by the user via the GUI.
  '##MODULE_REMARKS <P>To use this class, first create an object _
  of type AwudsParms. Properties can be assigned then retrieved using the Set _
  and Get routines of the respective properties.</P><P>If any of the 3 _
  databases have not been set when the property containing their name is _
  retrieved, that database will automatically be opened in the Get routine.</P> _
  <P>There are 3 Boolean Functions to ensure that the user has made valid _
    data selections. A fourth Boolean Function determines if a _
    specified data field is required.</P> _
  <P>There is one subroutine to initialize the variables pStateCode and _
    pUserOpt to nil values. There are 2 other subroutines that close the _
    State database and Categories.mdb, respectively.</P>
  '
  ' <><><><><><>< Objects Used by Section ><><><><><><><>
  '
  '##SUMMARY Variable Object -- State database object
Private pStateDB As Database
  '##SUMMARY Variable Object -- General.mdb object
Private pGenDB As Database
  '##SUMMARY Variable Object -- Categories.mdb object
Private pCatDB As Database
  '
  ' <><><><><><>< Variables Section ><><><><><><><>
  '
  '##SUMMARY Variable String -- contains state name
Private pState As String
  '##SUMMARY Variable String -- contains 'County', 'Huc - 8', 'Huc - 4', 'state', or 'Aquifer'
Private pUnitArea As String
  '##SUMMARY Variable String -- contains 'County', 'Huc - 8', 'Huc - 4', 'state', or 'Aquifer'; used when comparing by area
Private pUnitArea2 As String
  '##SUMMARY Variable String -- contains data storage option for current area/year selection.
Private pDataOpt As String
  '##SUMMARY Variable String -- contains data storage option for 2nd area/year selection; used when comparing by area or by year
Private pDataOpt2 As String
  '##SUMMARY Variable String -- contains data dictionary used for current area/year selection.
Private pDataDict As String
  '##SUMMARY Variable String -- contains data dictionary used for 2nd area/year selection; used when comparing by area or by year
Private pStateCode As String
  '##SUMMARY Variable String -- contains 'county', 'huc', 'state', or 'aquifer' as table name in General.mdb
Private pAreaTable As String
  '##SUMMARY Variable Long -- contains index for MainOpt OptionButton on 2nd tab of main form
Private pUserOpt As Long
  '##SUMMARY Variable Long -- number of areas selected on 3rd tab of main form
Private pNumAreas As Long
  '##SUMMARY Variable Long -- contains 4-digit year of data option on 3rd tab of main form
Private pYear1Opt As Long
  '##SUMMARY Variable Long -- contains 4-digit year of 2nd data option on 3rd tab; used when comparing by year
Private pYear2Opt As Long
  '##SUMMARY Variable Byte -- length of string value of pAreaTable
Private pLength As Byte
  '##SUMMARY Variable Byte -- length of string value of 2nd pAreaTable when comparing by year/area
Private PLength2 As Byte
  '##SUMMARY Variable String -- contains name of '____Fields_' table in Categories.mdb
Private pYearFields As String
  '##SUMMARY Variable String -- contains full pathname of General.mdb
Private pGenDBName As String
  '##SUMMARY Variable String -- contains full pathname of Categories.mdb
Private pCatDBName As String
  '##SUMMARY Variable String -- contains full pathname of State database
Private pStateDBName As String

Public Property Get GenDB() As Database
Attribute GenDB.VB_Description = "Opens General.mdb if not already opened."
  '##SUMMARY Opens General.mdb if not already opened.
  If pGenDB Is Nothing Then
    pGenDBName = AwudsDataPath & "General.mdb"
    Set pGenDB = OpenDatabase(pGenDBName, , True)
  End If
  Set GenDB = pGenDB
End Property

Public Property Get CatDB() As Database
Attribute CatDB.VB_Description = "Opens Categories.mdb if not already opened."
  '##SUMMARY Opens Categories.mdb if not already opened.
  If Len(pCatDBName) = 0 Then
    pCatDBName = AwudsDataPath & "Categories.mdb"
    Set pCatDB = OpenDatabase(pCatDBName, , True)
  End If
  Set CatDB = pCatDB
End Property

Public Function StateValid() As Boolean
Attribute StateValid.VB_Description = "Determines whether user has successfully chosen a state."
  '##SUMMARY Determines whether user has successfully chosen a state.
  '##RETURNS True if the variable pState has been set.
  If Len(pState) > 0 Then
    StateValid = True
  Else
    StateValid = False
  End If
End Function

Public Property Let State(ByVal NewValue As String)
Attribute State.VB_Description = "State name as derived from the 'state_nm' field of the 'state' table in General.mdb."
  '##SUMMARY State name as derived from the 'state_nm' field of _
    the 'state' table in General.mdb.
  '##PARAM NewValue (I) String to which State will be set
  pState = NewValue
End Property
Public Property Get State() As String
  State = pState
End Property

Public Sub StateStuff(ByVal newState As String, ByVal newStateCode As String)
Attribute StateStuff.VB_Description = "Sets pStateCode and pState to the incoming values. The values can be retrieved later via the StateCode and State properties."
  '##SUMMARY Sets pStateCode and pState to the incoming values. The values _
    can be retrieved later via the StateCode and State properties.
  '##PARAM newState (I) Name of the State as shown in the 'state_nm' _
    field of the 'state' table in General.mdb.
  '##PARAM newStateCode (I) FIPS code of the State as shown in the 'state_cd' _
    field of the 'state' table in General.mdb.
  Dim dbPath As String 'full pathname of the state database.
  If Len(pState) > 0 And newStateCode <> "00" Then
    pState = ""
  End If
  'open state db with county/huc/aquifer data
  dbPath = AwudsDataPath & newStateCode & ".mdb"
  If Len(Dir(dbPath)) > 0 Then
    pStateCode = newStateCode
    pState = newState
  ElseIf newStateCode = "00" Then
    pStateCode = newStateCode
    pState = newState
  Else
    pStateCode = ""
  End If
End Sub

Public Property Get stateDB() As Database
Attribute stateDB.VB_Description = "Opens the State database if not already opened then makes links to external tables."
  '##SUMMARY Opens the State database if not already opened then _
    makes links to external tables.
  Dim i As Long 'index for loop thru State database tables
  Dim dbPath As String 'full pathname of external table to which State DB has link
  If Len(pStateDBName) = 0 Then
    pStateDBName = AwudsDataPath & pStateCode & ".mdb"
    Set pStateDB = OpenDatabase(pStateDBName, False, False, "MS Access; pwd=B7Q6C9B752")
    For i = 0 To pStateDB.TableDefs.Count - 1
      dbPath = pStateDB.TableDefs(i).Connect
      If Len(dbPath) > 0 Then
        pStateDB.TableDefs(i).Connect = ";DATABASE=" & AwudsDataPath & FilenameNoPath(dbPath)
      End If
    Next i
  End If
  Set stateDB = pStateDB
End Property

Public Sub StateDBClose()
Attribute StateDBClose.VB_Description = "Closes State database."
  '##SUMMARY Closes State database.
  If Len(pStateDBName) > 0 Then pStateDB.Close
  pStateDBName = ""
End Sub

Public Sub CatDBClose()
Attribute CatDBClose.VB_Description = "Closes Categories.mdb."
  '##SUMMARY Closes Categories.mdb.
  If Len(pCatDBName) > 0 Then pCatDB.Close
  Set pCatDB = Nothing
  pCatDBName = ""
End Sub

Public Property Get stateCode() As String
Attribute stateCode.VB_Description = "State FIPS code as stored in the 'state_cd' field of the 'state' table in General.mdb."
  '##SUMMARY State FIPS code as stored in the 'state_cd' field of _
    the 'state' table in General.mdb.
  stateCode = pStateCode
End Property

Public Property Get UnitArea() As String
  '##SUMMARY Name of the land division chosen by the user; i.e., _
    'County', 'Huc - 8', 'Huc - 4', 'Aquifer', or 'State'.
  UnitArea = pUnitArea
End Property
Public Property Let UnitArea(ByVal NewValue As String)
  pUnitArea = NewValue
End Property

Public Property Get UnitArea2() As String
  '##SUMMARY Name of the 2nd land division chosen by the user _
    when comparing data by area.
  UnitArea2 = pUnitArea2
End Property
Public Property Let UnitArea2(ByVal NewValue As String)
  pUnitArea2 = NewValue
End Property

Public Property Get DataOpt() As Long
  '##SUMMARY Index representing data storage option for current area/year selection (1-6).
  DataOpt = pDataOpt
End Property
Public Property Let DataOpt(ByVal NewValue As Long)
  pDataOpt = NewValue
End Property

Public Property Get DataOpt2() As Long
  '##SUMMARY Index representing data storage option for 2nd area/year selection (1-6) _
    when comparing by year or by area.
  DataOpt2 = pDataOpt2
End Property
Public Property Let DataOpt2(ByVal NewValue As Long)
  pDataOpt2 = NewValue
End Property

Public Property Get DataDict() As Long
  '##SUMMARY Year representing data dictionary for current area/year selection _
      (1995, 2000, or 2005).
  DataDict = pDataDict
End Property
Public Property Let DataDict(ByVal NewValue As Long)
  pDataDict = NewValue
End Property

Public Property Get Length() As Byte
Attribute Length.VB_Description = "Length of the&nbsp;area_cd field of the appropriate table in General.mdb; i.e., county = 3, huc = 4 or 8, Aquifer = 10, and State = 2.  "
  '##SUMMARY Length of the&nbsp;<EM>area</EM>_cd field of the _
    appropriate table in General.mdb; i.e., county = 3, huc = 4 or 8, _
    Aquifer = 10, and State = 2.
  Length = pLength
End Property
Public Property Let Length(NewLength As Byte)
  pLength = NewLength
End Property

Public Property Get length2() As Byte
Attribute length2.VB_Description = "Retrieves the length of the 2nd&nbsp;area_cd field of the appropriate table in General.mdb when comparing data by area. Set to the length of the 2nd&nbsp;area_cd field of the appropriate table in General.mdb when comparing data by area."
  '##SUMMARY Retrieves the length of the 2nd&nbsp;<EM>area</EM>_cd field of the _
    appropriate table in General.mdb when comparing data by area.
  length2 = PLength2
End Property
Public Property Let length2(NewLength2 As Byte)
  '##SUMMARY Set to the length of the 2nd&nbsp;<EM>area</EM>_cd field of the _
    appropriate table in General.mdb when comparing data by area.
  PLength2 = NewLength2
End Property

Public Property Get UserOptValid() As Boolean
Attribute UserOptValid.VB_Description = "Boolean check to ensure the user made a valid Data Operation selection on the 2nd tab of main form."
  '##SUMMARY Boolean check to ensure the user made a valid Data Operation _
    selection on the 2nd tab of main form.
  If pUserOpt > 0 And pUserOpt <> 2 And pUserOpt <> 12 Then
    UserOptValid = True
  Else
    UserOptValid = False
  End If
End Property

Public Property Let UserOpt(ByVal newUserOpt As Long)
Attribute UserOpt.VB_Description = "Index of Data Operation selection made from MainOpt OptionButton array on 2nd tab of main form."
  '##SUMMARY Index of Data Operation selection made from <EM>MainOpt</EM> _
    OptionButton array on 2nd tab of main form.
  '##PARAM newUserOpt (I) Index to which UserOpt will be set.
  pUserOpt = newUserOpt
End Property
Public Property Get UserOpt() As Long
  UserOpt = pUserOpt
End Property

Public Property Let numAreas(newNumAreas As Long)
Attribute numAreas.VB_Description = "Number of Area selections made with lstArea AtcoSelectListSorted control on 3rd tab of main form."
  '##SUMMARY Number of Area selections made with <EM>lstArea</EM> _
    AtcoSelectListSorted control on 3rd tab of main form.
  '##PARAM newNumAreas (I) Number to which NumAreas will be set.
  pNumAreas = newNumAreas
End Property
Public Property Get numAreas() As Long
  numAreas = pNumAreas
End Property

Public Property Let AreaTable(newAreaTable As String)
Attribute AreaTable.VB_Description = "Name of _____Data table in State database; either County, Huc, or Aquifer."
  '##SUMMARY Name of <EM>_____Data</EM> table in State database; _
    either <EM>County</EM>, <EM>Huc</EM>, or <EM>Aquifer</EM>.
  '##PARAM newAreaTable (I) String to which AreaTable will be set.
  pAreaTable = newAreaTable
End Property
Public Property Get AreaTable() As String
  AreaTable = pAreaTable
End Property

Public Property Let YearFields(newYearFields As String)
Attribute YearFields.VB_Description = "Name of ____Fields_ table in Categories.mdb. Prefix is 1995 or 2000 depending on year of data selection. Suffix is 0 for year 1995 data, 1-6 for 2000 (depending on data storage options for that state), and A for Aquifer data."
  '##SUMMARY Name of <EM>____Fields_</EM> table in Categories.mdb. _
    Prefix is <EM>1995</EM> or <EM>2000</EM> depending on year of data _
    selection. Suffix is <EM>0</EM> for year 1995 data, <EM>1-6</EM> for 2000 _
    (depending on data storage options for that state), and <EM>A</EM> _
    for Aquifer data.
  '##PARAM newYearFields (I) String to which YearFields will be set.
  pYearFields = newYearFields
End Property
Public Property Get YearFields() As String
  YearFields = pYearFields
End Property

Public Function YearValid() As Boolean
Attribute YearValid.VB_Description = "Determines whether the user has selected a valid year when adding a new year of data."
  '##SUMMARY Determines whether the user has selected a valid year when adding a new year of data.
  '##RETURNS True if user has selected a year after 1800.
  If pUserOpt = 10 Then
    If pYear2Opt > 1800 Then
      YearValid = True
    Else
      YearValid = False
    End If
  Else
    If pYear1Opt > 1800 Then
      YearValid = True
    Else
      YearValid = False
    End If
  End If
End Function

Public Property Let Year1Opt(newYear1Opt As Long)
Attribute Year1Opt.VB_Description = "Year that user selected from lstYears ListBox control on 3rd tab of main form."
  '##SUMMARY Year that user selected from <EM>lstYears</EM> ListBox control _
    on 3rd tab of main form.
  '##PARAM newYear1Opt (I) Number to which Year1Opt will be set.
  pYear1Opt = newYear1Opt
End Property
Public Property Get Year1Opt() As Long
  Year1Opt = pYear1Opt
End Property

Public Property Let Year2Opt(newYear2Opt As Long)
Attribute Year2Opt.VB_Description = "2nd year that user selected from lstYears ListBox control on 3rd tab of main form when comparing 2 years of data."
  '##SUMMARY 2nd year that user selected from <EM>lstYears</EM> ListBox _
    control on 3rd tab of main form when comparing 2 years of data.
  '##PARAM newYear2Opt (I) Number to which Year2Opt will be set.
  pYear2Opt = newYear2Opt
End Property
Public Property Get Year2Opt() As Long
  Year2Opt = pYear2Opt
End Property

Private Sub Class_Initialize()
  '##SUMMARY Sets certain object variables to nil values.
  pStateCode = ""
  pUserOpt = 0
End Sub

Public Function Required(Field As Long, CodeSt As Long) As String
Attribute Required.VB_Description = "Determines if a specified data field is required for a specified state. All fields are mandatory for pre-2000 data.   2000 and later data are more complicated. The Mandatory field in the AWUDS data dictionary Categories.mdb table defined by property YearFields of this class contains information on whether a field is required or not. The value in Mandatory for the specified field is assigned to the local codeFld variable, which can be 0-4.&nbsp; When&nbsp;codeFld = 4, certain fields are required depending on the value of Required from the state table in General.mdb, which is passed into this function as codeSt. &nbsp;     codeFld Field Required?  0 no  1 County and Aquifer"
' ##SUMMARY <P>Determines if a specified data field is required for a specified state. _
          All fields are mandatory for pre-2000 data.</P> _
          <P></P> <P></P> _
          <P>2000 and later data are more complicated. The <EM>Mandatory00</EM> _
          field in the AWUDS data dictionary Categories.mdb table defined by property _
          <EM>YearFields</EM> of this class contains information on whether a field is _
          required or not. The value in <EM>Mandatory00</EM> for the specified field is _
          assigned to the local <EM>codeFld</EM> variable, which can be 0-4.&nbsp; _
          When&nbsp;<EM>codeFld</EM> = 4, certain fields are required depending on the _
          value of <EM>Required</EM> from the <EM>state</EM> table in General.mdb, which _
          is passed into this function as <EM>codeSt</EM>.</P> _
          <P></P> <P></P> _
          <P>2005 data is covered by the <EM>Mandatory05</EM> field in the data dictionary. _
          There are 2 types of mandatory fields for 2005 data: 1 = required, but nulls allowed; _
          and 2 = required, no nulls. There is no variation between states.<P>&nbsp;</P> _
          <P> <TABLE cellSpacing=0 cellPadding=10 border=1> _
          <TBODY vAlign=top> _
          <TR> <TH> <EM>codeFld</EM> <TH>Field Required? _
          <TR> <TD>0 <TD>no <TR> <TD>1 <TD>County and Aquifer
' ##SUMMARY
' ##SUMMARY _
          <TD>2 _
          <TD>County only _
          <TR> _
          <TD>4 _
          <TD><PRE><PRE>CodeSt=1: 146,147,149,150 required</PRE><PRE>CodeSt=2: 171,174 _
          required</PRE><PRE>CodeSt=3: 183,186 required</PRE><PRE>CodeSt=4: Options 1 _
          &amp; 2 combined</PRE><PRE>CodeSt=5: Options 1 &amp; 3 _
          combined</PRE><PRE><PRE>CodeSt=6: Options 2 &amp; 3 combined<PRE>CodeSt=7: _
          Options 1, 2 &amp; 3 _
          combined</PRE></PRE></PRE></PRE></TD></TR></TBODY></TABLE></P> _
          <P>&nbsp;</P> _
          <P>5 fields are required for Aquifer data: Field IDs 5, 57, 195, 278, and _
          288.</P>
  '##PARAM Field (I) Value of attribute FieldID from Data Dictionary as defined in _
          any of the Categories.mdb tables named ____Fields_; or value of attribute _
          ID as defined in Categories.mdb tables Field_ and AllFields.
  '##PARAM CodeSt (I) Value of <EM>Required</EM> from <EM>state</EM> table in General.mdb _
          for specified state.
  '##Returns True if data field is required for selected state.
  '##HISTORY PR 18637, 5/17/2007, prhummel Commented out conditional at the start _
             of this module since it was setting all fields to "red" (required) _
             if State totals (code 000) was selected along with other areas
  Dim fldRec As Recordset 'Fields recordset from table defined by property YearFields
  Dim codeFld As Byte 'Number read from Mandatory00 or Mandatory05 field of fldRec; indicates required fields
  
  'conditional removed, see History entry from 5/17/2007, prhummel
'  If Left(frmAwuds2.lstArea.RightItem(0), 3) = "000" Then 'state total field
'    If Left(MyP.UnitArea, 3) = "HUC" Then
'      Required = "white"
'    Else
'      Required = "red"
'    End If
'    Exit Function
'  End If
  Set fldRec = stateDB.OpenRecordset(Me.YearFields, dbOpenDynaset)
  fldRec.FindFirst "FieldID = " & Field
  If MyP.DataDict = 2005 Then
    If Left(MyP.UnitArea, 3) = "HUC" Or Left(MyP.UnitArea, 3) = "Aqu" Then
      Required = "white"
      Exit Function
    End If
    codeFld = fldRec("Mandatory05")
  ElseIf MyP.DataDict = 2000 Then
    If Left(MyP.UnitArea, 3) = "HUC" Then
      Required = "white"
      Exit Function
    End If
    codeFld = fldRec("Mandatory00")
  Else
    Required = "red"
    Exit Function
  End If
  If codeFld = 0 Then
    Required = "white"
  Else
    If MyP.DataDict = 2005 Then
      If MyP.Length = 3 Or MyP.Length = 2 Then
        If codeFld = 1 Then
          Required = "blue"
        ElseIf codeFld = 2 Then
          Required = "red"
        End If
      Else
          Required = "white"
      End If
    Else
      If (codeFld = 1 And (pLength = 2 Or pLength = 3 Or pLength = 10)) Or _
         (codeFld = 2 And (pLength = 2 Or pLength = 3)) Or _
         (codeFld = 3 And pLength = 3) Then
        Required = "red"
      ElseIf codeFld = 4 Then  'Required for some states.
        If CodeSt > 0 Then
          If ((Field = 146 Or Field = 147 Or Field = 149 Or Field = 150) _
                And (CodeSt = 1 Or CodeSt = 4 Or CodeSt = 5 Or CodeSt = 7)) Or _
              ((Field = 171 Or Field = 174) _
                And (CodeSt = 2 Or CodeSt = 4 Or CodeSt = 6 Or CodeSt = 7)) Or _
              ((Field = 183 Or Field = 186) _
                And (CodeSt = 3 Or CodeSt = 5 Or CodeSt = 6 Or CodeSt = 7)) Then
            Required = "red"
          Else
            Required = "white"
          End If
        Else
          Required = "white"
        End If
      Else
        Required = "white"
      End If
    End If
  End If
End Function

