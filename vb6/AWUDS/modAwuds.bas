Attribute VB_Name = "modAwuds"
Option Base 0
Option Explicit
'Copyright 2000 by AQUA TERRA Consultants

'##PROJECT_TITLE Awuds.exe
'##PROJECT_SUMMARY Primary project that drives the graphical user interface (GUI), _
  edits the State or National Database, and produces reports in Excel spreadsheet format.
'##PROJECT_OVERVIEW_BEFORE_DIAGRAM <P>AWUDS is the database management tool used by the _
        United States Geological Survey (USGS) to enter, store, and analyze water-use _
        data by county, state, HUC, or aquifer. The AwudsScreen project contains </P> _
        <P><P><STRONG>2 Forms</STRONG>:</P> _
        <P> _
        <UL> _
        <LI>frmAwuds2.frm - Provides the GUI that drives all database activity, _
        including report production. Most of the AWUDS code is in this module.</LI> _
        <LI>frmDialog.frm - Provides browser for user to locate files and _
        directories.<BR></LI></UL> _
        <P><P><STRONG>6 Standard Modules</STRONG>: _
        <P> _
        <UL> _
        <LI>ImpExp.bas - Imports and Exports data from a specially-formatted Excel _
        spreadsheet.</LI> _
        <LI>modAwuds.bas - Startup object when AWUDS is instantiated.</LI> _
        <LI>modAwudsArray.bas - Creates and evaluates arrays of database fields and _
        their values.</LI> _
        <LI>modShell.bas - Opens external files.</LI> _
        <LI>modWin32Api.bas - Contains declarations for miscellaneous Win32 API _
        functions.</LI> _
        <LI>UTILITY.bas - Contains general utility subroutines and _
        functions.<BR></LI></UL> _
        <P><P><STRONG>2 Class Modules</STRONG>:
'##PROJECT_OVERVIEW_BEFORE_DIAGRAM <UL> _
        <LI>AtcoValidateUser.cls - Allows users read-only or read-write access based on _
        domain, workgroup, and user ID.</LI> _
        <LI>AwudsParms.cls - Provides AWUDS-specific functions, properties, and _
        utilities (i.e., opening and closing the databases).<BR></LI></UL> _
        <P>and <STRONG>2 User Controls</STRONG>: _
        <P> _
        <UL> _
        <LI>ATCoSelectListSortByProp.ctl - A dual-paned (Available and Selected) list _
        box that sorts by a designated property of the listed items.</LI> _
        <LI>ATCoSelectListSorted.ctl - A dual-paned (Available and Selected) list box _
        that sorts the listed items numerically/alphabetically.<BR></LI></UL>
'##PROJECT_OVERVIEW_AFTER_DIAGRAM The User's Manual, which is included in the installation _
  package, contains documentation on terminology, user-interface architecture, database _
  structure, and more.

' ##MODULE_NAME modAwuds
' ##MODULE_DATE September 25, 2002
' ##MODULE_AUTHOR Robert Dusenbury and Mark Gray of AQUA TERRA CONSULTANTS
' ##MODULE_SUMMARY Contains set of global variable declarations and _
  startup object, <EM>Sub Main</EM>, when AWUDS is instantiated.
'
' <><><><><><>< Global Variables Section ><><><><><><><>
'
'##SUMMARY Interactive message box solicits user instruction
Global MyMsgBox As ATCoMessage
'##SUMMARY See class module
Global MyP As AwudsParms
'##SUMMARY Progress bar indicates status of external processes
Global AtcoLaunch1 As AtCoLaunch
'##SUMMARY Class object that saves/retrieves paths in/from the registry
Global Registry As ATCoRegistry
'##SUMMARY Path to directory with 'Awuds.exe'
Global ExePath As String
'##SUMMARY True if in development mode, False if running EXE
Global RunningVB As Boolean
'##SUMMARY Name of executable file (updated if in devel envir, frmAwuds2 load)
Global ExeName As String
'##SUMMARY Path to directory with databases and report templates
Global AwudsDataPath As String
'##SUMMARY Path to directory where reports are sent
Global ReportPath As String
'##SUMMARY True if '/DEBUG' set via Command member of VBA.Interaction
Global Debugging As Boolean
'##SUMMARY Stores level of user access as READ or WRITE
Global UserAccess As String

Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Sub Main()
Attribute Main.VB_Description = "Startup object when AWUDS is instantiated."
  '##SUMMARY Startup object when AWUDS is instantiated.
  '##DATE August 25, 2002
  '##AUTHOR Mark Gray and Robert Dusenbury, AQUA TERRA CONSULTANTS
  '##REMARKS Reads from registry and sets appropriate path/filenames. _
    If AwudsDataPath variable has not been set, the user will be prompted to browse _
    for the appropriate directory.  If ReportPath variable has not been set, a _
    directory titled 'AWUDSReports' will be created parallel to AwudsDataPath.
  '##ON_ERROR Errors will be raised if: AwudsDataPath is not the correct _
    path to a directory containing the required databases and report templates.
  Dim hdle As Long 'integer value returned by Function GetModuleHandle
  Dim binpos As Long 'position of '\BIN' or '\' in string ExeName
  Dim str As String * 80 'string value returned by Function GetModuleHandle
  Dim i As Long 'index to record user response
  Dim registryDataPath As String 'pathname to data directory stored in registry

  Set MyMsgBox = New ATCoMessage
  
  Set Registry = New ATCoRegistry

  'Check to see if in DEBUG mode
  If InStr(UCase(Command$), "/DEBUG") > 0 Then Debugging = True Else Debugging = False

  'Retrieve EXE name and info
  hdle = GetModuleHandle("Awuds")
  i = GetModuleFileName(hdle, str, 80)
  ExeName = UCase(Left(str, InStr(str, Chr(0)) - 1))
  
  If InStr(ExeName, "VB6.EXE") Then
    RunningVB = True
    ExeName = UCase("c:\VBExperimental\AWUDS\data")
  Else
    RunningVB = False
  End If
  
  'Reset ExePath for particular machine
  If Debugging Then MsgBox "ExeName = " & ExeName
  binpos = InStr(ExeName, "\BIN")
  If binpos < 1 Then binpos = InStrRev(ExeName, "\")
  If binpos < 1 Then
    ExePath = CurDir
  Else
    ExePath = Left(ExeName, binpos)
  End If
  If Right(ExePath, 1) <> "\" Then ExePath = ExePath & "\"
  If Debugging Then MsgBox "ExePath = " & ExePath

  Set MyP = New AwudsParms
  registryDataPath = GetSetting("AWUDS", "Defaults", "DataPath", "unknown")
  AwudsDataPath = registryDataPath
  If AwudsDataPath = "unknown" Or AwudsDataPath = "" Or Len(Dir(AwudsDataPath, vbDirectory)) = 0 Then
    MsgBox "AWUDS can not find the Data Directory." & vbCrLf & _
           "You must browse for an MDB file in a working Data Directory."
  End If
  SetDataPath

  'Retrieve Report path from registry
  ReportPath = GetSetting("AWUDS", "Defaults", "ReportPath", "unknown")
  If Len(Dir(ReportPath, vbDirectory)) = 0 Then
    ReportPath = PathNameOnly(Left(AwudsDataPath, Len(AwudsDataPath) - 1))
    ReportPath = ReportPath & "\AWUDSReports\"
  End If
  SaveSetting "AWUDS", "Defaults", "ReportPath", ReportPath
  If Len(Dir(ReportPath, vbDirectory)) = 0 Then
    'Report path does not exist; create dir parallel to Data dir to be default report directory
    MkDir ReportPath
    MsgBox "An 'AWUDSReports' directory has been created" & vbCrLf & _
           "parallel to the AWUDS data directory." & vbCrLf & _
           "All reports will be written to this new directory."
  End If
  
  'Start status monitor
  If Debugging Then MsgBox "StartMonitor: " & ExePath & "status.exe"
  Set AtcoLaunch1 = New AtCoLaunch
  AtcoLaunch1.StartMonitor (ExePath & "status.exe")
  frmAwuds2.Show
  
End Sub

Public Sub SetDataPath()
  Dim msg As String 'text for message in myMsgBox
  Dim tmp As String 'text buffer for message in myMsgBox
  Dim stRec As Recordset 'stores integer values returned from external calls
  Dim i As Long 'index in loop thru state recordset
  Dim j As Long 'path to directory with databases and report templates; stored in registry
  Dim dbPath As String 'full pathname of state DB
  Dim connectedDBpath As String 'full pathname of DB to which state DB has a link
  Dim connectedDBexpected As String 'expected full pathname of DB to which state DB has a link
  Dim stateDB As Database 'state DB object
  Dim stateName As String 'name of state stored in 'state' table of General.mdb
  Dim stateCode As String 'FIPS code of state stored in 'state' table of General.mdb
  
  Dim sBuffer As String * 255 'buffer to store user's computer name
  Dim workstationID As String 'name of user's Workstation ID
  
  'Retrieve Data path from registry
  If AwudsDataPath = "unknown" Or AwudsDataPath = "" Then
    i = MyMsgBox.Show("AWUDS has not run on this machine for this user before." & vbCr _
                   & "Before running, AWUDS needs to know the location of its data files." & vbCr _
                   & "If you do not yet have any AWUDS data files, " & vbCr _
                   & "press Exit now and download the data files before running AWUDS.", _
                     "AWUDS First Run", "+&Ok", "-&Exit")
    If i = 2 Then GoTo UnloadExit
  ElseIf MyP.State = "" And Len(Dir(AwudsDataPath)) > 0 Then 'just instantitated AWUDS
    GoTo CheckReadOnly
  End If

GetDataPath:
  
  On Error GoTo UnloadExit
  
  'Allow user to browse for Data path
  With frmDialog.cdlg
    .DialogTitle = "Locate any file in the data directory"
    .Filter = "*.mdb"
    .filename = "General.mdb"
    .ShowOpen
    i = 0
    If .filename = "" Then
      i = MyMsgBox.Show("Filename was blank.", "AWUDS data file problem", "+&Retry", "-&Exit")
    ElseIf Len(Dir(PathNameOnly(.filename) & "\General.mdb")) = 0 Then
      i = MyMsgBox.Show("The directory you specified does not contain the required files" & _
                        " 'General.mdb' and 'Categories.mdb'." & vbCrLf & _
                        "Either Reselect a different directory, or Exit and retain the previous directory", _
                        "Data Directory Problem", "&Reselect", "-&Exit")
    End If
    If i = 1 Then GoTo GetDataPath
    If i = 2 Then GoTo UnloadExit
    AwudsDataPath = PathNameOnly(.filename)
    If Left(AwudsDataPath, 2) <> "\\" Then
      If GetComputerNameA(sBuffer, 255&) > 0 Then
        workstationID = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      Else
        workstationID = "?"
      End If
      If Left(ExePath, 3) <> Left(AwudsDataPath, 3) Then
        i = MyMsgBox.Show("The path you specified does not start with '\\MachineName'," & vbCr _
                        & "which is fine if it is a local drive on '\\" & workstationID & "'." & vbCr & vbCr _
                        & "If the path leads to another computer/server on the network," & vbCr _
                        & "browse beginning with the 'Network', not the mapped network drive.", _
                          "AWUDS Data Directory Drive", "&Continue", "&Reselect", "-&Exit")
        If i = 2 Then GoTo GetDataPath
        If i = 3 Then GoTo UnloadExit
      End If
    End If
  End With
  Unload frmDialog
  
  If Len(Dir(AwudsDataPath, vbDirectory)) = 0 Then
    'Data path does not exist
    i = MyMsgBox.Show("Directory '" & AwudsDataPath & "' not found.", "AWUDS data file problem", "+&Retry", "-&Exit")
    If i = 1 Then GoTo GetDataPath
    If i = 2 Then GoTo UnloadExit
  End If
  If Right(AwudsDataPath, 1) <> "\" Then AwudsDataPath = AwudsDataPath & "\"
  i = 1
  If Len(AwudsDataPath) > i Then
    While Mid(AwudsDataPath, Len(AwudsDataPath) - i, 1) <> "\"
      i = i + 1
    Wend
  Else
    MsgBox "You must choose a directory containing AWUDS databases.", vbCritical, "Bad Directory Choice"
    GoTo GetDataPath
  End If

CheckForMDB:
  If Len(Dir(AwudsDataPath & "General.mdb")) = 0 Then
    msg = "Could not find General.mdb"
    If Len(Dir(AwudsDataPath & "Categories.mdb")) = 0 Then _
        msg = msg & " or Categories.mdb" & vbCr & "These" Else msg = msg & vbCr & "This"
ErrNoMDB:
    msg = msg & vbCr & " file must be with the data files that are expected to be in: "
    msg = msg & vbCr & AwudsDataPath
    i = MyMsgBox.Show(msg, "AWUDS data file problem", "+&Change Data Directory", "&Retry", "-&Exit")
    If i = 1 Then GoTo GetDataPath
    If i = 2 Then GoTo CheckForMDB
    If i = 3 Then GoTo UnloadExit
  End If
  If Len(Dir(AwudsDataPath & "Categories.mdb")) = 0 Then
    msg = "Could not find Categories.mdb" & vbCr & "This"
    GoTo ErrNoMDB
  End If
CheckReadOnly:
  On Error GoTo ErrorReadOnly
  
  'Check and possibly fix paths to General.mdb and Categories.mdb in the state/Nation databases
  Set stRec = MyP.GenDB.OpenRecordset("state", dbOpenDynaset)
  stRec.MoveLast
  stRec.MoveFirst
  'Loop thru states
  frmAwuds2.MousePointer = vbHourglass
  For j = 0 To stRec.RecordCount
    'Set name of DB
    If j = 0 Then stateCode = "Nation" Else stateCode = stRec("state_cd")
    If Len(stateCode) < 2 Then stateCode = "0" & stateCode
    dbPath = AwudsDataPath & stateCode & ".mdb"
    If Len(Dir(dbPath)) > 0 Then 'DB exists
      stateName = stRec("state_nm")
      Set stateDB = OpenDatabase(dbPath, False, False, "MS Access; pwd=B7Q6C9B752")
      'Loop thru each table in DB
      For i = 0 To stateDB.TableDefs.Count - 1
        'Check to ensure that path for each connection to external table is same as registry
        connectedDBpath = stateDB.TableDefs(i).Connect
        If Len(connectedDBpath) > 0 Then
          connectedDBexpected = ";DATABASE=" & AwudsDataPath & FilenameNoPath(connectedDBpath)
          If UCase(connectedDBpath) <> UCase(connectedDBexpected) Then
            stateDB.TableDefs(i).Connect = connectedDBexpected
            stateDB.TableDefs(i).RefreshLink
          End If
        End If
      Next i
      stateDB.Close
    End If
    If j > 0 Then stRec.MoveNext
  Next j
  frmAwuds2.MousePointer = vbDefault
  
  stRec.Close
  'Set registry entry for data path
  SaveSetting "AWUDS", "Defaults", "DataPath", AwudsDataPath
  Exit Sub
  
ErrorReadOnly:
  msg = Err.Description
  tmp = ""
  While Len(Trim(msg)) > 0
    tmp = tmp & vbCrLf & StrSplit(msg, "  ", "")
  Wend
  msg = "Error opening data directory '" & AwudsDataPath & "'" & vbCrLf & tmp
  Err.Clear
  Select Case MyMsgBox.Show(msg, "AWUDS data directory problem", _
                            "+&Exit", "&Change Directory", "&Retry this directory")
    Case 1: GoTo UnloadExit
    Case 2: GoTo GetDataPath
    Case 3: GoTo CheckReadOnly
  End Select

UnloadExit:
  If MyP.State = "" Then 'close up shop
    Set MyMsgBox = Nothing
    Set Registry = Nothing
    Set MyP = Nothing
    End
  End If
End Sub
