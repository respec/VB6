VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtcoValidateUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
  '##MODULE_NAME AtcoValidateUser
  '##MODULE_DATE June7, 2001
  '##MODULE_AUTHOR Jack Kittle and Mark Gray, AQUA TERRA CONSULTANTS
  '##MODULE_SUMMARY Establishes level of user access as either read/write or read only. _
                    Access is set to NONE if the access can not be determined.
  '##MODULE_REMARKS <P>The idea is to determine if the user is a member of an AWUDSWW _
          (READ/WRITE) or AWUDSWR (READ ONLY) group _
          (currently only AWUDSWW is being supported in this version of the code. Any _
          user not in the AWUDSWW group will have READ access to AWUDS).  </P> _
          <P>The parent routine in this module is GetUserAccess. All other _
          routines are daughter calls. The routine returns either local or domain _
          access; local access is supported for standalone versions of AWUDS _
          not connected to the network and to support WSCs as they migrate to _
          Active Directory. The process is as follows: _
          <P>If user is logged in under a domain, then determine the user's _
          access from the Active Directory domain controler. If the user is _
          not logged on under a domain or is logged in under a workgroup, _
          then try using the file services of the _
          the stand-alone computer and look for a local group named AWUDSWW to see _
          if the user is a member of that group. </P> <P> After Active Directory migration _
          the SA should be using DOI Active Directory group _
          names using the DOI naming standard as shown in the example below: _
          <br><br>IGSGOrlandoFL-S AWUDSWW<br> _
          Where the OrlandoFL-S part would be replaced with city state of _
          your organizational unit and the S would be replaced by your _
          discipline in this case it is S for Science center, but it _
          could be W for water, etc.  After migration to the domain GS -- _
          the NT domain AWUDSWW group should be renamed using the naming _
          convention shown above. </P> <P>The following program was tested and only _
          works on NT based operting systems. Windows 95/98 are not supported.</P>
  '##MODULE_HISTORY twaugens 8/7/2006; updated for DOI Active Directory implementation. _
                    Added module CGetDcNameUSGS, the module handles getting the _
                    domain controller that is nearest to the user's computer. _
                    The domain controller is then used to obtain a user object, _
                    the user object is then queried to determine if the user is _
                    a member of the AWUDSWW group. The AWUDSWR group is not _
                    currently being implemented. Any user not in the AWUDSWW _
                    group has READ access to AWUDS.
  '
  ' <><><><><><>< Variables Section ><><><><><><><>
  '
  ' <><><><><><>< External DLL Routine Declarations ><><><><><><><>
  '
  '##SUMMARY Determines name of user currently logged onto the system.
  '##PARAM lpBuffer O Name of user right buffered with spaces.
  '##PARAM nSize I Number of characters to return in lpBuffer.
  '##REMARKS <br> _
  http://msdn.microsoft.com/library/default.asp?url=/library/en-us/sysinfo/base/getusername.asp
  '##RETURNS If the function succeeds, the return value is nonzero, _
             and the variable pointed to by nSize contains the number of characters _
             copied to the buffer specified by lpBuffer, including the terminating _
             null character.
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
  (ByVal lpBuffer As String, ByRef nSize As Long) As Long
  '##SUMMARY Frees the memory that NetApiBufferAllocate allocates. _
             Use this function to free the memory that other Net functions return.
  '##PARAM buffer I Pointer to a buffer returned previously by another _
           network management function.
  '##REMARKS <br> _
  http://msdn.microsoft.com/library/default.asp?url=/library/en-us/netmgmt/netmgmt/netapibufferfree.asp
  '##RETURNS If the function succeeds, the return value is NERR_Success. _
             If the function fails, the return value is a Win32 API error code.
Private Declare Function NetApiBufferFree Lib "Netapi32.dll" _
  (ByRef Buffer As Any) As Long
  '##SUMMARY The GetVersionEx function obtains extended information about _
             the version of the operating system that is currently running.
  '##PARAM lpVersionInformation I Pointer to an OSVERSIONINFO data structure _
           that the function fills with operating system version information.
  '##REMARKS <br> _
  http://msdn.microsoft.com/library/default.asp?url=/library/en-us/sysinfo/base/getversionex.asp
  '##RETURNS If the function succeeds, the return value is nonzero.
Private Declare Function GetVersionExA Lib "kernel32" _
  (ByRef lpVersionInformation As OSVERSIONINFO) As Integer
  '##SUMMARY Retrieves information about the configuration elements for a workstation.
  '##PARAM strServer I Pointer to a Unicode string containing the name of the remote _
           server on which the function is to execute. A NULL pointer or string _
           specifies the local computer
  '##PARAM lLevel I Specifies one of the following values to return the level of _
           information provided: _
    <UL> _
    <li>100 = The bufptr parameter points to a WKSTA_INFO_100 structure, _
    <li>101 = The bufptr parameter points to a WKSTA_INFO_101 structure, _
    <li>102 = The bufptr parameter points to a WKSTA_INFO_102 structure, _
    <li>302 = The bufptr parameter points to a WKSTA_INFO_302 structure, _
    <li>402 = The bufptr parameter points to a WKSTA_INFO_402 structure, _
    <li>502 = The bufptr parameter points to a WKSTA_INFO_502 structure. _
    </UL>
  '##PARAM pbBuffer I Pointer to the buffer in which the data set with the _
           level parameter is stored.
  '##REMARKS <br> _
  http://msdn.microsoft.com/library/default.asp?url=/library/en-us/netmgmt/netmgmt/netwkstagetinfo.asp
  '##RETURNS Information about the configuration elements for a workstation.
Private Declare Function NetWkstaGetInfo& Lib "netapi32" _
  (ByRef strServer As Any, ByVal lLevel As Long, ByRef pbBuffer As Any)
  '##SUMMARY Determines current user's NT Domain; NT only.
  '##PARAM reserved I Send as 'ByVal 0&'
  '##PARAM lLevel I Specifies the information level of the data; send as '1'.
  '##PARAM pbBuffer I Pointer to the buffer that receives the data. _
           The format of this data depends on the value of the level parameter. _
           This buffer is allocated by the system and must be freed using the _
           NetApiBufferFree function.
  '##REMARKS <br> _
  http://msdn.microsoft.com/library/default.asp?url=/library/en-us/netmgmt/netmgmt/netwkstausergetinfo.asp
  '##RETURNS If the function succeeds, the return value is nonzero.
Private Declare Function NetWkstaUserGetInfo& Lib "netapi32" _
  (ByRef reserved As Any, ByVal lLevel As Long, ByRef pbBuffer As Any)
  '##SUMMARY Copies a string to a buffer; NT only.
  '##PARAM Dest I Points to a buffer to receive the contents of the _
      string pointed to by the lpString2 parameter. The buffer must _
      be large enough to contain the string, including the terminating _
      null character.
  '##PARAM src I Points to the null-terminated string to be copied.
  '##REMARKS <br> _
  http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/resources/strings/stringreference/stringfunctions/lstrcpy.asp
  '##RETURNS Pointer to buffer if successful; NULL if fails.
Private Declare Sub lstrcpyW Lib "kernel32" _
  (ByRef Dest As Any, ByVal src As Any)
  '##SUMMARY Copies a string to a buffer.
  '##PARAM Dest I Points to a buffer to receive the contents of the _
      string pointed to by the lpString2 parameter. The buffer must _
      be large enough to contain the string, including the terminating _
      null character.
  '##PARAM src I Points to the null-terminated string to be copied.
  '##REMARKS <br> _
  http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/resources/strings/stringreference/stringfunctions/lstrcpy.asp
  '##RETURNS Pointer to buffer if successful; NULL if fails.
Private Declare Sub lstrcpy Lib "kernel32" _
  (ByRef Dest As Any, ByVal src As Any)
  '##SUMMARY Moves a block of data directly to memory.
  '##PARAM Dest I Points to the starting address of the copied block’s destination.
  '##PARAM src I Points to the starting address of the block of memory to copy.
  '##PARAM size I Specifies the size, in bytes, of the block of memory to move.
  '##RETURNS Address of next block 4 bytes ahead if function works, 0 if fails.
Private Declare Sub RtlMoveMemory Lib "kernel32" _
  (ByRef Dest As Any, ByRef src As Any, ByVal size As Long)
  '
  ' <><><><><><>< Internal Type Declarations ><><><><><><><>
  '
Private Type OSVERSIONINFO
' ##SUMMARY Summary of information regarding operating system version.
' ##REMARKS <br> _
  http://msdn.microsoft.com/library/default.asp?url=/library/en-us/sysinfo/base/osversioninfo_str.asp
' ##UDT_MEMBER_DESCRIPTION dwOSVersionInfoSize Operating system version size (default to 148)
' ##UDT_MEMBER_DESCRIPTION dwMajorVersion Major version number
' ##UDT_MEMBER_DESCRIPTION dwMinorVersion Minor version number
' ##UDT_MEMBER_DESCRIPTION dwBuildNumber Build number
' ##UDT_MEMBER_DESCRIPTION dwPlatformId 1 = Windows 95/98, 2 = Windows NT and XP
' ##UDT_MEMBER_DESCRIPTION szCSDVersion CSD version size (default to Space$(128))
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type
Private Type WKSTA_USER_INFO_1
' ##SUMMARY Summary of information regarding user work station.
' ##REMARKS <br> _
http://windowssdk.msdn.microsoft.com/en-us/library/ms708179.aspx
' ##UDT_MEMBER_DESCRIPTION wkui1_username Pointer to user name on _
                account under which user is currently logged on.
' ##UDT_MEMBER_DESCRIPTION wkui1_logon_domain Pointer to domain name of _
                the user account of the user currently logged on.
' ##UDT_MEMBER_DESCRIPTION wkui1_oth_domains Pointer to list of other LAN _
                Manager domains browsed by the workstation.
' ##UDT_MEMBER_DESCRIPTION wkui1_logon_server Pointer to name of the computer _
                that authenticated the server.
   wkui1_username As Long
   wkui1_logon_domain As Long
   wkui1_oth_domains As Long
   wkui1_logon_server As Long
End Type
Private Type WKSTA_INFO_101
' ##SUMMARY Summary of information regarding user work station.
' ##REMARKS <br> _
  http://windowssdk.msdn.microsoft.com/en-us/library/ms708150.aspx
' ##UDT_MEMBER_DESCRIPTION wki101_platform_id Pointer to wki101_platform_id
' ##UDT_MEMBER_DESCRIPTION wki101_computername Pointer to wki101_computername
' ##UDT_MEMBER_DESCRIPTION wki101_langroup Pointer to wki101_langroup
' ##UDT_MEMBER_DESCRIPTION wki101_ver_major Pointer to wki101_ver_major
' ##UDT_MEMBER_DESCRIPTION wki101_ver_minor Pointer to wki101_ver_minor
' ##UDT_MEMBER_DESCRIPTION wki101_lanroot Pointer to wki101_lanroot
   wki101_platform_id As Long
   wki101_computername As Long
   wki101_langroup As Long
   wki101_ver_major As Long
   wki101_ver_minor As Long
   wki101_lanroot As Long
End Type

Public Function GetUserAccess() As String
Attribute GetUserAccess.VB_Description = "Determines username, version of Windows operating system, full pathname to computer, and then determines whether the user has write or read-only privelages."
  '##SUMMARY Determines whether the user has write or read-only privelages to _
             the AWUDS database.
  '##RETURNS "WRITE" if the user is a member of a group _
             where the group name ends in the text string AWUDSWW, _
             "READ" if the user is not a member of an AWUDSWW group or _
             "NONE" if the user is trying to run on a computer that _
             is not running a version of the NT operating system or if _
             the domain controller is not available (the user is running _
             off the network).
             
  Dim winVerName As String    'version of windows operating system
  Dim machineName As String   'name of computer on system
  Dim DomainName As String    'domain name where computer is assigned
  Dim workGroupName As String 'name of workgroup on system
  Dim tmpAccess As String     'tempory string used to hold user's access.
  Dim userName As String      'name user is logged on under
  Dim pathStr As String       'path to computer system
  Dim objCGetDcNameUSGS As CGetDcNameUSGS ' Network domain controller object.
  Dim tmpString As String     'temporary string
  Dim n As Long               'index
On Error GoTo ErrRoutine
  tmpAccess = "NONE"
  userName = NetworkUserName
  winVerName = WindowsVersion
  
  Select Case winVerName
    Case "Windows NT": 'Windows NT is used for Windows 2000, XP, and Vista.
      '
      ' The following call only works on NT machines
      '
      GetWorkStation machineName, DomainName, workGroupName
      '
      ' If the machine name and domain name are the same then the user
      ' is logged in locally. When the names differ the user is logged
      ' in under a domain.
      '
      If LCase(machineName) <> LCase(DomainName) Then
        '
        ' Bind to the domain controller that is closest to this
        ' machine. Creating this object automatically causes it to
        ' bind to the nearest machine. See the initialize code in
        ' CGetDcNameUSGS.
        '
        Set objCGetDcNameUSGS = New CGetDcNameUSGS
        
        If objCGetDcNameUSGS.DomainControllerName = "" Then
          '
          ' The domain controller could not be obtained,
          ' therefore, attempt to get the user's access locally.
          '
          pathStr = "WinNT://" & machineName
          tmpAccess = GetMachineAccess(pathStr, userName)
          GoTo ExitRoutine
        End If
        '
        ' Build the string that will be used to query Active Directory for
        ' a user object. The WinNt interface will be used.
        '
        tmpString = objCGetDcNameUSGS.DomainControllerName
        
        n = InStr(1, tmpString, "\\")
        pathStr = "WinNT://" & Mid(tmpString, n + 2) & "/" & userName & ",user"
        tmpAccess = GetAccess(pathStr)
        
        '
        ' If the awudsww or awudswr group did not exist in active directory
        ' then try looking on the user's local machine.
        '
        If tmpAccess = "NOGROUPS" Then
          '
          ' Attempt to get the user's access locally.
          '
          pathStr = "WinNT://" & machineName
          tmpAccess = GetMachineAccess(pathStr, userName)
          GoTo ExitRoutine
        End If
        
      Else
        pathStr = "WinNT://" & machineName
        tmpAccess = GetMachineAccess(pathStr, userName)
      End If
      
      
    Case Else
      MsgBox "AWUDS does not run under " & winVerName & ".", vbCritical
      tmpAccess = "NONE"
  End Select
  
ExitRoutine:
  If tmpAccess = "" Then tmpAccess = "NONE"
  GetUserAccess = tmpAccess
  
  If Not objCGetDcNameUSGS Is Nothing Then
     Set objCGetDcNameUSGS = Nothing
  End If
  
  Exit Function
  
ErrRoutine:
  MsgBox "Runtime error in GetUserAccess: " & Err.Description
  GoTo ExitRoutine
  
End Function

Private Function GetAccess(ByRef UserReference As String) As String
Attribute GetAccess.VB_Description = "Grants user access based on whether user is member of AWUDS usergroup with write access, read-only access, or if there is no AWUDS user group on the computer (read-only)."
  '##SUMMARY Grants user access based on whether a user is a member of an AWUDS _
    group with write access, read-only access, or if there is no AWUDS _
    user group on the computer (read-only).
  '##PARAM UserReference I Path to the domain user object on a domain (ADsPath) e.g. <br> _
           GetObject("WinNT://MYDOMAIN/jeffsmith,user")
  '##REMARKS Uses the ADSI WinNT Provider to determine if the user is a member _
             of the AWUDSWW group by checking the list of groups a user is a _
             member of by accessing the ADSI WinNT user object. See: <br> _
             http://msdn.microsoft.com/library/default.asp?url=/library/en-us/adsi/adsi/iadsuser.asp _
             <br>and the following: <br> _
             http://msdn.microsoft.com/library/default.asp?url=/library/en-us/adsi/adsi/adsi_objects_of_winnt.asp
  '##RETURNS WRITE if user is a member of the AWUDSWW group, or READ if the user is a _
             member of the AWUDSWR group or NOGROUPS if user is not a member of either _
             group or NONE if there was a runtime error.
  
  Dim myUser As IADsUser    'user object
  Dim myGroup As IADsGroup  'group object
  Dim ugName As String      'group name
  Dim tmpAccess As String   'Temporary string containing users access.
  
  On Error GoTo ErrRoutine
  
  tmpAccess = "NOGROUPS"
  
  If Not myUser Is Nothing Then
    Set myUser = Nothing
  End If
  
  Set myUser = GetObject(UserReference)
  
  '
  ' Determine if user is in an AWUDSWW domain group. The last 7 characters
  ' of the domain group should be spelled as AWUDSWW.
  '
  For Each myGroup In myUser.Groups
    ugName = Trim(LCase(myGroup.Name))
    If Len(ugName) >= 7 Then
        If Right(ugName, 7) = "awudsww" Then 'member of AWUDSWW
            tmpAccess = "WRITE"
            GoTo ExitRoutine
        ElseIf Right(ugName, 7) = "awudswr" Then 'member of AWUDSWR
            tmpAccess = "READ"
            GoTo ExitRoutine
        End If
    End If
  Next

ExitRoutine:
 
  GetAccess = tmpAccess
  
  If Not myGroup Is Nothing Then
    Set myGroup = Nothing
  End If
  
  If Not myUser Is Nothing Then
    Set myUser = Nothing
  End If
      
  Exit Function
  
ErrRoutine:
  tmpAccess = "NONE"
  MsgBox "Runtime error in GetAccess: " & Err.Description & vbCrLf & vbCrLf & _
         "Check your network connections, your computer may be disconnected from the network."
  GoTo ExitRoutine
End Function

Private Function GetMachineAccess(ByRef machineName As String, ByRef userName As String) As String
  '##SUMMARY Grants user access based on whether user is member of AWUDS _
    group with write access AWUDSWW. All other user's will receive _
    READ access to the AWUDS database. This is the original 1.2.1 code. The code _
    should only be called to obtain access based on a local AWUDSWW group. _
    Access by DOI Active Directory is obtained by calling GetAccess.
  '##PARAM machineName I The name of the local computer (ADsPath) e.g. <br> _
           GetObject("WinNT://MYCOMPUTER/jeffsmith,user")
  '##PARAM userName I Name user is logged on under.
  '##REMARKS Uses the ADSI WinNT Provider to determine if the user is a member _
             of the local AWUDSWW group by checking the list of groups a user is a _
             member of by accessing the ADSI WinNT user object. See: <br> _
             http://msdn.microsoft.com/library/default.asp?url=/library/en-us/adsi/adsi/iadsuser.asp _
             <br>and the following: <br> _
             http://msdn.microsoft.com/library/default.asp?url=/library/en-us/adsi/adsi/adsi_objects_of_winnt.asp
   '##RETURNS WRITE if user is a member of the AWUDSWW group, or READ if the user is not a _
             member of the AWUDSWW group.
  
  Dim myMachine As Object 'computer object
  Dim ugPath As String    'path to usergroup
  Dim ugName As String    'name of
  Dim item As Object      'member of collection belonging to computer object
  Dim tmpAccess As String 'local version of GetAccess
  
  On Error GoTo ErrRoutine
  
  tmpAccess = "READ"
  
  If Not myMachine Is Nothing Then
      Set myMachine = Nothing
  End If
  
  Set myMachine = GetObject(machineName)
  
  For Each item In myMachine
    If LCase(item.Class) = "group" Then
      ugName = LCase(item.Name)
      If InStr(1, ugName, "awudsww") > 0 Then 'member of AWUDSWW
        ugPath = machineName & "/" & item.Name
        If (IsUserMember(ugPath, userName)) Then
          tmpAccess = "WRITE"
          GoTo ExitRoutine
        End If
      End If
    End If
  Next item
  '
  ' Send back read or none access.
  '
ExitRoutine:
  
  GetMachineAccess = tmpAccess
  If Not myMachine Is Nothing Then
      Set myMachine = Nothing
  End If
  Exit Function
  
ErrRoutine:
  MsgBox "Runtime error in GetMachineAccess: " & Err.Description
  GoTo ExitRoutine
End Function

Private Function IsUserMember(ByRef GroupName As String, ByRef userName As String) As Boolean
Attribute IsUserMember.VB_Description = "Determines whether user is member of AWUDS work group."
  '##SUMMARY Determines whether user is member of AWUDS work group.
  '##PARAM GroupName I Name of User Group.
  '##PARAM userName I Name user is logged on under.
  '##RETURNS 'True' if user belongs to work group, 'False' if not.
  Dim myUserGroup As Object 'active user group object
  Dim user As Object        'active user object
  
  On Error GoTo ErrRoutine
  
  IsUserMember = False
  
  If Not myUserGroup Is Nothing Then
      Set myUserGroup = Nothing
  End If
  '
  ' See if user is in the group.
  '
  Set myUserGroup = GetObject(GroupName)
  For Each user In myUserGroup.Members
    If LCase(user.Name) = LCase(userName) And Len(Trim(userName)) > 0 Then 'a match
        IsUserMember = True
        GoTo ExitRoutine
    End If
  Next user
  
ExitRoutine:
  If Not myUserGroup Is Nothing Then
      Set myUserGroup = Nothing
  End If
  Exit Function
  
ErrRoutine:
  MsgBox "Runtime error in IsUserMember: " & Err.Description
  GoTo ExitRoutine

End Function

Private Function NetworkUserName() As String
Attribute NetworkUserName.VB_Description = "Determines name of user currently logged onto network."
  '##SUMMARY Determines name of user currently logged onto network.
  '##RETURNS Network user name.
  Dim lpBuff As String * 25 'Buffered string holding user name.
  Dim retval As Long        'return code from call to GetUserName; 1=successful.

  retval = GetUserName(lpBuff, 25)
  'trim off any trailing spaces found in the name
  NetworkUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
End Function

Private Function WindowsVersion() As String
Attribute WindowsVersion.VB_Description = "Determines version of Windows operating system."
  '##SUMMARY Determines version of Windows operating system.
  '##RETURNS Name of Windows Operating System; either "Windows 95/98", _
             "Windows NT", or "Unknown". Windows NT should represent _
              Windows Vista, Windows Server "Longhorn", _
              Windows Server 2003, Windows XP, Windows 2000, or Windows NT.
  
  Dim osinfo   As OSVERSIONINFO 'object storing Windows OS information.
  Dim retvalue As Integer       'return code from call to GetVersionExA
  
  On Error GoTo ErrRoutine
  
  osinfo.dwOSVersionInfoSize = 148
  osinfo.szCSDVersion = Space$(128)
  retvalue = GetVersionExA(osinfo)
  '
  ' http://windowssdk.msdn.microsoft.com/en-us/library/ms724834.aspx
  ' 1 - The operating system is Windows Me, Windows 98, or Windows 95.
  ' 2 - The operating system is Windows Vista, Windows Server "Longhorn",
  '     Windows Server 2003, Windows XP, Windows 2000, or Windows NT.
  '
  Select Case osinfo.dwPlatformId
    Case Is = 1: WindowsVersion = "Windows 95/98"
    Case Is = 2: WindowsVersion = "Windows NT"
    Case Else: WindowsVersion = "Unknown"
      MsgBox "Can not determine what operating system you are running.", vbCritical
  End Select
  Exit Function
  
ErrRoutine:
  MsgBox "Runtime error in WindowsVersion: " & Err.Description
  
End Function

Private Sub GetWorkStation(ByRef machineName As String, _
                           ByRef DomainName As String, _
                           ByRef workGroupName As String)
  '##SUMMARY Determines name of computer, domain, and workgroup via API calls; _
             only works for NT based operating systems.
  '##PARAM MachineName O Computer name.
  '##PARAM DomainName O Domain name.
  '##PARAM WorkGroupName O Workgroup name.
  '##REAMARKS Information summarized from a code walk through with Willot Wilson: _
              A computer can be assigned to a _
              work group in which case it is not part of the GS Active Directory _
              domain. In this case we think the following routine will return _
              the DomainName = machineName. In which case the security module _
              is designed to check for the local AWUDSWW group membership. _
              It is highly unlikely that a WSC will place a water-use user's coumputer in _
              a work group, therefore, NWIS is not planning on testing the _
              work group use case. User's computers are supposed to be members of the _
              DOI Active Directory domain GS. When a user's computer is in _
              the domain GS then DomainName <> machineName. In which case the _
              security module will check an Active Directory domain controller _
              to assign access based on the AWUDSWW group (having the correct _
              DOI naming convention applied to it, for example: _
              IGSGOrlandoFL-S AWUDSWW). If WSC's decide to implement work groups _
              then this module will require additional internal testing.
              
  Dim ret As Long              'return code for API calls
  Dim Buffer(512) As Byte      'byte array representing a string
  Dim i As Integer             'counter
  Dim wk101 As WKSTA_INFO_101  'Private Type: summary of work station information
  Dim pwk101 As Long           'pointer to buffer that receives the data.
  Dim wk1 As WKSTA_USER_INFO_1 'Private Type: summary of work station information
  Dim pwk1 As Long             'pointer to buffer that receives the data
  Dim ComputerName As String   'name of computer
  Dim lanGroup As String       'name of work group
  Dim logonDomain As String    'name of domain user is logged onto
  
  On Error GoTo ErrRoutine
     
  ' Clear all of the display values.
  lanGroup = ""
  logonDomain = ""

  'NT and XP only - call NetWkstaGetInfo to get computer name and lan group
  ret = NetWkstaGetInfo(ByVal 0&, 101, pwk101)
  RtlMoveMemory wk101, ByVal pwk101, Len(wk101)
  lstrcpyW Buffer(0), wk101.wki101_computername
  ' Get every other byte from Unicode string.
  i = 0
  Do While Buffer(i) <> 0
     ComputerName = ComputerName & Chr(Buffer(i))
     i = i + 2
  Loop
  lstrcpyW Buffer(0), wk101.wki101_langroup
  i = 0
  Do While Buffer(i) <> 0
     lanGroup = lanGroup & Chr(Buffer(i))
     i = i + 2
  Loop
  ret = NetApiBufferFree(pwk101)

  ' NT only - call NetWkstaUserGetInfo.
  ret = NetWkstaUserGetInfo(ByVal 0&, 1, pwk1)
  RtlMoveMemory wk1, ByVal pwk1, Len(wk1)
  lstrcpyW Buffer(0), wk1.wkui1_logon_domain
  i = 0
  Do While Buffer(i) <> 0
     logonDomain = logonDomain & Chr(Buffer(i))
     i = i + 2
  Loop
  ret = NetApiBufferFree(pwk1)
  machineName = ComputerName
  DomainName = logonDomain
  workGroupName = lanGroup
  
  Exit Sub

ErrRoutine:
  MsgBox "Runtime error in GetWorkStation: " & Err.Description
End Sub
