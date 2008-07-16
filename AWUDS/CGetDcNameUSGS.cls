VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGetDcNameUSGS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©2000 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************

'##MODULE_NAME CGetDcNameUSGS
'##MODULE_DATE Aquired: August 7, 2006
'##MODULE_AUTHOR Copyright ©2000 Karl E. Peterson <br> _
                 All Rights Reserved, http://www.mvps.org/vb <br> _
                 <p>NWIS can use this code free within applications, but NWIS _
                 is expressly forbidden from selling or otherwise distributing the _
                 source code, non-compiled, without prior written consent.</p> _
                 <p>Redistribution and Licensing information: <br> _
                 http://vb.mvps.org/license.asp _
                 <p>Code was modified by Todd Augenstein, 8/7/2006 as documented below.
'##MODULE_SUMMARY Obtain domain controller information on Win2000 (and higher) systems.
'##MODULE_REMARKS <p>Original code downloaded from (8/7/2006): <br> _
                  http://vb.mvps.org/samples/project.asp?id=NetDomain <br> _
                  The code was modified as follows: the debug, hexdump, _
                  PointerToDWord, PointerToPointerToStringW, _
                  and PointerToStringA modules were deleted because AWUDS did _
                  not specifically need these modules. _
                  For security reasons other public modules and property _
                  "let" statements and the flags functionality were commented out, _
                  but source code left in. The functionality that was commented out _
                  was NOT required by AWUDS and requires additional testing before use _
                  by any other applications, thus it was commented out. _
                  DocumentX style documentation was added to the remaining modules. _
                  A few other minor modifications were made as documented in HISTORY sections _
                  of the documentation. The original version of the code that _
                  was downloaded can be found in the zip file: NetDomainOriginal.zip
'##MODULE_REMARKS An object created from this class will immediately bind _
                  the object to a domain controller object _
                  based on the nearest domain controller on the _
                  Active Directory GS domain. Properties of the domain controller object _
                  can then be referenced to obtain controller properties such as _
                  the controllers name. AWUDS requires just the controller's name. _
                  For usage in AWUDS see module: AtcoValidateUser.GetUserAccess. _
                  <p>Programmer warning. _
                  The Unicode versions of the modules are used by AWUDS. _
                  <p>If you use other ANSI versions of the functions in this module _
                  please test them and then update or remove this warning as necessary.
  
Option Explicit

' Win32 APIs for networking...

'##SUMMARY Frees the memory that NetApiBufferAllocate allocates. _
           Use this function to free the memory that other Net functions return.
'##PARAM lpBuffer I Pointer to a buffer returned previously by another _
         network management function.
'##REMARKS <br> _
http://msdn.microsoft.com/library/default.asp?url=/library/en-us/netmgmt/netmgmt/netapibufferfree.asp
'##RETURNS If the function succeeds, the return value is NERR_Success. _
           If the function fails, the return value is a Win32 API error code.
Private Declare Function NetApiBufferFree Lib "Netapi32.dll" (ByVal lpBuffer As Long) As Long

'##SUMMARY Copies a block of memory from one location to another.
'##REMARKS See: <br> _
http://msdn.microsoft.com/library/default.asp?url=/library/en-us/memory/base/copymemory.asp
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)

'##SUMMARY Returns the length in bytes (ANSI version) or WCHAR values _
           (Unicode version) of the specified string (not including the _
           terminating null character).
'##REMARKS See: <br> _
http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/resources/strings/stringreference/stringfunctions/lstrlen.asp
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

'##SUMMARY Returns the length in bytes (ANSI version) or WCHAR values _
           (Unicode version) of the specified string (not including the _
           terminating null character).
'##REMARKS See: <br> _
http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/resources/strings/stringreference/stringfunctions/lstrlen.asp
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

'##SUMMARY Converts a globally unique identifier (GUID) into a string of _
           printable characters.
'##REMARKS See: <br> _
http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcedcom/html/cerefstringfromguid2.asp
Private Declare Function StringFromGUID2 Lib "ole32.dll" (pGuid As GUID, ByVal PointerToString As Long, ByVal MaxLength As Long) As Long

'DWORD DsGetDcName(
'  LPCTSTR ComputerName,
'  LPCTSTR DomainName,
'  GUID *DomainGuid,
'  LPCTSTR SiteName,
'  ULONG Flags,
'  PDOMAIN_CONTROLLER_INFO *DomainControllerInfo
');
 
'##SUMMARY Returns the name of a domain controller in a specified domain.
'##REMARKS See: <br> _
http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcecomm5/html/wce50lrfDsGetDcName.asp
Private Declare Function DsGetDcNameA Lib "netapi32" ( _
   ByVal ComputerName As String, _
   ByVal DomainName As String, _
   DomainGuid As Any, _
   ByVal SiteName As String, _
   ByVal Flags As Long, _
   lpDomainControllerInfo As Long) As Long
   
'##SUMMARY Returns the name of a domain controller in a specified domain.
'##REMARKS See: <br> _
http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcecomm5/html/wce50lrfDsGetDcName.asp
Private Declare Function DsGetDcNameW Lib "netapi32" ( _
   ByVal ComputerName As Long, _
   ByVal DomainName As Long, _
   DomainGuid As Any, _
   ByVal SiteName As Long, _
   ByVal Flags As Long, _
   lpDomainControllerInfo As Long) As Long
   
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

'typedef struct _DOMAIN_CONTROLLER_INFO
'{
'  LPTSTR DomainControllerName;
'  LPTSTR DomainControllerAddress;
'  ULONG  DomainControllerAddressType;
'  GUID   DomainGuid;
'  LPTSTR DomainName;
'  LPTSTR DnsForestName;
'  ULONG  Flags;
'  LPTSTR DcSiteName;
'  LPTSTR ClientSiteName;
'} DOMAIN_CONTROLLER_INFO, *PDOMAIN_CONTROLLER_INFO;

'
Private Type DOMAIN_CONTROLLER_INFO
'##SUMMARY Domain controller structure.
'##REMARKS See: <br> _
http://windowssdk.msdn.microsoft.com/en-us/library/ms675912.aspx
   DomainControllerName As Long
   DomainControllerAddress As Long
   DomainControllerAddressType As Long
   DomainGuid As GUID
   DomainName As Long
   DnsForestName As Long
   Flags As Long
   DcSiteName As Long
   ClientSiteName As Long
End Type

'
' VB displayable versions of what is in the DOMAIN_CONTROLLER_INFO structure.
'
Private Type DomainControllerInfo
'##SUMMARY Viewable version of the domain controller structure.
'##REMARKS See: <br> _
http://windowssdk.msdn.microsoft.com/en-us/library/ms675912.aspx
   DomainControllerName As String       ' the computer name of the discovered domain controller.
   DomainControllerAddress As String    ' address of the discovered domain controller. The address is prefixed with "\\".
   DomainControllerAddressType As Long  ' address is a string IP address or address is a NetBIOS name.
   DomainGuid As GUID                   ' the GUID of the domain
   DomainName As String                 ' the name of the domain
   DnsForestName As String              ' the name of the domain at the root of the Directory Service tree.
   Flags As Long                        ' contains a set of flags that describe the domain controller.
   DcSiteName As String                 ' the name of the site where the domain controller is located.
   ClientSiteName As String             ' name of the site that the computer belongs to.
End Type

'
' twaugens, 8/7/2006, the ability to change the flags was turned off since it
'                     it has not been tested.
'
' // Flags to passed to DsGetDcName
'Private Const DS_FORCE_REDISCOVERY = &H1
'Private Const DS_DIRECTORY_SERVICE_REQUIRED = &H10
'Private Const DS_DIRECTORY_SERVICE_PREFERRED = &H20
'Private Const DS_GC_SERVER_REQUIRED = &H40
'Private Const DS_PDC_REQUIRED = &H80
'Private Const DS_BACKGROUND_ONLY = &H100  '<-- undoc'd?
'Private Const DS_IP_REQUIRED = &H200
'Private Const DS_KDC_REQUIRED = &H400
'Private Const DS_TIMESERV_REQUIRED = &H800
'Private Const DS_WRITABLE_REQUIRED = &H1000
'Private Const DS_GOOD_TIMESERV_PREFERRED = &H2000
'Private Const DS_AVOID_SELF = &H4000
'Private Const DS_ONLY_LDAP_NEEDED = &H8000
'Private Const DS_IS_FLAT_NAME = &H10000
'Private Const DS_IS_DNS_NAME = &H20000
'Private Const DS_RETURN_DNS_NAME = &H40000000
'Private Const DS_RETURN_FLAT_NAME = &H80000000

'Public Enum DsGetDcNameFlags
'   ForceRediscovery = DS_FORCE_REDISCOVERY
'   DirectoryServiceRequired = DS_DIRECTORY_SERVICE_REQUIRED
'   DirectoryServicePreferred = DS_DIRECTORY_SERVICE_PREFERRED
'   GlobalCatalogServerRequired = DS_GC_SERVER_REQUIRED
'   PrimaryDomainControllerRequired = DS_PDC_REQUIRED
'   ' DS_PDC_REQUIRED   <--undoc'd?
'   IPRequired = DS_IP_REQUIRED
'   KerberosKeyDistributionCenterRequired = DS_KDC_REQUIRED
'   TimeServerRequired = DS_TIMESERV_REQUIRED
'   WritableRequired = DS_WRITABLE_REQUIRED
'   ReliableTimerServerPreferred = DS_GOOD_TIMESERV_PREFERRED
'   AvoidSelf = DS_AVOID_SELF
'   OnlyLdapNeeded = DS_ONLY_LDAP_NEEDED
'   IsFlatName = DS_IS_FLAT_NAME
'   IsDnsName = DS_IS_DNS_NAME
'   ReturnDnsName = DS_RETURN_DNS_NAME
'   ReturnFlatName = DS_RETURN_FLAT_NAME
'End Enum

' // Values for DomainControllerAddressType
Private Const DS_INET_ADDRESS = 1
Private Const DS_NETBIOS_ADDRESS = 2

Public Enum DcAddressTypes
   dsInetAddress = 1
   dsNetbiosAddress = 2
End Enum

' // Values for returned Flags
Private Const DS_PDC_FLAG = &H1                     '// DC is PDC of Domain
Private Const DS_GC_FLAG = &H4                      '// DC is a GC of forest
Private Const DS_LDAP_FLAG = &H8                    '// Server supports an LDAP server
Private Const DS_DS_FLAG = &H10                     '// DC supports a DS and is a Domain Controller
Private Const DS_KDC_FLAG = &H20                    '// DC is running KDC service
Private Const DS_TIMESERV_FLAG = &H40               '// DC is running time service
Private Const DS_CLOSEST_FLAG = &H80                '// DC is in closest site to client
Private Const DS_WRITABLE_FLAG = &H100              '// DC has a writable DS
Private Const DS_GOOD_TIMESERV_FLAG = &H200         '// DC is running time service (and has clock hardware)
Private Const DS_PING_FLAGS = &HFFFF                '// Flags returned on ping
Private Const DS_DNS_CONTROLLER_FLAG = &H20000000   '// DomainControllerName is a DNS name
Private Const DS_DNS_DOMAIN_FLAG = &H40000000       '// DomainName is a DNS name
Private Const DS_DNS_FOREST_FLAG = &H80000000       '// DnsForestName is a DNS name

' // Possible error values
Private Const NO_ERROR = 0&
Private Const ERROR_INVALID_DOMAINNAME = 1212&
Private Const ERROR_INVALID_FLAGS = 1004&
Private Const ERROR_NOT_ENOUGH_MEMORY = 8&
Private Const ERROR_NO_SUCH_DOMAIN = 1355&

Public Enum dsGetDcNameErrors
   NoError = NO_ERROR
   NotEnoughMemory = ERROR_NOT_ENOUGH_MEMORY
   InvalidFlags = ERROR_INVALID_FLAGS
   InvalidDomainName = ERROR_INVALID_DOMAINNAME
   NoSuchDomain = ERROR_NO_SUCH_DOMAIN
End Enum

' The following functions are used to determine if an API function is exported.
' See: http://msdn.microsoft.com/library/en-us/wcekernl/html/ceconDLLFunctions.asp

'##SUMMARY Returns a module handle for the specified module if the _
           file has been mapped into the address space of the calling process.
'##REMARKS See: <br> _
http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcekernl/html/_wcesdk_Win32_GetModuleHandle.asp
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

'##SUMMARY Maps the specified DLL file into the address space of the calling process.
'##REMARKS See: <br> _
http://msdn.microsoft.com/library/en-us/wcekernl/html/_wcesdk_Win32_LoadLibrary.asp
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'##SUMMARY Returns the address of the specified exported DLL function.
'##REMARKS See: <br> _
http://msdn.microsoft.com/library/en-us/wcekernl/html/_wcesdk_Win32_GetProcAddress.asp
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'##SUMMARY Decrements the reference count of the loaded DLL module.
'##REMARKS See: <br> _
http://msdn.microsoft.com/library/en-us/wcekernl/html/_wcesdk_Win32_FreeLibrary.asp
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

' Member vars

'##SUMMARY Domain controller structure.
Private m_DCI As DomainControllerInfo

'##SUMMARY TRUE if Unicode code is called, FALSE for ANSI.
Private m_Unicode As Boolean

'##SUMMARY Name of the local computer.
Private m_ComputerName As String

'##SUMMARY TRUE if Unicode call is exported and callable.
Private m_Supported As Boolean

'##SUMMARY Error number of the last error that occured. 0 if no error.
Private m_LastError As Long

'##SUMMARY Contains a set of flags that describe the domain controller. _
           Value is zero until the domain controller is obtained. After which _
           it has the set of flags used by that controller.
Private m_Flags As Long


' ********************************************
'  Initialize
' ********************************************
Private Sub Class_Initialize()
'##SUMMARY Initialize the domain controller object and get a reference to the nearest _
           domain controller on the network.
'##REMARKS The reference fetched by the refresh method.
'##HISTORY twaugens, 8/7/2006, set the domain and comupter names to blank so the _
                     nearest domain controller is forced to be fetched _
                     based on the computer the user is logged on to. _
                     Also the flags variable was set to 0, no additional flags _
                     are being processed, the default flags are used.
'
   m_DCI.DomainName = ""
   m_ComputerName = ""
   m_Flags = 0
   
   '
   ' Default values.
   '
   ' Unicode is set to TRUE so that the Unicode modules are only called.
   ' The ANSI modules have been left in the code in case they might be
   ' used at some other release. The ANSI modules have not been tested.
   '
   m_Unicode = True
   
   ' Make sure DsGetDcName is available!
   m_Supported = Exported("netapi32", "DsGetDcNameW")
   ' Get data for default domain.
   Me.Refresh
End Sub

' ********************************************
'  Public Properties: Read/Write
' ********************************************


'
' twaugens, 8/7/2006, the ability to change the computer name was turned off for AWUDS release.
'
'Public Property Let ComputerName(ByVal NewVal As String)
'   m_ComputerName = NewVal
'   'Call Refresh
'End Property

Public Property Get ComputerName() As String
'##SUMMARY Get value of the user's computer name.
   ComputerName = m_ComputerName
End Property
'
' twaugens, 8/7/2006, the ability to debug was turned off for AWUDS release.
'
'Public Property Let Debugging(ByVal NewVal As Boolean)
'   m_Debug = NewVal
'End Property
'
'Public Property Get Debugging() As Boolean
'   Debugging = m_Debug
'End Property

'
' twaugens, 8/7/2006, the ability to change the domain name was turned off for AWUDS release.
'
'Public Property Let DomainName(ByVal NewVal As String)
'   m_DCI.DomainName = NewVal
'   'Call Refresh
'End Property

Public Property Get DomainName() As String
'##SUMMARY Get value of the domain name of the domain the control is located on.
   DomainName = m_DCI.DomainName
End Property

'
' twaugens, 8/7/2006, the ability to change the flags was turned off since it
'                     it has not been tested.
'
'Public Property Let FlagsIn(ByVal WhichFlag As DsGetDcNameFlags, ByVal NewVal As Boolean)
'   If NewVal Then
'      m_Flags = m_Flags Or WhichFlag
'   Else
'      m_Flags = m_Flags And Not WhichFlag
'   End If
'End Property
'Public Property Get FlagsIn(ByVal WhichFlag As DsGetDcNameFlags) As Boolean
'   FlagsIn = m_Flags And WhichFlag
'End Property

'
' twaugens, 8/7/2006, the ability to change to ANSI was turned off since
'                     ANSI has not been tested.
'Public Property Let Unicode(ByVal NewVal As Boolean)
'   m_Unicode = NewVal
'   'Call Refresh
'End Property

Public Property Get Unicode() As Boolean
'##SUMMARY Get value of the unicode, TRUE if Unicode is used, FALSE if ANSI.
   Unicode = m_Unicode
End Property

' ********************************************
'  Public Properties: Read-only
' ********************************************
Public Property Get ClientSiteName() As String
'##SUMMARY Get value of the name of the site that the computer belongs to.
   ClientSiteName = m_DCI.ClientSiteName
End Property

Public Property Get DcSiteName() As String
'##SUMMARY Get value of the name of the site where the domain controller is located.
   DcSiteName = m_DCI.DcSiteName
End Property

Public Property Get DnsForestName() As String
'##SUMMARY Get value of the name of the domain at the root of the Directory Service tree.
   DnsForestName = m_DCI.DnsForestName
End Property

Public Property Get DomainControllerAddress() As String
'##SUMMARY Get value of address of the discovered domain controller. The address is prefixed with "\\".
   DomainControllerAddress = m_DCI.DomainControllerAddress
End Property

Public Property Get DomainControllerAddressType() As DcAddressTypes
'##SUMMARY Get value of the domain controller address type; address is a string IP address or address is a NetBIOS name.
   DomainControllerAddressType = m_DCI.DomainControllerAddressType
End Property

Public Property Get DomainControllerName() As String
'##SUMMARY Get value of the computer name of the discovered domain controller.
   DomainControllerName = m_DCI.DomainControllerName
End Property

Public Property Get DomainGuid() As String
'##SUMMARY Get value of the domain GUID.
   ' Max length for GUID string
   Const GUID_LENGTH As Long = 38
   ' Allocate storage space.
   DomainGuid = Space$(GUID_LENGTH)
   ' Convert raw GUID to formatted string.
   Call StringFromGUID2(m_DCI.DomainGuid, StrPtr(DomainGuid), GUID_LENGTH + 1)
End Property

Public Property Get FlagsOut() As Long
'##SUMMARY Get a value of the set of flags that describe the domain controller.

   FlagsOut = m_DCI.Flags
End Property

Public Property Get LastError() As dsGetDcNameErrors
'##SUMMARY Get the last error message by error number.

   LastError = m_LastError
End Property

Public Property Get LastErrorText() As String
'##SUMMARY Get the last error message as text.

   Dim sRet As String
   Select Case m_LastError
      Case NO_ERROR
         sRet = "No error."
      Case ERROR_NOT_ENOUGH_MEMORY
         sRet = "Insufficient memory is available."
      Case ERROR_INVALID_FLAGS
         sRet = "The Flags parameter has conflicting or superfluous bits set."
      Case ERROR_INVALID_DOMAINNAME
         sRet = "The format of the specified DomainName is invalid."
      Case ERROR_NO_SUCH_DOMAIN
         sRet = "No DC is available for the specified domain or the domain does not exist."
   End Select
   LastErrorText = sRet
End Property

' ********************************************
'  Public Methods
' ********************************************
Public Sub Refresh()
'##SUMMARY Get a domain controller reference to the nearest domain controller.
'##REMARKS Refresh gets called just as soon as you create an object of _
           type CGetDcNameUSGS.
'##HISTORY twaugens, 8/7/2006, Added error trap in case the Unicode version _
           of the code was not exported.
           
   ' Only try if call is available.
   If m_Supported Then
      ' Offer choice of Unicode/ANSI just
      ' to show different approaches.
      If m_Unicode Then
         Call GetDcNameW(m_DCI.DomainName, m_ComputerName)
      Else
         Call GetDcNameA(m_DCI.DomainName, m_ComputerName)
      End If
   Else
   ' twaugens, 8/7/2006, Added error trap in case the Unicode version _
     of the code was not exported.
      MsgBox "Runtime error in CGetDcNameUSGS.Refresh: " & "The interface to get the domain controller's name was not exported and can not be used."
   End If
End Sub

' ********************************************
'  Private Methods
' ********************************************
Private Function Exported(ByVal ModuleName As String, _
                          ByVal ProcName As String) As Boolean
'##SUMMARY Determine if a procedure is exported.
'##PARAM ModuleName I Name of DLL with out .dll suffix.
'##PARAM ProcName I Name of procedure in DLL.
'##RETURNS TRUE if the procedure has been exported and is callable.

   Dim hModule As Long
   Dim lpProc As Long
   Dim FreeLib As Boolean
   
   ' check first to see if the module is already
   ' mapped into this process.
   hModule = GetModuleHandle(ModuleName)
   If hModule = 0 Then
      ' need to load module into this process.
      hModule = LoadLibrary(ModuleName)
      FreeLib = True
   End If
   
   ' if the module is mapped, check procedure
   ' address to verify it's exported.
   If hModule Then
      lpProc = GetProcAddress(hModule, ProcName)
      Exported = (lpProc <> 0)
   End If
   
   ' unload library if we loaded it here.
   If FreeLib Then Call FreeLibrary(hModule)
End Function

Private Function GetDcNameA(Optional ByVal Domain As String = vbNullString, _
                            Optional ByVal Computer As String = vbNullString) As Boolean
'##SUMMARY ANSI call to locate a Windows NT Domain Controller.
'##REMARKS See: <br> _
           http://msdn.microsoft.com/library/default.asp?url=/library/en-us/ad/ad/dsgetdcname.asp
'##PARAM Domain I Domain name that computer is located on; if blank then I think the _
                  local computer's Netlogon service is used to find the domain controller name. _
                  AWUDS should use blank.
'##PARAM Computer I Name of the remote computer that should process the request using _
                    the Netlogon service. If blank the local computer is used. In the _
                    case of AWUDS blank should always be used.
'##REMARKS NOTE: GetDcNameA has not been tested. AWUDS uses the Unicode version which was _
                 hard coded in the initialization routine.
'##HISTORY twaugens, 8/7/2006 removed hexdump debug and added error message when getting _
           domain controller.

   Dim lpBuffer As Long
   Dim dci As DOMAIN_CONTROLLER_INFO
   
   ' Make call to API.
   m_LastError = DsGetDcNameA(Computer, Domain, ByVal 0&, vbNullString, m_Flags, lpBuffer)
   If m_LastError = NO_ERROR Then
      ' Recover structure into one we can deal with in VB
      Call CopyMemory(dci, ByVal lpBuffer, Len(dci))
            
      ' Transfer contents to member structure.
      m_DCI.DomainControllerName = PointerToStringA(dci.DomainControllerName)
      m_DCI.DomainControllerAddress = PointerToStringA(dci.DomainControllerAddress)
      m_DCI.DomainControllerAddressType = dci.DomainControllerAddressType
      m_DCI.DomainGuid = dci.DomainGuid
      m_DCI.DomainName = PointerToStringA(dci.DomainName)
      m_DCI.DnsForestName = PointerToStringA(dci.DnsForestName)
      m_DCI.Flags = dci.Flags
      m_DCI.DcSiteName = PointerToStringA(dci.DcSiteName)
      m_DCI.ClientSiteName = PointerToStringA(dci.ClientSiteName)
      
      ' Release API allocated memory.
      Call NetApiBufferFree(lpBuffer)
   
   Else
' twaugens, 8/7/2006 added error message when getting _
           domain controller.
      MsgBox "Error getting domain controller: " & vbCrLf & _
             Me.LastErrorText
   End If
   
   ' Return results
   GetDcNameA = CBool(m_LastError)
End Function

Private Function GetDcNameW(Optional ByVal Domain As String = "", _
                            Optional ByVal Computer As String = "") As Boolean
'##SUMMARY Unicode call to locate a Windows NT Domain Controller.
'##REMARKS See: <br> _
           http://msdn.microsoft.com/library/default.asp?url=/library/en-us/ad/ad/dsgetdcname.asp
'##PARAM Domain I Domain name that computer is located on; if blank then I think the _
                  local computer's Netlogon service is used to find the domain controller name. _
                  AWUDS should use blank.
'##PARAM Computer I Name of the remote computer that should process the request using _
                    the Netlogon service. If blank the local computer is used. In the _
                    case of AWUDS blank should always be used.
'##HISTORY twaugens, 8/7/2006 removed hexdump debug and added error message when getting _
           domain controller.
           
   Dim lpDomain As Long
   Dim lpComputer As Long
   Dim lpBuffer As Long
   Dim dci As DOMAIN_CONTROLLER_INFO
   
   ' Need pointers for domain and computer name;
   ' default to NULL if "".
   If Len(Domain) Then lpDomain = StrPtr(Domain)
   If Len(Computer) Then lpComputer = StrPtr(Computer)
   
   ' Make call to API.
   ' m_Flags = m_Flags Or DS_PDC_REQUIRED
   m_LastError = DsGetDcNameW(lpComputer, lpDomain, ByVal 0&, 0&, m_Flags, lpBuffer)
   If m_LastError = NO_ERROR Then
      ' Recover structure into one we can deal with in VB
      Call CopyMemory(dci, ByVal lpBuffer, Len(dci))
           
      ' Transfer contents to member structure.
      m_DCI.DomainControllerName = PointerToStringW(dci.DomainControllerName)
      m_DCI.DomainControllerAddress = PointerToStringW(dci.DomainControllerAddress)
      m_DCI.DomainControllerAddressType = dci.DomainControllerAddressType
      m_DCI.DomainGuid = dci.DomainGuid
      m_DCI.DomainName = PointerToStringW(dci.DomainName)
      m_DCI.DnsForestName = PointerToStringW(dci.DnsForestName)
      m_DCI.Flags = dci.Flags
      m_DCI.DcSiteName = PointerToStringW(dci.DcSiteName)
      m_DCI.ClientSiteName = PointerToStringW(dci.ClientSiteName)
      
      ' Release API allocated memory.
      Call NetApiBufferFree(lpBuffer)
      
   Else
' twaugens, 8/9/2006 added error message when getting _
           domain controller.
      MsgBox "Could not obtain domain controller (DC) to check user access, local access will be attempted. " & vbCrLf & _
             Me.LastErrorText, vbInformation, "Computer may be disconnected from DOI network."
   End If
   
   ' Return results
   GetDcNameW = CBool(m_LastError)
End Function

Private Function PointerToStringA(ByVal lpStringA As Long) As String
'##SUMMARY Converts pointer of a ANSI string to string that can be used in VB.
'##PARAM lpStringA I Pointer to the ANSI string.
'##RETURNS String.

   Dim Buffer() As Byte
   Dim nLen As Long
   
   If lpStringA Then
      nLen = lstrlenA(ByVal lpStringA)
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringA, nLen
         PointerToStringA = StrConv(Buffer, vbUnicode)
      End If
   End If
End Function

Private Function PointerToStringW(ByVal lpStringW As Long) As String
'##SUMMARY Converts pointer of a Unicode string to string that can be used in VB.
'##PARAM lpStringW I Pointer to the Unicode string.
'##RETURNS String.

   Dim Buffer() As Byte
   Dim nLen As Long

   If lpStringW Then
      nLen = lstrlenW(lpStringW) * 2
      If nLen Then
         ReDim Buffer(0 To (nLen - 1)) As Byte
         CopyMemory Buffer(0), ByVal lpStringW, nLen
         PointerToStringW = Buffer
      End If
   End If
End Function


