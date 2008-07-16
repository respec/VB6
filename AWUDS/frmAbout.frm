VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About AWUDS"
   ClientHeight    =   6045
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8430
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "System Info ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      TabIndex        =   2
      ToolTipText     =   "Only availabe to Administrators of this computer."
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      TabIndex        =   0
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   132
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit
'##MODULE_NAME frmAbout
'##MODULE_DATE unknown
'##MODULE_AUTHOR Todd W. Augenstein (USGS) and Mark Gray (Aqua Terra, Consultants)
'##MODULE_SUMMARY Displays technical information about the AWUDS applications.

'##SUMMARY Enter the date that AWUDS was compiled on for the current build:
Const CompileDate As String = "August 24, 2006"

' Reg Key Security Options..
Const KEY_ALL_ACCESS = &H2003F
                                          
' Reg Key ROOT Types..
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub cmdClose_Click()
Attribute cmdClose_Click.VB_Description = "Close the About SWUDS Window."
'##SUMMARY Close the About AWUDS Window.
  Unload Me
End Sub

Private Sub cmdSysInfo_Click()
Attribute cmdSysInfo_Click.VB_Description = "Open the System Information Window."
'##SUMMARY Open the System Information Window.

  Call StartSysInfo
End Sub

Private Sub Form_Load()
'##SUMMARY Load the About AWUDS Form and display technical information about AWUDS.
'##USECASE On the About AWUDS Window, display the version number.
'##USECASE On the About AWUDS Window, inform users where inquiries can be made.
'##USECASE On the About AWUDS Window, list the AWUDS help email address.
  
  Dim s As String
  
  s = "AWUDS" & vbCrLf
  s = s & "Version " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf
  s = s & CompileDate & vbCrLf
  s = s & "-----------" & vbCrLf & vbCrLf
  s = s & "Data Directory:" & vbCrLf & AwudsDataPath & vbCrLf & vbCrLf
  s = s & "Report Directory:" & vbCrLf & ReportPath & vbCrLf & vbCrLf
  s = s & "Inquiries about this software should be directed to:" & vbCrLf
  s = s & vbCrLf
  s = s & "U.S. Geological Survey" & vbCrLf
  s = s & "National Water Information System" & vbCrLf
  s = s & "MS 437" & vbCrLf
  s = s & "12201 Sunrise Valley Drive" & vbCrLf
  s = s & "Reston, VA 20192 " & vbCrLf & vbCrLf
  s = s & "To get help running AWUDS send mail to: GS-W help AWUDS" & vbCrLf
  s = s & "-------------- " & vbCrLf
  lblInfo.Caption = s
  
End Sub

Public Sub StartSysInfo()
Attribute StartSysInfo.VB_Description = "Starts the Micosoft, System Information application."
'##SUMMARY Starts the Micosoft, System Information application.
'##AUTHOR Mark Gray (Aqua Terra, Consultants)
'##REMARKS Find the location of the System Information application _
           in the registry and then launches it.
'##REMARKS On the About AWUDS Window, allow users access to system _
           information about their hardware.

  On Error GoTo SysInfoErr

  Dim rc As Long
  Dim SysInfoPath As String
  
  ' Try To Get System Info Program Path\Name From Registry..
  If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
  ' Try To Get System Info Program Path Only From Registry..
  ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
    ' Validate Existance Of Known 32 Bit File Version
    If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
      SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
    ' Error - File Can Not Be Found..
    Else
      GoTo SysInfoErr
    End If
  ' Error - Registry Entry Can Not Be Found..
  Else
    GoTo SysInfoErr
  End If

  Call Shell(SysInfoPath, vbNormalFocus)
  
  Exit Sub
SysInfoErr:
  MsgBox "System Information is only available to users that have Administrator rights.", vbOKOnly
End Sub

Public Function GetKeyValue(ByRef KeyRoot As Long, _
                            ByRef KeyName As String, _
                            ByRef SubKeyRef As String, _
                            ByRef KeyVal As String) As Boolean
Attribute GetKeyValue.VB_Description = "Get the associated value of a registry key."
'##SUMMARY Get the associated value of a registry key.
'##AUTHOR Mark Gray (Aqua Terra, Consultants)
'##PARAM KeyRoot (I) Handle of the root key to be opened/searched {HKEY_LOCAL_MACHINE..}
'##PARAM KeyName (I) Path name of the key to open.
'##PARAM SubKeyRef (I) Part of key to return a value for.
'##PARAM KeyVal (O) The value of the key.

  Dim i As Long                                           ' Loop Counter
  Dim rc As Long                                          ' Return Code
  Dim hKey As Long                                        ' Handle To An Open Registry Key
  Dim hDepth As Long                                      '
  Dim KeyValType As Long                                  ' Data Type Of A Registry Key
  Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
  Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
  '------------------------------------------------------------
  ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE..}
  '------------------------------------------------------------
  rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key

  If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error..

  tmpVal = String$(1024, 0)                               ' Allocate Variable Space
  KeyValSize = 1024                                       ' Mark Variable Size

  '------------------------------------------------------------
  ' Retrieve Registry Key Value..
  '------------------------------------------------------------
  rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                          
  If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors

  If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String..
    tmpVal = Left(tmpVal, KeyValSize - 1)                 ' Null Found, Extract From String
  Else                                                    ' WinNT Does NOT Null Terminate String..
    tmpVal = Left(tmpVal, KeyValSize)                     ' Null Not Found, Extract String Only
  End If
  '------------------------------------------------------------
  ' Determine Key Value Type For Conversion..
  '------------------------------------------------------------
  Select Case KeyValType                                  ' Search Data Types..
  Case REG_SZ                                             ' String Registry Key Data Type
       KeyVal = tmpVal                                    ' Copy String Value
  Case REG_DWORD                                          ' Double Word Registry Key Data Type
       For i = Len(tmpVal) To 1 Step -1                   ' Convert Each Bit
          KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
       Next
       KeyVal = Format$("&h" + KeyVal)                    ' Convert Double Word To String
  End Select

  GetKeyValue = True                                      ' Return Success
  rc = RegCloseKey(hKey)                                  ' Close Registry Key
  Exit Function                                           ' Exit

GetKeyError:    ' Cleanup After An Error Has Occured..
  KeyVal = ""                                             ' Set Return Val To Empty String
  GetKeyValue = False                                     ' Return Failure
  rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

