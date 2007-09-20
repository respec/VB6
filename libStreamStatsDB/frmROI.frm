VERSION 5.00
Begin VB.Form frmROI 
   Caption         =   "Form1"
   ClientHeight    =   4176
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   7212
   LinkTopic       =   "Form1"
   ScaleHeight     =   4176
   ScaleWidth      =   7212
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOptParms 
      Caption         =   "Optional Parameters"
      Height          =   1812
      Left            =   3840
      TabIndex        =   23
      Top             =   2280
      Width           =   2412
      Begin VB.CheckBox chkOptParms 
         Caption         =   "Shape"
         Height          =   252
         Index           =   3
         Left            =   960
         TabIndex        =   31
         Top             =   1440
         Width           =   1092
      End
      Begin VB.CheckBox chkOptParms 
         Caption         =   "Slope"
         Height          =   252
         Index           =   2
         Left            =   960
         TabIndex        =   30
         Top             =   1080
         Width           =   1092
      End
      Begin VB.CheckBox chkOptParms 
         Caption         =   "Distance"
         Height          =   252
         Index           =   1
         Left            =   960
         TabIndex        =   29
         Top             =   720
         Width           =   1092
      End
      Begin VB.CheckBox chkOptParms 
         Caption         =   "Drainage Area"
         Height          =   252
         Index           =   0
         Left            =   960
         TabIndex        =   28
         Top             =   360
         Value           =   1  'Checked
         Width           =   1332
      End
      Begin ATCoCtl.ATCoText txtSlope 
         Height          =   252
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   612
         _ExtentX        =   1080
         _ExtentY        =   445
         InsideLimitsBackground=   16777215
         OutsideHardLimitBackground=   8421631
         OutsideSoftLimitBackground=   8454143
         HardMax         =   1
         HardMin         =   0
         SoftMax         =   -999
         SoftMin         =   -999
         MaxWidth        =   -999
         Alignment       =   1
         DataType        =   0
         DefaultValue    =   "0"
         Value           =   "0"
         Enabled         =   -1  'True
      End
      Begin ATCoCtl.ATCoText txtShape 
         Height          =   252
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   612
         _ExtentX        =   1080
         _ExtentY        =   445
         InsideLimitsBackground=   16777215
         OutsideHardLimitBackground=   8421631
         OutsideSoftLimitBackground=   8454143
         HardMax         =   -999
         HardMin         =   -999
         SoftMax         =   -999
         SoftMin         =   -999
         MaxWidth        =   -999
         Alignment       =   1
         DataType        =   0
         DefaultValue    =   "0"
         Value           =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label lblShape 
         Alignment       =   2  'Center
         Caption         =   "Shape"
         Height          =   252
         Left            =   120
         TabIndex        =   26
         Top             =   1080
         Width           =   612
      End
      Begin VB.Label lblSlope 
         Alignment       =   2  'Center
         Caption         =   "Slope"
         Height          =   252
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   612
      End
   End
   Begin VB.TextBox txtState 
      Height          =   288
      Left            =   720
      TabIndex        =   22
      Top             =   240
      Width           =   372
   End
   Begin VB.Frame fraLong 
      Caption         =   "Station Longitude"
      Height          =   1332
      Left            =   1920
      TabIndex        =   13
      Top             =   2760
      Width           =   1692
      Begin ATCoCtl.ATCoText txtLong 
         Height          =   252
         Index           =   0
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   612
         _ExtentX        =   1080
         _ExtentY        =   445
         InsideLimitsBackground=   16777215
         OutsideHardLimitBackground=   8421631
         OutsideSoftLimitBackground=   8454143
         HardMax         =   360
         HardMin         =   0
         SoftMax         =   -999
         SoftMin         =   -999
         MaxWidth        =   -999
         Alignment       =   1
         DataType        =   0
         DefaultValue    =   "82"
         Value           =   "82"
         Enabled         =   -1  'True
      End
      Begin ATCoCtl.ATCoText txtLong 
         Height          =   252
         Index           =   1
         Left            =   960
         TabIndex        =   18
         Top             =   600
         Width           =   612
         _ExtentX        =   1080
         _ExtentY        =   445
         InsideLimitsBackground=   16777215
         OutsideHardLimitBackground=   8421631
         OutsideSoftLimitBackground=   8454143
         HardMax         =   60
         HardMin         =   0
         SoftMax         =   -999
         SoftMin         =   -999
         MaxWidth        =   -999
         Alignment       =   1
         DataType        =   0
         DefaultValue    =   "00"
         Value           =   "00"
         Enabled         =   -1  'True
      End
      Begin ATCoCtl.ATCoText txtLong 
         Height          =   252
         Index           =   2
         Left            =   960
         TabIndex        =   19
         Top             =   960
         Width           =   612
         _ExtentX        =   1080
         _ExtentY        =   445
         InsideLimitsBackground=   16777215
         OutsideHardLimitBackground=   8421631
         OutsideSoftLimitBackground=   8454143
         HardMax         =   60
         HardMin         =   0
         SoftMax         =   -999
         SoftMin         =   -999
         MaxWidth        =   -999
         Alignment       =   1
         DataType        =   0
         DefaultValue    =   "00"
         Value           =   "00"
         Enabled         =   -1  'True
      End
      Begin VB.Label lblLong 
         Alignment       =   1  'Right Justify
         Caption         =   "seconds:"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   732
      End
      Begin VB.Label lblLong 
         Alignment       =   1  'Right Justify
         Caption         =   "minutes:"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   732
      End
      Begin VB.Label lblLong 
         Alignment       =   1  'Right Justify
         Caption         =   "degrees:"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   732
      End
   End
   Begin VB.Frame fraLat 
      Caption         =   "Station Latitude"
      Height          =   1332
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1692
      Begin ATCoCtl.ATCoText txtLat 
         Height          =   252
         Index           =   1
         Left            =   960
         TabIndex        =   11
         Top             =   600
         Width           =   612
         _ExtentX        =   1080
         _ExtentY        =   445
         InsideLimitsBackground=   16777215
         OutsideHardLimitBackground=   8421631
         OutsideSoftLimitBackground=   8454143
         HardMax         =   60
         HardMin         =   0
         SoftMax         =   -999
         SoftMin         =   -999
         MaxWidth        =   -999
         Alignment       =   1
         DataType        =   0
         DefaultValue    =   "30"
         Value           =   "30"
         Enabled         =   -1  'True
      End
      Begin ATCoCtl.ATCoText txtLat 
         Height          =   252
         Index           =   2
         Left            =   960
         TabIndex        =   12
         Top             =   960
         Width           =   612
         _ExtentX        =   1080
         _ExtentY        =   445
         InsideLimitsBackground=   16777215
         OutsideHardLimitBackground=   8421631
         OutsideSoftLimitBackground=   8454143
         HardMax         =   60
         HardMin         =   0
         SoftMax         =   -999
         SoftMin         =   -999
         MaxWidth        =   -999
         Alignment       =   1
         DataType        =   0
         DefaultValue    =   "00"
         Value           =   "00"
         Enabled         =   -1  'True
      End
      Begin ATCoCtl.ATCoText txtLat 
         Height          =   252
         Index           =   0
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   612
         _ExtentX        =   1080
         _ExtentY        =   445
         InsideLimitsBackground=   16777215
         OutsideHardLimitBackground=   8421631
         OutsideSoftLimitBackground=   8454143
         HardMax         =   360
         HardMin         =   0
         SoftMax         =   -999
         SoftMin         =   -999
         MaxWidth        =   -999
         Alignment       =   1
         DataType        =   0
         DefaultValue    =   "35"
         Value           =   "35"
         Enabled         =   -1  'True
      End
      Begin VB.Label lblLat 
         Alignment       =   1  'Right Justify
         Caption         =   "seconds:"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   732
      End
      Begin VB.Label lblLat 
         Alignment       =   1  'Right Justify
         Caption         =   "minutes:"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   732
      End
      Begin VB.Label lblLat 
         Alignment       =   1  'Right Justify
         Caption         =   "degrees:"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   732
      End
   End
   Begin ATCoCtl.ATCoText txtUserParm 
      Height          =   252
      Index           =   0
      Left            =   3480
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   492
      _ExtentX        =   868
      _ExtentY        =   445
      InsideLimitsBackground=   16777215
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      HardMax         =   -999
      HardMin         =   0
      SoftMax         =   -999
      SoftMin         =   -999
      MaxWidth        =   -999
      Alignment       =   1
      DataType        =   0
      DefaultValue    =   ""
      Value           =   ""
      Enabled         =   -1  'True
   End
   Begin VB.Frame fraRegions 
      Caption         =   "State Region"
      Height          =   612
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   3012
      Begin VB.OptionButton rdoRegion 
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2772
      End
   End
   Begin VB.TextBox txtStaName 
      Height          =   288
      Left            =   120
      TabIndex        =   1
      Top             =   1044
      Width           =   1812
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   6360
      TabIndex        =   0
      Top             =   3480
      Width           =   732
   End
   Begin VB.Label lblState 
      Alignment       =   2  'Center
      Caption         =   "State"
      Height          =   252
      Left            =   120
      TabIndex        =   21
      Top             =   240
      Width           =   492
   End
   Begin VB.Label lblUserParm 
      Alignment       =   2  'Center
      Height          =   252
      Index           =   0
      Left            =   2040
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label lblStaName 
      Caption         =   "Station Name"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1932
   End
End
Attribute VB_Name = "frmROI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants
Dim StaName$ 'ungaged station name
Dim RegName$
Dim RegID&
Dim StateID&
Dim NumPeaks&            'number of return periods (NC=8, TX=6)
Dim StaCnt&              'number of stations in state
Dim ParmCnt&             'number of parameters in ROIstations table/ROIparmList
Dim StaLat(3) As Single, StaLong(3) As Single 'ungaged Lat/Long in Deg/Min/Sec
Dim Vars() As Single     'Values from ROIStations - This and next 2 have dim(StaCnt,ParmCnt)
Dim VarSDs() As Single   'Standard Deviation - calculated
Dim UserVars() As Single 'ungaged station values
Dim Lat() As Single      'decimal degrees latitude for gaged stations
Dim Lng() As Single      'decimal degrees longitude for gaged stations
Dim Distance() As Single 'dim(StaCnt) "distance" measurement calculated to ungaged station
Dim Sum() As Single      'Used in calc VarSDs dim(ParmCnt)
Dim Mcon() As Single     '2-d from StateMatrices table
Dim Rhoc() As Single     '2-d from rhotable for NC, calculated for TX
Dim UseFld() As Boolean  '(0, x) is field x in database (from ROIstations in database), (1, x) is it used in distance calculation from ROIparmsByState table
Dim INDX() As Integer    'dim(StaCnt) ranks Distance Distance(INDX(1)) = shortest distance, Vars(INDX(1), 13) = drainage area of closest station
Dim myDB As Database     'NSS.mdb
Dim RhoRec As Recordset  'reading rhoc matrix from RHo table
Const MaxSites& = 50     'maximum number of sites that can be selected
Const MaxInd& = 10       'maximum number of independent variables including the regression constant
Dim Ak(8) As Single      'a constant that applies to return periods
Dim Cf2!, Cf25!, Cf100!  'climate factors calculated in subroutine CFX
Dim Pklab(8) As String   'labels for return periods
Dim Ne&                  'number of independent parameters used in distance calculation (not inluding RegnID)
Dim Nsites&              'number of stations to be used in final flood frequency calculations
Dim outfile&             'used for dummy outfile in current stand-alone program

'remaining variables are used in statistical calculations
Const alpha_area = 1.5
Dim Yvar As Single, Sig As Single, Alpha As Single, _
    Theta As Single, UserDA As Single
Dim Sta(MaxSites) As Single             'time sampling error for stations
Dim X(MaxSites, MaxInd) As Single       'calculated x-coordinate for # of closest sites
Dim Y(MaxSites, 1) As Single            'calculated y-coordinate for # of closest sites
Dim Xt(MaxInd, MaxSites) As Single
Dim XtXinv(MaxInd, MaxInd) As Single
Dim Cov(MaxSites, MaxSites) As Single   'covariance array
Dim Wt(MaxSites, MaxSites) As Single
Dim e(MaxSites, 1) As Single
Dim ET(1, MaxSites) As Single
Dim c1(1, 1) As Single
Dim Bhat(MaxInd, 1) As Single
Dim Gamasq!, Atse!, Atscov!

Private Sub txtUserParm_LostFocus(index As Integer)
  Dim i&, j&
 
  If IsNumeric(txtUserParm(index).Value) Then
    If txtUserParm(index).Value < 0 Then
      MsgBox "The entered value must be a positive number"
      txtUserParm(index).Value = Mid(txtUserParm(index).Value, 2)
    Else
      j = index
      For i = 13 To ParmCnt
        If UseFld(1, i) And j = 0 Then
          UserVars(i) = txtUserParm(index).Value
        Else
          j = j - 1
        End If
      Next i
    End If
  Else
    MsgBox "The " & lblUserParm(index).Caption & " value must be a positive number"
    If IsEmpty(txtUserParm(index)) Then
      txtUserParm(index).Value = ""
    Else
      txtUserParm(index).Value = UserVars(index)
    End If
  End If
End Sub

Private Sub Form_Load()
  Set myDB = OpenDatabase("C:\vbExperimental\libNSS\NSS.mdb", False, True)
End Sub

Private Sub txtLat_Change(index As Integer)
  If IsNumeric(txtLat(index).Value) Then
    If txtLat(index).Value < 0 Then
      MsgBox "The coordinate must be a non-negative number"
      txtLat(index).Value = Mid(txtLat(index).Value, 2)
    ElseIf txtLat(index).Value > 360 And index = 0 Then
      MsgBox "The degrees coordinate must be between 0 and 360"
      txtLat(index).Value = ""
    ElseIf txtLat(index).Value > 60 And (index = 1 Or index = 2) Then
      MsgBox "The minutes and seconds coordinates must be between 0 and 60"
      txtLat(index).Value = ""
    Else
      StaLat(index + 1) = txtLat(index).Value
    End If
  Else
    MsgBox "The coordinate must be a non-negative number"
    txtLat(index).Value = ""
    StaLat(index + 1) = -999
  End If
End Sub

Private Sub txtLong_Change(index As Integer)
  If IsNumeric(txtLong(index).Value) Then
    If txtLong(index).Value < 0 Then
      MsgBox "The coordinate must be a non-negative number"
      txtLong(index).Value = Mid(txtLong(index).Value, 2)
    ElseIf txtLong(index).Value > 360 And index = 0 Then
      MsgBox "The degrees coordinate must be between 0 and 360"
      txtLat(index).Value = ""
    ElseIf txtLong(index).Value > 60 And (index = 1 Or index = 2) Then
      MsgBox "The minutes and seconds coordinates must be between 0 and 60"
      txtLat(index).Value = ""
    Else
      StaLong(index + 1) = txtLong(index).Value
    End If
  Else
    MsgBox "The coordinate must be a non-negative number"
    txtLong(index).Value = ""
    StaLong(index + 1) = -999
  End If
End Sub

'Private Sub txtSlope_Change()
'  If IsNumeric(txtSlope.Value) Then
'    If txtSlope.Value < 0 Or txtSlope.Value > 1# Then
'      MsgBox "The slope must be a decimal between 0.0 and 1.0"
'      txtSlope.Value = ""
'    Else
'      uservars(18) = txtSlope.Value
'    End If
'  Else
'    MsgBox "The slope must be a decimal between 0.0 and 1.0"
'    txtSlope.Value = ""
'  End If
'End Sub
'
'Private Sub txtShape_Change()
'  If IsNumeric(txtShape.Value) Then
'    If txtShape.Value < 0 Or txtShape.Value > 10# Then
'      MsgBox "The slope must be a decimal between 0.0 and 1.0"
'      txtShape.Value = ""
'    Else
'      uservars(17) = txtShape.Value
'    End If
'  Else
'    MsgBox "The shape must be a number between 0 and 10"
'    txtShape.Value = ""
'  End If
'End Sub

Private Sub txtStaName_Change()
  StaName = txtStaName.Text
End Sub

Private Sub cmdrun_click()

' Program to estimate flood frequency in North Carolina
  Dim i&, j&
  Dim selected As Boolean
  Dim staX As Single, staY As Single, dDist As Single, dSlope As Single, _
      dShape As Single, dArea As Single
  Dim RHO As Single, ss As Single, Years As Single, yhat As Single
  Dim icall As Integer, jpeak As Integer, ireg As Integer, ysav As Integer
  
  If txtUserParm(0).Value = "" Then MsgBox "You must enter a positive number for the drainage area"
  If txtStaName.Text = "" Then txtStaName = "No Name"
  For i = 0 To rdoRegion.Count - 1
    If rdoRegion(i) = True Then
      selected = True
      Exit For
    End If
  Next i
  If Not selected Then MsgBox "You must select a state region"
  selected = True
  For i = 1 To 3
    If txtLat(i - 1).Value = "" Then
      MsgBox "You must complete the '" & Left(lblLat(i - 1).Caption, Len(lblLat(i - 1).Caption) - 1) & "' latitude coordinates for the station"
      selected = False
    Else
      StaLat(i) = txtLat(i - 1).Value
    End If
    If txtLong(i - 1).Value = "" Then
      MsgBox "You must complete the '" & Left(lblLong(i - 1).Caption, Len(lblLong(i - 1).Caption) - 1) & "' longitude coordinates for the station"
      selected = False
    Else
      StaLong(i) = txtLong(i - 1).Value
    End If
  Next i
  If Not selected Or txtUserParm(0).Value = "" Then Exit Sub

  'Initialize parms for different states0
  Init
  NumPeaks = 8

  If StateID <> 27 Then BuildMatrix
  
  'compute standard deviation of various independent variables
  For j = 1 To ParmCnt
    Select Case j
      Case 13 To 14, 17 To 18:
          VarSDs(j) = Sum(j) / StaCnt  'actually calcing avg of variable here
          Sum(j) = 0
          For i = 1 To StaCnt
            Sum(j) = Sum(j) + (Vars(i, j) - VarSDs(j)) ^ 2
          Next i
          VarSDs(j) = (Sum(j) / (StaCnt - 1)) ^ 0.5
    End Select
  Next j
  'VarSDs(13) = DA standard deviation
  'VarSDs(14) = CF standard deviation
  'VarSDs(17) = Slope standard deviation
  'VarSDs(18) = Shape standard deviation

  'Initialize and read in ungaged site information
  Nsites = 30
  icall = 0  'set to read climate factors 1st time thru
  'compute climate factors from lat and long coordinates
  CFX icall, StaLat(), StaLong()
  icall = 1
  
  If UserDA > 0 Then UserDA = Log10(CDbl(UserDA))
  If Cf25 > 0 Then Cf25 = Log10(CDbl(Cf25))
  If UserVars(17) > 0 Then UserVars(17) = Log10(CDbl(UserVars(17)))
  If UserVars(18) > 0 Then UserVars(18) = Log10(CDbl(UserVars(18)))
  
' Compute distances
  ' calculate actual distance and SD if used
  If UseFld(1, 0) Then
    staX = StaLong(1) + (StaLong(2) + (StaLong(3) / 60#)) / 60#
    staY = StaLat(1) + (StaLat(2) + (StaLat(3) / 60#)) / 60#
    Sum(0) = 0#
    'calculate actual distances then their standard deviations
    For i = 1 To StaCnt
      Vars(i, 0) = TASKER_DISTANCE(staY, Lat(i), staX, Lng(i))
'      'Convert miles into kilometers
'      If metric = "M" Then Distance(i) = Distance(i) * 1.609344
      Sum(0) = Sum(0) + Vars(i, 0)
    Next i
    VarSDs(0) = Sum(0) / StaCnt  'calcing avg distance from stations here
    For i = 1 To StaCnt
      Sum(0) = Sum(0) + (Distance(i) - VarSDs(0)) ^ 2
    Next i
    VarSDs(0) = (VarSDs(0) / (StaCnt - 1)) ^ 0.5  'calcing StdDev of actual distance
  End If
  ' calculate respective "distances"
  For i = 1 To StaCnt
    Distance(i) = 0
    For j = 0 To ParmCnt
      If UseFld(1, j) Then
        Distance(i) = Distance(i) + ((UserVars(j) - Vars(i, j)) / VarSDs(j)) ^ 2
      End If
    Next j
    Distance(i) = Distance(i) ^ 0.5
    Distance(i) = Distance(i) + (RegID - Vars(i, 2)) / 0.001
  Next i
    
' Rank distances
  INDEXX StaCnt, Distance(), INDX()

' Select closest Nsites stations for regression
' loop through dependent variables and do regression

  ysav = 0#
  Sig = 0#

  'open file for output
  outfile = FreeFile
  Open CurDir & "\NSS_Output.txt" For Output As outfile
  'write header to file
  Print #outfile, " REGION OF INFLUENCE METHOD" & vbCrLf & vbCrLf & _
      " Flood frequency estimates for station " & txtStaName.Text & _
      " in region " & RegName & vbCrLf & vbCrLf & _
      " Drainage Area: " & UserVars(13) & vbCrLf & vbCrLf & _
      " Data used for ROI Method:" & vbCrLf
  Print #outfile, " StaID   REGION" & vbTab & "LAT" & vbTab & vbTab _
      & "LNG" & vbTab & vbTab & "LOG(DA)" & vbTab & "LOG(CF)" & vbTab & "LOG(P2)" & _
      vbTab & "LOG(P5)" & vbTab & "LOG(P10)" & vbTab & "LOG(P25)" & vbTab & _
      "LOG(50)" & vbTab & "LOG(100)" & vbTab & "LOG(P200)" & vbTab & "LOG(P500)"
  'loop thru sites - calc avg StdDev for peak flow across all sites
  'TX does not have following loop
  For i = 1 To Nsites
    X(i, 1) = 1#
    X(i, 2) = Vars(INDX(i), 13)
    X(i, 3) = Vars(INDX(i), 17)
    X(i, 4) = Vars(INDX(i), 18)
    'build the Xt-transpose matrix
    For j = 1 To Ne
      Xt(j, i) = X(i, j)
    Next j
    
    'write values for this station to outfile
    Print #outfile, Vars(INDX(i), 1) & vbTab & Vars(INDX(i), 2) & vbTab & _
        Vars(INDX(i), 3) & vbTab & Vars(INDX(i), 4) & vbTab & Vars(INDX(i), 13) & vbTab & _
        Vars(INDX(i, 14)) & vbTab & Vars(INDX(i), 5) & vbTab & Vars(INDX(i), 6) & _
        vbTab & Vars(INDX(i), 7) & vbTab & Vars(INDX(i), 8) & vbTab & _
        Vars(INDX(i), 9) & vbTab & Vars(INDX(i), 10) & vbTab & _
        Vars(INDX(i), 11) & vbTab & Vars(INDX(i), 12)
    
    Sig = Sig + Vars(INDX(i), 15)
  Next i
  Sig = Sig / Nsites  ' = avg StDev of peak flow across all stations
  Print #outfile, vbCrLf & "For " & StaName & vbCrLf & "area = " & Round(10 ^ UserDA, 2) & _
        "    : cf25 = " & Cf25 & vbCrLf & vbCrLf & "RI" & vbTab & " PREDICTED(cfs)" & _
        vbTab & "- SE (%)" & vbTab & "+ SE (%)" & vbTab & "90% PRED INT" & vbCrLf
  
  For jpeak = 1 To 8
    For i = 1 To Nsites
      Y(i, 1) = Vars(INDX(i), 4 + jpeak)
    Next i
    Sum(0) = 0#
    ss = 0#

    ' compute regional average standard deviation
    ' and time sampling error, sta(i), for each site
    For i = 1 To Nsites
      Sum(0) = Sum(0) + Y(i, 1)
      ss = ss + Y(i, 1) ^ 2
    Next i
    Yvar = (ss - Sum(0) ^ 2 / Nsites) / (Nsites - 1#)
    Atse = 0#
    Atscov = 0#
    For i = 1 To Nsites
      For j = 1 To i
        Years = Mcon(INDX(i), INDX(j)) _
                / (Mcon(INDX(i), INDX(i)) * Mcon(INDX(j), INDX(j)))
        RHO = Rhoc(INDX(i), INDX(j))
        Cov(i, j) = RHO * Sig ^ 2 * (1 + RHO * 0.5 * Ak(jpeak) ^ 2) * Years
        Cov(j, i) = Cov(i, j)
        If (i = j) Then
          Sta(i) = Cov(i, i)
          Atse = Atse + Sta(i)
        Else
          Atscov = Atscov + Cov(i, j)
        End If
      Next j
    Next i

    'do regression
    Secant
'
'    'check to see if predicted value is greater than previous prediction
'    If (yhat < ysav) Then
'      MsgBox "CAUTION: Predicted T-year flow is smaller" & vbCrLf & _
'             "than T-Year flow with lower recurrence interval." & vbCrLf & _
'             "See output."
'    End If
    'output final model
    OutPut 99, jpeak, yhat, outfile
    ysav = yhat
  Next jpeak
  Close outfile
End Sub

Public Sub mltply(Prod() As Single, X() As Single, Y() As Single, k1&, k2&, k3&, N1&, N2&, N3&)

  Dim i&, j&, k&
  Dim Sum!
' --------------------------------------------------------------
'  X IS K1*K2 MATRIX
'  Y IS K2*K3 MATRIX
'  PROD = X*Y IS A K1*K3 MATRIX
' --------------------------------------------------------------
  For i = 1 To k1
    For k = 1 To k3
      Sum = 0#
      For j = 1 To k2
        Sum = Sum + X(i, j) * Y(j, k)
      Next j
      Prod(i, k) = Sum
    Next k
  Next i

End Sub

Public Sub invert(n&, Ndim&, det!, CovInv() As Single, Cov() As Single)

  Dim i&, im&, j&, k&
  Dim detl!, Sum!, temp!
  Dim b() As Single, a() As Single
  '--------------------------------------------------------------
  '  COV IS AN N*N MATRIX
  '  SUBROUTINE COMPUTES DETERMINANT OF COV AS COVINV
  '  B IS THE LOWER TRIANGULAR DECOMPOSITION OF COV
  '--------------------------------------------------------------
  ReDim b(n, n)
  ReDim a(n, n)
  If n = 2 Then
    det = Cov(1, 1) * Cov(2, 2) - Cov(1, 2) ^ 2
    temp = Cov(1, 1) / det
    CovInv(1, 1) = Cov(2, 2) / det
    CovInv(2, 2) = temp
    CovInv(1, 2) = -Cov(1, 2) / det
    CovInv(2, 1) = CovInv(1, 2)
  Else
    decomp n, Ndim, Cov(), b()
    detl = b(1, 1)
    For i = 2 To n
      If detl > 5E+19 Then
        MsgBox "ERROR--Numerical overflow on B(i,i) product expansion series."
        Stop
      End If
      detl = detl * b(i, i)
    Next i
'   Following if statement is a questionable fix
    If detl > 5E+19 Then
      MsgBox "ERR0R--Determinant is too large." & vbCrLf & _
             "Numerical overflow." & vbCrLf & _
             "Bad fix, but try fewer stations."
      Stop
    Else
      det = detl ^ 2
    End If
    
    a(1, 1) = 1# / b(1, 1)
    a(2, 2) = 1# / b(2, 2)
    a(2, 1) = -b(2, 1) * a(1, 1) * a(2, 2)

    For i = 3 To n
      a(i, i) = 1# / b(i, i)
      im = i - 1
      For k = 1 To im
        Sum = 0#
        For j = k To im
          Sum = Sum + b(i, j) * a(j, k)
        Next j
        a(i, k) = -Sum * a(i, i)
      Next k
    Next i

    For i = 1 To n
      For j = 1 To i
        Sum = 0#
        For k = i To n
          Sum = Sum + a(k, i) * a(k, j)
        Next k
        CovInv(i, j) = Sum
        CovInv(j, i) = Sum
      Next j
    Next i
  End If

End Sub

Public Sub decomp(n&, Ndim&, XLAM() As Single, b() As Single)

  Dim iis&, ism&, js&, jsm&, ks&
  Dim bh!, bn!
  '--------------------------------------------------------------
  ' CHOLESKY DECOMPOSITION  BB-TRANSPOSE = XLAM
  '--------------------------------------------------------------
  If XLAM(1, 1) <= 0# Or XLAM(2, 2) <= 0# Then
    MsgBox "IN DECOMP/ NDIM,XLAM 1-1,2-1,2-2,1-2 = " & Ndim & XLAM(1, 1) & XLAM(2, 1) & XLAM(2, 2) & XLAM(1, 2) & vbCrLf & _
           " COVARIANCE MATRIX NOT POSITIVE DEFINITE"
  End If
  b(1, 1) = XLAM(1, 1) ^ 0.5
  b(1, 2) = 0#
  b(2, 1) = XLAM(2, 1) / b(1, 1)
  b(2, 2) = (XLAM(2, 2) - b(2, 1) ^ 2) ^ 0.5

  If n <= 2 Then '2x2 or 1x1 matrix, exit early
    Exit Sub
  End If
  'Main decomposition algorithm
  For iis = 3 To n
    b(iis, 1) = XLAM(iis, 1) / b(1, 1)
    bn = XLAM(iis, iis) - b(iis, 1) ^ 2
    ism = iis - 1
    For js = 2 To ism
      jsm = js - 1
      bh = XLAM(iis, js)
      For ks = 1 To jsm
        bh = bh - b(iis, ks) * b(js, ks)
      Next ks
      b(iis, js) = bh / b(js, js)
      bn = bn - b(iis, js) ^ 2
    Next js
    If bn <= 0# Then
      MsgBox "COVARIANCE MATRIX NOT POSITIVE DEFINITE BN=" & bn
    End If
    'b(iis, iis) = (AMAX1(bn, 0#)) ^ 0.5
    If bn > 0 Then
      b(iis, iis) = bn ^ 0.5
    Else
      b(iis, iis) = 0
    End If
  Next iis

End Sub

Public Function STUTP(X!, n&) As Single

  Dim a!, b!, t!, Y!, z!
  Dim j&, nn&, keeplooping%
  Const rhpi = 0.63661977
  'STUDENT T PROBABILITY
  'STUTP = PROB( STUDENT T WITH N DEG FR  .LT.  X )
  'NOTE  -  PROB(ABS(T).GT.X) = 2.*STUTP(-X,N) (FOR X .GT. 0.)
  'SUBPGM USED - GAUSCF
  'REF - G.W. HILL, ACM ALGOR 395, OCTOBER 1970.
  'USGS - WK 12/79.
  STUTP = 0.5
  If n < 1 Then Exit Function
  nn = n
  z = 1#
  t = X ^ 2
  Y = t / nn
  b = 1# + Y
  If Not (nn >= 20 And t < nn Or nn > 200) Then
    If nn < 20 And t < 4# Then 'nested summation of cosine series
      Y = Y ^ 0.5
      a = Y
      If nn = 1 Then a = 0#
    Else 'tail series for large t
      a = b ^ 0.5
      Y = a * nn
      j = 0
      keeplooping = 1
      While keeplooping = 1
        j = j + 2
        If (a = z) Then
          nn = nn + 2
          z = 0#
          Y = 0#
          a = -a
          keeplooping = 0
        Else
          z = a
          Y = Y * (j - 1) / (b * j)
          a = a + Y / (nn + j)
        End If
      Wend
    End If
    keeplooping = 1
    While keeplooping = 1
      nn = nn - 2
      If nn <= 1 Then
        If nn = 0 Then a = a / (b ^ 0.5)
        If nn <> 0 Then a = (Atn(Y) + a / b) * rhpi
        STUTP = 0.5 * (z - a)
        If X > 0# Then STUTP = 1# - STUTP
        Exit Function
      Else
        a = (nn - 1) / (b * nn) * a + Y
      End If
    Wend
  End If

  'asymptotic series for large or noninteger N
  If Y > 0.000001 Then Y = Log(b)
  a = nn - 0.5
  b = 48# * a ^ 2
  Y = a * Y
  Y = (((((-0.4 * Y - 3.3) * Y - 24#) * Y - 85.5) / _
      (0.8 * Y ^ 2 + 100# + b) + Y + 3#) / b + 1#) * Y ^ 0.5
  STUTP = gauscf(-Y)
  If X > 0# Then STUTP = 1# - STUTP
End Function

Public Sub Secant()
'function WLS was passed as argument in original Fortran

  Dim f1!, f2!, f3!, fnew!, X1!, X2!, x3!, xnew!
  Dim i&, j&
 
  X1 = 0#
  x3 = 0#
  X2 = Yvar * 2#
'  X2 = reg_var * 2#  ???????
  f2 = WLS(X2)
  f1 = WLS(X1)
  If f1 < 0# Then 'midpoint search for good starting point
    For j = 1 To 3
      xnew = (X1 + X2) / 2#
      fnew = WLS(xnew)
      If fnew < 0# Then
        X1 = xnew
        fnew = fnew
      Else
        X2 = xnew
        f2 = fnew
      End If
    Next j

  'search for gama sq using secant serch
    For i = 1 To 30
      If (f2 - f1) = 0# Then
        MsgBox "STATUS: F2-F1 = 0, about to divide by 0"
        Stop
      End If
      x3 = X1 - f1 * (X2 - X1) / (f2 - f1)
      If x3 < 0# Then
        'x3 = AMIN1(X2, X1) / 2#
        If X1 < X2 Then
          x3 = X1 / 2#
        Else
          x3 = X2 / 2#
        End If
        f3 = WLS(x3)
      Else
        f3 = WLS(x3)
      End If
      If Abs(f3) < 0.0001 Then Exit For
      If Abs(f1) < Abs(f2) Then
        X2 = x3
        f2 = f3
      Else
        X1 = x3
        f1 = f3
      End If
    Next i
  End If
  Gamasq = x3

End Sub

Public Function WLS(GAMA2!) As Single
  'computes the regression coefficients and
  'the weighted SSE (Sum of Square Error)

  Dim det!, xtx!(MaxInd, MaxInd), work!(MaxInd, MaxSites), _
      work2!(MaxInd, MaxSites), work1!(1, MaxSites)
  Dim k&

  'Compute the weighting matrix Wt by first generating the estimated
  '  GLS covariance matrix, Cov(k,k) by updating the diagonal with the
  '  sum of the model error gama2 estimated from first call of SECANT and
  '  the station time sampling error.
  For k = 1 To Nsites
    Cov(k, k) = GAMA2 + Sta(k)
    If Cov(k, k) <= 0# Then
      MsgBox "Diagonal of variance-covariance matrix is <= 0" & vbCrLf & _
             "Location (k,k) of " & k & ", " & k
      Stop
    End If
  Next k
  'Inverting the Variance-Covariance Matrix
  Call invert(Nsites, MaxSites, det, Wt(), Cov())
  'Multiply Xt-transpose by the inverted covariance array (Wt) and produce work
  Call mltply(work, Xt, Wt, Ne, Nsites, Nsites, MaxInd, MaxInd, MaxSites)
  'Multiply Xt*Wt (work) by X to produce XtX
  Call mltply(xtx, work, X, Ne, Nsites, Ne, MaxInd, MaxInd, MaxSites)

  'COMPUTE THE REGRESSION COEFFICIENTS
  'Invert the Xt*Wt-1*X matrix
  Call invert(Ne, MaxInd, det, XtXinv, xtx)
  'Multiply the inverted variance-covariance matrix by Xt
  Call mltply(work, XtXinv, Xt, Ne, Ne, Nsites, MaxInd, MaxInd, MaxInd)
  'Multiply the results of above the Wt matrix
  Call mltply(work2, work, Wt, Ne, Nsites, Nsites, MaxInd, MaxInd, MaxSites)
  'Estimate the coefficients Bhat
  Call mltply(Bhat, work2, Y, Ne, Nsites, 1, MaxInd, MaxInd, MaxSites)

  'ERROR ESTIMATION
  'Determine the Error matrix E from by first estimating E as
  '  product of the independent variables and the coefficients
  Call mltply(e, X, Bhat, Nsites, Ne, 1, MaxSites, MaxSites, MaxInd)
  'The fill the E matrix with the residuals and make the transpose
  For k = 1 To Nsites
    e(k, 1) = Y(k, 1) - e(k, 1)
    ET(1, k) = e(k, 1)
  Next k
  'Multiply Et by Wt to produce a vector of error weights
  Call mltply(work1, ET, Wt, 1, Nsites, Nsites, 1, 1, MaxSites)
  'Multiply work1 by the Errors to produce a single value of error
  Call mltply(c1, work1, e, 1, Nsites, 1, 1, 1, MaxSites)
  'WLS is the weighted SSE sum of square error
  WLS = (Nsites - Ne) / c1(1, 1) - 1#

End Function

Private Sub CFX(icall, StaLat() As Single, StaLong() As Single)
  Dim cfRec As Recordset
  Dim sql$
  Dim C2a(30, 26) As Single, C25a(30, 26) As Single, C100a(30, 26) As Single, _
      sx As Single, sy As Single
  Dim q As Double, q0 As Double, q1 As Double, q2 As Double, phi As Double, _
      phi0 As Double, phi1 As Double, phi2 As Double, e As Double, e2 As Double, _
      m1 As Double, m2 As Double, n As Double, RHO As Double, rho0 As Double, _
      a As Double, C As Double, theta1 As Double, lam As Double, lam0 As Double, _
      X As Double, Y As Double, xoff As Double, yoff As Double, dp As Double, _
      mp As Double, sp As Double, dm As Double, mm As Double, sm As Double, _
      one As Double, two As Double, tpi As Double
  Dim i As Integer, ix As Integer, iy As Integer, j As Integer
  
' subroutine computes climate factors from lat and long

'     COMPUTES X,Y COORDINATES(KILOMETERS) BASED ON LAT AND LONG,
'     USING ALBERS EQUAL-AREA CONIC PROJECTION--WITH EARTH AS ELLIPSOID.
'     EQUATIONS FROM BULLETIN 1532, PAGES 96-99.
'     A = a, equatorial radius of ellipsoid
'     E2 = e*2, square of eccentricity of ellipsoid (0.006768658).
'     PHI1,PHI2 = standard parallels of latitude (29.5,45.5).
'     PHI0 = middle latitude, or latitude chosen as the origin of y-coordinate (
'     38.0).
'     LAM0 = central meridian of the map, or longitude chosen as the origin
'            of x-coordinate (96.0).
'     PHI = latitude of desired y-coordinate.
'     LAM = longitude of desired x-coordinate.
'     XOFF = offset to x-coordinate to make all positive for KRIGING.
'     YOFF = offset to y-coordinate to make all positive for KRIGING.

  a = 6378206.4
  e2 = 0.006768658
  tpi = 6.2831853
  phi1 = 29.5
  phi2 = 45.5
  phi0 = 38#
  lam0 = 96#
  
  xoff = 1000#
  yoff = 1500#
  one = 1#
  two = 2#
  e = e2 ^ 0.5
  phi0 = phi0 * tpi / 360#
  phi1 = phi1 * tpi / 360#
  phi2 = phi2 * tpi / 360#
  q0 = (one - e2) * (Sin(phi0) / (one - e2 * Sin(phi0) * Sin(phi0)) _
      - (one / (two * e)) * Log((one - e * Sin(phi0)) / (one + e * Sin(phi0))))
  q1 = (one - e2) * (Sin(phi1) / (one - e2 * Sin(phi1) * Sin(phi1)) _
      - (one / (two * e)) * Log((one - e * Sin(phi1)) / (one + e * Sin(phi1))))
  q2 = (one - e2) * (Sin(phi2) / (one - e2 * Sin(phi2) * Sin(phi2)) _
      - (one / (two * e)) * Log((one - e * Sin(phi2)) / (one + e * Sin(phi2))))
  m1 = Cos(phi1) / (one - e2 * Sin(phi1) * Sin(phi1)) ^ 0.5
  m2 = Cos(phi2) / (one - e2 * Sin(phi2) * Sin(phi2)) ^ 0.5
  n = (m1 * m1 - m2 * m2) / (q2 - q1)
  C = m1 * m1 + n * q1
  rho0 = a * (C - n * q0) ^ 0.5 / n

  If icall = 0 Then 'init and read in Kriged climate factors
    sql = "SELECT * FROM ClimateFactor ORDER BY x, y;"
    Set cfRec = myDB.OpenRecordset("ClimateFactor", dbOpenDynaset)
    cfRec.MoveLast
    cfRec.MoveFirst
  
    For i = 1 To 30
      For j = 1 To 26
        C2a(i, j) = 0#
        C25a(i, j) = 0#
        C100a(i, j) = 0#
      Next j
    Next i
    With cfRec
      For i = 1 To .RecordCount
        ix = cfRec("x")
        iy = cfRec("y")
        C2a(ix, iy) = cfRec("C2")
        C25a(ix, iy) = cfRec("C25")
        C100a(ix, iy) = cfRec("C100")
        .MoveNext
      Next i
    End With
  End If
  
  dp = CDbl(StaLat(1))
  mp = CDbl(StaLat(2))
  sp = CDbl(StaLat(3))
  dm = CDbl(StaLong(1))
  mm = CDbl(StaLong(2))
  sm = CDbl(StaLong(3))
  mp = mp + sp / 60#
  phi = (dp + mp / 60#) * tpi / 360#
  mm = mm + sm / 60#
  lam = (dm + mm / 60#)
  q = (one - e2) * (Sin(phi) / (one - e2 * Sin(phi) * Sin(phi)) - (one / (two * e)) _
     * Log((one - e * Sin(phi)) / (one + e * Sin(phi))))
  theta1 = n * (lam0 - lam) * tpi / 360#
  RHO = a * (C - n * q) ^ 0.5 / n
  X = RHO * Sin(theta1) / 1000# + xoff
  Y = (rho0 - RHO * Cos(theta1)) / 1000# + yoff
  sx = CSng(X)
  sy = CSng(Y)
  
  wgt sx, sy, C2a(), C25a(), C100a()
      
End Sub

Private Sub wgt(X As Single, Y As Single, C2a() As Single, C25a() As Single, C100a() As Single)
  Dim cp As Single, r1 As Single, r2 As Single, r3 As Single, _
      r4 As Single, rx As Single, ry As Single, wr1 As Single, wr2 As Single, _
      wr3 As Single, wr4 As Single, xr As Single, yr As Single
  Dim ix As Integer, ixp As Integer, iy As Integer, iyp As Integer
      
  wr1 = 0#
  wr2 = 0#
  wr3 = 0#
  wr4 = 0#
'  ESTIMATE CLIMATE FACTOR FROM INTERPOLATION OF KRIGED VALUES AT 4 POINTS
'  SURROUNDING X,Y LOCATION OF INTEREST.  HOWEVER, USE KRIGED VALUE IF
'  DISTANCE TO NEAREST POINT IS 10KM OR LESS.
  ix = Fix(X / 100#)
  rx = ix
  xr = X - 100# * rx
  iy = Fix(Y / 100#)
  ry = iy
  yr = Y - 100# * ry
  ix = ix + 1
  iy = iy + 1
  r1 = (xr ^ 2 + yr ^ 2) ^ 0.5
  If r1 <= 10# Then
    wr1 = 1#
    GoTo 100
  End If
  r2 = (xr ^ 2 + (100# - yr) ^ 2) ^ 0.5
  If r2 <= 10# Then
    wr2 = 1#
    GoTo 100
  End If
  r3 = ((100# - xr) ^ 2 + (100# - yr) ^ 2) ^ 0.5
  If r3 <= 10# Then
    wr3 = 1#
    GoTo 100
  End If
  r4 = ((100# - xr) ^ 2 + yr ^ 2) ^ 0.5
  If r4 <= 10# Then
    wr4 = 1#
    GoTo 100
  End If
  cp = 1# / (1# / r1 + 1# / r2 + 1# / r3 + 1# / r4)
  wr1 = cp / r1
  wr2 = cp / r2
  wr3 = cp / r3
  wr4 = cp / r4
100:
  ixp = ix + 1
  iyp = iy + 1
  Cf2 = wr1 * C2a(ix, iy) + wr2 * C2a(ix, iyp) + _
       wr3 * C2a(ixp, iyp) + wr4 * C2a(ixp, iy)
  Cf25 = wr1 * C25a(ix, iy) + wr2 * C25a(ix, iyp) + _
        wr3 * C25a(ixp, iyp) + wr4 * C25a(ixp, iy)
  Cf100 = wr1 * C100a(ix, iy) + wr2 * C100a(ix, iyp) + _
         wr3 * C100a(ixp, iyp) + wr4 * C100a(ixp, iy)
End Sub

Sub INDEXX(n As Long, ARRIN() As Single, INDX() As Integer)
  Dim q As Double
  Dim i&, indxt&, ir&, j&, l&
  
' Subroutine INDEXX indexs an array ARRIN of length N, outputs the
'  array INDX such that ARRIN(INDX(J)) is in ascending order for
'  J=1,2,..,N. The input quantities ARRIN and N are not changed
'      (ref. Numerical Recipes, p. 233)
'

  For j = 1 To n
    INDX(j) = j
  Next j
  l = n / 2 + 1
  ir = n
20:
  If (l > 1) Then
    l = l - 1
    indxt = INDX(l)
    q = ARRIN(indxt)
  Else
    indxt = INDX(ir)
    q = ARRIN(indxt)
    INDX(ir) = INDX(1)
    ir = ir - 1
    If (ir = 1) Then
      INDX(1) = indxt
      Exit Sub
    End If
  End If
  i = l
  j = l + l
30:
  If (j <= ir) Then
    If (j < ir) Then
      If (ARRIN(INDX(j)) < ARRIN(INDX(j + 1))) Then j = j + 1
    End If
    If (q < ARRIN(INDX(j))) Then
      INDX(i) = INDX(j)
      i = j
      j = j + j
    Else
      j = ir + 1
    End If
    GoTo 30
  End If
  INDX(i) = indxt
  GoTo 20
End Sub

Public Function gauscf(xx!) As Single
  'cumulative probability function
  Dim ax!, t!, d!
  Const xlim! = 18.3

  ax = Abs(xx)
  gauscf = 1#
  If ax <= xlim Then
    t = 1# / (1# + 0.2316419 * ax)
    d = 0.3989423 * Exp(-xx * xx * 0.5)
    gauscf = 1# - _
             d * t * ((((1.330274 * t - 1.821256) * t + 1.781478) * t - 0.3565638) _
             * t + 0.3193815)
  End If
  If xx < 0 Then gauscf = 1# - gauscf
End Function

Public Function gausdy(xx!) As Single
  'cumulative probability function
  Const xlim! = 18.3

  gausdy = 0#
  If (Abs(xx) <= xlim) Then
    gausdy = 0.3989423 * Exp(-0.5 * xx * xx)
  End If
End Function

Private Sub EmboldenMe(o As Object, index As Integer)
  Dim objF As Font, i&
  
  For i = 0 To o.Count - 1
    Set objF = o(i).Font
    If i = index Then
      objF.Bold = True 'Embolden new selection
    Else
      objF.Bold = False 'disEmbolden new selection
    End If
  Next i
End Sub

Sub OutPut(IOUT As Integer, IPK As Integer, pru As Single, outfile As Long)
' subroutine outputs results to screen and file
  Dim eqyrs As Single, cl90 As Single, cookd As Single, cu90 As Single, _
      delres As Single, errmod As Single, hatdig As Single, hmax As Single, _
      pred As Single, press As Single, prx As Single, _
      pv4 As Single, resid As Single, samerr As Single, sdbeta As Single, _
      sepc As Single, sepu As Single, smod As Single, arhoc As Single
  Dim ssam As Single, stdres As Single, STUTP As Single, Sum As Single, _
      tbeta As Single, test1 As Single, test2 As Single, tstat As Single, _
      tv4 As Single, varres As Single, vpu As Single, asep As Single, splus As Single, sminu As Single
  Dim work1(50, 1) As Single, work3(10, 50) As Single, work2(10, 1) As Single, _
      xo(1, 10) As Single, xot(10, 1) As Single, ccc(1, 1) As Single, _
      hat(50, 50) As Single, hats(50, 50) As Single
  Dim i As Integer, iu As Integer, j As Integer, l As Integer, _
      ndf As Integer

  pru = 0#
  For iu = 1 To Ne
    Select Case iu
      Case 1: pru = pru + Bhat(iu, 1) * 1#
      Case 2: pru = pru + Bhat(iu, 1) * UserDA
      Case 3: pru = pru + Bhat(iu, 1) * UserVars(17)
      Case 4: pru = pru + Bhat(iu, 1) * UserVars(18)
    End Select
  Next iu
  If (IOUT > 0) Then
'    WRITE (16,9005) Siteid, Pklab(IPK)
  
  ' WRITE BETAS
'    IF (IOUT<99) then WRITE (16,9010) IOUT
'    WRITE (16,9015)
    sdbeta = XtXinv(1, 1) ^ 0.5
    tbeta = Bhat(1, 1) / sdbeta
'    WRITE (16,9020) Bhat(1,1), sdbeta, tbeta
    For i = 2 To Ne
      sdbeta = XtXinv(i, i) ^ 0.5
      tbeta = Bhat(i, 1) / sdbeta
      ndf = Nsites - Ne
      tv4 = Abs(tbeta)
'      pv4 = 2# * STUTP(-tv4, ndf)
      If (pv4 < 0.0001) Then pv4 = 0.0001
'      WRITE (16,9025) Ylab(i-1), Bhat(i,1), sdbeta, tbeta, pv4
    Next i
    Call mltply(work1, X, Bhat, Nsites, Ne, 1, 50, 50, 10)
  
  ' WRITE PREDICTED VALUES ETC.
    If (IOUT = 99) Then
      smod = 0#
      ssam = 0#
      press = 0#
      hmax = 0#
      Call mltply(work3, XtXinv, Xt, Ne, Ne, Nsites, 10, 10, 10)
      Call mltply(hats, X, work3, Nsites, Ne, Nsites, 50, 50, 10)
      Call mltply(hat, hats, Wt, Nsites, Nsites, Nsites, 50, 50, 50)
      For i = 1 To Nsites
        pred = work1(i, 1)
        resid = Y(i, 1) - work1(i, 1)
        samerr = hats(i, i)
        If (samerr > hmax) Then hmax = samerr
        hatdig = hat(i, i)
        delres = resid / (1# - hatdig)
        errmod = Gamasq
        varres = Gamasq + Sta(i) - samerr
        stdres = resid / varres ^ 0.5
        cookd = stdres ^ 2 * samerr / (Ne * (Gamasq + Sta(i) - samerr))
        test1 = 2# * Ne / Nsites
        test2 = 4# / Nsites
        ssam = ssam + samerr
        press = press + delres ^ 2
      Next i
  
      'WRITE AVG SAMPLING ERROR AND MODEL ERROR
      ssam = ssam / Nsites
      smod = smod / Nsites
      press = press / Nsites
      Atse = Atse / Nsites
      Atscov = Atscov / (Nsites / 2# * (Nsites - 1))
      arhoc = Atscov / Atse
  
      'Write out prediction for ungaged site
      For l = 1 To Ne
        xo(1, l) = 1#
        xot(l, 1) = 1#
      Next l
  
      Call mltply(work2, XtXinv, xot, Ne, Ne, 1, 10, 10, 10)
      Call mltply(ccc, xo, work2, 1, Ne, 1, 1, 1, 10)
      sepu = ccc(1, 1)
      vpu = Gamasq + sepu
      eqyrs = Sig ^ 2 * (1# + Ak(IPK) ^ 2 / 2#) / vpu
  
      'make adjustments
      prx = 10 ^ pru
  
      'convert to cfs and percent error
      sepc = 100# * (Exp(vpu * 5.302) - 1#) ^ 0.5
      asep = vpu ^ 0.5
      splus = 100# * (10 ^ (asep) - 1#)
      sminu = 100# * (10 ^ (-asep) - 1#)
      tstat = vpu ^ 0.5 * 1.65
      cu90 = 10 ^ (tstat + pru)
      cl90 = 10 ^ (pru - tstat)
      Call Round(prx)
      Call Round(cl90)
      Call Round(cu90)
      
    Print #outfile, Pklab(IPK) & vbTab & prx & vbTab & sminu & vbTab & splus & vbTab & _
        cl90 & " - " & cu90

      If (sepu > hmax) Then _
          Print #outfile, "WARNING: Prediction is outside range of observed data" & vbCrLf
      If (sepu > hmax) Then
        MsgBox "WARNING: Prediction is an extrapolation beyond'" & vbCrLf & _
            "observed data. Check for errors in input basin" & vbCrLf & _
            "characteristics. If no errors use results with caution."
      End If
    End If
  End If
End Sub

Private Sub txtState_Change()
  Dim stAbb$, sql$
  Dim i&, j&, k&
  Dim dataRec As Recordset, myRec As Recordset
  Dim stRec As Recordset, fldRec As Recordset, parmRec As Recordset
  'Dim myScenario as nssScenario
  
  If Len(txtState.Text) <> 2 Then Exit Sub
  stAbb = UCase(txtState.Text)
  Set stRec = myDB.OpenRecordset("States", dbOpenSnapshot)
  stRec.FindFirst "St='" & stAbb & "'"
  If stRec.NoMatch Then
    MsgBox "There is no state with the abbreviation '" & stAbb & "'." & _
        vbCrLf & "Enter the 2-letter symbol for the state of your choice."
    txtState.Text = ""
  Else
    StateID = stRec("ID")
    'open table with data records for specified state
    sql = "SELECT * From ROIStations WHERE StID=" & StateID
    Set dataRec = myDB.OpenRecordset(sql, dbOpenSnapshot)
    dataRec.MoveLast
    dataRec.MoveFirst
    StaCnt = dataRec.RecordCount
    'Check to make sure required selections have been made
    If StaCnt < 30 Then
      MsgBox "Not enough stations in this region to perform calculation"
      Exit Sub
    End If
    Set parmRec = myDB.OpenRecordset("ROIParmList", dbOpenSnapshot)
    parmRec.MoveLast
    parmRec.MoveFirst
    ParmCnt = parmRec.RecordCount
    ReDim UseFld(1, ParmCnt)
    '0th dimension for whether parm used. 1st dimension if used in distance calc.
    For i = 1 To dataRec.Fields.Count - 1
      If Not IsNull(dataRec(i)) Then UseFld(0, i) = True
    Next i
    
    sql = "SELECT * From ROIParmsByState WHERE StateID=" & StateID & _
          " ORDER BY ParmID;"
    Set myRec = myDB.OpenRecordset(sql, dbOpenSnapshot)
    With myRec
      While Not .EOF
        For i = 0 To ParmCnt
          If !ParmID = i Then UseFld(1, i) = True
        Next i
        .MoveNext
      Wend
    End With
    'redimension arrays to number of stations in state
    ReDim UserVars(ParmCnt)
    ReDim Vars(StaCnt, ParmCnt)
    ReDim VarSDs(ParmCnt)
    ReDim Sum(ParmCnt)
    ReDim Mcon(StaCnt, StaCnt)
    ReDim Rhoc(StaCnt, StaCnt)
    ReDim INDX(StaCnt)
    ReDim Distance(StaCnt)
    'open list of ROI parms and set parmCnt = number of parms
    'open state matrix table
    sql = "SELECT * FROM StateMatrices" & _
          " WHERE StID=" & StateID & _
          " ORDER BY RowID, ColID;"
    Set myRec = myDB.OpenRecordset(sql, dbOpenForwardOnly)
    sql = "SELECT * FROM RHO " & _
          "ORDER BY RowID, ColID;"
    Set RhoRec = myDB.OpenRecordset(sql, dbOpenForwardOnly)
  
  'Read in estimation data
    i = 1
    'The following loop will be set vals to classes, not arrays, in real program.
    With dataRec
      While Not .EOF
        For j = 1 To ParmCnt
          If UseFld(0, j) Then
            If j = 13 Or j = 14 Then
              Vars(i, j) = Log10(dataRec(j))
            Else
              Vars(i, j) = dataRec(j)
            End If
          ElseIf j = 15 Then  'StdDev not in DB -> calc it
            Vars(i, j) = (Vars(i, 10) - Vars(i, 5)) / Ak(6)
          End If
          If j = 13 Or j = 14 Or j = 17 Or j = 18 Then
            Sum(j) = Sum(j) + Vars(i, j)
          End If
        Next j
        dataRec.MoveNext
        'assign values from StateMatrices table to mcon()
        With myRec
          For j = 1 To i
            Mcon(i, j) = !Value
            .MoveNext
          Next j
        End With
        If StateID = 27 Then
          'assign values from RHO table to rhoc()
          With RhoRec
            For j = 1 To i
              Rhoc(i, j) = !Value
              .MoveNext
            Next j
          End With
        End If
        'transpose mcon and rhoc matrices
        For j = 1 To i
          Mcon(j, i) = Mcon(i, j)
          If StateID = 27 Then Rhoc(j, i) = Rhoc(i, j)
        Next j
        i = i + 1
      Wend
    End With
    sql = "SELECT * FROM Regions " & _
          "WHERE StateID=" & StateID & " AND ROI=True " & _
          "ORDER BY ID ASC;"
    Set fldRec = myDB.OpenRecordset(sql, dbOpenSnapshot)
    'Set FldRec = myscenario.project.DB.OpenRecordset(SQL, dbOpenSnapshot)
    fldRec.MoveLast
    fldRec.MoveFirst
    i = 0
    'load control arrays
    fraRegions.Visible = True
    txtUserParm(0).Visible = True
    lblUserParm(0).Visible = True
    While Not fldRec.EOF
      rdoRegion(i).Caption = fldRec("Name")
      If i < fldRec.RecordCount - 1 Then
        Load rdoRegion(i + 1)
        rdoRegion(i + 1).Top = 240 + (i + 1) * 360
        rdoRegion(i + 1).Visible = True
        fraRegions.Height = 612 + (i + 1) * 360
      End If
      i = i + 1
      fldRec.MoveNext
    Wend
    fldRec.Close
    sql = "SELECT * FROM ROIParmsByState " & _
          "WHERE StateID=" & StateID & " And ParmID>12 " & _
          "ORDER BY ParmID ASC;"
    Set fldRec = myDB.OpenRecordset(sql, dbOpenSnapshot)
    'Set FldRec = myscenario.project.DB.OpenRecordset(SQL, dbOpenSnapshot)
    fldRec.MoveLast
    fldRec.MoveFirst
    i = 0
    While Not fldRec.EOF
      parmRec.FindFirst "ID=" & fldRec("ParmID")
      lblUserParm(i).Caption = parmRec("Abbrev")
      If i < fldRec.RecordCount - 1 Then
        Load txtUserParm(i + 1)
        txtUserParm(i + 1).Top = 240 + (i + 1) * 360
        txtUserParm(i + 1).Visible = True
        Load lblUserParm(i + 1)
        lblUserParm(i + 1).Top = 240 + (i + 1) * 360
        lblUserParm(i + 1).Visible = True
      End If
      i = i + 1
      fldRec.MoveNext
    Wend
    fldRec.Close
    parmRec.Close
    myRec.Close
    End If
End Sub

Private Sub Init()
  Dim sql$
  
  Pklab(1) = "2-year"
  Pklab(2) = "5-year"
  Pklab(3) = "10-year"
  Pklab(4) = "25-year"
  Pklab(5) = "50-year"
  Pklab(6) = "100-year"
  Pklab(7) = "200-year"
  Pklab(8) = "500-year"

  Ak(1) = 0#
  Ak(2) = 0.84162
  Ak(3) = 1.28155
  Ak(4) = 1.75069
  Ak(5) = 2.05375
  Ak(6) = 2.32635
  Ak(7) = 2.575
  Ak(8) = 2.87816
  Alpha = 0.0025
  Theta = 0.983
  Select Case StateID
    Case 27
        Ne = 2
    Case 46
        Ne = 4
  End Select

End Sub

Private Sub rdoRegion_Click(index As Integer)
  EmboldenMe rdoRegion, index
  RegName = rdoRegion(index).Caption
  RegID = index
End Sub

Private Function TASKER_DISTANCE(latxin As Single, latyin As Single, _
            longxin As Single, longyin As Single)
' The latitudes and longitudes are inputted in decimal degrees
' Distance is returned in miles
' The projection is Albers Equal-Area Conic Projection (with Earth as an
' elipsoid) Ref. Snyder, J.P., 1982, "Map Projections Used by The U.S.
' Geological Survey", U. S. Geological Survey Bulletin 1532, p. 96-99.
  Dim radian As Single, a As Single, e2 As Single, latx As Single, _
      laty As Single, longx As Single, longy As Single, _
      latdiff As Single, longdiff As Single, latave As Single, _
      op1 As Single, op2 As Single, op3 As Single

  a = 6378206.4
  e2 = 0.00676866
  radian = 3.14159265 / 180
  latx = latxin * radian
  laty = latyin * radian
  longx = longxin * radian
  longy = longyin * radian

  latdiff = Abs(latx - laty)
  longdiff = Abs(longx - longy)
  latave = (latx + laty) / 2#
  op1 = a / (1# - e2 * Sin(latave) ^ 2) ^ 0.5
  op2 = ((1# - e2) ^ 2 * (latdiff) ^ 2)
  op2 = op2 / ((1# - e2 * Sin(latave) ^ 2) ^ 2)
  op3 = Cos(latave) ^ 2 * longdiff ^ 2
' Convert distance to miles by dividing the meters by 1000 to get
' kilometeres and then convert to miles by multiplying by 0.621
  TASKER_DISTANCE = 0.62137119 * (op1 * (op2 + op3) ^ 0.5) / 1000#
End Function

Private Sub BuildMatrix()
'-- Build the cross-correlation matrix on the number of stations selected
'--  This subroutine uses the arrays lat and long as basis
  Dim dist As Single, latx As Single, laty As Single, longx As Single, longy As Single
  Dim i&, j&

  If (Theta < 0 Or Theta <= 1#) Then
     MsgBox "ERROR-- Theta must be between 0 and 1"
     Exit Sub
  End If

  For i = 1 To StaCnt
    For j = 1 To StaCnt
      latx = Lat(i)
      laty = Lat(j)
      longx = Lng(i)
      longy = Lng(j)
      Rhoc(i, j) = TASKER_DISTANCE(latx, laty, longx, longy)
    Next j
  Next i

  For i = 1 To StaCnt
    For j = 1 To i
      If (i = j) Then
         Rhoc(i, j) = 1#
      Else
         Rhoc(i, j) = Theta ^ (Rhoc(i, j) / (Alpha * Rhoc(i, j) + 1#))
         If (Rhoc(i, j) < 0#) Then
           MsgBox "rhoc(i,j) = " & i & " " & j & " " & Rhoc(i, j) & vbCrLf & _
                  "ERROR--Cholesky Decomp requires >= 0"
           Exit Sub
         End If
      End If
    Next j
  Next i
End Sub
