VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDBMerge 
   Caption         =   "StreamStats Database Merge Utility"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMerge 
      Caption         =   "Merge"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.ComboBox cboState 
      Height          =   315
      Left            =   2280
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cdlDB 
      Left            =   3840
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDB 
      Caption         =   "Select Additional Database"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton cmdDB 
      Caption         =   "Select Master Database"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblState 
      Caption         =   "Select State To Merge"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label lblDBName 
      Caption         =   "Label1"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label lblDBName 
      Caption         =   "Label1"
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "frmDBMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Master As New nssDatabase
Private Adding As New nssDatabase
Private Renumb As New FastCollection 'of new stat label ids, with old id as key

Private Sub cboState_Click()
  Dim STKey As String
  If Len(Master.filename) > 0 And Len(Adding.filename) > 0 Then
    STKey = CStr(cboState.ItemData(cboState.ListIndex))
    If Len(STKey) = 1 Then STKey = "0" & STKey
    If Master.States(STKey).Name = Adding.States(STKey).Name Then
      cmdMerge.Enabled = True
      Set Master.State = Master.States(STKey)
      Set Adding.State = Adding.States(STKey)
    Else
      MsgBox "Problem!  State indices on databases don't match." & vbCrLf & vbCrLf & _
             "Can't merge databases.", vbExclamation, App.Title
    End If
  Else
    MsgBox "Please specify both databases before selecing state to merge.", vbInformation, App.Title
  End If

End Sub

Private Sub cmdDB_Click(Index As Integer)

  Dim i As Long
  Dim DBName As String

  If Index = 0 Then
    cdlDB.DialogTitle = "Open Master StreamStats Database"
  Else
    cdlDB.DialogTitle = "Open Additional StreamStats Database"
  End If
  cdlDB.ShowOpen
  DBName = cdlDB.filename
  If Index = 0 Then
    With Master
      .filename = DBName
      For i = 1 To .States.Count
        cboState.AddItem .States(i).Abbrev
        cboState.ItemData(i - 1) = CStr(.States(i).Code)
      Next i
    End With
  Else
    Set Adding = New nssDatabase
    Adding.filename = DBName
  End If
  lblDBName(Index).Caption = DBName
End Sub

Private Sub cmdMerge_Click()
  
  Dim i As Long

'  'first check for any new units
'  For i = 1 To Adding.Units.Count
'    If Not Master.Units.KeyExists(CStr(i)) Then 'these units not on master, add 'em
'    End If
'  Next i
  MousePointer = vbHourglass
  Set Renumb = Nothing
  MergeStatLabels
  MergeStations
  MousePointer = vbDefault
  MsgBox "Merging complete"
End Sub

Private Sub MergeStatLabels()

  Dim i As Long, fld As Long
  Dim myStatType As ssStatType
  Dim myStat As ssStatLabel
  Dim lSource As ssSource
  Dim mySource As ssSource
  Dim vStatLabel As Variant
  Dim myRec As Recordset
  Dim sql As String

  For i = 1 To Adding.StatisticTypes.Count
    If Not Master.StatisticTypes.KeyExists(Adding.StatisticTypes(i).Code) Then
      'add this stat type
      Set myStatType = New ssStatType
      myStatType.DB = Master
      myStatType.Add Adding.StatisticTypes(i).Code, Adding.StatisticTypes(i).Name
    End If
    'look through this types stat labels to see if any need to be added
    For Each vStatLabel In Adding.StatisticTypes(i).StatLabels
      If Not Master.StatisticTypes(Adding.StatisticTypes(i).Code).StatLabels.KeyExists(vStatLabel.Code) Then
        'need to add new stat
        sql = "SELECT STATLABEL.* FROM STATLABEL ORDER BY StatisticLabelCode;"
        Set myRec = Master.DB.OpenRecordset(sql, dbOpenDynaset)
        With myRec
          .MoveLast
          fld = !StatisticLabelCode + 1
          .AddNew
          .Fields("StatisticLabelCode") = fld
          If vStatLabel.id <> fld Then 'need to renumber this stat label
            Renumb.Add fld, vStatLabel.id 'store original stat label code as key
          End If
          .Fields("StatisticTypeCode") = vStatLabel.TypeCode
          .Fields("StatLabel") = vStatLabel.Code
          .Fields("StatisticLabel") = vStatLabel.Name
          .Fields("Units") = vStatLabel.Units
          'definition not available as property, use name without the underscores
          'this is how it currently is done in StreamStats
          .Fields("Definition") = ReplaceString(vStatLabel.Name, "_", " ")
          .Update
        End With
      End If
    Next
  Next i
  'now look through data sources being added
  For Each lSource In Adding.Sources
    'add all sources from adding database (existing ones won't be added)
    Set mySource = New ssSource
    Set mySource.DB = Master
    mySource.Add lSource.Name
  Next

End Sub

Private Sub MergeStations()

  Dim vStation As Variant
  Dim Stn As New ssStation
  Dim StnType As New ssStationType
  Dim vals(2, 1, 13) As String
  Dim sql As String, StnID As String
  Dim myRec As Recordset
  Dim myMsgBox As New ATCoMessage
  Dim rsp As Long, Ind As Long, ImportFg As Long
  Dim AddIt As Boolean

  rsp = 0
  For Each vStation In Adding.State.Stations
    AddIt = False 'assume not adding station
    Set Stn = vStation
    If Len(Stn.id) = 7 Then 'add preceeding 0 to make consistent with 8 digit ids
      StnID = "0" & Stn.id
    Else
      StnID = Stn.id
    End If
    'master DB should have fully defined station IDs
    sql = "SELECT * FROM [Station State] WHERE StaID='" & StnID & "'"
    Set myRec = Master.DB.OpenRecordset(sql, dbOpenDynaset)
    With myRec
      If .RecordCount > 0 Then 'station exists
        .FindFirst "StateCode='" & Master.State.Code & "'"
        If .NoMatch Then 'add to station state table
          .AddNew
          .Fields("StaID") = StnID
          .Fields("StateCode") = Master.State.Code
          .Fields("ROI") = Stn.ROIIndex
          .Update
        End If
        'see what user wants to do
        If rsp = 0 Then
          rsp = myMsgBox.Show("Station " & StnID & " already exists on the master database." & vbCrLf & _
                              "Do you want to Replace it or Keep the master version?", _
                              "Station Exists", "&Replace", "Replace &All", "+&Keep", "K&eep All", "-&Cancel")
          If rsp = 5 Then Exit Sub 'user wants to cancel merge
          MousePointer = vbHourglass
        End If
        If rsp < 3 Then AddIt = True 'new station or replace existing
        If rsp = 1 Or rsp = 3 Then rsp = 0 'reset to ask again next time
      Else 'new station, add it
        AddIt = True
      End If
      If AddIt Then 'station not on master, just add it
        vals(2, 1, 1) = StnID
        vals(2, 1, 2) = Stn.Name
        Set StnType = Stn.StationType
        vals(2, 1, 3) = StnType.Name
        vals(2, 1, 4) = Stn.IsRegulated
        vals(2, 1, 5) = Stn.Period
        vals(2, 1, 6) = Stn.Remarks
        vals(2, 1, 7) = Stn.Latitude
        vals(2, 1, 8) = Stn.Longitude
        vals(2, 1, 9) = Stn.HUCCode
        vals(2, 1, 10) = Stn.StatebasinCode
        vals(2, 1, 11) = Stn.CountyCode
        vals(2, 1, 12) = Stn.MCDCode
        vals(2, 1, 13) = Stn.Directions
        If Stn.Statistics.Count > 0 Then 'has data
          ImportFg = 1
        Else
          ImportFg = 0
        End If
        Stn.id = StnID 'update to 8-digit name for adding/editing stats
        Set Stn.DB = Master 'change to update master
        If .RecordCount > 0 Then 'edit station
          Stn.Edit vals, 1
        Else 'add station
          If Stn.IsROI Then
            Ind = -Stn.ROIIndex
          Else
            Ind = 1
          End If
          Stn.Add vals, Ind, ImportFg
        End If
        Stn.id = StnID 'update to 8-digit name for adding/editing stats
        AddStats Stn
      End If
    End With
  Next
End Sub

Private Sub AddStats(Stn As ssStation)

  Dim atts() As String
  Dim sql As String
  Dim myRec As Recordset
  Dim AddFg As Boolean
  Dim vStat As Variant
  Dim myStat As New ssStatistic

  ReDim atts(2, 1, 7)
'  Set Stn.DB = Adding 'set back to database containing stats to be added
  sql = "SELECT * FROM Statistic WHERE StaID='" & Stn.id & "'"
  Set myRec = Master.DB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    For Each vStat In Stn.Statistics
      Set myStat = vStat
      AddFg = False 'assume editing, not adding
      If .RecordCount > 0 Then 'statistics exist for this station
        .FindFirst "StatisticLabelCode=" & myStat.Code
        If .NoMatch Then AddFg = True 'this stat not on master, add to station's statistics
      Else 'no exising stats, add all stats for this station
        AddFg = True
      End If
      If Renumb.KeyExists(myStat.Code) Then 'need to renumber this stat's code
        atts(2, 1, 2) = Renumb(CStr(myStat.Code))
      Else
        atts(2, 1, 2) = myStat.Code
      End If
      atts(2, 1, 4) = myStat.Value
      atts(2, 1, 6) = myStat.RecDate
      atts(2, 1, 7) = myStat.Source
      Set myStat.DB = Master
      If AddFg Then
        myStat.Add atts, 1
      Else
        myStat.Edit atts, 1
      End If
    Next
  End With
  
End Sub
