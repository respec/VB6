VERSION 5.00
Begin VB.Form frmCreate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinHSPF - Create Project"
   ClientHeight    =   4260
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   26
   Icon            =   "frmCreate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMet 
      Caption         =   "Initial Met Station"
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   4335
      Begin VB.ListBox lstMet 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   1008
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdOkayCancel 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   14
      Top             =   3600
      Width           =   852
   End
   Begin VB.CommandButton cmdOkayCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   13
      Top             =   3600
      Width           =   852
   End
   Begin VB.Frame fraScheme 
      Caption         =   "Model Segmentation"
      Height          =   1695
      Left            =   4560
      TabIndex        =   10
      Top             =   1680
      Width           =   3375
      Begin VB.OptionButton opnScheme 
         Caption         =   "Grouped"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   12
         ToolTipText     =   "Each Perlnd/Implnd connects to multiple Rchres"
         Top             =   600
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton opnScheme 
         Caption         =   "Individual"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   11
         ToolTipText     =   "Each Perlnd/Implnd connects to only one Rchres"
         Top             =   960
         Width           =   2775
      End
   End
   Begin VB.Frame fraFiles 
      Caption         =   "Files"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.ListBox lstWDM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   216
         Left            =   3480
         TabIndex        =   4
         Top             =   600
         Width           =   4212
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "Select"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "Select"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   732
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "Select"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   732
      End
      Begin VB.Label lblName 
         Caption         =   "Met WDM Files"
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   9
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblName 
         Caption         =   "Project WDM File"
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   8
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblName 
         Caption         =   "BASINS Watershed File"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblFile 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<none>"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label lblFile 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<none>"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   5
         Top             =   960
         Width           =   4215
      End
   End
   Begin MSComDlg.CommonDialog CDFile 
      Left            =   7440
      Top             =   3480
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      FontSize        =   4.09255e-38
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Copyright 2002 AQUA TERRA Consultants - Royalty-free use permitted under open source license

Dim WDMId$(4), MetDetails$()

Private Sub cmdFile_Click(Index As Integer)
  Dim iwdm&, i&, s$, f$, fun&, numMetSeg&, arrayMetSegs$(), wid$
  Dim tmetseg, lMetDetails$(), lMetDescs$()
  
  If Index = 0 And lblFile(0).Caption <> "<none>" Then
    Call MsgBox("Only one project WDM file may be included.", _
                vbOKOnly, "Create Project Problem")
  ElseIf Index = 1 And lstWDM.listcount > 2 Then
    Call MsgBox("No more than three Met WDM files may be included in a project.", _
                vbOKOnly, "Create Project Problem")
  Else
    If Index = 0 Then
      'project wdm file
      If FileExists(BASINSPath & "\modelout", True, False) Then
        ChDriveDir BASINSPath & "\modelout"
      End If
      CDFile.flags = &H8806&
    ElseIf Index = 1 Then
      'met wdm file
      If FileExists(BASINSPath & "\data\met_data", True, False) Then
        ChDriveDir BASINSPath & "\data\met_data"
      End If
      CDFile.flags = &H1806&
    End If
    If Index = 0 Or Index = 1 Then
      CDFile.Filter = "WDM files (*.wdm)|*.wdm"
      CDFile.Filename = "*.wdm"
      If Index = 0 Then
        CDFile.DialogTitle = "Select Project WDM File"
      Else
        CDFile.DialogTitle = "Select Met WDM File"
      End If
      On Error GoTo 40
      CDFile.CancelError = True
      CDFile.Action = 1
      If Index = 1 Then
        'met wdms
        f = CDFile.Filename
        If InList(f, lstWDM) Then
          GoTo 25
        Else
          If lstWDM.List(0) = "<none>" Then lstWDM.RemoveItem 0
          lstWDM.AddItem f
          iwdm = 0
          wid = "WDM" & lstWDM.listcount + 1
          myUci.openWDM iwdm, f, fun, wid
          WDMId(lstWDM.listcount) = wid
          If fun < 1 Then
            Call MsgBox("Problem opening the Met WDM file.", _
                        vbOKOnly, "Create Project Problem")
          Else
            myUci.GetMetSegNames fun, numMetSeg, arrayMetSegs, lMetDetails, lMetDescs
            If numMetSeg > 0 Then
              'add the met segs from this wdm to the list
              If lstWDM.listcount = 1 Then
                tmetseg = 0
              Else
                tmetseg = UBound(MetDetails)
              End If
              ReDim Preserve MetDetails(tmetseg + numMetSeg)
              For i = 0 To numMetSeg - 1
                lstMet.AddItem arrayMetSegs(i) & ":" & lMetDescs(i)
                MetDetails(tmetseg + i) = lMetDetails(i) & "," & wid
              Next i
              lstMet.Enabled = True
              lstMet.BackColor = &H80000005
              If lstMet.SelCount = 0 Then
                lstMet.Selected(0) = True
              End If
            'Else
            '  Call MsgBox("This WDM file has no location attributes.  These attributes " & vbCrLf & _
            '    "would have been updated had the associated .inf file been present.", vbOKOnly, "Met WDM Problem")
            End If
          End If
        End If
25      ' continue here on cancel
      Else
        'project wdm
        lblFile(Index).Caption = CDFile.Filename
        'does wdm exist?
        On Error GoTo 20
        Open lblFile(0).Caption For Input As #1
        'yes, it exists
        Close #1
        iwdm = 0
        GoTo 30
20      'no, it does not exist, create it
        iwdm = 2
30      'open wdm file
        wid = "WDM1"
        myUci.openWDM iwdm, lblFile(0).Caption, fun, wid
        If iwdm = 2 And fun < 1 Then
          Call MsgBox("Problem creating the project WDM file.", _
                        vbOKOnly, "Create Project Problem")
          lblFile(Index).Caption = "<none>"
        ElseIf iwdm = 0 And fun < 1 Then
          Call MsgBox("Problem opening the project WDM file.", _
                        vbOKOnly, "Create Project Problem")
          lblFile(Index).Caption = "<none>"
        Else
          WDMId(0) = wid
        End If
      End If
40        'continue here on cancel
    ElseIf Index = 2 Then
      If FileExists(BASINSPath & "\modelout", True, False) Then
        ChDriveDir BASINSPath & "\modelout"
      End If
      CDFile.flags = &H8806&
      CDFile.Filter = "BASINS Watershed Files (*.wsd)"
      CDFile.Filename = "*.wsd"
      CDFile.DialogTitle = "Select BASINS Watershed File"
      On Error GoTo 50
      CDFile.CancelError = True
      CDFile.Action = 1
      lblFile(Index).Caption = CDFile.Filename
50        'continue here on cancel
    End If
  End If
End Sub

Private Sub cmdOkayCancel_Click(Index As Integer)
    Dim i&, s$, wdmname$(3), outwdm$, tmpuci$, iresp&, lmetdetail$, continuefg As Boolean
    Dim dsn&, lwdmid$, lId&, sjday#, ejday#, sdat&(6), edat&(6)
    Dim tmetseg As HspfMetSeg
    
    If Index = 0 And lblFile(0) = "<none>" Then
      'no project file specified, don't allow to okay
      Call MsgBox("A project WDM file must be specified.", vbOKOnly, "WinHSPF Create Problem")
    Else
      'specifications okay, continue
      If lblFile(0) <> "<none>" Then
        outwdm = lblFile(0)
      End If
      For i = 1 To 3
        wdmname(i) = ""
      Next i
      For i = 1 To lstWDM.listcount
        If lstWDM.List(i - 1) <> "<none>" Then
          wdmname(i) = lstWDM.List(i - 1)
        End If
      Next i
      If Index = 0 Then
        'okay to create new
        If lblFile(2) <> "<none>" Then
          tmpuci = Mid(lblFile(2), 1, Len(lblFile(2)) - 3) & "uci"
          iresp = 6
          If FileExists(tmpuci) Then
            'this file already exists, warn user
            iresp = MsgBox("A uci file by this name already exists." & vbCrLf & vbCrLf & _
                    "Do you want to overwrite it?", vbExclamation + vbYesNo + vbDefaultButton1, "Create Problem")
          End If
          If iresp = 6 Then
            If lstMet.ListIndex < 0 Then
              lmetdetail = ""
            Else
              lmetdetail = MetDetails(lstMet.ListIndex)
            End If
            'check to see if using BASINS wdm
            
            continuefg = True
            If Not IsBASINSMetWDM(lmetdetail) Then
              'do window to specify met data details
              frmAddMet.Init "", 2
              frmAddMet.Show vbModal
              If myUci.MetSegs.Count = 0 Then
                'user clicked cancel
                continuefg = False
              Else
                'specified met segment
                Set tmetseg = myUci.MetSegs(1)
                dsn = tmetseg.MetSegRec(1).Source.volid
                lwdmid = tmetseg.MetSegRec(1).Source.volname
                If Len(lwdmid) > 0 Then
                  lId = CInt(Mid(lwdmid, 4, 1))
                End If
                If lId > 0 And dsn > 0 Then
                  sjday = myUci.GetDataSetFromDsn(lId, dsn).dates.Summary.sjday
                  ejday = myUci.GetDataSetFromDsn(lId, dsn).dates.Summary.ejday
                  Call J2Date(sjday, sdat)
                  Call J2Date(ejday, edat)
                End If
                lmetdetail = CStr(-1 * dsn) & "," & _
                   CStr(sdat(0)) & "," & CStr(sdat(1)) & "," & CStr(sdat(2)) & "," & CStr(sdat(3)) & "," & CStr(sdat(4)) & "," & CStr(sdat(5)) & "," & _
                   CStr(edat(0)) & "," & CStr(edat(1)) & "," & CStr(edat(2)) & "," & CStr(edat(3)) & "," & CStr(edat(4)) & "," & CStr(edat(5)) & "," & _
                   lwdmid
              End If
            Else
              'default first met seg from basins data
              DefaultBASINSMetseg lmetdetail
            End If
            
            If continuefg Then
              Me.MousePointer = vbHourglass
              Call HSPFMain.DoCreate(lblFile(2), outwdm, wdmname, WDMId, lmetdetail, opnScheme(0).Value)
            
              setDefault myUci, defUci
              setDefaultML myUci, defUci
              myUci.save
              Me.MousePointer = vbNormal
              Unload Me
            End If
          End If
        Else
          Call MsgBox("User must specify a BASINS Watershed File.", _
                        vbOKOnly, "Create Project Problem")
        End If
        'add files to files block
      ElseIf Index = 1 Then 'cancel
        Unload Me
      End If
    End If
End Sub

Private Sub Form_Load()
    lstWDM.AddItem "<none>"
    'Set myUci = Nothing
    'Set myUci = New HspfUci
    myUci.HelpFile = App.HelpFile
    myUci.MsgWDMName = HSPFMain.W_HSPFMSGWDM
    myUci.MessageUnit = HSPFMain.MessageUnit
    'myUci.InitWDMArray
    Set myUci.Icon = HSPFMain.Icon
End Sub

Public Sub Init(wsdfile$)
  lblFile(2) = wsdfile
End Sub

Private Function IsBASINSMetWDM(MetDetails$)
  Dim dsn&, i&, loc$, j&, tempsj As Double, tempej As Double
  Dim sen$, con$, SDate&(6), EDate&(6), checkcount&
  Dim lts As Collection 'of atcotimser
  Dim ldate As ATCclsTserDate, sj As Double, ej As Double
  Dim llocts As Collection 'of atcotimser
  Dim delim$, quote$, basedsn&, metwdmid$, lMetDetails$, lunit&, WDMId$

  IsBASINSMetWDM = False
  lMetDetails = MetDetails
  If Len(lMetDetails) > 0 Then
    'get details from the met data details
    delim = ","
    quote = """"
    basedsn = StrSplit(lMetDetails, delim, quote)
    For i = 0 To 5
      SDate(i) = StrSplit(lMetDetails, delim, quote)
    Next i
    For i = 0 To 5
      EDate(i) = StrSplit(lMetDetails, delim, quote)
    Next i
    metwdmid = lMetDetails
    
    'look for matching WDM datasets
    Call myUci.FindTimSer("OBSERVED", "", "", lts)
    lunit = 0
    For i = 1 To lts.Count
      If lts(i).Header.Id = basedsn Then
        myUci.GetWDMIDFromUnit lts(i).File.FileUnit, WDMId
        If WDMId = metwdmid Then
          lunit = lts(i).File.FileUnit
          Exit For
        End If
      End If
    Next i
    checkcount = 0
    For i = 1 To lts.Count
      If lts(i).Header.Id = basedsn And lts(i).File.FileUnit = lunit Then
        If lts(i).attrib("TSTYPE") = "PREC" Then
          checkcount = checkcount + 1
        End If
      ElseIf lts(i).Header.Id = basedsn + 1 And lts(i).File.FileUnit = lunit Then
        If lts(i).attrib("TSTYPE") = "EVAP" Then
          checkcount = checkcount + 1
        End If
      ElseIf lts(i).Header.Id = basedsn + 2 And lts(i).File.FileUnit = lunit Then
        If lts(i).attrib("TSTYPE") = "ATEM" Then
          checkcount = checkcount + 1
        End If
      ElseIf lts(i).Header.Id = basedsn + 3 And lts(i).File.FileUnit = lunit Then
        If lts(i).attrib("TSTYPE") = "WIND" Then
          checkcount = checkcount + 1
        End If
      ElseIf lts(i).Header.Id = basedsn + 4 And lts(i).File.FileUnit = lunit Then
        If lts(i).attrib("TSTYPE") = "SOLR" Then
          checkcount = checkcount + 1
        End If
      ElseIf lts(i).Header.Id = basedsn + 5 And lts(i).File.FileUnit = lunit Then
        If lts(i).attrib("TSTYPE") = "PEVT" Then
          checkcount = checkcount + 1
        End If
      ElseIf lts(i).Header.Id = basedsn + 6 And lts(i).File.FileUnit = lunit Then
        If lts(i).attrib("TSTYPE") = "DEWP" Then
          checkcount = checkcount + 1
        End If
      ElseIf lts(i).Header.Id = basedsn + 7 And lts(i).File.FileUnit = lunit Then
        If lts(i).attrib("TSTYPE") = "CLOU" Then
          checkcount = checkcount + 1
        End If
      End If
    Next i
    If checkcount = 8 Then
      IsBASINSMetWDM = True
    End If
  End If

End Function

Private Sub DefaultBASINSMetseg(MetDetails As String)
  Dim r&, i&
  Dim SDate&(6), EDate&(6)
  Dim delim$, quote$, basedsn&, metwdmid$, lMetDetails$
  Dim lMetSeg As HspfMetSeg
  
  lMetDetails = MetDetails
  If Len(lMetDetails) > 0 Then
    'get details from the met data details
    delim = ","
    quote = """"
    basedsn = StrSplit(lMetDetails, delim, quote)
    For i = 0 To 5
      SDate(i) = StrSplit(lMetDetails, delim, quote)
    Next i
    For i = 0 To 5
      EDate(i) = StrSplit(lMetDetails, delim, quote)
    Next i
    metwdmid = lMetDetails
    
    Set lMetSeg = New HspfMetSeg
    Set lMetSeg.Uci = myUci
    For r = 1 To 8
      lMetSeg.MetSegRec(r).Source.volname = metwdmid
      lMetSeg.MetSegRec(r).Sgapstrg = ""
      lMetSeg.MetSegRec(r).Ssystem = "ENGL"
      lMetSeg.MetSegRec(r).Tran = "SAME"
      lMetSeg.MetSegRec(r).typ = r
      Select Case r
        Case 1:
          lMetSeg.MetSegRec(r).Source.volid = basedsn
          lMetSeg.MetSegRec(r).Source.member = "PREC"
          lMetSeg.MetSegRec(r).MFactP = 1
          lMetSeg.MetSegRec(r).MFactR = 1
          lMetSeg.MetSegRec(r).Sgapstrg = "ZERO"
        Case 2:
          lMetSeg.MetSegRec(r).Source.volid = basedsn + 2
          lMetSeg.MetSegRec(r).Source.member = "ATEM"
          lMetSeg.MetSegRec(r).MFactP = 1
          lMetSeg.MetSegRec(r).MFactR = 1
        Case 3:
          lMetSeg.MetSegRec(r).Source.volid = basedsn + 6
          lMetSeg.MetSegRec(r).Source.member = "DEWP"
          lMetSeg.MetSegRec(r).MFactP = 1
          lMetSeg.MetSegRec(r).MFactR = 1
        Case 4:
          lMetSeg.MetSegRec(r).Source.volid = basedsn + 3
          lMetSeg.MetSegRec(r).Source.member = "WIND"
          lMetSeg.MetSegRec(r).MFactP = 1
          lMetSeg.MetSegRec(r).MFactR = 1
        Case 5:
          lMetSeg.MetSegRec(r).Source.volid = basedsn + 4
          lMetSeg.MetSegRec(r).Source.member = "SOLR"
          lMetSeg.MetSegRec(r).MFactP = 1
          lMetSeg.MetSegRec(r).MFactR = 1
        Case 6:
          lMetSeg.MetSegRec(r).Source.volid = basedsn + 7
          lMetSeg.MetSegRec(r).Source.member = "CLOU"
          lMetSeg.MetSegRec(r).MFactP = 0
          lMetSeg.MetSegRec(r).MFactR = 1
        Case 7:
          lMetSeg.MetSegRec(r).Source.volid = basedsn + 5
          lMetSeg.MetSegRec(r).Source.member = "PEVT"
          lMetSeg.MetSegRec(r).MFactP = 1
          lMetSeg.MetSegRec(r).MFactR = 0
        Case 8:
          lMetSeg.MetSegRec(r).Source.volid = basedsn + 1
          lMetSeg.MetSegRec(r).Source.member = "EVAP"
          lMetSeg.MetSegRec(r).MFactP = 0
          lMetSeg.MetSegRec(r).MFactR = 1
      End Select
    Next r
    lMetSeg.ExpandMetSegName metwdmid, basedsn
    lMetSeg.Id = myUci.MetSegs.Count + 1
    myUci.MetSegs.Add lMetSeg
  End If
End Sub
