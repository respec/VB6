VERSION 5.00
Object = "{872F11D5-3322-11D4-9D23-00A0C9768F70}#1.10#0"; "ATCoCtl.ocx"
Begin VB.Form frmEditScenario 
   Caption         =   "Edit Scenario"
   ClientHeight    =   3690
   ClientLeft      =   1800
   ClientTop       =   3000
   ClientWidth     =   11190
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   24
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3690
   ScaleWidth      =   11190
   Begin VB.Frame fraTotalArea 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   360
      Left            =   2880
      TabIndex        =   14
      Top             =   0
      Width           =   7092
      Begin VB.TextBox txtScenario 
         Height          =   288
         HelpContextID   =   24
         Left            =   960
         TabIndex        =   15
         Text            =   "txtScenario"
         Top             =   0
         Width           =   1932
      End
      Begin ATCoCtl.ATCoText txtBasinArea 
         Height          =   300
         HelpContextID   =   24
         Left            =   5520
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   732
         _ExtentX        =   1296
         _ExtentY        =   529
         InsideLimitsBackground=   16777215
         OutsideHardLimitBackground=   255
         OutsideSoftLimitBackground=   65535
         HardMax         =   -999
         HardMin         =   0
         SoftMax         =   -999
         SoftMin         =   -999
         MaxWidth        =   5
         Alignment       =   1
         DataType        =   2
         DefaultValue    =   "0"
         Value           =   "0"
         Enabled         =   0   'False
      End
      Begin VB.Label lblScenario 
         Caption         =   "&Scenario:"
         Height          =   252
         Left            =   0
         TabIndex        =   16
         Top             =   50
         Width           =   1092
      End
      Begin VB.Label lblBasinArea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Basin Drainage Area:"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   3000
         TabIndex        =   3
         Top             =   48
         Width           =   2412
      End
      Begin VB.Label lblUnits 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "mi"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   6240
         TabIndex        =   5
         Top             =   48
         Width           =   372
      End
      Begin VB.Label lblUnitExponent 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   6624
         TabIndex        =   6
         Top             =   0
         Width           =   372
      End
   End
   Begin VB.Frame fraCribu 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   708
      Left            =   2880
      TabIndex        =   8
      Top             =   2280
      Width           =   7092
      Begin VB.CommandButton cmdMap 
         Caption         =   "&Map..."
         Height          =   252
         HelpContextID   =   24
         Left            =   5160
         TabIndex        =   11
         Top             =   360
         Width           =   972
      End
      Begin VB.ComboBox comboRegion 
         Height          =   315
         HelpContextID   =   24
         ItemData        =   "frmEditScenario.frx":0000
         Left            =   3120
         List            =   "frmEditScenario.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Crippen and Bue region for optional maximum flood-envelope computation"
         Top             =   360
         Width           =   1932
      End
      Begin VB.Label lblRange 
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   6975
      End
      Begin VB.Label lblRegion 
         BackStyle       =   0  'Transparent
         Caption         =   "&Crippen && Bue (1977) flood region:"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   405
         Width           =   3135
      End
   End
   Begin VB.ListBox lstRegion 
      Height          =   1620
      HelpContextID   =   24
      ItemData        =   "frmEditScenario.frx":0004
      Left            =   120
      List            =   "frmEditScenario.frx":0006
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      HelpContextID   =   24
      Left            =   4560
      TabIndex        =   13
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      HelpContextID   =   24
      Left            =   3000
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin ATCoCtl.ATCoGrid agd 
      Height          =   1812
      HelpContextID   =   24
      Left            =   2880
      TabIndex        =   7
      Top             =   360
      Width           =   7092
      _ExtentX        =   12515
      _ExtentY        =   3201
      SelectionToggle =   0   'False
      AllowBigSelection=   0   'False
      AllowEditHeader =   0   'False
      AllowLoad       =   0   'False
      AllowSorting    =   -1  'True
      Rows            =   1
      Cols            =   2
      ColWidthMinimum =   300
      gridFontBold    =   0   'False
      gridFontItalic  =   0   'False
      gridFontName    =   "MS Sans Serif"
      gridFontSize    =   8
      gridFontUnderline=   0   'False
      gridFontWeight  =   400
      gridFontWidth   =   0
      Header          =   ""
      FixedRows       =   1
      FixedCols       =   1
      ScrollBars      =   3
      SelectionMode   =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorBkg    =   -2147483632
      BackColorSel    =   -2147483634
      ForeColorSel    =   16777215
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      InsideLimitsBackground=   16777215
      OutsideHardLimitBackground=   8421631
      OutsideSoftLimitBackground=   8454143
      ComboCheckValidValues=   -1  'True
   End
   Begin VB.Frame sashV 
      BorderStyle     =   0  'None
      Height          =   1692
      Left            =   2760
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      Top             =   360
      Width           =   60
   End
   Begin VB.Label lblRegions 
      BackStyle       =   0  'Transparent
      Caption         =   "&Regions"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditScenario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants
Private pScenario As nssScenario
Private pMetric As Boolean

Private RegUseCnt As Integer
Private InAgd As Boolean

Private SashVdragging As Boolean

Public Property Set Scenario(newValue As nssScenario)
  Set pScenario = newValue
  pMetric = pScenario.Project.Metric
  If pMetric Then lblUnits.Caption = "km"
  txtScenario.Text = pScenario.Name
  txtBasinArea.Value = pScenario.GetArea(pMetric)
  comboRegion.ListIndex = pScenario.RegCrippenBue
  PopulateRegions
End Property

Private Sub PopulateRegions()
  Dim myRegion As Variant, myUserRegion As Variant
  lstRegion.Clear
  If pScenario.Urban Then
    lstRegion.AddItem pScenario.Project.NationalUrban.Name
    If pScenario.UserRegions.Count > 0 Then
      If pScenario.Project.NationalUrban.Name = pScenario.UserRegions(1).Region.Name Then
        lstRegion.Selected(lstRegion.ListCount - 1) = True
      End If
    End If
  End If
  For Each myRegion In pScenario.Project.DB.States(pScenario.Project.State.Code).Regions
    If myRegion.Urban = pScenario.Urban Then
      lstRegion.AddItem myRegion.Name
      For Each myUserRegion In pScenario.UserRegions
        If myUserRegion.Region.Name = myRegion.Name Then
          lstRegion.Selected(lstRegion.ListCount - 1) = True
          Exit For
        End If
      Next
    End If
  Next myRegion
End Sub

Private Sub agd_Click()
  agd_RowColChange
End Sub

Private Sub agd_CommitChange(ChangeFromRow As Long, ChangeToRow As Long, _
                             ChangeFromCol As Long, ChangeToCol As Long)
  Dim oldrow&, oldcol&
  Dim row&, col&
  Dim Max As Double
  Dim Min As Double
  Dim Val As Double
  Dim ParmName As String, parmVal As String
  If Not InAgd Then
    InAgd = True
    oldrow = agd.row
    oldcol = agd.col
    For col = ChangeFromCol To ChangeToCol
      agd.col = col
      'parmIndex = 0
      For row = 1 To ChangeToRow 'Start with first row so we can count parms
        agd.row = row
        If agd.CellBackColor <> agd.BackColorFixed Then
          'parmIndex = parmIndex + 1
          If row >= ChangeFromRow Then 'Only check rows that have been changed
            parmVal = agd.Text
            ParmName = VarName(row)
            If IsNumeric(parmVal) And Len(ParmName) > 0 Then
              With pScenario.UserRegions(col).UserParms(ParmName)
                Val = CDbl(parmVal)
                .SetValue Val, pMetric
                Max = .Parameter.GetMax(pMetric)
                Min = .Parameter.GetMin(pMetric)
                If Max > Min Then
                  If Val > Max Or Val < Min Then
                    agd.CellBackColor = agd.OutsideSoftLimitBackground
                  Else
                    agd.CellBackColor = agd.InsideLimitsBackground
                  End If
                Else
                  agd.CellBackColor = agd.InsideLimitsBackground
                End If
              End With
            Else
              agd.CellBackColor = agd.OutsideHardLimitBackground
            End If
          End If
        End If
      Next
    Next
    agd.row = oldrow
    agd.col = oldcol
    InAgd = False
    UpdateArea
  End If
End Sub

Private Sub UpdateArea()
  Dim col As Long
  Dim txt As String
  Dim totalArea As Double
  If InStr(LCase(agd.TextMatrix(1, 0)), "area") > 0 Then
    totalArea = 0
    For col = 1 To agd.cols - 1
      txt = agd.TextMatrix(1, col)
      If IsNumeric(txt) Then
        totalArea = totalArea + CDbl(txt)
      End If
    Next
    txtBasinArea.Value = totalArea
    pScenario.SetArea totalArea, pMetric
  End If
End Sub

Private Sub agd_RowColChange()
  Dim ParmName As String
  Dim Parameter As nssParameter
  Dim minString As String
  Dim maxString As String
  On Error GoTo NeverMind
  ParmName = VarName(agd.row)
  If Len(ParmName) > 0 Then
    On Error GoTo NotEditable
    Set Parameter = pScenario.UserRegions(agd.col).UserParms(ParmName).Parameter
    agd.ColEditable(agd.col) = True
    If Parameter.GetMax(pMetric) > Parameter.GetMin(pMetric) Then
      maxString = SignificantDigits(Parameter.GetMax(pMetric), 3)
      minString = SignificantDigits(Parameter.GetMin(pMetric), 3)
      agd.ColSoftMax(agd.col) = maxString
      agd.ColSoftMin(agd.col) = minString
      If InStr(maxString, ".") > 0 Then
        While Right(maxString, 1) = "0"
          maxString = Left(maxString, Len(maxString) - 1)
        Wend
        If Right(maxString, 1) = "." Then maxString = Left(maxString, Len(maxString) - 1)
      End If
      If InStr(minString, ".") > 0 Then
        While Right(minString, 1) = "0"
          minString = Left(minString, Len(minString) - 1)
        Wend
        If Right(minString, 1) = "." Then minString = Left(minString, Len(minString) - 1)
      End If
      lblRange.Caption = "Range for " & ParmName & " in " & agd.ColTitle(agd.col) & ": " _
                       & minString & " to " & maxString
    Else
      agd.ColSoftMax(agd.col) = -999
      agd.ColSoftMin(agd.col) = -999
    End If
  Else
NotEditable:
    agd.ColEditable(agd.col) = False
    agd.ColSoftMax(agd.col) = -999
    agd.ColSoftMin(agd.col) = -999
    lblRange.Caption = ""
  End If
NeverMind:
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdMap_Click()
  frmCrippenBue.Show
End Sub

Private Sub cmdOk_Click()
  Dim row As Long, col As Long
  Dim ParmName As String, parmVal As String
  Dim WarnMessage As String, nWarnings As Long
  Dim HaveBlankValues As Boolean
  Dim parmsOk As VbMsgBoxResult
  Dim Replace As VbMsgBoxResult
  Dim oldScenario As nssScenario
  Dim scenColl As FastCollection
  Dim key As String
  Dim Val As Double
  Dim Max As Double
  Dim Min As Double
  Dim vRegion As Variant
  
  nWarnings = 0
  WarnMessage = ""
  
  If pScenario.UserRegions.Count < 1 Then
    Unload Me
    Exit Sub
  End If
  Me.MousePointer = vbHourglass
  For col = 1 To agd.cols - 1
    agd.col = col
    'parmIndex = 0
    For row = 1 To agd.rows
      agd.row = row
      If agd.CellBackColor <> agd.BackColorFixed Then
        'parmIndex = parmIndex + 1
        parmVal = agd.Text
        ParmName = VarName(row)
        If Len(ParmName) > 0 Then
          With pScenario.UserRegions(col).UserParms(ParmName)
            If Not IsNumeric(parmVal) Then
              nWarnings = nWarnings + 1
              HaveBlankValues = True
              WarnMessage = WarnMessage & vbCr _
                          & "Region '" & agd.ColTitle(col) _
                          & "', Parameter '" & .Parameter.NSSName _
                          & "' value is not a number."
            Else
              .SetValue CDbl(parmVal), pMetric
              Val = .GetValue(pMetric)
              Max = .Parameter.GetMax(pMetric)
              Min = .Parameter.GetMin(pMetric)
              If Max > Min Then
                If Val > Max Then
                  nWarnings = nWarnings + 1
                  WarnMessage = WarnMessage & vbCr _
                              & "Region '" & agd.ColTitle(col) _
                              & "', Parameter '" & .Parameter.NSSName _
                              & "' value " & Val & " > suggested max of " & Max
                ElseIf Val < Min Then
                  nWarnings = nWarnings + 1
                  WarnMessage = WarnMessage & vbCr _
                              & "Region '" & agd.ColTitle(col) _
                              & "', Parameter '" & .Parameter.NSSName _
                              & "' value " & Val & " < suggested min of " & Min
                End If
              End If
            End If
          End With
        End If
      End If
    Next
  Next
  
  pScenario.ROI = False
  pScenario.UsePredInts = True
  pScenario.LowFlow = True
  pScenario.ProbEqtn = True
  For Each vRegion In pScenario.UserRegions
    If Not vRegion.Region.PredInt Then pScenario.UsePredInts = False
    If vRegion.Region.LowFlowRegnID <= 0 Then pScenario.LowFlow = False
    If vRegion.Region.LowFlowRegnID >= 0 Then pScenario.ProbEqtn = False
    If vRegion.Region.ROIRegnID >= 0 Then
      pScenario.ROI = True
      pScenario.UsePredInts = True 'ROI always uses prediction intervals
    End If
  Next
  
  If pScenario.ROI Then
    If pScenario.UserRegions.Count > 1 Then
      ssMessageBox "If a ROI region is specified, it must be the only region.", vbCritical, _
             "Cannot continue with more than one ROI region"
      parmsOk = vbCancel
      nWarnings = -1
    End If
  End If
  
  If nWarnings = 0 Then
    parmsOk = vbOK
  ElseIf nWarnings > 0 Then
    If HaveBlankValues Then
      ssMessageBox "All values must be specified before this form is complete." & vbCr _
            & WarnMessage, vbCritical, "Cannot continue until all values are specified"
      parmsOk = vbCancel
    Else
      If nWarnings = 1 Then
        WarnMessage = "Warning: One parameter is outside the suggested range" & vbCr _
                    & "Estimates will be extrapolations with unknown errors." & vbCr _
                    & WarnMessage & vbCr & vbCr & "Proceed with this value anyway?"
      Else
        WarnMessage = "Warning: " & nWarnings & " parameters are outside the suggested range" & vbCr _
                    & "Estimates will be extrapolations with unknown errors." & vbCr _
                    & WarnMessage & vbCr & vbCr & "Proceed with these values anyway?"
      End If
      parmsOk = ssMessageBox(WarnMessage, vbOKCancel, "Warning")
    End If
  End If
  
  If parmsOk = vbOK Then
    If pScenario.Urban Then
      Set scenColl = pScenario.Project.UrbanScenarios
    Else
      Set scenColl = pScenario.Project.RuralScenarios
    End If
    
    key = LCase(txtScenario.Text)
    Replace = vbYes
    
    On Error GoTo AddScenario
    Set oldScenario = scenColl(key) 'This will create an error if no existing scenario has this name
    If key <> LCase(pScenario.Name) Then
      'Confirm replacing existing scenario if it is different than the one we started with
      Replace = ssMessageBox("Replace existing scenario named '" _
                      & oldScenario.Name _
                      & "' with this one?", vbOKCancel, "Replacing Scenario")
    End If
    If Replace = vbYes Then
      scenColl.Remove key
AddScenario:
      pScenario.Name = txtScenario.Text
      pScenario.RegCrippenBue = comboRegion.ListIndex
      scenColl.Add pScenario, key
      If pScenario.Urban Then
        pScenario.Project.CurrentUrbanScenario = scenColl.Count
      Else
        pScenario.Project.CurrentRuralScenario = scenColl.Count
      End If
      pScenario.Project.RaiseEdited
      Unload Me
    End If
  End If
  Me.MousePointer = vbDefault
  
End Sub

Private Sub Form_Load()
  Dim rgn&
  
  lstRegion.Clear
  
  txtScenario.Text = ""
  
  comboRegion.Clear
  comboRegion.AddItem "None"
  For rgn = 1 To 17
    comboRegion.AddItem rgn
  Next rgn
  comboRegion.ListIndex = 0

End Sub

Private Sub Form_Resize()
  Dim fh&, fw&, newRightWidth&
  Static lastHeight&
  fw = Width
  fh = Height
  'sashV.Left = fw - agd.Width - sashV.Width - 100
  'If sashV.Left < 200 Then
  '  sashV.Left = 100
  '  sashV_MouseMove 0, 0, 0, 0
  'Else
  'If Width > 3500 Then agd.Width = Width - 3150
  newRightWidth = fw - (sashV.Left + sashV.Width + 100)
  If newRightWidth > 0 And agd.Width <> newRightWidth Then
    agd.Width = newRightWidth
    fraCribu.Width = newRightWidth
    lblRange.Width = newRightWidth - lblRange.Left
  End If
  agd.Left = sashV.Left + sashV.Width
  fraTotalArea.Left = agd.Left
  fraCribu.Left = agd.Left
  If Height > 2400 Then
    'If rurfg Then
      agd.Height = Height - 2200 '1800
      fraCribu.Top = agd.Top + agd.Height + comboRegion.Height / 2
      cmdOk.Top = fraCribu.Top + fraCribu.Height + cmdOk.Height / 3
    'Else
    '  agd.Height = Height - 1655
    '  cmdOk.Top = agd.Top + agd.Height + cmdOk.Height / 3
    'End If
    If agd.Height <> lastHeight Then
      'eliminate flashing because list height snaps to full line heights
      lastHeight = agd.Height
      lstRegion.Height = lastHeight
      sashV.Height = lastHeight
    End If
    cmdCancel.Left = Width / 2 + cmdCancel.Width / 4
    cmdOk.Left = cmdCancel.Left - cmdOk.Width * 1.5
  End If

  'lblRegion.Top = comboRegion.Top
  cmdCancel.Top = cmdOk.Top
End Sub

Private Sub sashV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  SashVdragging = True
End Sub

Private Sub sashV_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If SashVdragging And (sashV.Left + X) > 100 And (sashV.Left + X < Width - 100) Then
    Dim newLeftWidth&
    sashV.Left = sashV.Left + X
    If sashV.Left < lstRegion.Left + 200 Then sashV.Left = lstRegion.Left + 200
    newLeftWidth = sashV.Left - lstRegion.Left
    If newLeftWidth > 0 And lstRegion.Width <> newLeftWidth Then lstRegion.Width = newLeftWidth
    Form_Resize
  End If
End Sub

Private Sub sashV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  SashVdragging = False
End Sub

Private Sub lstRegion_Click()
  Dim i&

  Static InClick As Boolean
  If Not InClick Then
    InClick = True
    If pScenario.Urban Then 'Can only have one urban equation, unselect others
      For i = 0 To lstRegion.ListCount - 1
        If i <> lstRegion.ListIndex Then lstRegion.Selected(i) = False
      Next i
    End If
    SetSelectedRegionsFromList
    InClick = False
  End If
End Sub

Private Sub SetSelectedRegionsFromList()
  Static SettingSelected As Boolean
  Dim lstIndx As Long
  Dim newRegion As userRegion
  Dim myUserRegion As Variant
  
  'If lstRegion.SelCount > MAX_REG_USE And pScenario.Urban = False Then
  '  MsgBox "A maximum of " & MAX_REG_USE & " regions may be used for rural calculations.", _
  '         48, "NSS Limit"
  '  lstRegion.Selected(lstRegion.ListIndex) = False
  'Else
  If lstRegion.SelCount > 1 And pScenario.Urban Then
    ssMessageBox "Only one equation may be used for urban calculations.", 48, "NSS Limit"
  Else 'ok to add the region
    If pScenario.Urban Then
      If pScenario.UserRegions.Count = 1 Then
        If pScenario.UserRegions(1).Region.Name = lstRegion.List(lstRegion.ListIndex) Then
          GoTo AlreadyHaveRegion
        Else
          pScenario.UserRegions.Remove 1
        End If
      End If
      If pScenario.UserRegions.Count = 0 Then
        Set newRegion = New userRegion
        If lstRegion.List(lstRegion.ListIndex) = pScenario.Project.NationalUrban.Name Then
          Set newRegion.Region = pScenario.Project.NationalUrban
        Else
          Set newRegion.Region = pScenario.Project.DB.States(pScenario.Project.State.Code).Regions(lstRegion.List(lstRegion.ListIndex))
        End If
        If newRegion.Region.UrbanNeedsRural Then
          If pScenario.Project.CurrentRuralScenario > 0 Then
            If pScenario.Project.RuralScenarios(pScenario.Project.CurrentRuralScenario).LowFlow Or _
               pScenario.Project.RuralScenarios(pScenario.Project.CurrentRuralScenario).LowFlow Then
              ssMessageBox "This equation requires a rural scenario result," & vbCr _
                 & "but the current rural scenario is for Low Flow estimates.", vbOKOnly, "Urban Requires Rural"
            Else
              Set pScenario.RuralScenario = pScenario.Project.RuralScenarios(pScenario.Project.CurrentRuralScenario)
              txtBasinArea.Value = pScenario.GetArea(pMetric)
            End If
          Else
            ssMessageBox "This equation requires a rural scenario result," & vbCr _
                 & "but no rural scenario is selected.", vbOKOnly, "Urban Requires Rural"
          End If
        End If
        pScenario.UserRegions.Add newRegion, newRegion.Region.Name
      End If
    Else  'rural
      On Error Resume Next
      'Errors occur when adding existing keys or deleting non-existing keys from collections
      If lstRegion.Selected(lstRegion.ListIndex) Then
        For Each myUserRegion In pScenario.UserRegions
          If myUserRegion.Region.Name = lstRegion.List(lstRegion.ListIndex) Then GoTo AlreadyHaveRegion
        Next
        Set newRegion = New userRegion
        Set newRegion.Region = pScenario.Project.DB.States(pScenario.Project.State.Code).Regions(lstRegion.List(lstRegion.ListIndex))
        pScenario.UserRegions.Add newRegion, newRegion.Region.Name
        Set newRegion = Nothing
        If Left(newRegion.Region.Name, 3) = "ROI" Then
          pScenario.ROI = True
        Else
          pScenario.ROI = False
        End If
      Else
        pScenario.UserRegions.Remove lstRegion.List(lstRegion.ListIndex)
      End If
    End If
  End If
AlreadyHaveRegion:
  'If pScenario.Urban And pScenario.UserRegions.Count > 0 Then
  '  If pScenario.UserRegions(1).UrbanNeedsRural Then _
  '      ssMessageBox "Before this urban equation can be used, a rural equation for the area must be computed."
  'Else
    SetComputeGrid
  'End If
End Sub

Private Sub SetComputeGrid()
  Dim row&, col&, UnitString As String, desiredHeight&
  Dim curReg As Long
  Dim curVarName As String
  Dim Metric As Boolean, eqnMetric As Boolean
  Dim myRegion As userRegion
  Dim myParm As Variant
  
  Metric = pScenario.Project.State.Metric
'  If pScenario.Urban Then
'    SetUrbEstimate  'init urban info
'  Else
'    SetRurEstimate  'init rural info
'  End If
  RegUseCnt = pScenario.UserRegions.Count
  
  'initialize grid parms
  agd.rows = 1 'will be expanded as necessary below
  agd.TextMatrix(agd.rows, 0) = ""
  agd.cols = RegUseCnt + 1
  agd.ColType(0) = ATCoTxt
  agd.ColTitle(0) = "Variable"

  For curReg = 1 To RegUseCnt
    agd.ColType(curReg) = ATCoSng
    agd.ColEditable(curReg) = True
    Set myRegion = pScenario.UserRegions(curReg)
    agd.ColTitle(curReg) = myRegion.Region.Name
    agd.col = curReg
    For row = 1 To agd.rows
      agd.row = row
      agd.Text = ""
      agd.CellBackColor = agd.BackColorFixed
    Next
    For Each myParm In myRegion.UserParms
      curVarName = myParm.Parameter.NSSName
      'Find row that already has this variable if it already exists
      row = 1
      While row < agd.rows And VarName(row) <> curVarName
        row = row + 1
      Wend
      If VarName(row) <> curVarName Then
        While VarName(row) <> ""
          row = row + 1
        Wend
        UnitString = myParm.Parameter.Units.Label(pMetric)
        If Len(UnitString) > 0 And UnitString <> "-" Then
          agd.TextMatrix(row, 0) = curVarName & " (" & UnitString & ")"
        Else
          agd.TextMatrix(row, 0) = curVarName
        End If
        agd.row = row
        For col = 1 To curReg - 1
          agd.col = col
          agd.CellBackColor = agd.BackColorFixed
        Next
      End If
      agd.row = row
      agd.col = curReg
      If myParm.GetValue(pMetric) = -999 Then
        agd.Text = ""
        agd.CellBackColor = agd.OutsideSoftLimitBackground
      Else
        agd.CellBackColor = agd.InsideLimitsBackground
        agd.Text = SignificantDigits(myParm.GetValue(pMetric), 3)
      End If
    Next myParm
  Next curReg

'  For row = 1 To agd.rows
'    For col = 1 To agd.cols - 1
'      agd.row = row
'      agd.col = col
'      ival = agdVarInd(row, col)
'      If ival >= 0 Then
'        If rval(ival, col - 1) = -999 Then
'          agd.Text = ""
'          agd.CellBackColor = agd.OutsideSoftLimitBackground
'        Else
'          agd.Text = rval(ival, col - 1)
'          If InRange(ival, col - 1) Then
'            agd.CellBackColor = agd.InsideLimitsBackground
'          Else
'            agd.CellBackColor = agd.OutsideSoftLimitBackground
'          End If
'        End If
'      Else
'        agd.CellBackColor = agd.BackColorFixed
'        agd.Text = ""
'      End If
'    Next col
'  Next row
'  txtBasinArea_UpdateValue
  agd.row = 1
  agd.col = 1
  agd.ColsSizeByContents
'  agd_RowColChange
  If agd.gridWidth + agd.Left + 100 > Me.Width Then Me.Width = agd.gridWidth + agd.Left + 100
'  If rurfg Then
'    desiredHeight = 2100 + 253 * (agd.rows + 1)
'  Else
'    desiredHeight = 1655 + 253 * (agd.rows + 1)
'  End If
'  If desiredHeight > Height Then Height = desiredHeight
End Sub

Private Function VarName(ByVal row As Long) As String
  Dim retval As String
  Dim UnitsPos As String
  retval = agd.TextMatrix(row, 0)
  UnitsPos = InStr(retval, " (")
  If UnitsPos > 0 Then retval = Left(retval, UnitsPos - 1)
  VarName = retval
End Function
