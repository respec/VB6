VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "*\A..\..\ATCoCtl\ATCoCtl.vbp"
Begin VB.Form frmLowFlow 
   Caption         =   "Streamflow Equation Editor"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   9690
   Icon            =   "frmLowFlow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraUnits 
      Caption         =   "Units"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   38
      Top             =   480
      Width           =   2175
      Begin VB.OptionButton rdoUnits 
         Caption         =   "English"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   40
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton rdoUnits 
         Caption         =   "Metric"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1170
         TabIndex        =   39
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdDatabase 
      Caption         =   "Se&lect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8880
      TabIndex        =   32
      Top             =   7440
      Width           =   732
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   8880
      TabIndex        =   29
      Top             =   600
      Width           =   732
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Ex&port"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7920
      TabIndex        =   28
      Top             =   600
      Width           =   732
   End
   Begin VB.OptionButton rdoMainOpt 
      Caption         =   "Low Flow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   1560
      TabIndex        =   27
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton rdoMainOpt 
      Caption         =   "Peak Flow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   26
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8880
      TabIndex        =   25
      Top             =   8040
      Width           =   732
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   8880
      TabIndex        =   21
      Top             =   4320
      Width           =   732
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   8880
      TabIndex        =   20
      Top             =   3840
      Width           =   732
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   8880
      TabIndex        =   19
      Top             =   4800
      Width           =   732
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   8880
      TabIndex        =   18
      Top             =   5280
      Width           =   732
   End
   Begin VB.Frame fraEdit 
      Caption         =   "Return Period Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   8655
      Begin ATCoCtl.ATCoGrid grdComps 
         CausesValidation=   0   'False
         Height          =   1515
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2672
         SelectionToggle =   0   'False
         AllowBigSelection=   -1  'True
         AllowEditHeader =   0   'False
         AllowLoad       =   0   'False
         AllowSorting    =   0   'False
         Rows            =   2
         Cols            =   2
         ColWidthMinimum =   300
         gridFontBold    =   0   'False
         gridFontItalic  =   0   'False
         gridFontName    =   "MS Sans Serif"
         gridFontSize    =   8
         gridFontUnderline=   0   'False
         gridFontWeight  =   400
         gridFontWidth   =   0
         Header          =   "Equation Components"
         FixedRows       =   2
         FixedCols       =   0
         ScrollBars      =   2
         SelectionMode   =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorBkg    =   -2147483632
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         InsideLimitsBackground=   -2147483643
         OutsideHardLimitBackground=   8421631
         OutsideSoftLimitBackground=   8454143
         ComboCheckValidValues=   -1  'True
      End
      Begin VB.CommandButton cmdComponent 
         Caption         =   "Remove Component"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   34
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton cmdComponent 
         Caption         =   "Add Component"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   33
         Top             =   3120
         Width           =   2055
      End
      Begin ATCoCtl.ATCoGrid grdMatrix 
         CausesValidation=   0   'False
         Height          =   1635
         Left            =   120
         TabIndex        =   31
         Top             =   3960
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2884
         SelectionToggle =   0   'False
         AllowBigSelection=   -1  'True
         AllowEditHeader =   0   'False
         AllowLoad       =   0   'False
         AllowSorting    =   0   'False
         Rows            =   353
         Cols            =   2
         ColWidthMinimum =   300
         gridFontBold    =   0   'False
         gridFontItalic  =   0   'False
         gridFontName    =   "MS Sans Serif"
         gridFontSize    =   8
         gridFontUnderline=   0   'False
         gridFontWeight  =   400
         gridFontWidth   =   0
         Header          =   "Covariance Matrix"
         FixedRows       =   0
         FixedCols       =   0
         ScrollBars      =   3
         SelectionMode   =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorBkg    =   -2147483632
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         InsideLimitsBackground=   -2147483643
         OutsideHardLimitBackground=   8421631
         OutsideSoftLimitBackground=   8454143
         ComboCheckValidValues=   0   'False
      End
      Begin ATCoCtl.ATCoGrid grdInterval 
         CausesValidation=   0   'False
         Height          =   1185
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2090
         SelectionToggle =   0   'False
         AllowBigSelection=   -1  'True
         AllowEditHeader =   0   'False
         AllowLoad       =   0   'False
         AllowSorting    =   0   'False
         Rows            =   2
         Cols            =   2
         ColWidthMinimum =   300
         gridFontBold    =   0   'False
         gridFontItalic  =   0   'False
         gridFontName    =   "MS Sans Serif"
         gridFontSize    =   8
         gridFontUnderline=   0   'False
         gridFontWeight  =   400
         gridFontWidth   =   0
         Header          =   "Return Interval"
         FixedRows       =   2
         FixedCols       =   0
         ScrollBars      =   0
         SelectionMode   =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorBkg    =   -2147483632
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         InsideLimitsBackground=   -2147483643
         OutsideHardLimitBackground=   8421631
         OutsideSoftLimitBackground=   8454143
         ComboCheckValidValues=   0   'False
      End
      Begin VB.Label lblEquation 
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   3480
         Width           =   6735
      End
   End
   Begin VB.Frame fraEdit 
      Caption         =   "Parameter Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3012
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   8655
      Begin ATCoCtl.ATCoGrid grdParms 
         Height          =   1695
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   2990
         SelectionToggle =   0   'False
         AllowBigSelection=   -1  'True
         AllowEditHeader =   0   'False
         AllowLoad       =   0   'False
         AllowSorting    =   0   'False
         Rows            =   2
         Cols            =   6
         ColWidthMinimum =   300
         gridFontBold    =   0   'False
         gridFontItalic  =   0   'False
         gridFontName    =   "MS Sans Serif"
         gridFontSize    =   8
         gridFontUnderline=   0   'False
         gridFontWeight  =   400
         gridFontWidth   =   0
         Header          =   ""
         FixedRows       =   2
         FixedCols       =   0
         ScrollBars      =   3
         SelectionMode   =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorBkg    =   -2147483632
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         InsideLimitsBackground=   -2147483643
         OutsideHardLimitBackground=   8421631
         OutsideSoftLimitBackground=   8454143
         ComboCheckValidValues=   0   'False
      End
   End
   Begin VB.Frame fraSelections 
      Caption         =   "Selections"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   9495
      Begin VB.ListBox lstRetPds 
         Height          =   1230
         Left            =   7440
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.ListBox lstParms 
         Height          =   1230
         Left            =   3960
         MultiSelect     =   1  'Simple
         TabIndex        =   7
         Top             =   600
         Width           =   3375
      End
      Begin VB.ListBox lstRegions 
         Height          =   1230
         ItemData        =   "frmLowFlow.frx":030A
         Left            =   120
         List            =   "frmLowFlow.frx":0311
         TabIndex        =   6
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label lblParms 
         Caption         =   "Parameters:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblRetPds 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         TabIndex        =   4
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lblRegions 
         Caption         =   "Regions:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1212
      End
   End
   Begin VB.Frame fraEdit 
      Caption         =   "Region Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3012
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   7815
      Begin VB.Frame fraRegType 
         Caption         =   "Region Type"
         Height          =   1332
         Left            =   4080
         TabIndex        =   14
         Top             =   240
         Width           =   3012
         Begin VB.CheckBox chkPredInt 
            Caption         =   "use prediction intervals"
            Height          =   252
            Left            =   120
            TabIndex        =   30
            Top             =   960
            Width           =   2652
         End
         Begin VB.CheckBox chkRuralInput 
            Caption         =   "with Rural Input"
            Height          =   252
            Left            =   1080
            TabIndex        =   17
            Top             =   480
            Visible         =   0   'False
            Width           =   1752
         End
         Begin VB.OptionButton rdoRegOpt 
            Caption         =   "Urban"
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   852
         End
         Begin VB.OptionButton rdoRegOpt 
            Caption         =   "Rural"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1092
         End
      End
      Begin VB.TextBox txtRegName 
         Height          =   288
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label lblRegName 
         Caption         =   "Region Name:"
         Height          =   252
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1212
      End
   End
   Begin VB.ComboBox cboState 
      Height          =   315
      ItemData        =   "frmLowFlow.frx":0321
      Left            =   3720
      List            =   "frmLowFlow.frx":0323
      TabIndex        =   0
      Text            =   "cboState"
      Top             =   600
      Width           =   1452
   End
   Begin MSComDlg.CommonDialog cdlgFileSel 
      Left            =   -120
      Top             =   360
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label lblDatabase 
      Caption         =   "Database:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   37
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label lblState 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&State:"
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
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   600
      Width           =   570
   End
End
Attribute VB_Name = "frmLowFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBPath As String
Dim RDO As Long
Dim MyParm As nssParameter
Dim MyDepVar As nssDepVar
Dim MyComp As nssComponent
Dim NotNew As Boolean, Skip As Boolean, _
    ChoseParms As Boolean, ChoseReturns As Boolean
Dim Changes() As String, CompChanges() As String, _
    MatrixChanges() As String, OldMatrix() As String
Dim SelParms As FastCollection  'of NSSParameter
Dim Metric As Boolean

Private Sub cmdCancel_Click()
  Dim row&, col&, i&

  If fraEdit(0).Visible Then
    For row = 1 To DB.State.Regions.Count
      If DB.State.Regions(row).IsNew Then
        DB.State.Regions.Remove row
        lstRegions.RemoveItem (row - 1)
      End If
    Next row
    If MyRegion.IsNew Then
      fraEdit(0).Visible = False
      Set MyRegion = Nothing
    ElseIf Not Skip Then
      ResetRegion
    End If
  ElseIf fraEdit(1).Visible Then
    For row = 1 To SelParms.Count
      If SelParms(row - i).IsNew Then
        SelParms.RemoveByIndex (row - i)
        grdParms.DeleteRow (row - i)
        lstParms.RemoveItem lstParms.ListCount - 1
        i = i + 1
      ElseIf Not Skip Then
        For col = 0 To ParmFlds
          With grdParms
            Select Case col
              Case 0: .TextMatrix(row - i, col) = SelParms(row - i).statTypeCode
              Case 1: .TextMatrix(row - i, col) = SelParms(row - i).Name
              Case 2: .TextMatrix(row - i, col) = SelParms(row - i).Abbrev
              Case 3: .TextMatrix(row - i, col) = SelParms(row - i).GetMin(False)
              Case 4: .TextMatrix(row - i, col) = SelParms(row - i).GetMax(False)
              Case 5: .TextMatrix(row - i, col) = CInt(SelParms(row - i).Units.id)
            End Select
          End With
        Next col
      End If
    Next row
    If SelParms.Count = 0 Then
      cmdSave.Enabled = False
      cmdDelete.Enabled = False
      cmdCancel.Enabled = False
    End If
    If Not MyParm Is Nothing Then MyParm.IsNew = False
  ElseIf fraEdit(2).Visible = True Then
    ChoseReturns = True
    If MyRegion.depVars.Count > 0 Then SetGrid "DepVars"
    ChoseReturns = False
    If lstRetPds.ListIndex >= 0 Then
      cmdSave.Enabled = True
      cmdDelete.Enabled = True
      cmdCancel.Enabled = True
    Else
      cmdSave.Enabled = False
      cmdDelete.Enabled = False
      cmdCancel.Enabled = False
    End If
  End If
End Sub

Private Sub cmdComponent_Click(Index As Integer)
  Dim Resp As Integer, i As Integer

  Resp = vbYes
  With grdComps
    If Index = 0 Then 'add row
      For i = 1 To .Rows 'make sure there isn't a blank row already
        If Len(.TextMatrix(i, 0)) = 0 Then
          Resp = MsgBox("There is already a blank component row available for adding a component." & vbCrLf & _
                 "Are you sure you want to add another component row?", vbYesNo + vbExclamation, "Add Component")
          Exit For
        End If
      Next i
      If Resp = vbYes Then
        .Rows = .Rows + 1
        .TextMatrix(.Rows, 1) = "0"
        .TextMatrix(.Rows, 2) = "1"
        .TextMatrix(.Rows, 4) = "none"
        .TextMatrix(.Rows, 5) = "0"
        .TextMatrix(.Rows, 6) = "0"
      End If
    ElseIf .Rows > 1 Then 'delete row
      If Len(.TextMatrix(.row, 0)) > 0 Then 'confirm delete of existing component
        Resp = MsgBox("Are you sure you want to delete the component " & .TextMatrix(.row, 0) & "?", vbYesNo + vbInformation, "Remove Component")
      End If
      If Resp = vbYes Then .Rows = .Rows - 1
    End If
  End With
  With grdMatrix
    .Rows = grdComps.Rows + 1
    .cols = .Rows
    For i = 0 To .cols - 1
      .ColEditable(i) = True
      .colWidth(i) = 1000
    Next i
  End With
End Sub

Private Sub cmdDatabase_Click()
  Dim lDBFName As String
  lDBFName = DB.FileName

  SetDB (True)
  If lDBFName <> DB.FileName Then 'database changed
    rdoMainOpt(0).Value = False
    rdoMainOpt(1).Value = False
    fraEdit(0).Visible = False
    fraEdit(1).Visible = False
    fraEdit(2).Visible = False
  End If

End Sub

Private Sub cmdExit_Click()
  Dim Resp As Integer
  On Error GoTo x
  Resp = vbYes
  If ChangesMade Then
    Resp = MsgBox("You have unsaved values.  Are you sure you want to exit without saving them?", vbExclamation + vbYesNo, "Exit Confirmation")
  End If
  If Resp = vbYes Then
    If Len(Dir(DBPath)) > 0 Then MyRegion.DB.DB.Close
x:
    Unload Me
  End If
End Sub

Private Sub cmdHelp_Click()
  Dim helpFilePath As String
  
  On Error GoTo x
  
  helpFilePath = GetSetting("SEE", "Defaults", "HelpPath", App.path & "\SEE.chm")
  If Len(Dir(helpFilePath)) = 0 Then
    With cdlgFileSel
BadFile:
      .DialogTitle = "Select the help file"
      .FileName = App.path
      .Filter = "(*.chm)|*.chm"
      .FilterIndex = 1
      .CancelError = True
      .ShowOpen
      helpFilePath = .FileName
      If Len(Dir(helpFilePath)) = 0 Then
        MsgBox "Could not find '" & helpFilePath & "'."
        GoTo BadFile
      End If
    End With
    SaveSetting "SEE", "Defaults", "HelpPath", helpFilePath
  End If
x:
  If Len(Dir(helpFilePath)) > 0 Then
    OpenFile helpFilePath, cdlgFileSel
  Else
    MsgBox "Help file not available"
  End If
End Sub

Private Sub cmdImport_Click()
  Dim i&, j&, k&, m&, inFile&, regnCnt&, parmCnt&, depVarCnt&, _
      compCnt&, flds&, DepVarID&, response&
  Dim FileName$, str$, regnName$, flowFlag$
  Dim urban As Boolean, isReturn As Boolean, isLowFlow As Boolean
  Dim regnVals() As Integer
  Dim parmVals() As String, depVarVals() As String, compVals() As String
  Dim covArray() As String

  On Error GoTo x
  
  response = myMsgBox.Show("Importing peak-flow or low-flow data will replace " & _
        "all such data in the database for that state." & vbCrLf & vbCrLf & _
        "Do you wish to continue?", _
        "User Action Verification", "+&Yes", "-&Cancel")
  If response = 2 Then Exit Sub

TryAgain:
  With cdlgFileSel
    .DialogTitle = "Select import file"
    If RDO = 0 Then
      FileName = GetSetting("SEE", "Defaults", "NSSExportFile", FileName)
    ElseIf RDO = 1 Then
      FileName = GetSetting("SEE", "Defaults", "LowFlowExportFile", FileName)
    End If
    If Len(Dir(FileName, vbDirectory)) = 0 Then
      FileName = CurDir & "\Import.txt"
    Else
      FileName = FileName & "\Import.txt"
    End If
    .FileName = FileName
    .Filter = "(*.txt)|*.txt"
    .FilterIndex = 1
    .CancelError = True
    .ShowOpen
    FileName = .FileName
  End With
  If Len(Dir(FileName, vbDirectory)) = 0 Then
    MsgBox "The filename you selected does not exist." & vbCrLf & _
           "Try again or cancel out of the dialog box."
    GoTo TryAgain
  End If

  Me.MousePointer = vbHourglass
  inFile = FreeFile
  Open FileName For Input As inFile
  'read in state info
  Line Input #inFile, str
  Set DB.State = DB.States(StrRetRem(str))
  While cboState.ItemData(i) <> CLng(DB.State.code)
    i = i + 1
  Wend
  cboState.ListIndex = i
  flowFlag = Right(str, 1)
  If flowFlag = "0" Then
    DB.State.ClearState "ReturnPeriods"
    isReturn = True
    isLowFlow = False
  ElseIf flowFlag = "1" Then
    DB.State.ClearState "Statistics"
    isReturn = False
    isLowFlow = True
  End If

  DepVarFlds = 10 'always set to import all possible DepVar fields
  
  'read in number of regions and metric flag
  Line Input #inFile, str
  regnCnt = CLng(StrRetRem(str))
  If str = "1" Then Metric = True Else Metric = False
  
  'loop thru regions
  ReDim regnVals(RegionFlds - 3)
  For i = 1 To regnCnt
    'read in region info
    lstParms.Clear
    lstRetPds.Clear
    Line Input #inFile, str
    regnName = StrSplit(str, vbTab, "")
    If StrRetRem(str) = "0" Then urban = False Else urban = True
    regnVals(0) = StrRetRem(str)
    regnVals(1) = StrRetRem(str)
    Set MyRegion = New nssRegion
    Set MyRegion.DB = DB
    MyRegion.Add isReturn, regnName, urban, regnVals(0), regnVals(1), -1
    DB.State.PopulateRegions
    Set MyRegion = DB.State.Regions(regnName)
    'Read in parameters info
    Line Input #inFile, str
    parmCnt = StrRetRem(str)  'number or parameters for this region
    ReDim parmVals(parmCnt - 1, ParmFlds)
    depVarCnt = StrSplit(str, vbTab, "")  'number or RetPds/Statistics for this region
    If depVarCnt > 0 Then ReDim depVarVals(depVarCnt - 1, DepVarFlds + 1)
    'Loop thru parameters
    For j = 0 To parmCnt - 1
      If j > 0 Then
        Line Input #inFile, str
        str = Mid(str, 2) 'gets rid or initial vbtab
      End If
      str = Trim(str)
      'Read in fields for each parm
      For k = 0 To ParmFlds
        If k < 2 Then
          parmVals(j, k) = StrSplit(str, vbTab, "")
        Else
          parmVals(j, k) = StrRetRem(str)
        End If
      Next k
      If parmVals(j, 0) <> "RDA" And parmVals(j, 0) <> "CRD" Then
        'Write values to DB
        Set MyParm = New nssParameter
        MyParm.Add MyRegion, parmVals(j, 0), parmVals(j, 2), _
            parmVals(j, 3), parmVals(j, 4)
      End If
    Next j
    MyRegion.PopulateParameters
    For k = 0 To lstRegions.ListCount - 1
      If lstRegions.List(k) = MyRegion.Name Then Exit For
    Next k
    If k = lstRegions.ListCount Then 'region not in list - add it (importing?)
      lstRegions.List(k) = MyRegion.Name
      lstRegions.ItemData(k) = MyRegion.id
    Else
      lstRegions.ListIndex = k
    End If
    ResetRegion
    'Loop thru Return Periods/Statistics
    For j = 1 To depVarCnt
      'Read in values
      Line Input #inFile, str
      str = Mid(str, 2)  'gets rid of initial vbtab
      For k = 0 To DepVarFlds + 1 '+1 is for component count at end of line
        If Len(str) > 0 Then
          depVarVals(j - 1, k) = StrRetRem(str)
        Else
          Exit For
        End If
      Next k
      compCnt = depVarVals(j - 1, k - 1)
      If compCnt > 0 Then ReDim compVals(depVarCnt - 1, compCnt - 1, CompFlds)
      'Write values to DB
      Set MyDepVar = New nssDepVar
      DepVarID = MyDepVar.Add(isReturn, MyRegion, depVarVals(j - 1, 0), depVarVals(j - 1, 1), _
          depVarVals(j - 1, 2), depVarVals(j - 1, 3), depVarVals(j - 1, 4), depVarVals(j - 1, 5), _
          depVarVals(j - 1, 6), depVarVals(j - 1, 7), depVarVals(j - 1, 8), depVarVals(j - 1, 9))
      MyRegion.PopulateDepVars
      'Loop thru Components
      For k = 0 To compCnt - 1
        'Read in values
        Line Input #inFile, str
        str = Mid(str, 3)  'gets rid of initial 2 tabs
        For m = 0 To CompFlds
          If m = 0 Or m = 4 Then
            compVals(j - 1, k, m) = GetCode(StrSplit(str, vbTab, ""))
          Else
            compVals(j - 1, k, m) = StrRetRem(str)
          End If
        Next m
        'Get rid of instructional equation on end of component string
        compVals(j - 1, k, m - 1) = StrSplit(compVals(j - 1, k, m - 1), vbTab, "")
        'Write values to DB
        Set MyComp = New nssComponent
        MyComp.Add MyRegion, DepVarID, CLng(compVals(j - 1, k, 0)), compVals(j - 1, k, 1), _
            compVals(j - 1, k, 2), compVals(j - 1, k, 3), CLng(compVals(j - 1, k, 4)), _
            compVals(j - 1, k, 5), compVals(j - 1, k, 6)
      Next k
      'Loop thru Covariance Matrix
      If compCnt > 0 And MyRegion.PredInt Then
        ReDim covArray(1 To compCnt + 1, 1 To compCnt + 1)
        For k = 1 To compCnt + 1
          Line Input #inFile, str
          While Left(str, 1) = vbTab Or Left(str, 1) = " "
            str = Mid(str, 2)  'gets rid of initial tabs/spaces
          Wend
          For m = 1 To compCnt + 1
            covArray(k, m) = StrRetRem(str)
          Next m
        Next k
        MyDepVar.AddMatrix MyRegion, DepVarID, covArray()
      End If
    Next j
  Next i
  Close inFile
'  ResetRegion
  cboState_Click
  Me.MousePointer = vbDefault
  MsgBox "Completed import from file " & FileName, , "SEE Import"
  Exit Sub
x:
  Me.MousePointer = vbDefault
  If Err.Number = 32755 Then Exit Sub
  MsgBox "The format of the import file is not correct." & vbCrLf & _
      "Make sure the indicated number of Regions, Parameters, " & _
      "Return Periods/Statistics, and Components are correct", _
      vbCritical, "Import Error"
End Sub

Private Sub cmdExport_Click()
  Dim i&, j&, k&, OutFile&, tmpCnt&, compCnt&, row&, col&
  Dim FileName$, str$
  Dim covArray() As String
  
  On Error GoTo x
  
  If cboState.ListIndex < 0 Then
    MsgBox "You must select a state before exporting"
    Exit Sub
  ElseIf ChangesMade Then
    SaveChanges
  End If
  
  i = 1
  With cdlgFileSel
    .DialogTitle = "Assign name of export file"
    If RDO = 0 Then
      FileName = GetSetting("SEE", "Defaults", "NSSExportFile")
    ElseIf RDO = 1 Then
      FileName = GetSetting("SEE", "Defaults", "LowFlowExportFile")
    End If
    If Len(Dir(FileName, vbDirectory)) <= 1 Then
      FileName = CurDir & "\" & DB.State.Abbrev & "_Export"
    Else
      FileName = FileName & "\" & DB.State.Abbrev & "_Export"
    End If
    If RDO = 0 Then
      FileName = FileName & "-PeakFlow"
    ElseIf RDO = 1 Then
      FileName = FileName & "-LowFlow"
    End If
    'Increment output file name if files already exported for state
    While Len(Dir(FileName & ".txt")) > 0
      i = i + 1
      If i > 2 Then FileName = Left(FileName, Len(FileName) - 2)
      FileName = FileName & "-" & i
    Wend
    .FileName = FileName
    .Filter = "(*.txt)|*.txt"
    .FilterIndex = 1
    .CancelError = True
    .ShowSave
    FileName = .FileName
    If RDO = 0 Then
      SaveSetting "SEE", "Defaults", "NSSExportFile", PathNameOnly(FileName)
      j = 0
    ElseIf RDO = 1 Then
      SaveSetting "SEE", "Defaults", "LowFlowExportFile", PathNameOnly(FileName)
      j = 1
    End If
  End With
  
  OutFile = FreeFile
  Open FileName For Output As OutFile
  Print #OutFile, DB.State.code & " " & DB.State.Name & " " & j
  If DB.State.Metric Then j = 1 Else j = 0
  Print #OutFile, lstRegions.ListCount & " " & j
  'Loop thru Regions
  For i = 1 To lstRegions.ListCount
    Set MyRegion = DB.State.Regions(lstRegions.List(i - 1))
    If MyRegion.ROIRegnID <> "-1" Then GoTo nextRegion
    If MyRegion.urban Then j = 1 Else j = 0
    str = MyRegion.Name & vbTab & j
    If MyRegion.UrbanNeedsRural Then j = 1 Else j = 0
    str = str & " " & j
    If MyRegion.PredInt Then j = 1 Else j = 0
    str = str & " " & j
    Print #OutFile, str
    'Loop thru Parameters
    For j = 1 To MyRegion.Parameters.Count
      Set MyParm = MyRegion.Parameters(j)
      'If MyParm.Abbrev <> "RDA" And MyParm.Abbrev <> "CRD" Then
        tmpCnt = MyRegion.depVars.Count
        str = ""
        If j = 1 Then
          str = MyRegion.Parameters.Count & " " & tmpCnt
        End If
        str = str & vbTab & MyParm.Abbrev & vbTab & MyParm.Name & vbTab & _
            MyParm.GetMin(False) & " " & MyParm.GetMax(False) & " " & MyParm.Units.id
        Print #OutFile, str
      'End If
    Next j
    'Loop thru Return Periods/Statistics
    For j = 1 To tmpCnt
      Set MyDepVar = MyRegion.depVars(j)
      compCnt = MyDepVar.Components.Count
      Print #OutFile, vbTab & MyDepVar.Name & " " & _
          Round(MyDepVar.StdErr, 1) & " "; Round(MyDepVar.EstErr, 1) & " " _
          ; Round(MyDepVar.PreErr, 1) & " " & Round(MyDepVar.EquivYears, 1) & " " & _
          MyDepVar.Constant & " " & MyDepVar.BCF & " " & _
          Round(MyDepVar.tdist, 4) & " " & Round(MyDepVar.Variance, 4) & _
          " " & Round(MyDepVar.ExpDA, 4) & " " & compCnt
      If MyRegion.PredInt Then covArray = MyDepVar.PopulateMatrix
      'Loop thru Components
      For k = 1 To compCnt
        Set MyComp = MyDepVar.Components(k)
        str = BldEqtn(MyComp)
        Print #OutFile, vbTab & vbTab & GetAbbrev(MyComp.ParmID) & vbTab & _
            MyComp.BaseMod & " " & MyComp.BaseCoeff & " " & _
            MyComp.BaseExp & " " & GetAbbrev(MyComp.expID) & vbTab & _
            MyComp.ExpMod & " " & MyComp.ExpExp & str
      Next k
      If MyRegion.PredInt Then  'using prediction intervals
        If UBound(covArray, 1) > 1 Then  'this Return/Stat has a covariance matrix
          'Loop thru Covariance Matrix
          For row = 1 To UBound(covArray, 1)
            str = vbTab & vbTab & vbTab
            For col = 1 To UBound(covArray, 2)
              str = str & " " & covArray(row, col)
            Next col
            Print #OutFile, str
          Next row
        End If
      End If
    Next j
nextRegion:
  Next i
  Close OutFile
x:
  Me.MousePointer = vbDefault
  MsgBox "Completed Export to file " & FileName, vbOKOnly, "SEE Export"
  If lstRegions.SelCount > 0 Then
    Set MyRegion = DB.State.Regions(lstRegions.List(lstRegions.ListIndex))
  Else
    Set MyRegion = Nothing
  End If
End Sub

Private Function BldEqtn(MyComp As nssComponent) As String
  Dim str
  
  With MyComp
    'Set base portion of equation
    str = vbTab & "#" & vbTab & .BaseCoeff & "(" & .BaseMod & "+" & _
          GetAbbrev(.ParmID) & ")^" & .BaseExp
    'Set exponent portion of equation, if applicable
    If GetAbbrev(.expID) <> "none" Then
      str = str & "(" & .ExpMod & "+" & GetAbbrev(.expID) & ")"
    End If
    'Set exponent of exponent, if applicable
    If .ExpExp <> 0 Then
      str = str & "^" & .ExpExp
    End If
    BldEqtn = str
  End With
End Function

Private Sub Form_Load()
  
  SetDB

End Sub

Private Sub grdInterval_RowColChange()
  Dim i&, j&, ipos&, retCnt&
  Dim returns As New FastCollection
  
  If DB Is Nothing Then Exit Sub
  With grdInterval
    .ClearValues
    If .col = 0 Then
      If RDO = 0 Then  'order then add return intervals to drop-down list
        retCnt = DB.returns.Count
'        ReDim returns(1 To retCnt)
        For i = 1 To retCnt
'          rank = 1
'          For j = 1 To retCnt
'            If IsNumeric(DB.returns(j).Name) Then
'              If CSng(DB.returns(i).Name) > CSng(DB.returns(j).Name) Then rank = rank + 1
'            End If
'          Next j
'          returns(rank) = DB.returns(i).Name
        
          ipos = 0
          If IsNumeric(DB.returns(i).Name) Then 'put it in sorted position
            j = 1
            While ipos = 0 And j <= returns.Count
              If IsNumeric(returns.ItemByIndex(j)) Then
                If CSng(DB.returns(i).Name) > CSng(returns.ItemByIndex(j)) Then
                  j = j + 1
                Else
                  ipos = j
                End If
              Else
                j = j + 1
              End If
            Wend
          End If
          If ipos > 0 Then 'insert in proper sorted position
            returns.Add DB.returns(i).Name, , ipos
          Else 'just add it to end
            returns.Add DB.returns(i).Name
          End If
        Next i
        For i = 1 To retCnt
          .addValue returns.ItemByIndex(i)
        Next i
      ElseIf RDO = 1 Then  'add statistics to drop-down list
        For i = 1 To DB.LFStats.Count
          .addValue DB.LFStats(i).Name
        Next i
      End If
      .ComboCheckValidValues = False
    ElseIf .col = .cols - 1 Then
      If Metric Then
        For i = 1 To DB.Units.Count
          .addValue DB.Units(i).MetricLabel
        Next i
      Else
        For i = 1 To DB.Units.Count
          .addValue DB.Units(i).EnglishLabel
        Next i
      End If
      .ComboCheckValidValues = True
    End If
  End With
End Sub

Private Sub grdMatrix_CommitChange(ChangeFromRow As Long, ChangeToRow As Long, ChangeFromCol As Long, ChangeToCol As Long)
  If ChangeToCol <> ChangeToRow - 1 Then
  grdMatrix.TextMatrix(ChangeToCol + 1, ChangeToRow - 1) = _
      grdMatrix.TextMatrix(ChangeToRow, ChangeToCol)
  End If
End Sub

Private Sub grdParms_CommitChange(ChangeFromRow As Long, ChangeToRow As Long, _
                                  ChangeFromCol As Long, ChangeToCol As Long)
  Dim i&
  'Make sure entries for 1st 2 columns are in sync
  With grdParms
    If ChangeFromCol = 0 Then
      'Clear the Code, Name, and Units fields when new Type selected
      For i = 1 To DB.StatisticTypes.Count
        If .TextMatrix(ChangeFromRow, 0) = DB.StatisticTypes(i).code Then
          If .TextMatrix(ChangeFromRow, 0) <> SelParms(ChangeFromRow).statTypeCode Then
            .TextMatrix(ChangeFromRow, 1) = ""
            .TextMatrix(ChangeFromRow, 2) = ""
            .TextMatrix(ChangeFromRow, 5) = ""
            Exit Sub
          End If
        End If
      Next i
    ElseIf ChangeFromCol = 1 Then
      For i = 1 To DB.Parameters.Count
        If .TextMatrix(ChangeFromRow, 1) = DB.Parameters(i).Name Then
          .TextMatrix(ChangeFromRow, 2) = DB.Parameters(i).Abbrev
          '.TextMatrix(ChangeFromRow, 5) = DB.Parameters(i).ConvFlag
          'use real unit names
          If Metric Then
            .TextMatrix(ChangeFromRow, 5) = DB.Units.ItemByKey(DB.Parameters(i).ConvFlag).MetricLabel
          Else
            .TextMatrix(ChangeFromRow, 5) = DB.Units.ItemByKey(DB.Parameters(i).ConvFlag).EnglishLabel
          End If
          Exit Sub
        End If
      Next i
    ElseIf ChangeFromCol = 2 Then
      For i = 1 To DB.Parameters.Count
        If .TextMatrix(ChangeFromRow, 2) = DB.Parameters(i).Abbrev Then
          .TextMatrix(ChangeFromRow, 1) = DB.Parameters(i).Name
          '.TextMatrix(ChangeFromRow, 5) = DB.Parameters(i).ConvFlag
          'use real unit names
          If Metric Then
            .TextMatrix(ChangeFromRow, 5) = DB.Units.ItemByKey(DB.Parameters(i).ConvFlag).MetricLabel
          Else
            .TextMatrix(ChangeFromRow, 5) = DB.Units.ItemByKey(DB.Parameters(i).ConvFlag).EnglishLabel
          End If
          Exit Sub
        End If
      Next i
    End If
  End With
End Sub

Private Sub grdParms_RowColChange()
  Dim i&
  Dim statTypeCode$
  
  If DB Is Nothing Then Exit Sub
  
  With grdParms
    'If .row = 0 Then Exit Sub
    'lblStatSel(1).Caption = .TextMatrix(.row, 2)
    .ClearValues
    'SizeGrid
    Select Case .col
      'Fill list of Statistic Types in first column
      Case 0:
        For i = 1 To DB.StatisticTypes.Count
          .addValue DB.StatisticTypes(i).code
        Next i
        .ComboCheckValidValues = True
      'Fill list of Stat Abbreviations in second column
      Case 1:
        If Len(Trim(.TextMatrix(.row, 0))) > 0 Then
          statTypeCode = .TextMatrix(.row, 0)
          For i = 1 To DB.StatisticTypes(statTypeCode).StatLabels.Count
            .addValue DB.StatisticTypes(statTypeCode).StatLabels(i).Name
          Next i
          .ComboCheckValidValues = True
        End If
      'Fill list of Statistic Names in third column
      Case 2:
        If Len(Trim(.TextMatrix(.row, 0))) > 0 Then
          statTypeCode = .TextMatrix(.row, 0)
          For i = 1 To DB.StatisticTypes(statTypeCode).StatLabels.Count
            .addValue DB.StatisticTypes(statTypeCode).StatLabels(i).code
          Next i
          .ComboCheckValidValues = True
        End If
      Case 5:
        If Metric Then
          For i = 1 To DB.Units.Count
            .addValue DB.Units(i).MetricLabel
          Next i
        Else
          For i = 1 To DB.Units.Count
            .addValue DB.Units(i).EnglishLabel
          Next i
        End If
        .ComboCheckValidValues = True
    End Select
  End With
End Sub

Private Sub rdoMainOpt_Click(Index As Integer)
  Dim stIndex&, selState&, i&
  
  If NotNew Then Exit Sub
  On Error GoTo x

  Set MyRegion = Nothing
  
  'Clear previous selections and hide frames
  lstParms.Clear
  lstRetPds.Clear
  For i = 0 To 2
    fraEdit(i).Visible = False
  Next i

  'Populate state listbox
  State = GetSetting("SEE", "Defaults", "StateName")
  cboState.Clear
  For stIndex = 0 To DB.States.Count - 1
    cboState.List(stIndex) = DB.States(stIndex + 1).Name
    cboState.ItemData(stIndex) = DB.States(stIndex + 1).code
    If DB.States(stIndex + 1).Name = State Then selState = stIndex
  Next
  If selState > 0 Then
    Set DB.State = DB.States(selState + 1)
    cboState.ListIndex = selState
  Else
    lstRegions.Clear
    lstParms.Clear
    lstRetPds.Clear
  End If
  NotNew = False
  RDO = Index
  If RDO = 0 Then
    lblRetPds.Caption = "Return Periods:"
  ElseIf RDO = 1 Then
    lblRetPds.Caption = vbCrLf & "Statistics:"
  End If
  FocusOnRegions
  Exit Sub
x:
  If RDO > -1 Then
    NotNew = True
    rdoMainOpt(RDO) = True
    NotNew = False
  Else
    For i = 0 To rdoMainOpt.Count - 1
      rdoMainOpt(i) = False
    Next i
  End If
End Sub

Private Function DBCheck(dbName As String) As Boolean
  If Len(Dir(DBPath)) = 0 Then
    MsgBox "'" & dbName & "' was not found." & vbCrLf & _
        "Please select another database.", vbCritical, "File not found"
    Exit Function
  ElseIf DB.States.Count = 0 Then
    MsgBox "'" & dbName & "' does not contain a proper table of state names." & _
        vbCrLf & "Please select another database.", vbCritical, "File not found"
    Exit Function
  End If
  If rdoMainOpt(0) Then
    If DB.StationTypes.Count = 0 Then
      MsgBox "'" & dbName & "' is not an NSS database." & vbCrLf & _
          "Please select another database.", vbCritical, "Wrong database"
      Exit Function
    End If
  ElseIf rdoMainOpt(1) Then
    If DB.LFStats.Count = 0 Then
      MsgBox "'" & dbName & "' is not a LowFlow database." & _
          vbCrLf & "Please select another database.", vbCritical, "Wrong database"
      Exit Function
    End If
  End If
  DBCheck = True
End Function

Private Sub cboState_Click()
  Dim stID$
  Dim regnIndex&, i&, regnCount&

  stID = CStr(cboState.ItemData(cboState.ListIndex))
  If Len(stID) = 1 Then stID = "0" & stID
  If DB.States.IndexFromKey(stID) > 0 Then
    Set DB.State = DB.States(cboState.ListIndex + 1)
    lstRegions.Clear
    lstParms.Clear
    lstRetPds.Clear
    fraEdit(0).Visible = False
    fraEdit(1).Visible = False
    fraEdit(2).Visible = False
    If Not DB.State.Regions Is Nothing Then
      DB.State.Regions.Clear
      Set DB.State.Regions = Nothing
    End If
    'Remove Regions from collection if not right type
    regnCount = DB.State.Regions.Count
    i = 0
    For regnIndex = 1 To regnCount
      If DB.State.Regions(regnIndex - i).ROIRegnID > 0 Or _
          rdoMainOpt(0) And DB.State.Regions(regnIndex - i).LowFlowRegnID > 0 Or _
          rdoMainOpt(1) And DB.State.Regions(regnIndex - i).LowFlowRegnID < 0 Then
        DB.State.Regions.RemoveByIndex (regnIndex - i)
        i = i + 1
      End If
    Next regnIndex
    'set metric button
    If DB.State.Metric Then
      rdoUnits(1).Value = True
    Else
      rdoUnits(0).Value = True
    End If
    'Populate Regions list box
    regnCount = DB.State.Regions.Count
    i = 0
    For regnIndex = 1 To regnCount
      i = i + 1
      lstRegions.List(i - 1) = DB.State.Regions(regnIndex).Name
      lstRegions.ItemData(i - 1) = regnIndex
    Next
    State = DB.State.Name
    SaveSetting "SEE", "Defaults", "StateName", State
    cmdExport.Enabled = True
    cmdImport.Enabled = True
  End If
End Sub

Private Sub grdComps_RowColChange()
  Dim i&
  Dim EqtnStr As String, BaseVar As String, BaseStr As String
  Dim ExpStr As String
  Dim BaseMod As Single, ExpMod As Single
  
  With grdComps
    .ClearValues
    If .col = 0 Or .col = 4 Then
      For i = 0 To lstParms.ListCount - 1
        If Not (.col = 4 And lstParms.ItemData(i) < 0) Then _
            .addValue GetAbbrev(lstParms.ItemData(i))
      Next i
      If .col = 4 Then
        For i = 0 To lstParms.ListCount - 1
          If lstParms.ItemData(i) >= 0 Then _
            .addValue "log(" & GetAbbrev(lstParms.ItemData(i)) & ")"
        Next i
      End If
      .addValue "none"
    End If

    'update equation display
    If Len(grdInterval.TextMatrix(1, 5)) > 0 And grdInterval.TextMatrix(1, 5) <> "1" Then
      EqtnStr = lstRetPds.List(lstRetPds.ListIndex) & " = " & grdInterval.TextMatrix(1, 5)
    Else
      EqtnStr = lstRetPds.List(lstRetPds.ListIndex) & " ="
    End If
    For i = 1 To .Rows
      If Len(.TextMatrix(i, 0)) > 0 Then
        If .TextMatrix(i, 0) = "none" Then
          BaseVar = ""
        Else
          BaseVar = .TextMatrix(i, 0)
        End If
        If Len(.TextMatrix(i, 1)) > 0 Then
          BaseMod = CSng(.TextMatrix(i, 1))
        Else
          BaseMod = 0
        End If
        Select Case BaseMod
          Case Is > 0: BaseStr = "(" & BaseVar & "+" & BaseMod & ")"
          Case Is < 0: BaseStr = "(" & BaseVar & "-" & BaseMod & ")"
          Case Else: BaseStr = "(" & BaseVar & ")"
        End Select
        If Len(.TextMatrix(i, 2)) > 0 And .TextMatrix(i, 2) <> "1" Then
          BaseStr = "(" & .TextMatrix(i, 2) & BaseStr & ")"
        End If
      End If
      EqtnStr = EqtnStr & "  " & BaseStr
      If Len(.TextMatrix(i, 3)) > 0 And .TextMatrix(i, 3) <> "1" Then
        EqtnStr = EqtnStr & "^" & .TextMatrix(i, 3)
      End If
      If Len(.TextMatrix(i, 4)) > 0 Then
        If .TextMatrix(i, 4) = "none" Then
          ExpStr = ""
        Else
          ExpStr = .TextMatrix(i, 4)
          If Len(.TextMatrix(i, 5)) > 0 Then
            ExpMod = CSng(.TextMatrix(i, 5))
          Else
            ExpMod = 0
          End If
          Select Case ExpMod
            Case Is > 0: ExpStr = "^" & "(" & ExpStr & "+" & ExpMod & ")"
            Case Is < 0: ExpStr = "^" & "(" & ExpStr & "-" & ExpMod & ")"
            Case Else: ExpStr = "^" & "(" & ExpStr & ")"
          End Select
          If Len(.TextMatrix(i, 6)) > 0 And .TextMatrix(i, 6) <> "1" Then
            ExpStr = ExpStr & "^" & .TextMatrix(i, 6)
          End If
        End If
        EqtnStr = EqtnStr & ExpStr
      End If
    Next i
    lblEquation.Caption = EqtnStr
  End With
End Sub

'Private Sub grdComps_TextChange(ChangeFromRow As Long, ChangeToRow As Long, _
'                                ChangeFromCol As Long, ChangeToCol As Long)
'  Dim col&
'  With grdComps
'    If ChangeToRow <> .Rows Then Exit Sub
'    For col = 0 To .cols - 1
'      If Len(Trim(.TextMatrix(.Rows, col))) = 0 Then Exit Sub
'    Next col
'    .Rows = .Rows + 1
'        grdMatrix.Rows = .Rows
'        grdMatrix.cols = .Rows
'  End With
'  With grdMatrix
'    For col = 0 To .cols - 1
'      .colWidth(col) = 800
'      .ColEditable(col) = True
'    Next col
'  End With
'End Sub

Private Sub lstParms_GotFocus()
  Dim i&
  
  If MyRegion Is Nothing Then Exit Sub
  
  If SelParms Is Nothing Then Set SelParms = New FastCollection
  If lstParms.ListCount > 0 Then
    If lstParms.Selected(lstParms.ListIndex) Then  'possible selection made
      For i = 1 To SelParms.Count
        If SelParms(i).Name = lstParms.List(lstParms.ListIndex) Then Exit For
      Next i
      If i > SelParms.Count Then
        lstParms_Click  'selecting a parm
      Else
        If ChangesMade Then SaveChanges
        FocusOnParms
      End If
'    Else  'possible deselection made
'      For i = 1 To SelParms.Count
'        If SelParms(i).Name = lstParms.List(lstParms.ListIndex) Then Exit For
'      Next i
'      If i <= SelParms.Count Then
'        lstParms_Click  'deselecting a parm
'      Else
'        If ChangesMade Then SaveChanges
'        FocusOnParms
'      End If
    End If
  Else
    If ChangesMade Then SaveChanges
    FocusOnParms
  End If
End Sub

Private Sub lstParms_Click()
  Dim i&, j&, thisIndex&
  Dim theseParms() As String
  
  If MyRegion Is Nothing Then Exit Sub
  
  If Not (ChoseReturns Or fraEdit(1).Visible) Then
    If ChangesMade Then
      j = -1
      ReDim theseParms(0)
      For i = 0 To lstParms.ListCount - 1
        If lstParms.Selected(i) Then
          j = j + 1
          If j > 0 Then ReDim Preserve theseParms(j)
          theseParms(j) = lstParms.List(i)
        End If
      Next i
      SaveChanges
      If lstParms.SelCount = 0 Then
        For i = 0 To UBound(theseParms)
          For j = 0 To lstParms.ListCount - 1
            If lstParms.List(j) = theseParms(i) Then lstParms.Selected(j) = True
          Next j
        Next i
        Exit Sub
      End If
    End If
  End If
  
  Me.MousePointer = vbHourglass
  thisIndex = lstParms.ListIndex
  If SelParms Is Nothing Then Set SelParms = New FastCollection
  If lstParms.Selected(thisIndex) Then  'adding selection
    If SelParms.IndexFromKey(CStr(MyRegion.Parameters(thisIndex + 1).id)) = -1 Then
      SelParms.Add MyRegion.Parameters(thisIndex + 1), _
          CStr(MyRegion.Parameters(thisIndex + 1).id)
    End If
  Else  'removing selection
    For i = 1 To grdParms.Rows
      If lstParms.List(thisIndex) = grdParms.TextMatrix(i, 1) Then Exit For
    Next i
    If i <= grdParms.Rows Then
      SelParms.RemoveByIndex i
      With grdParms
        If .Rows = 1 Then
          .Rows = 0
        Else
          .DeleteRow i
        End If
      End With
    End If
  End If
  If SelParms.Count > 0 Then
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = True
  Else
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = False
  End If
  If Not ChoseReturns Then FocusOnParms
  Me.MousePointer = vbDefault
End Sub

Private Sub FocusOnParms()
  Dim i&
  
  ChoseParms = True
  For i = 0 To lstRetPds.ListCount - 1
    lstRetPds.Selected(i) = False
  Next i
  fraEdit(0).Visible = False
  fraEdit(1).Visible = True
  fraEdit(1).Caption = "Parameter Values"
  fraEdit(2).Visible = False
  SetGrid "Parameters"
  ChoseParms = False
End Sub

Private Sub lstRegions_Click()
  
  fraEdit(0).Enabled = True
  txtRegName.BackColor = &H80000005
  fraRegType.Enabled = True
  lblRegName.Enabled = True
  If Not MyRegion Is Nothing Then
    If MyRegion.IsNew Then
      lstRegions.Selected(lstRegions.ListCount - 1) = True
      Exit Sub
    End If
  End If
  If lstRegions.ListIndex >= 0 Then
    Set MyRegion = DB.State.Regions(lstRegions.List(lstRegions.ListIndex))
    MyRegion.IsNew = False
    ResetRegion
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = True
  Else  'no region is selected
    Set MyRegion = Nothing
    fraEdit(0).Visible = True
    fraEdit(1).Visible = False
    fraEdit(2).Visible = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = False
    txtRegName.Text = ""
    rdoRegOpt(0) = False
    rdoRegOpt(1) = False
    chkPredInt.Value = 0
  End If
End Sub

Private Sub lstRegions_GotFocus()

  If Not MyRegion Is Nothing Then
    If ChangesMade Then SaveChanges
  End If
  
  With lstRegions
    If .ListCount > 0 Then
'      If Not MyRegion Is Nothing Then
'        If .List(.ListIndex) = MyRegion.Name Then Exit Sub
'      End If
      If .ListIndex > -1 Then
        lstRegions_Click  'selecting a new region
      Else
        FocusOnRegions
      End If
    Else
      FocusOnRegions
    End If
  End With
End Sub

Private Sub FocusOnRegions()
  Dim i&
  
  For i = 0 To 1
    rdoRegOpt(i) = False
  Next i
  chkRuralInput.Value = 0
  chkRuralInput.Visible = False
  chkPredInt.Value = 0
  txtRegName = ""
  fraEdit(0).Visible = True
  fraEdit(0).Enabled = False
  txtRegName.BackColor = &H8000000F
  fraRegType.Enabled = False
  lblRegName.Enabled = False
End Sub

Private Sub ResetDB()
  DB.DB.Close
  
  Set DB = Nothing
  Set DB = New nssDatabase
  DB.FileName = DBPath
  Set DB.State = DB.States.ItemByKey(CStr(cboState.ItemData(cboState.ListIndex)))
  DB.State.Regions.Clear
  Set DB.State.Regions = Nothing
  If fraEdit(0).Visible Then
    Set MyRegion = DB.State.Regions(NoSpaces(lstRegions.List(lstRegions.ListIndex)))
  Else
    Set MyRegion = DB.State.Regions(MyRegion.Name)  'Resets same region with new collections
  End If
End Sub

Private Sub ResetRegion()
  Dim i&
  
  fraEdit(0).Visible = True
  fraEdit(1).Visible = False
  fraEdit(2).Visible = False
  If Not MyRegion.depVars Is Nothing Then
    MyRegion.depVars.Clear
    Set MyRegion.depVars = Nothing
  End If
  If Not MyRegion.Parameters Is Nothing Then
    MyRegion.Parameters.Clear
    Set MyRegion.Parameters = Nothing
  End If
  txtRegName.Text = MyRegion.Name
  If MyRegion.urban Then
    rdoRegOpt(1) = True
    chkRuralInput.Visible = True
    If MyRegion.UrbanNeedsRural Then
      chkRuralInput.Value = 1
    Else
      chkRuralInput.Value = 0
    End If
  Else
    rdoRegOpt(0) = True
  End If
  If MyRegion.PredInt Then chkPredInt.Value = 1 Else chkPredInt.Value = 0
  cmdDelete.Enabled = True
  cmdSave.Enabled = True
  cmdCancel.Enabled = True
  PopulateParms
  PopulateDepVars
End Sub

Private Sub cmdAdd_Click()
  Dim i&, j&, mDim&
  Dim str$
  
  If fraEdit(0).Visible Then
    If Not MyRegion Is Nothing Then
      If MyRegion.IsNew Then Exit Sub
    End If
    For i = 0 To 1
      rdoRegOpt(i) = False
    Next i
    chkRuralInput.Value = 0
    chkRuralInput.Visible = False
    chkPredInt.Value = 0
    Set MyRegion = New nssRegion
    MyRegion.IsNew = True
    Set MyRegion.DB = DB
    Set MyRegion.State = DB.State
    DB.State.Regions.Add MyRegion, "0"
    lstRegions.AddItem "New Region"
    lstRegions.Selected(lstRegions.ListCount - 1) = True
    txtRegName.Text = "New Region"
    txtRegName.SetFocus
    lstParms.Clear
    If Not SelParms Is Nothing Then
      SelParms.Clear
      Set SelParms = Nothing
    End If
    lstRetPds.Clear
  ElseIf fraEdit(1).Visible Then
    If Not MyParm Is Nothing Then
      If MyParm.IsNew And grdParms.Rows > 0 Then Exit Sub
    End If
    Set MyParm = New nssParameter
    MyParm.IsNew = True
    Set MyParm.Region = MyRegion
    MyRegion.Parameters.Add MyParm, "0"
    lstParms.AddItem "New Parameter"
    lstParms.Selected(lstParms.ListCount - 1) = True
  ElseIf fraEdit(2).Visible Then
    If grdInterval.Rows > 0 Then
      If MyDepVar.IsNew Then Exit Sub
      Skip = True
      For i = 0 To lstRetPds.ListCount - 1
        lstRetPds.Selected(i) = False
      Next i
      Skip = False
    End If
    Set MyDepVar = New nssDepVar
    MyDepVar.IsNew = True
    Set MyDepVar.Region = MyRegion
    If MyRegion.depVars.IndexFromKey("0") < 0 Then
      lstRetPds.AddItem "New"
    End If
    ChoseParms = True  'workaround to avoid triggering selection event
    lstRetPds.Selected(lstRetPds.ListCount - 1) = True
    ChoseParms = False
    With grdInterval
      .ClearData
      .Rows = 1
      .col = 2  'avoids conflict with change event that creates listboxes
    End With
    With grdComps
      .ClearData
      .Rows = 1
      .cols = 7
      .TextMatrix(1, 1) = "0"
      .TextMatrix(1, 2) = "1"
      .TextMatrix(1, 4) = "none"
      .TextMatrix(1, 5) = "0"
      .TextMatrix(1, 6) = "0"
      .col = 2  'avoids conflict with change event that creates listboxes
    End With
    With grdMatrix
      .ClearData
      .Rows = 2
      .cols = 2
      For i = 0 To .cols - 1
        .colWidth(i) = 800
        .ColEditable(i) = True
      Next i
    End With
  End If
  cmdSave.Enabled = True
  cmdDelete.Enabled = True
  cmdCancel.Enabled = True
End Sub

Private Sub cmdDelete_Click()
  Dim i&, j&, response&
  Dim tmpDepVar As nssDepVar
  Dim tmpComp As nssComponent
  
  If MyRegion Is Nothing Then Exit Sub
  
  If (fraEdit(0).Visible And lstRegions.SelCount = 0) Or _
     (fraEdit(1).Visible And lstParms.SelCount = 0) Or _
     (fraEdit(2).Visible And lstRetPds.SelCount = 0) Then Exit Sub
  If fraEdit(0).Visible Then
    If MyRegion.IsNew Then
      cmdCancel_Click
    Else
      response = myMsgBox.Show("Are you certain you want to delete the region" & _
          vbCrLf & MyRegion.Name & " from the database for " & State & "?", _
          "User Action Verification", "+&Yes", "-&Cancel")
      If response = 1 Then
        For i = 1 To MyRegion.Parameters.Count
          Set MyParm = MyRegion.Parameters(i)
          MyParm.Delete
        Next i
        For i = 1 To MyRegion.depVars.Count
          Set MyDepVar = MyRegion.depVars(i)
          For j = 1 To MyDepVar.Components.Count
            Set MyComp = MyDepVar.Components(j)
            MyComp.Delete
          Next j
          MyDepVar.Delete
        Next i
        MyRegion.Delete
        DB.State.Regions.RemoveByKey (MyRegion.Name)
        lstRegions.RemoveItem (lstRegions.ListIndex)
        Set MyRegion = Nothing
        lstParms.Clear
        lstRetPds.Clear
        fraEdit(0).Visible = False
      End If
    End If
  ElseIf fraEdit(1).Visible Then
    response = myMsgBox.Show("Are you certain you want to delete the selected" & _
        vbCrLf & "Parameters from the database for " & State & "?", _
        "User Action Verification", "+&Yes", "-&Cancel")
    If response = 1 Then
      CheckRuralInput  'to make sure we don't delete RDA or CRD
      For i = 1 To SelParms.Count
        For Each tmpDepVar In MyRegion.depVars
          For Each tmpComp In tmpDepVar.Components
            If tmpComp.ParmID = SelParms(i).id Or Abs(tmpComp.expID) = SelParms(i).id Then tmpComp.Delete
          Next tmpComp
        Next tmpDepVar
        Set SelParms(i).Region.DB = DB
        SelParms(i).Delete
        lstParms.RemoveItem (j)
      Next i
      ResetDB
      grdParms.ClearData
      grdParms.Rows = 0
    End If
  ElseIf fraEdit(2).Visible Then
    If MyDepVar.IsNew Then
      cmdCancel_Click
    Else
      If RDO = 0 Then
        response = myMsgBox.Show("Are you certain you want to delete the " & _
            lstRetPds.List(lstRetPds.ListIndex) & "-year" & vbCrLf & _
            "Return Period from the database for " & MyRegion.Name & ", " & State & "?", _
            "User Action Verification", "+&Yes", "-&Cancel")
      ElseIf RDO = 1 Then
        response = myMsgBox.Show("Are you certain you want to delete the " & _
            lstRetPds.List(lstRetPds.ListIndex) & " statistic" & vbCrLf & _
            "from the database for " & MyRegion.Name & ", " & State & "?", _
            "User Action Verification", "+&Yes", "-&Cancel")
      End If
      If response = 1 Then
        j = lstRetPds.ItemData(lstRetPds.ListIndex)
        Set MyDepVar = MyRegion.depVars(CStr(j))
        For i = 1 To MyDepVar.Components.Count
          Set MyComp = MyDepVar.Components(i)
          MyComp.Delete
        Next i
        MyDepVar.Delete
        lstRetPds.RemoveItem lstRetPds.ListIndex
        ResetDB
        ResetRegion
      End If
    End If
  End If
End Sub

Private Sub cmdSave_Click()
  Dim i&, j&, k&, response&, baseID&, expID&, tmpID&
  Dim str$, depVarName$
  Dim covArray() As String, parmNames() As String
  Dim isReturn As Boolean
  Dim BCF As String, tdist As String, Variance As String, ExpDA As String
  Dim UnitID As Long
  
  On Error GoTo 0
  
  If RDO = 0 Then
    isReturn = True
  ElseIf RDO = 1 Then
    isReturn = False
  End If

  'QA check
  If fraEdit(0).Visible Then
    If Len(Trim(txtRegName)) = 0 Then
      MsgBox "You must enter a name for the region."
      Exit Sub
    Else
      For i = 0 To rdoRegOpt.Count - 1
        If rdoRegOpt(i) Then Exit For
      Next i
      If i = rdoRegOpt.Count Then
        MsgBox "You must select the type of region."
        Exit Sub
      End If
    End If
  ElseIf fraEdit(1).Visible Then
    For i = 1 To grdParms.Rows
      For j = 1 To grdParms.cols - 1
        If grdParms.TextMatrix(i, j) = "" Then
          If j = 1 Then  'no name entered for parameter
            MsgBox "You must enter a name for the parameter in row " & _
            i & " of the grid."
          Else
            MsgBox "You must enter a value in the " & _
                grdParms.TextMatrix(0, j) & " field for the parameter " & _
                grdParms.TextMatrix(i, 1)
          End If
          Exit Sub
        End If
      Next j
    Next i
  ElseIf fraEdit(2).Visible Then
    'Make sure user has entered values for each field in both grids
    For i = 0 To grdInterval.cols - 1
      If grdInterval.TextMatrix(1, i) = "" Then
        MsgBox "You must enter a value for the field '" & _
            grdInterval.TextMatrix(-1, i) & " " & grdInterval.TextMatrix(0, i) & _
            "' for the Return Interval."
        Exit Sub
      End If
    Next i
    For i = 1 To grdComps.Rows - 1
      For j = 0 To grdComps.cols - 1
        If grdComps.TextMatrix(i, j) = "" Then
          MsgBox "You must enter a value for the field '" & grdComps.TextMatrix(0, j) & vbCrLf _
                  & "' in row " & i & " of the grid.", , "Missing Data"
          Exit Sub
        End If
      Next j
    Next i
  End If
  
  'Check for changes and write them to an array
  If Not ChangesMade Then
    GoTo x
  End If
  
  'Make sure user wants to overwrite existing values
  If Skip Then
    response = 1
  Else
    str = "Are you certain you want to "
    If fraEdit(0).Visible Then  'editing Region
      If MyRegion.IsNew Then
        str = str & "add a new Region to the database for " & State & "?"
      Else
        str = str & "overwrite the existing values for " & MyRegion.Name & "?"
      End If
    ElseIf fraEdit(1).Visible Then  'editing Parameter
      str = str & "edit the Parameter values in the database for " & MyRegion.Name & "?"
    ElseIf fraEdit(2).Visible Then  'editing Return Period / Statistic
      If RDO = 0 Then
        depVarName = "Return Period"
      ElseIf RDO = 1 Then
        depVarName = "Statistic"
      End If
      If MyDepVar.IsNew Then
        str = str & "add a new " & depVarName & " to the database for " & MyRegion.Name & "?"
      Else
        str = str & "overwrite the existing values for this " & depVarName & " in " & MyRegion.Name & "?"
      End If
    End If
    response = myMsgBox.Show(str, "User Action Verification", "+&Yes", "-&Cancel")
  End If
  If response = 1 Then
  'Overwrite values in DB
    frmUserInfo.Show vbModal, Me
    If Not UserInfoOK Then GoTo x
    Me.MousePointer = vbHourglass
    If fraEdit(0).Visible Then 'editing region
      If MyRegion.IsNew Then
        Set MyRegion.DB = DB
        If Not MyRegion.Add(isReturn, txtRegName.Text, rdoRegOpt(1), _
            chkRuralInput.Value, chkPredInt.Value, -1) Then GoTo x
'?????? add possible 2 parms to DB?
        ResetDB
        lstRegions.ListIndex = lstRegions.ListCount - 1
      Else
        MyRegion.Edit txtRegName.Text, rdoRegOpt(1), _
            chkRuralInput.Value, chkPredInt.Value
        lstRegions.List(lstRegions.ListIndex) = txtRegName.Text
        ResetDB
      End If
      'Write changes to DetailedLog table
      For i = 0 To UBound(Changes, 2)
        If Len(Changes(1, i)) > 0 Then
          MyRegion.DB.RecordChanges TransID, "Regions", i + 2, _
              CStr(MyRegion.id), Changes(0, i), Changes(1, i)
        End If
      Next i
      ResetRegion
    ElseIf fraEdit(1).Visible Then 'editing parameter(s)
      ReDim parmNames(1 To grdParms.Rows)
      For i = 1 To grdParms.Rows
        parmNames(i) = grdParms.TextMatrix(i, 1)
        UnitID = UnitIDFromLabel(grdParms.TextMatrix(i, 5), DB.Units)
        Set MyParm = SelParms(i)
        If MyParm.IsNew Then
          With grdParms
            If Not MyParm.Add(MyRegion, .TextMatrix(i, 2), _
                .TextMatrix(i, 3), .TextMatrix(i, 4), UnitID) Then GoTo x
          End With
          'Write changes to DetailedLog table
          For k = 0 To UBound(Changes, 3)
            If Len(Changes(1, i - 1, k)) > 0 Then
              MyRegion.DB.RecordChanges TransID, "Parameters", k + 2, _
                  CStr(MyParm.id), Changes(0, i - 1, k), Changes(1, i - 1, k)
            End If
          Next k
        Else
          MyParm.Edit grdParms.TextMatrix(i, 2), grdParms.TextMatrix(i, 3), _
              grdParms.TextMatrix(i, 4), UnitID
          'Write changes to DetailedLog table
          For k = 0 To UBound(Changes, 3)
            If Len(Changes(1, i - 1, k)) > 0 Then
              MyRegion.DB.RecordChanges TransID, "Parameters", k + 2, _
                  CStr(MyParm.id), Changes(0, i - 1, k), Changes(1, i - 1, k)
            End If
          Next k
        End If
      Next i
      ResetDB
      ResetRegion
      fraEdit(0).Visible = False
      fraEdit(1).Visible = True
      fraEdit(2).Visible = False
      grdParms.ClearData
      grdParms.Rows = 0
      For i = 1 To UBound(parmNames)
        For k = 1 To lstParms.ListCount
          If parmNames(i) = lstParms.List(k - 1) Then
            lstParms.Selected(k - 1) = True
            Exit For
          End If
        Next k
      Next i
      Set MyParm = SelParms(CStr(SelParms(UBound(parmNames)).id))
    ElseIf fraEdit(2).Visible Then 'editing return period/statistic, components, and matrix
      If MyRegion.PredInt Then
        BCF = grdInterval.TextMatrix(1, 6)
        tdist = grdInterval.TextMatrix(1, 7)
        Variance = grdInterval.TextMatrix(1, 8)
        ExpDA = grdInterval.TextMatrix(1, 9)
      Else
        BCF = ""
        tdist = ""
        Variance = ""
        ExpDA = grdInterval.TextMatrix(1, 6)
      End If
      If MyDepVar.IsNew Then
        'Add new Return or Statistic
        If lstRetPds.List(lstRetPds.ListIndex) = "New" Then
          depVarName = grdInterval.TextMatrix(1, 0)
        Else
          depVarName = lstRetPds.List(lstRetPds.ListIndex)
        End If
        Set MyDepVar = New nssDepVar
        tmpID = MyDepVar.Add(isReturn, MyRegion, grdInterval.TextMatrix(1, 0), _
            grdInterval.TextMatrix(1, 1), grdInterval.TextMatrix(1, 2), _
            grdInterval.TextMatrix(1, 3), grdInterval.TextMatrix(1, 4), _
            grdInterval.TextMatrix(1, 5), BCF, tdist, Variance, ExpDA)
        If tmpID = -1 Then GoTo x
        ResetDB
        lstRetPds.Clear
        PopulateDepVars
        MyDepVar.IsNew = True
      Else
        'Edit existing Return or Statistic
        MyDepVar.Edit grdInterval.TextMatrix(1, 0), grdInterval.TextMatrix(1, 1), _
            grdInterval.TextMatrix(1, 2), grdInterval.TextMatrix(1, 3), grdInterval.TextMatrix(1, 4), grdInterval.TextMatrix(1, 5), BCF, tdist, Variance, ExpDA
        MyDepVar.ClearOldComponents
        ResetDB
        tmpID = MyDepVar.id
      End If
      'Add values in Components grid and Covariance grid to DB
      Set MyComp = New nssComponent
      ReDim covArray(1 To grdMatrix.Rows, 1 To grdMatrix.cols)
      For i = 1 To grdComps.Rows
        If i <= grdComps.Rows Then
          baseID = GetCode(grdComps.TextMatrix(i, 0))
          expID = GetCode(grdComps.TextMatrix(i, 4))
          'Add components from grid to DB
          MyComp.Add MyRegion, tmpID, baseID, grdComps.TextMatrix(i, 1), _
              grdComps.TextMatrix(i, 2), grdComps.TextMatrix(i, 3), expID, _
              grdComps.TextMatrix(i, 5), grdComps.TextMatrix(i, 6)
          'Write Component changes to DetailedLog table
          str = CStr(tmpID) & " " & CStr(baseID) & " " & CStr(expID)
          For k = 0 To UBound(CompChanges, 3)
            If Len(CompChanges(1, i - 1, k)) > 0 Then
              MyRegion.DB.RecordChanges TransID, "Components", k + 1, _
                  str, CompChanges(0, i - 1, k), CompChanges(1, i - 1, k)
            End If
          Next k
        End If
        If MyRegion.PredInt Then
          'Add matrix values from 'i'th row of grid to DB
          For k = 1 To grdMatrix.cols
            covArray(i, k) = grdMatrix.TextMatrix(i, k - 1)
            If i = grdComps.Rows Then 'include final row of matrix
              covArray(i + 1, k) = grdMatrix.TextMatrix(i + 1, k - 1)
            End If
            'Write Covariance Matrix changes to DetailedLog table
            If Len(MatrixChanges(1, i, k)) > 0 Then
              MyRegion.DB.RecordChanges TransID, "Covariance", 3, CStr(tmpID) & " " & _
                   i & " " & k, MatrixChanges(0, i, k), MatrixChanges(1, i, k)
            End If
          Next k
        End If
      Next i
      If MyRegion.PredInt Then MyDepVar.AddMatrix MyRegion, tmpID, covArray()
      ResetDB
      i = 0
      'always reselect the return period to update the depvars/components
'      If MyDepVar.IsNew Then
        For i = 1 To lstRetPds.ListCount
          If lstRetPds.List(i - 1) = depVarName Then
            ChoseParms = True
            lstRetPds.Selected(i - 1) = True
            ChoseParms = False
          End If
        Next i
        Set MyDepVar = MyRegion.depVars(CStr(lstRetPds.ItemData(lstRetPds.ListIndex)))
'      End If
      'Write Return Period/Statistic changes to DetailedLog table
      For k = 0 To UBound(Changes, 2)
        If Len(Changes(1, k)) > 0 Then
          MyRegion.DB.RecordChanges TransID, "DepVars", k + 2, _
              CStr(tmpID), Changes(0, k), Changes(1, k)
        End If
      Next k
      fraEdit(0).Visible = False
      fraEdit(1).Visible = False
      fraEdit(2).Visible = True
    End If
  End If
x:
  Me.MousePointer = vbDefault
End Sub

Private Sub SetGrid(Table As String)
  Dim i&, row&, col&, compCnt&
  Dim tmp As Single, AddSpace As Single
  Dim str$
  Dim covArray() As String
  Dim newParm As Boolean
  
  i = -1
  Select Case Table
    Case "Parameters":
      With grdParms
        If SelParms Is Nothing Then Set SelParms = New FastCollection
        .Rows = SelParms.Count
        .cols = 6
        .ColType(0) = ATCoTxt
        .TextMatrix(-1, 0) = "Parm"
        .TextMatrix(0, 0) = "Type"
        .colWidth(0) = 900
        .ColType(1) = ATCoTxt
        .TextMatrix(-1, 1) = ""
        .TextMatrix(0, 1) = "Parameter"
        .colWidth(1) = 3170
        .ColType(2) = ATCoTxt
        .TextMatrix(-1, 2) = ""
        .TextMatrix(0, 2) = "Abbreviation"
        .colWidth(2) = 1300
        .ColType(3) = ATCoSng
        .TextMatrix(-1, 3) = "Min"
        .TextMatrix(0, 3) = "Value"
        .colWidth(3) = 800
        .ColType(4) = ATCoSng
        .TextMatrix(-1, 4) = "Max"
        .TextMatrix(0, 4) = "Value"
        .colWidth(4) = 800
        '.ColType(5) = ATCoInt
        .CausesValidation = False
        .TextMatrix(-1, 5) = "Conversion"
        .TextMatrix(0, 5) = "Flag"
        .colWidth(5) = 1100
        '.CellBackColor = &H80000005
        For row = 1 To SelParms.Count
          For col = 0 To .cols - 1
            .row = row
            .col = col
            'If col < .cols - 1 Then .ColEditable(col) = True Else .ColEditable(col) = False
            'all columns, including units, now editable, PRH 2/2005
            .ColEditable(col) = True
            If Not SelParms(row).IsNew Then
              newParm = True
              Select Case col
                Case 0: .Text = SelParms(row).statTypeCode
                Case 1: .Text = SelParms(row).Name
                Case 2: .Text = SelParms(row).Abbrev
                Case 3: .Text = SelParms(row).GetMin(False)
                Case 4: .Text = SelParms(row).GetMax(False)
                'Case 5: .Text = CInt(SelParms(row).Units.id)
                Case 5:
                  If rdoUnits(0).Value Then
                    .Text = DB.Units.ItemByKey(SelParms(row).Units.id).EnglishLabel
                  Else
                    .Text = DB.Units.ItemByKey(SelParms(row).Units.id).MetricLabel
                  End If
              End Select
            End If
          Next col
        Next row
      End With
    Case "DepVars":
      i = lstRetPds.ListIndex
      If i >= 0 Then
        If lstRetPds.ItemData(i) > 0 Then
          Set MyDepVar = MyRegion.depVars(CStr(lstRetPds.ItemData(i)))
          compCnt = MyDepVar.Components.Count
          If MyRegion.PredInt And MyDepVar.Components.Count > 0 Then
            covArray = MyDepVar.PopulateMatrix()
          End If
        End If
      End If
      With grdInterval
        .Rows = lstRetPds.SelCount
        If MyRegion.PredInt Then
          DepVarFlds = 10
          AddSpace = 0
        Else
          DepVarFlds = 7
          AddSpace = 250
        End If
        .cols = DepVarFlds + 1
        .ColType(0) = ATCoTxt
        If RDO = 0 Then
          .TextMatrix(-1, 0) = "Return"
          .TextMatrix(0, 0) = "Interval"
          grdInterval.header = "Return Interval"
        ElseIf RDO = 1 Then
          .TextMatrix(-1, 0) = ""
          .TextMatrix(0, 0) = "Statistic"
          grdInterval.header = "Statistic"
        End If
        .colWidth(0) = 620 + AddSpace
        .ColType(1) = ATCoSng
        .TextMatrix(-1, 1) = "Standard"
        .TextMatrix(0, 1) = "Error"
        .colWidth(1) = 730 + AddSpace
        .ColType(2) = ATCoSng
        .TextMatrix(-1, 2) = "Estimate"
        .TextMatrix(0, 2) = "Error"
        .colWidth(2) = 700 + AddSpace
        .ColType(3) = ATCoSng
        .TextMatrix(-1, 3) = "Prediction"
        .TextMatrix(0, 3) = "Error"
        .colWidth(3) = 800 + AddSpace
        .ColType(4) = ATCoSng
        .TextMatrix(-1, 4) = "Equivalent"
        .TextMatrix(0, 4) = "Years"
        .colWidth(4) = 850 + AddSpace
        .ColType(5) = ATCoSng
        .TextMatrix(-1, 5) = "Regression"
        .TextMatrix(0, 5) = "Constant"
        .colWidth(5) = 910 + AddSpace
        If MyRegion.PredInt Then
          .ColType(6) = ATCoSng
          .TextMatrix(-1, 6) = ""
          .TextMatrix(0, 6) = "BCF"
          .colWidth(6) = 480
          .ColType(7) = ATCoSng
          .TextMatrix(-1, 7) = "t -"
          .TextMatrix(0, 7) = "Distribution"
          .colWidth(7) = 880
          .ColType(8) = ATCoSng
          .TextMatrix(-1, 8) = ""
          .TextMatrix(0, 8) = "Variance"
          .colWidth(8) = 740
        End If
        .ColType(.cols - 2) = ATCoSng
        .TextMatrix(-1, .cols - 2) = "Drn. Area"
        .TextMatrix(0, .cols - 2) = "Exponent"
        .colWidth(.cols - 2) = 850 + AddSpace
        .TextMatrix(0, .cols - 1) = "Units"
        .colWidth(.cols - 1) = 850 + AddSpace
        For col = 0 To .cols - 1
          .ColEditable(col) = True
          For row = 1 To lstRetPds.SelCount
            .row = row
            .col = col
            If Not MyRegion.depVars(row).IsNew Then
              Select Case col
                Case 0: .Text = MyDepVar.Name
                Case 1: .Text = MyDepVar.StdErr
                Case 2: .Text = MyDepVar.EstErr
                Case 3: .Text = MyDepVar.PreErr
                Case 4: .Text = MyDepVar.EquivYears
                Case 5: .Text = MyDepVar.Constant
                Case 6:
                        If MyRegion.PredInt Then
                          .Text = MyDepVar.BCF
                        Else
                          .Text = MyDepVar.ExpDA
                        End If
                Case .cols - 1:
                  If rdoUnits(0).Value Then
                    .Text = MyDepVar.Units.EnglishLabel
                  Else
                    .Text = MyDepVar.Units.MetricLabel
                  End If
                Case 7: .Text = MyDepVar.tdist
                Case 8: .Text = MyDepVar.Variance
                Case 9: .Text = MyDepVar.ExpDA
              End Select
            End If
          Next row
        Next col
      End With
      With grdComps
        If lstRetPds.SelCount > 0 Then 'DepVar is selected
          .Rows = compCnt
        Else
          .Rows = 0
        End If
        .cols = 7
        .col = 2
        .ColType(0) = ATCoTxt
        .TextMatrix(-1, 0) = "Base"
        .TextMatrix(0, 0) = "Variable"
        .ColType(1) = ATCoSng
        .TextMatrix(-1, 1) = "Base"
        .TextMatrix(0, 1) = "Modifier"
        .ColType(2) = ATCoSng
        .TextMatrix(-1, 2) = "Base"
        .TextMatrix(0, 2) = "Coefficient"
        .ColType(3) = ATCoSng
        .TextMatrix(-1, 3) = "Base"
        .TextMatrix(0, 3) = "Exponent"
        .ColType(4) = ATCoTxt
        .TextMatrix(-1, 4) = "Exponent"
        .TextMatrix(0, 4) = "Variable"
        .ColType(5) = ATCoSng
        .TextMatrix(-1, 5) = "Exponent"
        .TextMatrix(0, 5) = "Modifier"
        .ColType(6) = ATCoSng
        .TextMatrix(-1, 6) = "Exponent"
        .TextMatrix(0, 6) = "Exponent"
        For col = 0 To .cols - 1
          .ColEditable(col) = True
          .colWidth(0) = 1080
        Next col
        If lstRetPds.SelCount > 0 Then
          For row = 1 To compCnt
            Set MyComp = MyDepVar.Components(row)
            str = GetAbbrev(MyComp.ParmID)
            For col = 0 To .cols - 1
              .row = row
              .col = col
              If Not MyDepVar.Components(row).IsNew Then
                Select Case col
                  Case 0: .Text = str
                  Case 1: .Text = MyComp.BaseMod
                  Case 2: .Text = MyComp.BaseCoeff
                  Case 3: .Text = MyComp.BaseExp
                  Case 4: str = GetAbbrev(MyComp.expID)
                          .Text = str
                  Case 5: .Text = MyComp.ExpMod
                  Case 6: .Text = MyComp.ExpExp
                End Select
              End If
            Next col
          Next row
        End If
        'If Not newParm And grdInterval.Rows > 0 Then .Rows = .Rows + 1
      End With
      If Not MyRegion.PredInt Then Exit Sub
      With grdMatrix
        If i < 0 Or compCnt = 0 Then
          .Rows = 0
          .cols = 0
        Else
          .Rows = compCnt + 1
          .cols = compCnt + 1
          For col = 1 To .cols
            .ColType(col - 1) = ATCoTxt
            .colWidth(col - 1) = 800
            For row = 1 To .Rows
              .TextMatrix(row, col - 1) = covArray(row, col)
            Next row
          Next col
        End If
        For col = 0 To .cols - 1
          .ColEditable(col) = True
        Next col
        .ColsSizeByContents
      End With
  End Select
End Sub

Private Sub CheckRuralInput()
  Dim i
  If MyRegion.UrbanNeedsRural Then
    ChoseParms = True
    For i = 0 To lstParms.ListCount - 1
      If lstParms.ItemData(i) < 0 And lstParms.Selected(i) Then
        lstParms.Selected(i) = False
        MsgBox "'" & lstParms.List(i) & "' is not an editable field." & _
            vbCrLf & "It has been deselected and removed from the grid."
      End If
    Next i
    ChoseParms = False
  End If
End Sub

Private Sub lstRetPds_Click()
  If Not (ChoseParms Or Skip) Then FocusOnReturns
End Sub

Private Sub lstRetPds_GotFocus()
  FocusOnReturns
End Sub

Private Sub FocusOnReturns()
  Dim i&
  Dim depVarName$
  
  If MyRegion Is Nothing Then Exit Sub
  
  If ChangesMade Then
    depVarName = lstRetPds.List(lstRetPds.ListIndex)
    SaveChanges
    If Not depVarName = "" Then
      If lstRetPds.SelCount = 0 Then
        For i = 1 To lstRetPds.ListCount
          If lstRetPds.List(i - 1) = depVarName Then lstRetPds.Selected(i - 1) = True
          Exit Sub
        Next i
      End If
    End If
  End If

  Me.MousePointer = vbHourglass
  ChoseReturns = True
  If RDO = 0 Then
    fraEdit(2).Caption = "Return Period Values"
  ElseIf RDO = 1 Then
    fraEdit(2).Caption = "Statistic Values"
  End If
  fraEdit(0).Visible = False
  fraEdit(1).Visible = False
  fraEdit(2).Visible = True
  If MyRegion.PredInt Then
    grdMatrix.Visible = True
  Else
    grdMatrix.Visible = False
  End If
  For i = 0 To lstParms.ListCount - 1
    lstParms.Selected(i) = True
  Next i
  SetGrid "DepVars"
  ChoseReturns = False
  If lstRetPds.ListIndex >= 0 Then
    cmdSave.Enabled = True
    cmdDelete.Enabled = True
    cmdCancel.Enabled = True
  Else
    cmdSave.Enabled = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = False
  End If
  Me.MousePointer = vbDefault
End Sub

Private Sub rdoRegOpt_Click(Index As Integer)

  If Index = 1 Then
    chkRuralInput.Visible = True
  Else
    chkRuralInput.Visible = False
  End If
  chkRuralInput.Value = 0
  If MyRegion Is Nothing Then
    rdoRegOpt(0) = False
    rdoRegOpt(1) = False
    chkRuralInput.Visible = False
    chkPredInt.Value = 0
  End If
End Sub

Private Sub PopulateParms()
  Dim parmIndex&, i&

  If MyRegion.UrbanNeedsRural Then
    AddUrbanNeedsRuralParms
  End If
  If MyRegion.id > 0 Then
    lstParms.Clear
    If Not SelParms Is Nothing Then
      SelParms.Clear
      Set SelParms = Nothing
    End If
    For parmIndex = 1 To MyRegion.Parameters.Count
      lstParms.List(parmIndex - 1) = MyRegion.Parameters(parmIndex).Name
      lstParms.ItemData(parmIndex - 1) = MyRegion.Parameters(parmIndex).id
    Next
  End If
End Sub
    
Private Sub AddUrbanNeedsRuralParms()
    'add Rural Drainage Area and Computed Rural Discharge parameters
    Set MyParm = New nssParameter
    With MyParm
      Set .Region = MyRegion
      .id = -1
      .Abbrev = "RDA"
      .Name = "Rural Drainage Area"
      Set .Units = DB.Units("1")
      If MyRegion.Parameters.Count > 0 Then
        'Set min and max values = those for drainage area in region
        .SetMin MyRegion.Parameters(1).GetMin(MyRegion.State.Metric), MyRegion.State.Metric
        .SetMax MyRegion.Parameters(1).GetMax(MyRegion.State.Metric), MyRegion.State.Metric
      End If
    End With
    MyRegion.Parameters.Add MyParm, CStr("-1")
    Set MyParm = New nssParameter
    With MyParm
      Set .Region = MyRegion
      .id = -2
      .Abbrev = "CRD"
      .Name = "Computed Rural Discharge"
      Set .Units = DB.Units("13")
      .SetMin 0.01, MyRegion.State.Metric
      .SetMax 1000000, MyRegion.State.Metric
    End With
    MyRegion.Parameters.Add MyParm, "-2"
    Set MyParm = Nothing
End Sub
    
    
Private Sub PopulateDepVars()
  Dim depVarIndex&, cntr&, thisDepVar!, rank&, intCnt
  Dim rankedDepVars() As String
  Dim lVarNames() As String

  If MyRegion.id > 0 Then
    lstRetPds.Clear
    intCnt = MyRegion.depVars.Count
    If intCnt = 0 Then Exit Sub
    ReDim rankedDepVars(1, intCnt - 1)
    ReDim lVarNames(intCnt)
    If RDO = 0 Then 'remove PK from DepVar names so they can be sorted numerically
      For depVarIndex = 1 To intCnt
        If Left(MyRegion.depVars(depVarIndex).Name, 2) = "PK" Then
          lVarNames(depVarIndex) = Mid(MyRegion.depVars(depVarIndex).Name, 3)
        Else
          lVarNames(depVarIndex) = MyRegion.depVars(depVarIndex).Name
        End If
      Next depVarIndex
    End If
    For depVarIndex = 1 To intCnt
      If RDO = 0 Then  'peak flow
        rank = 0
        thisDepVar = CSng(lVarNames(depVarIndex))
        For cntr = 1 To intCnt
          If thisDepVar > CSng(lVarNames(cntr)) Then
            rank = rank + 1
          End If
        Next cntr
      ElseIf RDO = 1 Then  'low flow
        If depVarIndex > 1 Then rank = rank + 1
      End If
      rankedDepVars(0, depVarIndex - 1) = MyRegion.depVars(rank + 1).Name
      rankedDepVars(1, depVarIndex - 1) = MyRegion.depVars(rank + 1).id
    Next depVarIndex
    For depVarIndex = 0 To intCnt - 1
      lstRetPds.List(depVarIndex) = rankedDepVars(0, depVarIndex)
      lstRetPds.ItemData(depVarIndex) = rankedDepVars(1, depVarIndex)
    Next depVarIndex
  End If
End Sub

Private Function GetAbbrev(ByVal Parm As Long) As String
  Dim takeLog As Boolean
  Select Case Parm
    Case -2: GetAbbrev = "rural Dis"
    Case -1: GetAbbrev = "rural DA"
    Case 0: GetAbbrev = "none"
    Case Else:
      If Parm < 0 Then
        takeLog = True
        Parm = -Parm
      End If
      Set MyParm = MyRegion.Parameters(CStr(Parm))
      GetAbbrev = MyParm.Abbrev
      If takeLog Then GetAbbrev = "log(" & GetAbbrev & ")"
  End Select
End Function

Private Function GetCode(ByVal Parm As String) As Long
  Dim i&
  Dim takeLog As Boolean
  Select Case Parm
    Case "rural Dis": GetCode = -2
    Case "rural DA": GetCode = -1
    Case "none": GetCode = 0
    Case Else:
      If Left(Parm, 4) = "log(" Then
        Parm = Mid(Parm, 5, Len(Parm) - 5)
        takeLog = True
      End If
      Set MyParm = MyRegion.Parameters(CStr(lstParms.ItemData(i)))
      While Parm <> MyParm.Abbrev
        Set MyParm = MyRegion.Parameters(CStr(lstParms.ItemData(i)))
        i = i + 1
      Wend
      GetCode = MyParm.id
      If takeLog Then GetCode = -GetCode
  End Select
End Function

Private Function ChangesMade() As Boolean
  Dim row&, col&, i&, j&, k&, baseID&, expID&, compCnt&
  Dim tmpParm As nssParameter, baseParm As nssParameter, expParm As nssParameter
  Dim tmpComp As nssComponent
  Dim oldVals() As String
  Dim newval As String
  Dim oldStatType As String
  Dim tmpstr As String
  
  ChangesMade = False
  
  If fraEdit(0).Visible And Not MyRegion Is Nothing Then
    ReDim Changes(1, 3)
    If MyRegion.IsNew Then
      Changes(1, 0) = txtRegName.Text
      If rdoRegOpt(0) Then
        Changes(1, 1) = "0"
      ElseIf rdoRegOpt(1) Then
        Changes(1, 1) = "-1"
      End If
      If chkRuralInput.Value = 0 Then
        Changes(1, 2) = "0"
      ElseIf chkRuralInput.Value = 1 Then
        Changes(1, 2) = "-1"
      End If
      If chkPredInt.Value = 0 Then
        Changes(1, 3) = "0"
      ElseIf chkPredInt.Value = 1 Then
        Changes(1, 3) = "-1"
      End If
      ChangesMade = True
    Else
      If txtRegName.Text <> MyRegion.Name Then
        Changes(0, 0) = MyRegion.Name
        Changes(1, 0) = txtRegName.Text
        ChangesMade = True
      End If
      If rdoRegOpt(0) And MyRegion.urban Then
        Changes(0, 1) = "-1"
        Changes(1, 1) = "0"
        ChangesMade = True
      ElseIf rdoRegOpt(1) And Not MyRegion.urban Then
        Changes(0, 1) = "0"
        Changes(1, 1) = "-1"
        ChangesMade = True
      End If
      If chkRuralInput.Value = 0 And MyRegion.UrbanNeedsRural Then
        Changes(0, 2) = "-1"
        Changes(1, 2) = "0"
        ChangesMade = True
      ElseIf chkRuralInput.Value = 1 And Not MyRegion.UrbanNeedsRural Then
        Changes(0, 2) = "0"
        Changes(1, 2) = "-1"
        ChangesMade = True
      End If
      If chkPredInt.Value = 0 And MyRegion.PredInt Then
        Changes(0, 3) = "-1"
        Changes(1, 3) = "0"
        ChangesMade = True
      ElseIf chkPredInt.Value = 1 And Not MyRegion.PredInt Then
        Changes(0, 3) = "0"
        Changes(1, 3) = "-1"
        ChangesMade = True
      End If
    End If
  ElseIf fraEdit(1).Visible Then
    If grdParms.Rows = 0 Then Exit Function
    ReDim Changes(1, grdParms.Rows - 1, ParmFlds)
    i = -1
    ReDim oldVals(ParmFlds)
    For row = 1 To grdParms.Rows
      Set tmpParm = SelParms(row)
      If Not tmpParm.IsNew Then
        oldStatType = tmpParm.statTypeCode
        oldVals(0) = tmpParm.Name
        oldVals(1) = tmpParm.Abbrev
        oldVals(2) = tmpParm.GetMin(False)
        oldVals(3) = tmpParm.GetMax(False)
        oldVals(4) = tmpParm.Units.id
      Else
        ReDim oldVals(ParmFlds)
      End If
      For j = 1 To ParmFlds - 1
        If j = 4 Then 'convert unit label to index for comparison
          newval = UnitIDFromLabel(grdParms.TextMatrix(row, j + 1), DB.Units)
        Else
          newval = grdParms.TextMatrix(row, j + 1)
        End If
        If newval <> oldVals(j) Then
          Changes(0, row - 1, j) = oldVals(j)
          Changes(1, row - 1, j) = newval
          If grdParms.TextMatrix(row, 2) = "CRD" Then
            MsgBox "The Computed Rural Discharge is not an editable field." & _
                vbCrLf & "No changes will be saved for this parameter.", _
                vbCritical, "Not an editable parameter"
            grdParms.TextMatrix(row, 0) = oldStatType
            For k = 0 To ParmFlds
              grdParms.TextMatrix(row, k + 1) = oldVals(k)
            Next k
            Exit For
          ElseIf grdParms.TextMatrix(row, 2) = "RDA" Then
            MsgBox "The Rural Drainage Area is not an editable field." & _
                vbCrLf & "No changes will be saved for this parameter.", _
                vbCritical, "Not an editable field"
            grdParms.TextMatrix(row, 0) = oldStatType
            For k = 0 To ParmFlds
              grdParms.TextMatrix(row, k + 1) = oldVals(k)
            Next k
            Exit For
          End If
          ChangesMade = True
        End If
      Next j
    Next row
  ElseIf fraEdit(2).Visible And Not MyDepVar Is Nothing Then
    'First record changes to Return Period or Statistic
    ReDim Changes(1, DepVarFlds)
    ReDim oldVals(DepVarFlds)
    If MyDepVar.IsNew Then
      For i = 0 To DepVarFlds
        Changes(1, i) = grdInterval.TextMatrix(1, i)
      Next i
      ChangesMade = True
    Else
      compCnt = MyDepVar.Components.Count
      oldVals(0) = MyDepVar.Name
      oldVals(1) = MyDepVar.StdErr
      oldVals(2) = MyDepVar.EstErr
      oldVals(3) = MyDepVar.PreErr
      oldVals(4) = MyDepVar.EquivYears
      oldVals(5) = MyDepVar.Constant
      If MyRegion.PredInt Then
        oldVals(6) = MyDepVar.BCF
        oldVals(7) = MyDepVar.tdist
        oldVals(8) = MyDepVar.Variance
      End If
      oldVals(DepVarFlds - 1) = MyDepVar.ExpDA
      oldVals(DepVarFlds) = MyDepVar.Units.id
      Set MyDepVar.DB = DB
      If compCnt > 0 And MyRegion.PredInt Then
        OldMatrix = MyDepVar.PopulateMatrix()
        grdMatrix.Rows = MyDepVar.Components.Count + 1
        grdMatrix.cols = grdMatrix.Rows
      End If
      For i = 0 To DepVarFlds
        tmpstr = grdInterval.TextMatrix(1, i)
        If tmpstr = "" Then tmpstr = "0"
        If i = DepVarFlds Then 'convert unit label to index for comparison
          tmpstr = CStr(UnitIDFromLabel(grdInterval.TextMatrix(1, i), DB.Units))
        End If
        If tmpstr <> oldVals(i) Then
          Changes(0, i) = oldVals(i)
          Changes(1, i) = grdInterval.TextMatrix(1, i)
          ChangesMade = True
        End If
      Next i
    End If
    'Now record changes to Components
    ReDim CompChanges(1, grdComps.Rows, grdComps.cols - 1)
    For row = 1 To grdComps.Rows
      If row > compCnt Then
        CompChanges(1, row - 1, 0) = grdComps.TextMatrix(row, 0)
        CompChanges(1, row - 1, 1) = grdComps.TextMatrix(row, 1)
        CompChanges(1, row - 1, 2) = grdComps.TextMatrix(row, 2)
        CompChanges(1, row - 1, 3) = grdComps.TextMatrix(row, 3)
        CompChanges(1, row - 1, 4) = grdComps.TextMatrix(row, 4)
        CompChanges(1, row - 1, 5) = grdComps.TextMatrix(row, 5)
        CompChanges(1, row - 1, 6) = grdComps.TextMatrix(row, 6)
        ChangesMade = True
      Else
        Set tmpComp = MyDepVar.Components(row)
        baseID = GetCode(grdComps.TextMatrix(row, 0))
        If baseID <> tmpComp.ParmID Then
          CompChanges(0, row - 1, 0) = GetAbbrev(tmpComp.ParmID)
          CompChanges(1, row - 1, 0) = grdComps.TextMatrix(row, 0)
          ChangesMade = True
        End If
        If grdComps.TextMatrix(row, 1) <> tmpComp.BaseMod Then
          CompChanges(0, row - 1, 1) = tmpComp.BaseMod
          CompChanges(1, row - 1, 1) = grdComps.TextMatrix(row, 1)
          ChangesMade = True
        End If
        If grdComps.TextMatrix(row, 2) <> tmpComp.BaseCoeff Then
          CompChanges(0, row - 1, 2) = tmpComp.BaseCoeff
          CompChanges(1, row - 1, 2) = grdComps.TextMatrix(row, 2)
          ChangesMade = True
        End If
        If grdComps.TextMatrix(row, 3) <> tmpComp.BaseExp Then
          CompChanges(0, row - 1, 3) = tmpComp.BaseExp
          CompChanges(1, row - 1, 3) = grdComps.TextMatrix(row, 3)
          ChangesMade = True
        End If
        expID = GetCode(grdComps.TextMatrix(row, 4))
        If expID <> tmpComp.expID Then
          CompChanges(0, row - 1, 4) = GetAbbrev(tmpComp.expID)
          CompChanges(1, row - 1, 4) = grdComps.TextMatrix(row, 4)
          ChangesMade = True
        End If
        If grdComps.TextMatrix(row, 5) <> tmpComp.ExpMod Then
          CompChanges(0, row - 1, 5) = tmpComp.ExpMod
          CompChanges(1, row - 1, 5) = grdComps.TextMatrix(row, 5)
          ChangesMade = True
        End If
        If grdComps.TextMatrix(row, 6) <> tmpComp.ExpExp Then
          CompChanges(0, row - 1, 6) = tmpComp.ExpExp
          CompChanges(1, row - 1, 6) = grdComps.TextMatrix(row, 6)
          ChangesMade = True
        End If
      End If
    Next row
    If MyRegion.PredInt And grdMatrix.Rows > 1 Then
      'Record changes to Matrix
      ReDim MatrixChanges(1, 1 To grdMatrix.Rows, 1 To grdMatrix.cols)
      For row = 1 To grdMatrix.Rows
        For col = 1 To grdMatrix.cols
          If Not (compCnt > 0 And (row <= compCnt + 1 Or col <= compCnt + 1)) Then
            'new matrix
            MatrixChanges(1, row, col) = grdMatrix.TextMatrix(row, col - 1)
            ChangesMade = True
          Else
            'editing existing value in matrix
            If row <= UBound(OldMatrix, 1) Or col <= UBound(OldMatrix, 2) Then
              If IsNumeric(grdMatrix.TextMatrix(row, col - 1)) And IsNumeric(OldMatrix(row, col)) Then
                If CDbl(grdMatrix.TextMatrix(row, col - 1)) <> CDbl(OldMatrix(row, col)) Then
                  MatrixChanges(0, row, col) = OldMatrix(row, col)
                  MatrixChanges(1, row, col) = grdMatrix.TextMatrix(row, col - 1)
                  ChangesMade = True
                End If
              ElseIf grdMatrix.TextMatrix(row, col - 1) <> OldMatrix(row, col) Then
                MatrixChanges(0, row, col) = OldMatrix(row, col)
                MatrixChanges(1, row, col) = grdMatrix.TextMatrix(row, col - 1)
                ChangesMade = True
              End If
            Else 'new row or col
'              MatrixChanges(0, row, col) = "0"
              MatrixChanges(1, row, col) = grdMatrix.TextMatrix(row, col - 1)
              ChangesMade = True
            End If
          End If
        Next col
      Next row
    End If
  End If
End Function

Private Sub rdoUnits_Click(Index As Integer)
  If Index = 0 Then
    Metric = False
  Else
    Metric = True
  End If
End Sub

Private Sub txtRegName_Change()
  If Not MyRegion Is Nothing Then
    If MyRegion.IsNew Then
      lstRegions.List(lstRegions.ListIndex) = NoSpaces(txtRegName.Text)
    End If
  End If
End Sub

Private Sub SaveChanges()
  Dim i&
  Dim str$
  
  If MyRegion Is Nothing Then Exit Sub

  If fraEdit(0).Visible Then
    On Error GoTo x
    i = MsgBox("Do you want to save the new information for " & _
        vbCrLf & txtRegName.Text & ", " & State & " to the database?", _
        vbYesNo, "User Action Verification")
    Skip = True
    If i = vbYes Then
      cmdSave_Click
    Else
x:
      cmdCancel_Click
    End If
    Skip = False
  ElseIf fraEdit(1).Visible Then
    ChoseParms = True
      On Error GoTo y
      i = MsgBox("Do you want to save the Parameter changes " & _
          vbCrLf & "to the database for " & MyRegion.Name & "?", _
          vbYesNo, "User Action Verification")
      Skip = True
      If i = vbYes Then
        cmdSave_Click
      Else
y:
        cmdCancel_Click
      End If
      Skip = False
    ChoseParms = False
  ElseIf fraEdit(2).Visible Then
    ChoseReturns = True
      On Error GoTo Z
      If RDO = 0 Then
        str = "Return Period"
      ElseIf RDO = 1 Then
        str = "Statistic"
      End If
      i = MsgBox("Do you want to save the new " & str & vbCrLf & _
           "'" & grdInterval.TextMatrix(1, 0) & "' information to the database for " _
          & MyRegion.Name & "?", vbYesNo, "User Action Verification")
      Skip = True
      If i = vbYes Then
        cmdSave_Click
      Else
Z:
        cmdCancel_Click
      End If
      Skip = False
    ChoseReturns = False
  End If
End Sub

Private Function NoSpaces(ObjName As String) As String
  Dim str$
  
  str = StrRetRem(ObjName)
  NoSpaces = str
  While Len(Trim(ObjName)) > 0
    str = StrRetRem(ObjName)
    NoSpaces = NoSpaces & "_" & str
  Wend
End Function

Private Function GetStatTypeCode(StatName As String)
  Dim i&
  For i = 1 To DB.StatisticTypes.Count
    If StatName = DB.StatisticTypes(i).Name Then
      GetStatTypeCode = DB.StatisticTypes(i).code
      Exit Function
    End If
  Next i
End Function

Public Function UnitIDFromLabel(UnitLabel As String, Units As FastCollection) As Long
  Dim k As Long
  Dim id As Long
  Dim lUnitLabel As String

  id = -1
  For k = 1 To Units.Count
    If Metric Then
      lUnitLabel = Units.ItemByIndex(k).MetricLabel
    Else
      lUnitLabel = Units.ItemByIndex(k).EnglishLabel
    End If
    If UnitLabel = lUnitLabel Then
      id = Units.ItemByIndex(k).id
      Exit For
    End If
  Next k
  UnitIDFromLabel = id

End Function

Private Sub SetDB(Optional lVerify As Boolean = False)
  Dim ff As New ATCoFindFile

  On Error GoTo NoDB

  Set myMsgBox = New ATCoMessage

  'Open Stream Stats Database
  
FindDB:
  On Error GoTo NoDB
  ff.SetDialogProperties "Please locate NSS or StreamStats database version 4", "NSSv4.mdb"
  ff.SetRegistryInfo "StreamStatsDB", "Defaults", "NSSDatabaseV4"
  DBPath = ff.GetName(lVerify)
  
  If Len(DBPath) > 0 Then
    Set DB = New nssDatabase
    DB.FileName = DBPath
    If Not DBCheck(DBPath) Then
      GoTo FindDB
    End If
    
    RDO = -1
    RegionFlds = 5
    ParmFlds = 5
    DepVarFlds = 8
    CompFlds = 6
    cboState.Clear
    lstRegions.Clear
    lstParms.Clear
    lstRetPds.Clear
    grdMatrix.Rows = 1
    grdMatrix.cols = 1
    lblDatabase.Caption = "Database: " & DB.FileName
    Me.Show
  End If
  Exit Sub

NoDB:
  If MsgBox("Could not open database" & vbCr & vbCr _
        & Err.Description & vbCr & vbCr _
        & "Search for current database?", vbOKCancel, "NSS Database Problem") = vbOK Then
    SaveSetting "StreamStatsDB", "Defaults", "nssDatabaseV3", "NSSv3.mdb"
    GoTo FindDB
  Else
    End
  End If

End Sub
