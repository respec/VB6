VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "*\A..\..\ATCoCtl\ATCoCtl.vbp"
Begin VB.Form frmLowFlow 
   Caption         =   "Streamflow Equation Editor"
   ClientHeight    =   10140
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   10725
   Icon            =   "frmLowFlow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10140
   ScaleWidth      =   10725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Update Order"
      Height          =   495
      Left            =   9840
      TabIndex        =   48
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton rdoMainOpt 
      Caption         =   "Probability"
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
      Index           =   2
      Left            =   2880
      TabIndex        =   36
      Top             =   600
      Width           =   1335
   End
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
      Left            =   6720
      TabIndex        =   33
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
         TabIndex        =   35
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
         TabIndex        =   34
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
      TabIndex        =   31
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
      Left            =   9840
      TabIndex        =   30
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
      Left            =   9840
      TabIndex        =   28
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
      Left            =   9000
      TabIndex        =   27
      Top             =   600
      Width           =   732
   End
   Begin VB.OptionButton rdoMainOpt 
      Caption         =   "Other Flow"
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      Left            =   9840
      TabIndex        =   24
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
      Left            =   9840
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
      Left            =   9840
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
      Left            =   9840
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
      Left            =   9840
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
      Height          =   6735
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   9615
      Begin VB.Frame fraPIs 
         Caption         =   "Prediction Interval Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         TabIndex        =   44
         Top             =   3600
         Width           =   9375
         Begin VB.CommandButton cmdTestXi 
            Caption         =   "Test"
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
            Left            =   5280
            TabIndex        =   49
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtVector 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   47
            Top             =   600
            Width           =   9135
         End
         Begin ATCoCtl.ATCoGrid grdMatrix 
            CausesValidation=   0   'False
            Height          =   1995
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Width           =   9135
            _ExtentX        =   16113
            _ExtentY        =   3519
            SelectionToggle =   0   'False
            AllowBigSelection=   -1  'True
            AllowEditHeader =   0   'False
            AllowLoad       =   0   'False
            AllowSorting    =   0   'False
            Rows            =   590
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
         Begin VB.Label lblStatusXi 
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6120
            TabIndex        =   50
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label lblVector 
            Caption         =   "Enter Xi Vector Elements (separated by "":"")"
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
            TabIndex        =   46
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame fraEquation 
         Caption         =   "Equation Definition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   37
         Top             =   1680
         Width           =   9375
         Begin VB.CommandButton cmdTest 
            Caption         =   "Test"
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
            Left            =   5280
            TabIndex        =   43
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtEquation 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   2160
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   40
            Top             =   600
            Width           =   7095
         End
         Begin VB.ListBox lstEqtnVars 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   120
            TabIndex        =   38
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblStatus 
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6120
            TabIndex        =   42
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label lblEquation 
            Caption         =   "Label1"
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
            Left            =   2160
            TabIndex        =   41
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblEqtnVars 
            Caption         =   "Variables:"
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
            TabIndex        =   39
            Top             =   360
            Width           =   1695
         End
      End
      Begin ATCoCtl.ATCoGrid grdInterval 
         CausesValidation=   0   'False
         Height          =   1185
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
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
         Width           =   9375
         _ExtentX        =   16536
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
      Width           =   10455
      Begin VB.ListBox lstRetPds 
         Height          =   1425
         Left            =   8160
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.ListBox lstParms 
         Height          =   1425
         Left            =   4200
         MultiSelect     =   1  'Simple
         TabIndex        =   7
         Top             =   480
         Width           =   3855
      End
      Begin VB.ListBox lstRegions 
         Height          =   1425
         ItemData        =   "frmLowFlow.frx":030A
         Left            =   120
         List            =   "frmLowFlow.frx":0311
         TabIndex        =   6
         Top             =   480
         Width           =   3975
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
         Left            =   4200
         TabIndex        =   5
         Top             =   240
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
         Height          =   255
         Left            =   8160
         TabIndex        =   4
         Top             =   240
         Width           =   1455
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
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
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
            TabIndex        =   29
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
      Left            =   5160
      List            =   "frmLowFlow.frx":0323
      TabIndex        =   0
      Text            =   "cboState"
      Top             =   600
      Width           =   1455
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
      TabIndex        =   32
      Top             =   120
      Width           =   9735
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
      Left            =   4560
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
'Dim MyComp As nssComponent
Dim lXiVector As FastCollection
Dim NotNew As Boolean, Skip As Boolean, _
    ChoseParms As Boolean, ChoseReturns As Boolean
Dim Changes() As String, CompChanges() As String, _
    MatrixChanges() As String, OldMatrix() As String
Dim SelParms As FastCollection  'of NSSParameter
Dim Metric As Boolean
Dim lMath As New clsMathParser

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
              Case 3: .TextMatrix(row - i, col) = SelParms(row - i).GetMin(DB.State.Metric)
              Case 4: .TextMatrix(row - i, col) = SelParms(row - i).GetMax(DB.State.Metric)
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
    If MyRegion.DepVars.Count > 0 Then SetGrid "DepVars"
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

'Private Sub cmdConvert_Click()
'  Dim vState As nssState
'  Dim vRegion As nssRegion
'  Dim vDepVar As nssDepVar
'  Dim lEqtnStr As String
'  Dim myRec As Recordset
'  Dim sql As String
'
'  Me.MousePointer = vbHourglass
'  For Each vState In DB.States
'    For Each MyRegion In vState.Regions
'      For Each vDepVar In MyRegion.DepVars
'        lEqtnStr = BuildEquation(vDepVar)
'        sql = "SELECT * FROM DepVars WHERE DepVarID=" & vDepVar.id
'        Set myRec = DB.DB.OpenRecordset(sql, dbOpenDynaset)
'        With myRec
'          If Not .NoMatch Then
'            .Edit
'            If Len(lEqtnStr) > 255 Then
'              StrTrim lEqtnStr
'            End If
'            !Equation = lEqtnStr
'            .Update
'          Else
'            MsgBox "There is no dependent variable with the ID= " & vDepVar.id & "."
'          End If
'        End With
'      Next
'    Next
'  Next
'  Me.MousePointer = vbDefault
'End Sub
'
'Private Sub cmdComponentConvert_Click()
'  Dim vState As nssState
'  Dim vRegion As nssRegion
'  Dim vDepVar As nssDepVar
'  Dim vComp As nssComponent
'  Dim lEqtnStr As String
'  Dim myRec As Recordset
'  Dim sql As String
'
'  Me.MousePointer = vbHourglass
'  For Each vState In DB.States
'    For Each MyRegion In vState.Regions
'      If MyRegion.PredInt Then
'        For Each vDepVar In MyRegion.DepVars
'          lEqtnStr = ""
'          If vDepVar.Components.Count > 0 Then
'            For Each vComp In vDepVar.Components
'              lEqtnStr = lEqtnStr & BldPredIntComponent(vComp) & " : "
'            Next
'            lEqtnStr = Left(lEqtnStr, Len(lEqtnStr) - 2)
'            sql = "SELECT * FROM DepVars WHERE DepVarID=" & vDepVar.id
'            Set myRec = DB.DB.OpenRecordset(sql, dbOpenDynaset)
'            With myRec
'              If Not .NoMatch Then
'                .Edit
'                If Len(lEqtnStr) > 255 Then
'                  StrTrim lEqtnStr
'                End If
'                !XiVector = lEqtnStr
'                .Update
'              Else
'                MsgBox "There is no dependent variable with the ID= " & vDepVar.id & "."
'              End If
'            End With
'          End If
'        Next
'      End If
'    Next
'  Next
'  Me.MousePointer = vbDefault
'
'End Sub

Private Sub cmdConvert_Click()
  Dim vState As nssState
  Dim vRegion As nssRegion
  Dim vDepVar As nssDepVar
  Dim lEqtnStr As String
  Dim myRec As Recordset
  Dim sql As String
  Dim lDepVars As FastCollection
  Dim lVarNames() As String
  Dim lInd As Integer
  Dim i As Integer
  Dim lRank As Integer
  Dim AllNumeric As Boolean

  Me.MousePointer = vbHourglass
  For Each vState In DB.States
    For Each MyRegion In vState.Regions
      ReDim lVarNames(MyRegion.DepVars.Count)
      AllNumeric = True
      For lInd = 1 To MyRegion.DepVars.Count
        If MyRegion.id < 10000 Then 'peak flow region, try to sort return intervals
          lVarNames(lInd) = ReplaceString(MyRegion.DepVars(lInd).Name, "_", ".")
          If Left(lVarNames(lInd), 2) = "PK" Then
            lVarNames(lInd) = Mid(lVarNames(lInd), 3)
          End If
          If Not IsNumeric(lVarNames(lInd)) Then
            AllNumeric = False
          End If
        Else
          AllNumeric = False
        End If
      Next lInd
      
      lInd = 0
      lRank = 0
      For Each vDepVar In MyRegion.DepVars
        If AllNumeric Then 'apply sorted indices
          lInd = lInd + 1
          Dim thisDepVar As Single
          lRank = 1
          thisDepVar = CSng(lVarNames(lInd))
          For i = 1 To MyRegion.DepVars.Count
            If thisDepVar > CSng(lVarNames(i)) Then
              lRank = lRank + 1
            End If
          Next i
        Else
          lRank = lRank + 1
        End If
        sql = "SELECT * FROM DepVars WHERE DepVarID=" & vDepVar.id
        Set myRec = DB.DB.OpenRecordset(sql, dbOpenDynaset)
        With myRec
          If Not .NoMatch Then
            .Edit
            !OrderIndex = lRank
            .Update
          Else
            MsgBox "There is no dependent variable with the ID= " & vDepVar.id & "."
          End If
        End With
      Next
    Next
  Next
  Me.MousePointer = vbDefault
End Sub

Private Sub cmdDatabase_Click()
  Dim lDBFName As String
  lDBFName = DB.FileName

  SetDB (True)
  If lDBFName <> DB.FileName Then 'database changed
    rdoMainOpt(0).Value = False
    rdoMainOpt(1).Value = False
    rdoMainOpt(2).Value = False
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
  Dim FileName$, str$, regnName$, flowFlag$, lstr$
  Dim urban As Boolean, isReturn As Boolean
  Dim regnVals() As Integer
  Dim parmVals() As String, depVarVals() As String, compVals() As String
  Dim covArray() As String
  Dim lXiVector As FastCollection
  Dim lVarNotFound As Integer
  Dim lRegID As Long

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
    ElseIf RDO = 2 Then
      FileName = GetSetting("SEE", "Defaults", "ProbabilityExportFile", FileName)
    End If
    If Len(Dir(FileName, vbDirectory)) = 0 Then
      FileName = CurDir & "\Import.csv"
    Else
      FileName = FileName & "\Import.csv"
    End If
    .FileName = FileName
    .Filter = "(*.csv)|*.csv"
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
  lstr = StrRetRem(str)
  If Len(lstr) < 2 Then lstr = "0" & lstr
  Set DB.State = DB.States(lstr)
  While cboState.ItemData(i) <> CLng(DB.State.code)
    i = i + 1
  Wend
  cboState.ListIndex = i
  flowFlag = StrRetRem(str) 'this skips the state name
  flowFlag = StrRetRem(str) ' Right(str, 1)
  If flowFlag = "0" Then
    DB.State.ClearState "ReturnPeriods"
    isReturn = True
  ElseIf flowFlag = "1" Then
    DB.State.ClearState "Statistics"
    isReturn = False
  ElseIf flowFlag = "2" Then
    DB.State.ClearState "Probability"
    isReturn = False
  End If

  DepVarFlds = 11 'always set to import all possible DepVar fields
  
  'read in number of regions and metric flag
  regnCnt = CLng(StrRetRem(str))
  If str = "1" Then Metric = True Else Metric = False
  
  'loop thru regions
  ReDim regnVals(RegionFlds - 3)
  For i = 1 To regnCnt
    'read in region info
    lstParms.Clear
    lstRetPds.Clear
    Line Input #inFile, str
    regnName = ""
    While Len(regnName) = 0 And Len(str) > 0
      regnName = StrRetRem(str)
    Wend
    lstr = StrRetRem(str)
    If Len(lstr) > 0 Then 'Region ID found
      lRegID = CInt(lstr)
    Else
      lRegID = -1
    End If
    If StrRetRem(str) = "0" Then urban = False Else urban = True
    regnVals(0) = StrRetRem(str)
    regnVals(1) = StrRetRem(str)
    Set MyRegion = New nssRegion
    Set MyRegion.DB = DB
    MyRegion.Add RDO, regnName, urban, regnVals(0), regnVals(1), 0, , lRegID
    DB.State.PopulateRegions
    Set MyRegion = DB.State.Regions(regnName)
    'Read in parameters info - NOT ANY MORE, records merged 1/21/2010, prh
    parmCnt = StrRetRem(str)  'number or parameters for this region
    ReDim parmVals(parmCnt - 1, ParmFlds)
    depVarCnt = StrRetRem(str)  'number or RetPds/Statistics for this region
    If depVarCnt > 0 Then ReDim depVarVals(depVarCnt - 1, DepVarFlds)
    'Loop thru parameters
    For j = 0 To parmCnt - 1
      Line Input #inFile, str
      parmVals(j, 0) = ""
      While Len(parmVals(j, 0)) = 0 And Len(str) > 0
        parmVals(j, 0) = StrRetRem(str)
      Wend
      'Read in fields for each parm
      For k = 1 To ParmFlds - 1
        parmVals(j, k) = StrRetRem(str)
      Next k
      If parmVals(j, 0) <> "RURAL_DA" And parmVals(j, 0) <> "RURAL_DIS" Then
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
    For j = 0 To depVarCnt - 1
      'Read in values
      Line Input #inFile, str
      depVarVals(j, 0) = ""
      While Len(depVarVals(j, 0)) = 0 And Len(str) > 0
        depVarVals(j, 0) = StrRetRem(str)
      Wend
      For k = 1 To DepVarFlds
        If Len(str) > 0 Then
          If k >= 9 And Left(str, 1) = """" Then 'strip quotes from Equation and XiVector fields
            depVarVals(j, k) = StrSplit(str, """", "") 'finds first quote
            depVarVals(j, k) = StrSplit(str, """", "") 'finds equations up to 2nd quote
            If (k = 9 Or k = 10) And Len(str) > 0 Then 'another field, find comma
              depVarVals(j, k + 1) = StrSplit(str, ",", "")
            End If
'            If Left(depVarVals(j, k), 1) = """" Then
'              depVarVals(j, k) = Mid(depVarVals(j, k), 2)
'              If Right(depVarVals(j, k), 1) = """" Then
'                depVarVals(j, k) = Left(depVarVals(j, k), Len(depVarVals(j, k)) - 1)
'              End If
'            End If
          Else 'just use next comma as field separator
            depVarVals(j, k) = StrSplit(str, ",", "")
          End If
        Else
          Exit For
        End If
      Next k
      'Write values to DB
      Set MyDepVar = New nssDepVar
      DepVarID = MyDepVar.Add(isReturn, MyRegion, depVarVals(j, 0), depVarVals(j, 1), _
          depVarVals(j, 2), depVarVals(j, 3), depVarVals(j, 4), depVarVals(j, 5), _
          depVarVals(j, 6), depVarVals(j, 7), depVarVals(j, 8), depVarVals(j, 9), _
          depVarVals(j, 10), depVarVals(j, 11))
      MyRegion.PopulateDepVars
      'check equation being imported
      If lMath.StoreExpression(depVarVals(j, 9)) Then
        'compCnt = lMath.VarTop
        If MyRegion.PredInt And Len(depVarVals(j, 10)) > 0 Then
          Set lXiVector = ParseXiVector(depVarVals(j, 10))
          compCnt = lXiVector.Count
        End If
        lVarNotFound = 0
        For m = 1 To lMath.VarTop
          k = 0
          While k < parmCnt
            If UCase(lMath.VarName(m)) = UCase(parmVals(k, 0)) Then
              k = parmCnt
            End If
            k = k + 1
          Wend
          If k = parmCnt Then 'didn't find variable in parameter list
            lVarNotFound = m
            Exit For
          End If
        Next m
        If lVarNotFound > 0 Then
          MsgBox "In Region " & MyRegion.Name & ", invalid Parameter in equation for DepVar: " & _
                 depVarVals(j, 0) & vbCrLf & "Equation:  " & depVarVals(j, 9) & vbCrLf & _
                 "Invalid Parameter is: " & lMath.VarName(m) & vbCrLf & _
                 "Import will be stopped; correct Import file and try again.", vbCritical, "Import Error"
          Err.Raise 32755
        End If
      Else
        MsgBox "In Region " & MyRegion.Name & ", invalid equation format for DepVar: " & _
               depVarVals(j, 0) & vbCrLf & "Equation:  " & depVarVals(j, 9) & vbCrLf & _
               "Problem occurs at position: " & lMath.ErrorPos & _
               "Import will be stopped; correct Import file and try again.", vbCritical, "Import Error"
        Err.Raise 32755
      End If
      If MyRegion.PredInt Then
        'ReDim covArray(1 To compCnt + 1, 1 To compCnt + 1)
        ReDim covArray(1 To compCnt, 1 To compCnt)
        For k = 1 To compCnt '+ 1
          Line Input #inFile, str
          While Left(str, 1) = "," Or Left(str, 1) = " "
            str = Mid(str, 2)  'gets rid of initial separators
          Wend
          For m = 1 To compCnt '+ 1
            covArray(k, m) = StrRetRem(str)
          Next m
        Next k
        MyDepVar.AddMatrix MyRegion, DepVarID, covArray()
      End If
    Next j
  Next i
  Close inFile
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
  Dim i&, j&, k&, OutFile&, tmpCnt&, row&, col&
  Dim FileName$, str$
  Dim lBlankPredsStr As String
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
    ElseIf RDO = 2 Then
      FileName = GetSetting("SEE", "Defaults", "ProbabilityExportFile")
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
    ElseIf RDO = 2 Then
      FileName = FileName & "-Probability"
    End If
    'Increment output file name if files already exported for state
    While Len(Dir(FileName & ".csv")) > 0
      i = i + 1
      If i > 2 Then FileName = Left(FileName, Len(FileName) - 2)
      FileName = FileName & "-" & i
    Wend
    .FileName = FileName
    .Filter = "(*.csv)|*.csv"
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
    ElseIf RDO = 2 Then
      SaveSetting "SEE", "Defaults", "ProbabilityExportFile", PathNameOnly(FileName)
      j = 2
    End If
  End With
  
  OutFile = FreeFile
  Open FileName For Output As OutFile
  If DB.State.Metric Then
    str = DB.State.code & "," & DB.State.Name & "," & j & "," & lstRegions.ListCount & ",1,,,,,,,,,,,,,,,,,,,,,"
  Else
    str = DB.State.code & "," & DB.State.Name & "," & j & "," & lstRegions.ListCount & ",0,,,,,,,,,,,,,,,,,,,,,"
  End If
  Print #OutFile, str
  'Loop thru Regions
  For i = 1 To lstRegions.ListCount
    Set MyRegion = DB.State.Regions(lstRegions.List(i - 1))
    If MyRegion.ROIRegnID <> "0" Then GoTo nextRegion
    If MyRegion.urban Then
      j = 1
      AddUrbanNeedsRuralParms
    Else
      j = 0
    End If
    str = ",,,,," & MyRegion.Name & "," & MyRegion.id & "," & j
    If MyRegion.UrbanNeedsRural Then j = 1 Else j = 0
    str = str & "," & j
    If MyRegion.PredInt Then j = 1 Else j = 0
    str = str & "," & j
    'append blank fields at end of string to match Excel formatting
    str = str & "," & MyRegion.Parameters.Count & "," & MyRegion.DepVars.Count & ",,,,,,,,,,,,,,"
    If MyRegion.PredInt Then
      lBlankPredsStr = ""
      For j = 1 To MyRegion.Parameters.Count + 1
        lBlankPredsStr = lBlankPredsStr & ","
      Next j
      str = str & lBlankPredsStr
    End If
    Print #OutFile, str
    'Loop thru Parameters
    For j = 1 To MyRegion.Parameters.Count
      Set MyParm = MyRegion.Parameters(j)
      str = ",,,,,,,,,,,," & MyParm.Abbrev & "," & MyParm.Name & "," & _
            MyParm.GetMin(DB.State.Metric) & "," & MyParm.GetMax(DB.State.Metric) & "," & _
            MyParm.Units.id & ",,,,,,,,,,"
      If MyRegion.PredInt Then str = str & lBlankPredsStr
      Print #OutFile, str
    Next j
    'Loop thru Return Periods/Statistics
    For j = 1 To MyRegion.DepVars.Count
      Set MyDepVar = MyRegion.DepVars(j)
      str = ",,,,,,,,,,,,,,,,," & MyDepVar.Name & "," & Round(MyDepVar.StdErr, 1) & "," & _
            Round(MyDepVar.EstErr, 1) & "," & Round(MyDepVar.PreErr, 1) & "," & _
            Round(MyDepVar.EquivYears, 1) & "," & MyDepVar.BCF & "," & _
            Round(MyDepVar.tdist, 4) & "," & Round(MyDepVar.Variance, 4) & "," & _
            Round(MyDepVar.ExpDA, 4)
      If InStr(MyDepVar.Equation, ",") > 0 Then 'commas in equation, surround w/quotes
        str = str & ",""" & MyDepVar.Equation & """"
      Else 'no quotes needed
        str = str & "," & MyDepVar.Equation
      End If
      If MyRegion.PredInt Then
        If InStr(MyDepVar.XiVectorText, ",") > 0 Then 'commas in XiVector, surround w/quotes
          str = str & ",""" & MyDepVar.XiVectorText & """"
        Else 'no quotes needed
          str = str & "," & MyDepVar.XiVectorText
        End If
      Else 'include blank XiVector field for proper importing of order index in next field
        str = str & ","
      End If
      str = str & "," & MyDepVar.OrderIndex & lBlankPredsStr
      Print #OutFile, str
      If MyRegion.PredInt Then  'using prediction intervals
        covArray = MyDepVar.PopulateMatrix
        If UBound(covArray, 1) > 1 Then  'this Return/Stat has a covariance matrix
          'Loop thru Covariance Matrix
          For row = 1 To UBound(covArray, 1)
            str = "" 'vbTab & vbTab & vbTab
            For col = 1 To UBound(covArray, 2)
              str = str & covArray(row, col) & ","
            Next col
            Print #OutFile, ",,,,,,,,,,,,,,,,,,,,,,,,,,,,," & Left(str, Len(str) - 1)
          Next row
        End If
      End If
    Next j
nextRegion:
  Next i
  Close OutFile
  
  MsgBox "Completed Export to file " & FileName, vbOKOnly, "SEE Export"
  cboState_Click
  If lstRegions.SelCount > 0 Then
    Set MyRegion = DB.State.Regions(lstRegions.List(lstRegions.ListIndex))
  Else
    Set MyRegion = Nothing
  End If

x:
  Me.MousePointer = vbDefault

End Sub

Private Sub cmdTest_Click()
  Dim lPos As Integer
  Dim i As Integer
  Dim j As Integer
  Dim lVarNotFound As Integer
'  Dim lVarCount As Integer
  
  'check equation status
  If lMath.StoreExpression(txtEquation.Text) Then
    lblStatus.Caption = "Good!"
    lblStatus.BackColor = vbGreen
  Else 'problem with equation
    lPos = lMath.ErrorPos
    lblStatus.Caption = "Problem with equation" & vbCrLf & "at position " & lPos
    lblStatus.BackColor = vbRed
  End If
  If lMath.VarTop > 0 Then
    lVarNotFound = 0
'    lVarCount = 0
    For i = 1 To lMath.VarTop
      j = 0
      While j < lstEqtnVars.ListCount
        If UCase(lMath.VarName(i)) = UCase(lstEqtnVars.List(j)) Then
          j = lstEqtnVars.ListCount
'          lVarCount = lVarCount + 1
        End If
        j = j + 1
      Wend
      If j = lstEqtnVars.ListCount Then 'didn't find variable in parameter list
        lVarNotFound = i
        Exit For
      End If
    Next i
    If lVarNotFound > 0 Then
      lblStatus.Caption = "Var not found:" & vbCrLf & lMath.VarName(lVarNotFound)
      lblStatus.BackColor = vbRed
    End If
  Else
    lblStatus.Caption = "No Variables"
    lblStatus.BackColor = vbRed
  End If
  
End Sub

Private Sub cmdTestXi_Click()
  Dim lPos As Integer
  Dim i As Integer
  Dim j As Integer
  Dim lVarNotFound As Integer
  Dim lstr As String
  Dim lEqtn As String
  
  Set lXiVector = New FastCollection
  lXiVector.Add "1"
  lstr = txtVector.Text
  lEqtn = StrSplit(lstr, ":", "")
  While Len(lEqtn) > 0
    'check equation status
    lXiVector.Add lEqtn
    Set lMath = New clsMathParser
    If lMath.StoreExpression(lEqtn) Then
      lblStatusXi.Caption = "Good!"
      lblStatusXi.BackColor = vbGreen
    Else 'problem with equation
      lPos = lMath.ErrorPos
      lblStatusXi.Caption = "Problem with Element " & lXiVector.Count - 1 & vbCrLf & "at position " & lPos
      lblStatusXi.BackColor = vbRed
      lstr = ""
    End If
    If Len(lstr) > 0 Then
      If lMath.VarTop > 0 Then
        lVarNotFound = 0
        For i = 1 To lMath.VarTop
          j = 0
          While j < lstEqtnVars.ListCount
            If UCase(lMath.VarName(i)) = UCase(lstEqtnVars.List(j)) Then
              j = lstEqtnVars.ListCount
            End If
            j = j + 1
          Wend
          If j = lstEqtnVars.ListCount Then 'didn't find variable in parameter list
            lVarNotFound = i
            lstr = ""
            Exit For
          End If
        Next i
      Else
        lblStatusXi.Caption = "No Variables"
        lblStatusXi.BackColor = vbRed
        lstr = ""
      End If
    End If
    lEqtn = StrSplit(lstr, ":", "")
  Wend

  If lVarNotFound > 0 Then
    lblStatusXi.Caption = "Variable not found:" & vbCrLf & lMath.VarName(lVarNotFound)
    lblStatusXi.BackColor = vbRed
    lstr = ""
  ElseIf MyRegion.PredInt Then 'set covariance matrix grid size
    grdMatrix.Rows = lXiVector.Count
    grdMatrix.cols = lXiVector.Count
    For i = 0 To grdMatrix.cols - 1
      grdMatrix.ColEditable(i) = True
    Next i
  End If

End Sub

Private Sub Form_Load()
  
  SetDB

End Sub

Private Sub grdInterval_LostFocus()
  lblEquation.Caption = grdInterval.TextMatrix(1, 0) & " ="
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
        For i = 1 To retCnt
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
      ElseIf RDO >= 1 Then  'add statistics to drop-down list
        For i = 1 To DB.LFStats.Count
          .addValue DB.LFStats(i).Name
        Next i
      End If
      .ComboCheckValidValues = False
    ElseIf .col = .cols - 2 Then
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
  For stIndex = 0 To DB.States.Count - 2 'use -2 to remove Dummy state from list
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
  ElseIf RDO >= 1 Then
    lblRetPds.Caption = "Statistics:"
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
  ElseIf rdoMainOpt(1) Or rdoMainOpt(2) Then
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
      If DB.State.Regions(regnIndex - i).ROIRegnID <> 0 Or _
          rdoMainOpt(0) And Abs(DB.State.Regions(regnIndex - i).LowFlowRegnID) > 0 Or _
          rdoMainOpt(1) And DB.State.Regions(regnIndex - i).LowFlowRegnID <= 0 Or _
          rdoMainOpt(2) And DB.State.Regions(regnIndex - i).LowFlowRegnID >= 0 Then
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
'  Dim i&
'  Dim EqtnStr As String, BaseVar As String, BaseStr As String
'  Dim EqtnModStr As String
'  Dim ExpStr As String
'  Dim BaseMod As Single, ExpMod As Single
'  Dim InMultExp As Boolean
'
'  With grdComps
'    .ClearValues
'    If .col = 0 Or .col = 4 Then 'build list of parameters for base or exponent variable
'      For i = 0 To lstParms.ListCount - 1
'        If Not (.col = 4 And lstParms.ItemData(i) < 0) Then _
'            .addValue GetAbbrev(lstParms.ItemData(i))
'      Next i
'      If .col = 0 And RDO = 2 Then 'allow natural log for probability equations
'        For i = 0 To lstParms.ListCount - 1
'          If lstParms.ItemData(i) >= 0 Then _
'            .addValue "ln(" & GetAbbrev(lstParms.ItemData(i)) & ")"
'        Next i
'      End If
'      If .col = 4 Then 'allow log transformations in exponent
'        For i = 0 To lstParms.ListCount - 1
'          If lstParms.ItemData(i) >= 0 Then _
'            .addValue "log(" & GetAbbrev(lstParms.ItemData(i)) & ")"
'        Next i
'      End If
'      .addValue "none"
'    ElseIf .col = 7 And Len(.TextMatrix(.row, .col)) > 0 Then 'set allowable indices for exponent
'      .addValue "0"
'      .addValue "1" 'allows for first exponent with multiple parms to be selected
'      For i = 1 To .Rows
'        If Len(.TextMatrix(i, 7)) > 0 And .TextMatrix(i, 7) <> "0" Then
'          'allows for next exponent with mutliple parms to be entered
'          .addValue CStr(CInt(.TextMatrix(i, 7)) + 1)
'        End If
'      Next i
'    End If
'
'    'update equation display
'    EqtnModStr = grdInterval.TextMatrix(1, 5)
'    If Len(EqtnModStr) > 0 And EqtnModStr <> "1" Then
'      EqtnStr = lstRetPds.List(lstRetPds.ListIndex) & " = " & EqtnModStr
'    Else
'      EqtnStr = lstRetPds.List(lstRetPds.ListIndex) & " ="
'    End If
'    InMultExp = False
'    For i = 1 To .Rows
'      If Not InMultExp Then
'        If Len(.TextMatrix(i, 0)) > 0 Then
'          If .TextMatrix(i, 0) = "none" Then
'            BaseVar = ""
'          Else
'            BaseVar = .TextMatrix(i, 0)
'          End If
'          If Len(.TextMatrix(i, 1)) > 0 Then
'            BaseMod = CSng(.TextMatrix(i, 1))
'          Else
'            BaseMod = 0
'          End If
'          Select Case BaseMod
'            Case Is > 0: BaseStr = "(" & BaseVar & "+" & BaseMod & ")"
'            Case Is < 0: BaseStr = "(" & BaseVar & BaseMod & ")"
'            Case Else: BaseStr = "(" & BaseVar & ")"
'          End Select
'          If Len(.TextMatrix(i, 2)) > 0 And .TextMatrix(i, 2) <> "1" Then
'            BaseStr = "(" & .TextMatrix(i, 2) & BaseStr & ")"
'          End If
'        End If
'        If RDO = 2 Then
'          EqtnStr = EqtnStr & "+ " & BaseStr
'        Else
'          EqtnStr = EqtnStr & "  " & BaseStr
'        End If
'      End If
'      If Len(.TextMatrix(i, 3)) > 0 And .TextMatrix(i, 3) <> "1" Then
'        If Not InMultExp And .TextMatrix(i, 7) <> "0" Then '1st instance of multiple parms in exponent
'          EqtnStr = EqtnStr & "^(" & .TextMatrix(i, 3)
'        ElseIf InMultExp Then
'          EqtnStr = EqtnStr & .TextMatrix(i, 3)
'        Else
'          EqtnStr = EqtnStr & "^" & .TextMatrix(i, 3)
'        End If
'      End If
'      If Len(.TextMatrix(i, 4)) > 0 Then
'        If .TextMatrix(i, 4) = "none" Then
'          ExpStr = ""
'        Else
'          ExpStr = .TextMatrix(i, 4)
'          If Len(.TextMatrix(i, 5)) > 0 Then
'            ExpMod = CSng(.TextMatrix(i, 5))
'          Else
'            ExpMod = 0
'          End If
'          Select Case ExpMod
'            Case Is > 0: ExpStr = "(" & ExpStr & "+" & ExpMod & ")" '"^" & "(" & ExpStr & "+" & ExpMod & ")"
'            Case Is < 0: ExpStr = "(" & ExpStr & ExpMod & ")" '"^" & "(" & ExpStr & "-" & ExpMod & ")"
'            Case Else: ExpStr = "(" & ExpStr & ")" '"^" & "(" & ExpStr & ")"
'          End Select
'          If Len(.TextMatrix(i, 6)) > 0 And .TextMatrix(i, 6) <> "1" Then
'            ExpStr = ExpStr & "^" & .TextMatrix(i, 6)
'          End If
'        End If
'        EqtnStr = EqtnStr & ExpStr
'        If .TextMatrix(i, 7) <> "0" Then 'indicates multiple parms in exponent
'          If i < .Rows Then 'see if next row's parm also belongs in this exponent
'            If .TextMatrix(i + 1, 7) = .TextMatrix(i, 7) Then 'yes, it's part of this exponent
'              InMultExp = True
'            Else 'no, this is the last parm in this exponent
'              InMultExp = False
'              EqtnStr = EqtnStr & ")"
'            End If
'          Else 'must be last of multiple parms in exponent
'            EqtnStr = EqtnStr & ")"
'          End If
'        End If
'      End If
'    Next i
'    If RDO = 2 Then
'      i = InStr(EqtnStr, "=")
'      lblEquation.Caption = Left(EqtnStr, i + 1) & "e^(" & Mid(EqtnStr, i + 2) & ")/1+e^(" & Mid(EqtnStr, i + 2) & ")"
'    Else
'      lblEquation.Caption = EqtnStr
'    End If
'  End With
End Sub

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
  Set DB.State = DB.States(cboState.ListIndex + 1) 'DB.States.ItemByKey(CStr(cboState.ItemData(cboState.ListIndex)))
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
  If Not MyRegion.DepVars Is Nothing Then
    MyRegion.DepVars.Clear
    Set MyRegion.DepVars = Nothing
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
  
  If fraEdit(0).Visible Then 'regions
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
  ElseIf fraEdit(1).Visible Then 'parameters
    If Not MyParm Is Nothing Then
      If MyParm.IsNew And grdParms.Rows > 0 Then Exit Sub
    End If
    Set MyParm = New nssParameter
    MyParm.IsNew = True
    Set MyParm.Region = MyRegion
    MyRegion.Parameters.Add MyParm, "0"
    lstParms.AddItem "New Parameter"
    lstParms.Selected(lstParms.ListCount - 1) = True
  ElseIf fraEdit(2).Visible Then 'depvars
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
    If MyRegion.DepVars.IndexFromKey("0") < 0 Then
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
    lblEquation.Caption = "New ="
    txtEquation.Text = ""
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
'  Dim tmpDepVar As nssDepVar
'  Dim tmpComp As nssComponent
  
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
        For i = 1 To MyRegion.DepVars.Count
          Set MyDepVar = MyRegion.DepVars(i)
'          For j = 1 To MyDepVar.Components.Count
'            Set MyComp = MyDepVar.Components(j)
'            MyComp.Delete
'          Next j
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
      CheckRuralInput  'to make sure we don't delete RURAL_DA or RURAL_DIS
      For i = 1 To SelParms.Count
'        For Each tmpDepVar In MyRegion.DepVars
'          For Each tmpComp In tmpDepVar.Components
'            If tmpComp.ParmID = SelParms(i).id Or Abs(tmpComp.expID) = SelParms(i).id Then tmpComp.Delete
'          Next tmpComp
'        Next tmpDepVar
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
      ElseIf RDO >= 1 Then
        response = myMsgBox.Show("Are you certain you want to delete the " & _
            lstRetPds.List(lstRetPds.ListIndex) & " statistic" & vbCrLf & _
            "from the database for " & MyRegion.Name & ", " & State & "?", _
            "User Action Verification", "+&Yes", "-&Cancel")
      End If
      If response = 1 Then
        j = lstRetPds.ItemData(lstRetPds.ListIndex)
        Set MyDepVar = MyRegion.DepVars(CStr(j))
'        For i = 1 To MyDepVar.Components.Count
'          Set MyComp = MyDepVar.Components(i)
'          MyComp.Delete
'        Next i
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
  ElseIf RDO >= 1 Then
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
    cmdTest_Click
    If lblStatus.BackColor = vbRed Then
      MsgBox "You must enter a valid equation." & vbCrLf & _
             "See the Status field for more detail.", , "Equation Problem"
    End If
    If fraPIs.Visible Then
      cmdTestXi_Click
      If lblStatusXi.BackColor = vbRed Then
        MsgBox "You must enter valid Xi Vector elements." & vbCrLf & _
               "See the Status field for more detail.", , "Prediction Intervals Problem"
      End If
    End If
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
      ElseIf RDO >= 1 Then
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
        If Not MyRegion.Add(RDO, txtRegName.Text, rdoRegOpt(1), _
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
          Set MyParm.DB = DB
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
    ElseIf fraEdit(2).Visible Then 'editing return period/statistic, equation, and matrix
      If MyRegion.PredInt Then
        BCF = grdInterval.TextMatrix(1, 5)
        tdist = grdInterval.TextMatrix(1, 6)
        Variance = grdInterval.TextMatrix(1, 7)
        ExpDA = grdInterval.TextMatrix(1, 8)
      Else
        BCF = ""
        tdist = ""
        Variance = ""
        ExpDA = grdInterval.TextMatrix(1, 5)
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
            BCF, tdist, Variance, ExpDA, txtEquation.Text, txtVector.Text, _
            grdInterval.TextMatrix(1, grdInterval.cols - 1))
            'grdInterval.TextMatrix(1, 5), BCF, tdist, Variance, ExpDA)
        If tmpID = -1 Then GoTo x
        ResetDB
        lstRetPds.Clear
        PopulateDepVars
        MyDepVar.IsNew = True
      Else
        'Edit existing Return or Statistic
        MyDepVar.Edit grdInterval.TextMatrix(1, 0), grdInterval.TextMatrix(1, 1), _
                      grdInterval.TextMatrix(1, 2), grdInterval.TextMatrix(1, 3), _
                      grdInterval.TextMatrix(1, 4), 1, BCF, tdist, Variance, ExpDA, _
                      txtEquation.Text, txtVector.Text, _
                      grdInterval.TextMatrix(1, grdInterval.cols - 1)
                      'grdInterval.TextMatrix(1, 5), BCF, tdist, Variance, ExpDA
        MyDepVar.ClearOldMatrix 'ClearOldComponents
        ResetDB
        tmpID = MyDepVar.id
      End If
      If MyRegion.PredInt Then
        ReDim covArray(1 To grdMatrix.Rows, 1 To grdMatrix.cols)
        For i = 1 To grdMatrix.Rows
          'Add matrix values from 'i'th row of grid to DB
          For k = 1 To grdMatrix.cols
            covArray(i, k) = grdMatrix.TextMatrix(i, k - 1)
            'Write Covariance Matrix changes to DetailedLog table
            If Len(MatrixChanges(1, i, k)) > 0 Then
              MyRegion.DB.RecordChanges TransID, "Covariance", 3, CStr(tmpID) & " " & _
                   i & " " & k, MatrixChanges(0, i, k), MatrixChanges(1, i, k)
            End If
          Next k
        Next i
        'MyDepVar.ClearOldMatrix
        MyDepVar.AddMatrix MyRegion, tmpID, covArray()
      End If
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
        Set MyDepVar = MyRegion.DepVars(CStr(lstRetPds.ItemData(lstRetPds.ListIndex)))
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
  Dim i&, j&, row&, col&, compCnt&
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
                Case 3: .Text = SelParms(row).GetMin(DB.State.Metric)
                Case 4: .Text = SelParms(row).GetMax(DB.State.Metric)
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
          Set MyDepVar = MyRegion.DepVars(CStr(lstRetPds.ItemData(i)))
          If MyRegion.PredInt Then 'And MyDepVar.Components.Count > 0 Then
            covArray = MyDepVar.PopulateMatrix()
          End If
        End If
      End If
      With grdInterval
        .Rows = lstRetPds.SelCount
        If MyRegion.PredInt Then
          DepVarFlds = 10
          AddSpace = 100
        Else
          DepVarFlds = 7
          AddSpace = 350
        End If
        .cols = DepVarFlds + 1
        .ColType(0) = ATCoTxt
        If RDO = 0 Then
          .TextMatrix(-1, 0) = "Return"
          .TextMatrix(0, 0) = "Interval"
          grdInterval.header = "Return Interval"
        ElseIf RDO >= 1 Then
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
        If MyRegion.PredInt Then
          .ColType(5) = ATCoSng
          .TextMatrix(-1, 5) = ""
          .TextMatrix(0, 5) = "BCF"
          .colWidth(5) = 480
          .ColType(6) = ATCoSng
          .TextMatrix(-1, 6) = "t -"
          .TextMatrix(0, 6) = "Distribution"
          .colWidth(6) = 880
          .ColType(7) = ATCoSng
          .TextMatrix(-1, 7) = ""
          .TextMatrix(0, 7) = "Variance"
          .colWidth(7) = 740
        End If
        .ColType(.cols - 3) = ATCoSng
        .TextMatrix(-1, .cols - 3) = "Drn. Area"
        .TextMatrix(0, .cols - 3) = "Exponent"
        .colWidth(.cols - 3) = 850 + AddSpace
        .TextMatrix(0, .cols - 2) = "Units"
        .colWidth(.cols - 2) = 850 + AddSpace
        .ColType(.cols - 1) = ATCoTxt
        .TextMatrix(-1, .cols - 1) = "Order"
        .TextMatrix(0, .cols - 1) = "Index"
        .ColType(.cols - 1) = ATCoInt
        .ColMin(.cols - 1) = 1
        .ColMax(.cols - 1) = MyRegion.DepVars.Count
        For col = 0 To .cols - 1
          .ColEditable(col) = True
          For row = 1 To lstRetPds.SelCount
            .row = row
            .col = col
            If Not MyRegion.DepVars(row).IsNew Then
              Select Case col
                Case 0: .Text = MyDepVar.Name
                Case 1: .Text = MyDepVar.StdErr
                Case 2: .Text = MyDepVar.EstErr
                Case 3: .Text = MyDepVar.PreErr
                Case 4: .Text = MyDepVar.EquivYears
                Case 5:
                        If MyRegion.PredInt Then
                          .Text = MyDepVar.BCF
                        Else
                          .Text = MyDepVar.ExpDA
                        End If
                Case .cols - 2:
                  If rdoUnits(0).Value Then
                    .Text = MyDepVar.Units.EnglishLabel
                  Else
                    .Text = MyDepVar.Units.MetricLabel
                  End If
                Case .cols - 1: .Text = MyDepVar.OrderIndex
                Case 6: .Text = MyDepVar.tdist
                Case 7: .Text = MyDepVar.Variance
                Case 8: .Text = MyDepVar.ExpDA
              End Select
            End If
          Next row
        Next col
      End With
      lstEqtnVars.Clear
      For j = 0 To lstParms.ListCount - 1
        lstEqtnVars.AddItem GetAbbrev(lstParms.ItemData(j))
      Next j
      lblEquation.Caption = MyDepVar.Name & " ="
      txtEquation.Text = MyDepVar.Equation
      If MyRegion.PredInt Then
        txtVector.Text = MyDepVar.XiVectorText
        compCnt = MyDepVar.XiVector.Count
        With grdMatrix
          If i < 0 Or compCnt = 0 Then
            .Rows = 0
            .cols = 0
          Else
            .Rows = compCnt '+ 1
            .cols = compCnt '+ 1
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
      End If
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
  ElseIf RDO >= 1 Then
    fraEdit(2).Caption = "Statistic Values"
  End If
  fraEdit(0).Visible = False
  fraEdit(1).Visible = False
  fraEdit(2).Visible = True
  If MyRegion.PredInt Then
    'grdMatrix.Visible = True
    fraPIs.Visible = True
  Else
    'grdMatrix.Visible = False
    fraPIs.Visible = False
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
      .Abbrev = "RURAL_DA"
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
      .Abbrev = "RURAL_DIS"
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
    intCnt = MyRegion.DepVars.Count
    If intCnt = 0 Then Exit Sub
    ReDim rankedDepVars(1, intCnt - 1)
    ReDim lVarNames(intCnt)
    rank = 0
    For depVarIndex = 1 To intCnt
      rankedDepVars(0, rank) = MyRegion.DepVars(depVarIndex).Name
      rankedDepVars(1, rank) = MyRegion.DepVars(depVarIndex).id
      rank = rank + 1
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
    Case -2: GetAbbrev = "RURAL_DIS"
    Case -1: GetAbbrev = "RURAL_DA"
    Case -999, -4, -3, 0: GetAbbrev = "none"
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
  Select Case UCase(Parm)
    Case "RURAL DIS": GetCode = -2
    Case "RURAL DA": GetCode = -1
    Case "NONE": GetCode = 0
    Case "NONE-3": GetCode = -3
    Case "NONE-4": GetCode = -4
    Case Else:
      If Left(UCase(Parm), 3) = "LN(" Then
        Parm = Mid(Parm, 4, Len(Parm) - 4)
        takeLog = True
      End If
      If Left(UCase(Parm), 4) = "LOG(" Then
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
'  Dim tmpComp As nssComponent
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
        oldVals(2) = tmpParm.GetMin(DB.State.Metric)
        oldVals(3) = tmpParm.GetMax(DB.State.Metric)
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
          If grdParms.TextMatrix(row, 2) = "RURAL_DIS" Then
            MsgBox "The Computed Rural Discharge is not an editable field." & _
                vbCrLf & "No changes will be saved for this parameter.", _
                vbCritical, "Not an editable parameter"
            grdParms.TextMatrix(row, 0) = oldStatType
            For k = 0 To ParmFlds
              grdParms.TextMatrix(row, k + 1) = oldVals(k)
            Next k
            Exit For
          ElseIf grdParms.TextMatrix(row, 2) = "RURAL_DA" Then
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
    ReDim Changes(1, DepVarFlds + 2)
    ReDim oldVals(DepVarFlds)
    If MyDepVar.IsNew Then
      For i = 0 To DepVarFlds
        Changes(1, i) = grdInterval.TextMatrix(1, i)
      Next i
      ChangesMade = True
    Else
      compCnt = MyDepVar.XiVector.Count '.VarCount ' MyDepVar.Components.Count
      oldVals(0) = MyDepVar.Name
      oldVals(1) = MyDepVar.StdErr
      oldVals(2) = MyDepVar.EstErr
      oldVals(3) = MyDepVar.PreErr
      oldVals(4) = MyDepVar.EquivYears
      If MyRegion.PredInt Then
        oldVals(5) = MyDepVar.BCF
        oldVals(6) = MyDepVar.tdist
        oldVals(7) = MyDepVar.Variance
      End If
      oldVals(DepVarFlds - 2) = MyDepVar.ExpDA
      oldVals(DepVarFlds - 1) = MyDepVar.Units.id
      oldVals(DepVarFlds) = MyDepVar.OrderIndex
      Set MyDepVar.DB = DB
      If compCnt > 0 And MyRegion.PredInt Then
        OldMatrix = MyDepVar.PopulateMatrix()
        grdMatrix.Rows = MyDepVar.XiVector.Count 'VarCount + 1 ' MyDepVar.Components.Count + 1
        grdMatrix.cols = grdMatrix.Rows
      End If
      For i = 0 To DepVarFlds
        tmpstr = grdInterval.TextMatrix(1, i)
        If tmpstr = "" Then tmpstr = "0"
        If i = DepVarFlds - 1 Then 'convert unit label to index for comparison
          tmpstr = CStr(UnitIDFromLabel(grdInterval.TextMatrix(1, i), DB.Units))
        End If
        If tmpstr <> oldVals(i) Then
          Changes(0, i) = oldVals(i)
          Changes(1, i) = grdInterval.TextMatrix(1, i)
          ChangesMade = True
        End If
      Next i
      If txtEquation.Text <> MyDepVar.Equation Then
        Changes(0, DepVarFlds + 1) = MyDepVar.Equation
        Changes(1, DepVarFlds + 1) = txtEquation.Text
        ChangesMade = True
      End If
      If MyRegion.PredInt And txtVector.Text <> MyDepVar.XiVectorText Then
        Changes(0, DepVarFlds + 2) = MyDepVar.XiVectorText
        Changes(1, DepVarFlds + 2) = txtVector.Text
        ChangesMade = True
      End If
    End If
    If MyRegion.PredInt And grdMatrix.Rows > 1 Then
      'Record changes to Matrix
      ReDim MatrixChanges(1, 1 To grdMatrix.Rows, 1 To grdMatrix.cols)
      For row = 1 To grdMatrix.Rows
        For col = 1 To grdMatrix.cols
          'If Not (compCnt > 0 And (row <= compCnt + 1 Or col <= compCnt + 1)) Then
          If Not (compCnt > 0 And (row <= compCnt Or col <= compCnt)) Then
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

Private Sub txtEquation_Change()
  lblStatus.Caption = ""
  lblStatus.BackColor = vbButtonFace
  lblStatusXi.Caption = ""
  lblStatusXi.BackColor = vbButtonFace
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
      ElseIf RDO >= 1 Then
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
  ff.SetDialogProperties "Please locate NSS or StreamStats database version 5", "NSSv5.mdb"
  ff.SetRegistryInfo "StreamStatsDB", "Defaults", "NSSDatabaseV5"
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

Private Function ParseXiVector(ByVal aXiVectorText As String) As FastCollection
  Dim lEqtn As String
  Dim lXiVector As FastCollection
  Set lXiVector = New FastCollection

  lXiVector.Add "1"
  lEqtn = StrSplit(aXiVectorText, ":", "")
  While Len(lEqtn) > 0
    lXiVector.Add lEqtn
    lEqtn = StrSplit(aXiVectorText, ":", "")
  Wend

  Set ParseXiVector = lXiVector
End Function
