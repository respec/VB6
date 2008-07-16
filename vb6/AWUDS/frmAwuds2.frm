VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{872F11D5-3322-11D4-9D23-00A0C9768F70}#1.10#0"; "ATCoCtl.ocx"
Begin VB.Form frmAwuds2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aggregate Water-Use Data System (AWUDS)"
   ClientHeight    =   6570
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   9360
   Icon            =   "frmAwuds2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRetrieve 
      Caption         =   "Retrieve Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6600
      TabIndex        =   76
      ToolTipText     =   "Retrieve previously saved group of user selections"
      Top             =   6120
      Width           =   972
   End
   Begin VB.CommandButton cmdSaveGroup 
      Caption         =   "Save Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6600
      TabIndex        =   75
      ToolTipText     =   "Save current group of user selections"
      Top             =   5640
      Width           =   972
   End
   Begin VB.Frame fraWhat2Do 
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   120
      TabIndex        =   63
      Top             =   5520
      Width           =   6372
      Begin VB.Label lblInstructs 
         Caption         =   "Next Step Is ..."
         Height          =   612
         Left            =   120
         TabIndex        =   64
         Top             =   240
         Width           =   6132
      End
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
      Height          =   612
      Left            =   8520
      TabIndex        =   78
      ToolTipText     =   "Close the AWUDS application"
      Top             =   5760
      Width           =   732
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
      Height          =   612
      Left            =   7680
      TabIndex        =   77
      ToolTipText     =   "Access on-line help system"
      Top             =   5760
      Width           =   732
   End
   Begin MSComDlg.CommonDialog cdlgFileSel 
      Left            =   120
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   5352
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   9132
      _ExtentX        =   16113
      _ExtentY        =   9446
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "&State"
      TabPicture(0)   =   "frmAwuds2.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdNationalDB"
      Tab(0).Control(1)=   "fraDomainName"
      Tab(0).Control(2)=   "fraStSel"
      Tab(0).Control(3)=   "cdmAbout"
      Tab(0).Control(4)=   "fraDataPath"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Operation"
      TabPicture(1)   =   "frmAwuds2.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDataOpts"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Area"
      TabPicture(2)   =   "frmAwuds2.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblDate"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraAreaUnits2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdExeOpts"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fraAreaUnits"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lstYears"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lstArea"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "fraAreaUnitID"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "&Category"
      TabPicture(3)   =   "frmAwuds2.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraCats"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&Values"
      TabPicture(4)   =   "frmAwuds2.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraDataFlds"
      Tab(4).ControlCount=   1
      Begin VB.Frame fraDataPath 
         Caption         =   "Pathname to the Current AWUDS Data Directory"
         ClipControls    =   0   'False
         Height          =   1215
         Left            =   -71640
         TabIndex        =   95
         Top             =   600
         Width           =   5652
         Begin VB.CommandButton cmdDataPath 
            Caption         =   "Change Database Location"
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
            Left            =   2760
            TabIndex        =   81
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txtDataPath 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   96
            Top             =   360
            Width           =   5412
         End
      End
      Begin VB.CommandButton cdmAbout 
         Caption         =   "About AWUDS"
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
         Left            =   -67560
         TabIndex        =   94
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Frame fraAreaUnitID 
         Caption         =   "Area Unit ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1092
         Left            =   3910
         TabIndex        =   74
         Top             =   360
         Width           =   1300
         Begin VB.OptionButton rdoID 
            Caption         =   "both"
            Height          =   252
            Index           =   2
            Left            =   120
            TabIndex        =   71
            Top             =   720
            Width           =   1092
         End
         Begin VB.OptionButton rdoID 
            Caption         =   "by name"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   1092
         End
         Begin VB.OptionButton rdoID 
            Caption         =   "by code"
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   70
            Top             =   480
            Width           =   1092
         End
      End
      Begin Awuds.ATCoSelectListSorted lstArea 
         Height          =   2892
         Left            =   120
         TabIndex        =   73
         Top             =   1200
         Width           =   8892
         _ExtentX        =   15690
         _ExtentY        =   5106
      End
      Begin VB.ListBox lstYears 
         Height          =   840
         Left            =   3600
         MultiSelect     =   1  'Simple
         TabIndex        =   61
         Top             =   4440
         Width           =   1452
      End
      Begin VB.Frame fraDataOpts 
         Caption         =   "Data Operations"
         Height          =   4870
         Left            =   -74880
         TabIndex        =   34
         Top             =   360
         Width           =   8892
         Begin VB.OptionButton MainOpt 
            Caption         =   "&Import Data from External File"
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   5
            Top             =   960
            Width           =   2950
         End
         Begin VB.OptionButton MainOpt 
            Caption         =   "Interactive &Data Input/Edit"
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   720
            Width           =   2950
         End
         Begin VB.OptionButton MainOpt 
            Caption         =   "Basic Tables by Ca&tegory"
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   6
            Top             =   1560
            Width           =   3000
         End
         Begin VB.OptionButton MainOpt 
            Caption         =   "Basic Tables by A&rea"
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   7
            Top             =   1800
            Width           =   2976
         End
         Begin VB.OptionButton MainOpt 
            Caption         =   "&Entered Data Elements"
            Height          =   252
            Index           =   4
            Left            =   240
            TabIndex        =   8
            Top             =   2040
            Width           =   2950
         End
         Begin VB.OptionButton MainOpt 
            Caption         =   "Ca&lculated Tables"
            Height          =   252
            Index           =   5
            Left            =   240
            TabIndex        =   9
            Top             =   2280
            Width           =   2950
         End
         Begin VB.OptionButton MainOpt 
            Caption         =   "&Facility Tables"
            Height          =   252
            Index           =   6
            Left            =   240
            TabIndex        =   10
            Top             =   2520
            Width           =   2950
         End
         Begin VB.OptionButton MainOpt 
            Caption         =   "&Quality-Assurance Program"
            Height          =   252
            Index           =   7
            Left            =   240
            TabIndex        =   11
            Top             =   3120
            Width           =   2950
         End
         Begin VB.OptionButton MainOpt 
            Caption         =   "C&ompare State Totals by Area"
            Height          =   252
            Index           =   8
            Left            =   240
            TabIndex        =   12
            Top             =   3360
            Width           =   2952
         End
         Begin VB.OptionButton MainOpt 
            Caption         =   "Compare Data for 2 &Years"
            Height          =   252
            Index           =   9
            Left            =   240
            TabIndex        =   13
            Top             =   3600
            Width           =   2950
         End
         Begin VB.OptionButton MainOpt 
            Caption         =   "E&xport Data"
            Height          =   252
            Index           =   10
            Left            =   240
            TabIndex        =   14
            Top             =   4200
            Width           =   2892
         End
         Begin VB.Frame fraImport 
            Caption         =   "Import Data File"
            Height          =   1572
            Left            =   3240
            TabIndex        =   50
            Top             =   240
            Visible         =   0   'False
            Width           =   5532
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "Browse"
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
               Left            =   120
               TabIndex        =   52
               ToolTipText     =   "Select file for import"
               Top             =   360
               Width           =   852
            End
            Begin VB.CommandButton cmdImport 
               Caption         =   "Execute"
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
               Height          =   492
               Left            =   4560
               TabIndex        =   55
               ToolTipText     =   "Execute import for selected year and file"
               Top             =   960
               Width           =   852
            End
            Begin VB.TextBox txtCurFile 
               Height          =   288
               Left            =   1080
               TabIndex        =   53
               Top             =   600
               Width           =   4332
            End
            Begin ATCoCtl.ATCoText txtYear 
               Height          =   252
               Left            =   1080
               TabIndex        =   54
               Top             =   1080
               Width           =   732
               _ExtentX        =   1296
               _ExtentY        =   450
               InsideLimitsBackground=   16777215
               OutsideHardLimitBackground=   8421631
               OutsideSoftLimitBackground=   8454143
               HardMax         =   2100
               HardMin         =   1900
               SoftMax         =   -999
               SoftMin         =   -999
               MaxWidth        =   -999
               Alignment       =   1
               DataType        =   0
               DefaultValue    =   ""
               Value           =   ""
               Enabled         =   -1  'True
            End
            Begin VB.Label lblImpYear 
               Alignment       =   2  'Center
               Caption         =   "Year of Data:"
               Height          =   252
               Left            =   120
               TabIndex        =   56
               Top             =   1080
               Width           =   972
            End
            Begin VB.Label lblCurFile 
               Caption         =   "Current File:"
               Height          =   252
               Left            =   1080
               TabIndex        =   51
               Top             =   360
               Width           =   4332
            End
         End
         Begin VB.Frame fraNewYear 
            Caption         =   "Add New Year of Data"
            Height          =   3510
            Left            =   3240
            TabIndex        =   36
            Top             =   1200
            Visible         =   0   'False
            Width           =   5412
            Begin VB.Frame FraDataDict 
               Caption         =   "Data Dictionary"
               Height          =   576
               Left            =   120
               TabIndex        =   91
               Top             =   240
               Width           =   1695
               Begin VB.OptionButton rdoDataDict 
                  Caption         =   "2005"
                  Height          =   300
                  Index           =   1
                  Left            =   910
                  TabIndex        =   93
                  Top             =   240
                  Width           =   750
               End
               Begin VB.OptionButton rdoDataDict 
                  Caption         =   "2000"
                  Height          =   300
                  Index           =   0
                  Left            =   120
                  TabIndex        =   92
                  Top             =   240
                  Width           =   750
               End
            End
            Begin VB.Frame fraPSPop 
               Caption         =   "PS Population Served"
               Height          =   1520
               Left            =   120
               TabIndex        =   84
               Top             =   960
               Width           =   2532
               Begin VB.OptionButton rdoPSPop 
                  Caption         =   "State Total (Total)"
                  Height          =   252
                  Index           =   2
                  Left            =   120
                  TabIndex        =   89
                  Top             =   720
                  Width           =   2220
               End
               Begin VB.OptionButton rdoPSPop 
                  Caption         =   "by Unit Area (GW/SW)"
                  Height          =   252
                  Index           =   0
                  Left            =   120
                  TabIndex        =   87
                  Top             =   240
                  Width           =   2380
               End
               Begin VB.OptionButton rdoPSPop 
                  Caption         =   "by Unit Area (Total)"
                  Height          =   252
                  Index           =   1
                  Left            =   120
                  TabIndex        =   86
                  Top             =   480
                  Width           =   2295
               End
               Begin ATCoCtl.ATCoText txtPSTotal 
                  Height          =   252
                  Left            =   720
                  TabIndex        =   85
                  ToolTipText     =   "Enter the state total for PS - population served"
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   852
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  InsideLimitsBackground=   16777215
                  OutsideHardLimitBackground=   8421631
                  OutsideSoftLimitBackground=   8454143
                  HardMax         =   9999
                  HardMin         =   0
                  SoftMax         =   -999
                  SoftMin         =   0
                  MaxWidth        =   -999
                  Alignment       =   2
                  DataType        =   0
                  DefaultValue    =   ""
                  Value           =   ""
                  Enabled         =   -1  'True
               End
               Begin VB.Label lblPSTotal 
                  Alignment       =   2  'Center
                  Caption         =   "Pop Value"
                  Height          =   252
                  Left            =   720
                  TabIndex        =   88
                  ToolTipText     =   "Enter the state total for PS - population served"
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   852
               End
            End
            Begin VB.Frame fraDO 
               Caption         =   "DO Withdrawals"
               Height          =   1520
               Left            =   2760
               TabIndex        =   72
               Top             =   960
               Width           =   2532
               Begin ATCoCtl.ATCoText txtDOTotalSW 
                  Height          =   252
                  Left            =   1440
                  TabIndex        =   45
                  ToolTipText     =   "Enter the state total for DO - surface water withdrawals"
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   852
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  InsideLimitsBackground=   16777215
                  OutsideHardLimitBackground=   8421631
                  OutsideSoftLimitBackground=   8454143
                  HardMax         =   99999
                  HardMin         =   0
                  SoftMax         =   -999
                  SoftMin         =   0
                  MaxWidth        =   -999
                  Alignment       =   2
                  DataType        =   0
                  DefaultValue    =   ""
                  Value           =   ""
                  Enabled         =   -1  'True
               End
               Begin ATCoCtl.ATCoText txtDOTotalGW 
                  Height          =   252
                  Left            =   240
                  TabIndex        =   44
                  ToolTipText     =   "Enter the state total for DO - groundwater withdrawals"
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   852
                  _ExtentX        =   1508
                  _ExtentY        =   450
                  InsideLimitsBackground=   16777215
                  OutsideHardLimitBackground=   8421631
                  OutsideSoftLimitBackground=   8454143
                  HardMax         =   99999
                  HardMin         =   0
                  SoftMax         =   -999
                  SoftMin         =   0
                  MaxWidth        =   -999
                  Alignment       =   2
                  DataType        =   0
                  DefaultValue    =   ""
                  Value           =   ""
                  Enabled         =   -1  'True
               End
               Begin VB.OptionButton rdoDO 
                  Caption         =   "State Total (GW/SW)"
                  Height          =   252
                  Index           =   1
                  Left            =   120
                  TabIndex        =   43
                  Top             =   480
                  Width           =   2292
               End
               Begin VB.OptionButton rdoDO 
                  Caption         =   "by Unit Area (GW/SW)"
                  Height          =   252
                  Index           =   0
                  Left            =   120
                  TabIndex        =   42
                  Top             =   240
                  Width           =   2300
               End
               Begin VB.Label lblDOTotalSW 
                  Alignment       =   2  'Center
                  Caption         =   "SW Value"
                  Height          =   252
                  Left            =   1440
                  TabIndex        =   83
                  ToolTipText     =   "Enter the state total for DO - surface water withdrawals"
                  Top             =   1020
                  Visible         =   0   'False
                  Width           =   852
               End
               Begin VB.Label lblDOTotalGW 
                  Alignment       =   2  'Center
                  Caption         =   "GW Value"
                  Height          =   252
                  Left            =   240
                  TabIndex        =   82
                  ToolTipText     =   "Enter the state total for DO - groundwater withdrawals"
                  Top             =   1020
                  Visible         =   0   'False
                  Width           =   852
               End
            End
            Begin VB.Frame fraIR 
               Caption         =   "Irrigation"
               Height          =   775
               Left            =   120
               TabIndex        =   68
               Top             =   2520
               Width           =   2532
               Begin VB.OptionButton rdoIR 
                  Caption         =   "Total"
                  Height          =   252
                  Index           =   1
                  Left            =   120
                  TabIndex        =   47
                  Top             =   480
                  Width           =   1572
               End
               Begin VB.OptionButton rdoIR 
                  Caption         =   "Crops/Golf"
                  Height          =   252
                  Index           =   0
                  Left            =   120
                  TabIndex        =   46
                  Top             =   240
                  Width           =   1572
               End
            End
            Begin VB.CommandButton cmdAddYear 
               Caption         =   "Execute"
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
               Left            =   4320
               TabIndex        =   49
               Top             =   2760
               Width           =   855
            End
            Begin VB.Frame fraNewUnitArea 
               Caption         =   "Unit Area"
               Height          =   576
               Left            =   1920
               TabIndex        =   38
               Top             =   250
               Width           =   3375
               Begin VB.OptionButton rdoNewAreaUnit 
                  Caption         =   "Aquifer"
                  Height          =   252
                  Index           =   2
                  Left            =   2280
                  TabIndex        =   41
                  Top             =   240
                  Width           =   975
               End
               Begin VB.OptionButton rdoNewAreaUnit 
                  Caption         =   "HUC - 8"
                  Height          =   252
                  Index           =   1
                  Left            =   1200
                  TabIndex        =   40
                  Top             =   240
                  Width           =   1000
               End
               Begin VB.OptionButton rdoNewAreaUnit 
                  Caption         =   "County"
                  Height          =   252
                  Index           =   0
                  Left            =   120
                  TabIndex        =   39
                  Top             =   240
                  Width           =   950
               End
            End
            Begin ATCoCtl.ATCoText txtNewYear 
               Height          =   255
               Left            =   3000
               TabIndex        =   48
               Top             =   3000
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               InsideLimitsBackground=   16777215
               OutsideHardLimitBackground=   8421631
               OutsideSoftLimitBackground=   8454143
               HardMax         =   2100
               HardMin         =   1980
               SoftMax         =   2100
               SoftMin         =   1900
               MaxWidth        =   -999
               Alignment       =   1
               DataType        =   0
               DefaultValue    =   ""
               Value           =   ""
               Enabled         =   -1  'True
            End
            Begin VB.Label lblNewYear 
               Alignment       =   1  'Right Justify
               Caption         =   "Year of Data: "
               Height          =   255
               Left            =   3000
               TabIndex        =   62
               Top             =   2760
               Width           =   975
            End
         End
         Begin VB.OptionButton MainOpt 
            Caption         =   "Create &New Year of Data"
            Height          =   252
            Index           =   11
            Left            =   240
            TabIndex        =   15
            Top             =   4440
            Width           =   2892
         End
         Begin VB.Label lblEdit 
            Caption         =   "Data Entry and Editing"
            Height          =   252
            Left            =   120
            TabIndex        =   60
            Top             =   480
            Width           =   1812
         End
         Begin VB.Label lblReports 
            Caption         =   "Reports"
            Height          =   252
            Left            =   120
            TabIndex        =   59
            Top             =   1320
            Width           =   1812
         End
         Begin VB.Label lblQC 
            Caption         =   "Quality-Control Reports"
            Height          =   252
            Left            =   120
            TabIndex        =   58
            Top             =   2880
            Width           =   1812
         End
         Begin VB.Label lblOther 
            Caption         =   "Other"
            Height          =   252
            Left            =   120
            TabIndex        =   57
            Top             =   3960
            Width           =   1812
         End
      End
      Begin VB.Frame fraAreaUnits 
         Caption         =   "Area Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   3732
         Begin VB.OptionButton rdoAreaUnit 
            Caption         =   "County"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   852
         End
         Begin VB.OptionButton rdoAreaUnit 
            Caption         =   "HUC - 8"
            Height          =   252
            Index           =   1
            Left            =   960
            TabIndex        =   33
            Top             =   360
            Width           =   972
         End
         Begin VB.OptionButton rdoAreaUnit 
            Caption         =   "HUC - 4"
            Height          =   252
            Index           =   2
            Left            =   1920
            TabIndex        =   35
            Top             =   360
            Width           =   972
         End
         Begin VB.OptionButton rdoAreaUnit 
            Caption         =   "Aquifer"
            Height          =   252
            Index           =   3
            Left            =   2880
            TabIndex        =   37
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.Frame fraStSel 
         Caption         =   "State"
         Height          =   4092
         Left            =   -74760
         TabIndex        =   0
         Top             =   600
         Width           =   3012
         Begin VB.ListBox lstStates 
            Height          =   2595
            Left            =   120
            MultiSelect     =   1  'Simple
            TabIndex        =   2
            Top             =   240
            Width           =   2772
         End
      End
      Begin VB.Frame fraDomainName 
         Caption         =   "DomainName"
         Height          =   612
         Left            =   -71640
         TabIndex        =   29
         ToolTipText     =   "Enter domain name of user machine"
         Top             =   4560
         Visible         =   0   'False
         Width           =   2415
         Begin VB.CommandButton cmdDomain 
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Left            =   1800
            TabIndex        =   90
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtDomainName 
            Height          =   288
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   55
            TabIndex        =   30
            Text            =   "NT"
            Top             =   240
            Width           =   1572
         End
      End
      Begin VB.Frame fraCats 
         Caption         =   "Data Categories"
         Height          =   4932
         Left            =   -74880
         TabIndex        =   28
         Top             =   360
         Width           =   8892
         Begin VB.CommandButton cmdCatOpt 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   972
            Left            =   3480
            TabIndex        =   80
            Top             =   3600
            Width           =   1332
         End
         Begin Awuds.ATCoSelectListSortByProp lstDataCats 
            Height          =   4572
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   8652
            _ExtentX        =   15266
            _ExtentY        =   8070
         End
      End
      Begin VB.CommandButton cmdExeOpts 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   7560
         TabIndex        =   27
         Top             =   4440
         Width           =   1332
      End
      Begin VB.Frame fraDataFlds 
         Height          =   4932
         Left            =   -74880
         TabIndex        =   16
         Top             =   360
         Width           =   8892
         Begin VB.TextBox txtDataFlds 
            BackColor       =   &H00FFFFFF&
            Height          =   288
            Index           =   0
            Left            =   2160
            TabIndex        =   97
            Top             =   360
            Width           =   972
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete All Fields in Category"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   662
            Left            =   120
            TabIndex        =   17
            Top             =   4200
            Width           =   1092
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   612
            Left            =   7680
            TabIndex        =   19
            ToolTipText     =   "Change values in database to reflect current edits"
            Top             =   4200
            Width           =   1092
         End
         Begin VB.CommandButton cmdReturn 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   612
            Left            =   6240
            TabIndex        =   18
            ToolTipText     =   "Return to Category tab without saving changes"
            Top             =   4200
            Width           =   1092
         End
         Begin VB.Label lblDataFlds 
            Alignment       =   1  'Right Justify
            Height          =   288
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1932
         End
         Begin VB.Label lblDataUnits 
            Height          =   288
            Index           =   0
            Left            =   3240
            TabIndex        =   25
            Top             =   360
            Width           =   5532
         End
      End
      Begin VB.CommandButton cmdNationalDB 
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
         Left            =   -74640
         TabIndex        =   3
         Top             =   4800
         Width           =   2772
      End
      Begin VB.Frame fraAreaUnits2 
         Caption         =   "Second Area Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Left            =   5280
         TabIndex        =   20
         Top             =   360
         Width           =   3732
         Begin VB.OptionButton rdoAreaUnit2 
            Caption         =   "County"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   852
         End
         Begin VB.OptionButton rdoAreaUnit2 
            Caption         =   "HUC - 8"
            Height          =   252
            Index           =   1
            Left            =   960
            TabIndex        =   22
            Top             =   360
            Width           =   972
         End
         Begin VB.OptionButton rdoAreaUnit2 
            Caption         =   "HUC - 4"
            Height          =   252
            Index           =   2
            Left            =   1920
            TabIndex        =   23
            Top             =   360
            Width           =   972
         End
         Begin VB.OptionButton rdoAreaUnit2 
            Caption         =   "Aquifer"
            Height          =   252
            Index           =   3
            Left            =   2880
            TabIndex        =   24
            Top             =   360
            Width           =   825
         End
      End
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
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
         Left            =   3600
         TabIndex        =   67
         Top             =   4080
         Width           =   1452
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   396
      Left            =   0
      TabIndex        =   65
      Top             =   0
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   372
      Left            =   4200
      TabIndex        =   66
      Top             =   3120
      Width           =   972
   End
End
Attribute VB_Name = "frmAwuds2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Copyright 2000 by AQUA TERRA Consultants

' ##MODULE_NAME frmAwuds2
' ##MODULE_DATE December 12, 2003
' ##MODULE_AUTHOR Robert Dusenbury of AQUA TERRA CONSULTANTS
' ##MODULE_SUMMARY This form provides the main interface that allows the user to access _
          the state and national databases.

Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

Private Sub cdmAbout_Click()
  frmAbout.Show 0, Me
End Sub

Private Sub cmdBrowse_Click()
Attribute cmdBrowse_Click.VB_Description = "Raises commom dialog box so user can browse for&nbsp;Excel spreadsheet to be imported."
' ##SUMMARY Raises commom dialog box so user can browse for&nbsp;Excel spreadsheet to be _
          imported.
' ##HISTORY 5/21/2007, prhummel Use file filter/extension determined from version of Excel

  On Error GoTo x
  
  With cdlgFileSel
    .DialogTitle = "Select a file for import"
    .filename = ReportPath & "Import" & XLFileExt '.xls"
    .Filter = XLFileFilter '"(*.xls)|*.xls"
    .FilterIndex = 1
    .CancelError = True
    .ShowOpen
    If Len(Dir(.filename)) > 0 Then txtCurFile.Text = .filename
  End With
  MyP.Year1Opt = txtYear.value
x:
End Sub

Private Sub cmdDataPath_Click()
  '##SUMMARY Allows user to browse for data directory.
  '##DATE November 29, 2005
  '##AUTHOR Mark Gray and Robert Dusenbury, AQUA TERRA CONSULTANTS
  '##REMARKS Allows AwudsDataPath variable to be reset.  The user is prompted _
    to browse for an appropriate directory.
  '##ON_ERROR Errors will be raised if: AwudsDataPath is not the correct _
    path to a directory containing the required databases and report templates.
  SetDataPath
  If txtDataPath.Text <> AwudsDataPath Then
    txtDataPath.Text = AwudsDataPath
    ClearOpts
    MyP.StateDBClose
    ListStates
    If lstStates.SelCount = 0 Then
      tabMain.Tab = 0
      tabMain.TabEnabled(1) = False
    End If
  End If
End Sub

Private Sub cmdDelete_Click()
Attribute cmdDelete_Click.VB_Description = "Deletes all data in selected category when user editing data."
  Dim response As Long 'records user response to ATCoMessageBox
' ##SUMMARY Deletes all data in selected category when user editing data.
' ##REMARKS Button titled 'Delete All Fields in Category' displayed on 5th tab of main _
          form.
  
  response = MyMsgBox.Show("Are you certain you want to delete all of the data fields" & _
      vbCrLf & "for the category '" & lstDataCats.RightItem(0) & "' in " & _
      lstArea.RightItem(0) & vbCrLf & "for the year " & MyP.Year1Opt & "?", _
      "User Action Verification", "+&Cancel", "-&Yes")
  If response = 2 Then  'delete all records for the selected category
    DeleteOK = True
    cmdEdit_Click
    DeleteOK = False
    CmdReturn_Click
  End If
End Sub

Private Sub cmdDomain_Click()
Attribute cmdDomain_Click.VB_Description = "Submits domain name for current computer."
' ##SUMMARY Submits domain name for current computer.
' ##REMARKS User types domain name, which is needed to set user access,&nbsp;into text _
          box 'txtDomainName'.
  Dim DomainName As String  'string holds user-entered domain name
  Dim response As Long      'records user response to ATCoMessageBox
  Dim v As AtcoValidateUser 'Class Module object
  
  Set v = New AtcoValidateUser
  
  DomainName = txtDomainName.Text
'  UserAccess = v.GetUserAccess("AWUDSWW", "AWUDSWR", DomainName)
  UserAccess = v.GetUserAccess()
  If UserAccess = "" Then
tryagain:
    response = MyMsgBox.Show("The domain name was not successful." & _
        vbCrLf & "Re-enter a domain name, continue with read-only access, or exit the program." & _
        vbCrLf & "The system administrator should know the domain name.", _
        "Bad Domain Name", "+&Re-enter", "-&Read-only", "-&Quit")
    If response = 2 Then
      UserAccess = "Read"
      fraDomainName.Visible = False
      Ok2DoMore
    ElseIf response = 3 Then
      cmdExit_Click
    End If
    Exit Sub
  Else
    If UserAccess = "WRITE" Then 'user may edit database
      MainOpt(0).Enabled = True
      MainOpt(1).Enabled = True
      MainOpt(10).Enabled = True
      MainOpt(11).Enabled = True
    End If
    fraDomainName.Visible = False
    SaveSetting "AWUDS", "Defaults", "DomainName", DomainName
    Ok2DoMore
  End If
  
End Sub

Private Sub cmdEdit_Click()
Attribute cmdEdit_Click.VB_Description = "Saves edits made to data in selected category to database."
' ##SUMMARY Saves edits made to data in selected category to database.
' ##REMARKS Button titled 'Save' displayed on the 5th tab of the main form.
  Dim i As Long
  Dim j As Long
  Dim response As Long
  Dim sql As String
  Dim locn As String
  Dim stRec As Recordset
  Dim editRec As Recordset
  Dim areaRec As Recordset
  Dim bumValue As Boolean
  
  On Error GoTo x
  
  response = 0
  
  If Left(lstArea.RightItem(0), 3) = "000" Then  'editing state values
    'Use SQL language to create recordset with all unit areas (i.e., counties,
    ' hucs, or aquifers) within selected state.
    ' For example:
    '  SELECT County_cd, County_nm From [County]
    '  WHERE state_cd='09' AND Len(Trim(County_cd))=3;
    sql = "SELECT " & TableName & "_cd, " & TableName & "_nm " & _
          "From [" & TableName & _
          "] WHERE state_cd='" & MyP.stateCode & _
          "' AND Len(Trim(" & TableName & "_cd))=" & MyP.Length & ";"
    Set areaRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    areaRec.MoveLast
    areaRec.MoveFirst
    'Use SQL language to create recordset with all data for selected year and unit area type.
    ' For example:
    '  Select * from [CountyData]
    '  Where Date=2000
    '  ORDER BY FieldID, Location;
    sql = "Select * from [" & MyP.AreaTable & _
          "] Where Date=" & MyP.Year1Opt & _
          " ORDER BY FieldID, Location;"
    Set stRec = MyP.stateDB.OpenRecordset(sql, dbOpenDynaset)
    'Use SQL language to create recordset with distinct set of fields from
    ' the Data Dictionary for the selected category.
    sql = "Select DISTINCT FieldID from [LastEdit]" & _
          " where CategoryID=" & lstDataCats.RightItemData(0)
    Set editRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    editRec.MoveLast
    editRec.MoveFirst
    With stRec
      For i = 1 To editRec.RecordCount
        .FindFirst "FieldID=" & editRec("FieldID")
        'State level totals are stored in data field for each unit area.
        ' When retrieved, value for any one unit area represents state total.
        For j = 1 To areaRec.RecordCount  'change values one area at a time
          .Edit
          If DeleteOK Or Len(Trim((txtDataFlds(i - 1)))) = 0 Then
            !value = Null
          ElseIf IsNumeric(txtDataFlds(i - 1)) Then
            If txtDataFlds(i - 1) < 0 Then
              MyMsgBox.Show _
                  "The value for " & lblDataFlds(i - 1) & " " & vbCrLf _
                  & "is a negative number and will not be entered in the database." _
                  & vbCrLf & vbCrLf & "All data values must be positive numbers.", _
                  "Bad data value", "+-&OK"
              bumValue = True
              Exit For
            Else
              !value = txtDataFlds(i - 1)
            End If
          Else
            MyMsgBox.Show _
                "The value for " & lblDataFlds(i - 1) & " " & vbCrLf _
                & "is non-numeric and will not be entered in the database.", _
                "Bad data value", "+-&OK"
            bumValue = True
            Exit For
          End If
          .Update
          .MoveNext
        Next j
        editRec.MoveNext
      Next i
    End With
  Else  'editing county/huc/aquifer values
    'Use SQL language to create recordset with all data for
    ' selected year and unit area (i.e., county, huc, or aquifer).
    ' For example:
    '  Select * from [CountyData]
    '  Where Date=2000 AND Location='005'
    sql = "Select * from [" & MyP.AreaTable & _
          "] Where Date=" & MyP.Year1Opt & _
          " AND Location='" & LocnArray(0, 0) & "'"
    Set stRec = MyP.stateDB.OpenRecordset(sql, dbOpenDynaset)
    'Use SQL language to create recordset with all records for selected
    ' category from "LastEdit" table in Categories.mdb
    ' For example:
    '  Select * from [LastEdit] Where CategoryID=2
    sql = "Select * from [LastEdit" & _
          "] Where CategoryID=" & lstDataCats.RightItemData(0)
    Set editRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    editRec.MoveLast
    editRec.MoveFirst
    For i = 0 To editRec.RecordCount - 1
      stRec.FindFirst "FieldID=" & editRec("FieldID")
      If txtDataFlds(i) <> stRec("Value") Or IsNull(stRec("Value")) Then
        If EditOK = False And DeleteOK = False Then
          response = MyMsgBox.Show("Are you certain you want to overwrite" & _
              vbCrLf & "the data fields that you have edited?", _
              "User Action Verification", "+&Yes", "-&Cancel")
          If response = 1 Then
            EditOK = True
          Else
            i = editRec.RecordCount - 1
          End If
        End If
        If EditOK = True Then
          With stRec
            .Edit
            If Len(Trim((txtDataFlds(i)))) = 0 Then
              !value = Null
            ElseIf IsNumeric(txtDataFlds(i)) Then
              If txtDataFlds(i) < 0 And !FieldID <> 218 Then
                MyMsgBox.Show _
                    "The value for " & lblDataFlds(i) & " " & vbCrLf _
                    & "is a negative number and will not be entered in the database." _
                    & vbCrLf & vbCrLf & "All data values must be positive numbers.", _
                    "Bad data value", "+-&OK"
                bumValue = True
              Else
                !value = txtDataFlds(i)
              End If
            Else
              MyMsgBox.Show _
                  "The value for " & lblDataFlds(i) & " " & vbCrLf _
                  & "is non-numeric and will not be entered in the database.", _
                  "Bad data value", "+-&OK"
              bumValue = True
            End If
            .Update
          End With
        End If
      End If
      If DeleteOK = True Then
        With stRec
          .Edit
          !value = Null
          txtDataFlds(i) = ""
          .Update
        End With
      End If
      editRec.MoveNext
    Next i
  End If
  stRec.Close
  editRec.Close
  cmdEdit.Enabled = False
  If Not bumValue Then CmdReturn_Click
x:
End Sub

Private Sub cmdExeOpts_Click()
Attribute cmdExeOpts_Click.VB_Description = "This routine logs the unit area, label, and year&nbsp;selections made by the user on the third tab"
' ##SUMMARY This routine logs the unit area, label, and year&nbsp;selections made by the _
          user on the third tab
' ##REMARKS Creates array with all selected unit areas/states. Also creates preliminary _
          report table containing user-selected data. Calls 3 daughter subroutines: SetAreas, _
          CreateTable, and FillCats.
  
  Me.MousePointer = vbHourglass
  lstDataCats.ClearRight
  lstDataCats.ClearLeft
  IRinTwo = False
  IRinTwo2 = False
  TwoAreas = False
  TwoYears = False
  MyP.DataOpt = 0
  MyP.DataOpt2 = 0

  If lstYears.SelCount = 1 Then
    Years = "(" & TableName & "Data.Date=" & MyP.Year1Opt & ")"
  Else
    Years = "(" & TableName & "Data.Date=" & MyP.Year1Opt & " Or " & TableName & "Data.Date=" & MyP.Year2Opt & ")"
  End If

  If (MyP.Length = 3 Or MyP.Length = 10) And MyP.UserOpt <> 9 And _
      Not NationalDB And Left(lstArea.RightItem(0), 3) <> "000" Then CheckAreas

  SetAreas
  CreateTable
  FillCats
  If lstDataCats.LeftCount > 0 Or lstDataCats.RightCount > 0 Then
    tabMain.Tab = 3
    tabMain.TabEnabled(3) = True
  End If
  Me.MousePointer = vbDefault
End Sub

Private Sub SetAreas()
Attribute SetAreas.VB_Description = "Creates array with names and codes of user-selected areas.&nbsp; Also sets module variable Areas&nbsp;to be used as part of sql when creating report table in database."
' ##REMARKS First daughter subroutine called by cmdExeOpts_Click.
' ##SUMMARY Creates array with names and codes of user-selected areas.&nbsp; Also sets _
          module variable Areas&nbsp;to be used as part of sql when creating report table _
          in database.
' ##HISTORY 5/21/2007, prhummel Adjusted conditional to include all unit areas _
            if State Totals selected for compare years report (MyP.UserOpt=10)
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim selArea As String
  Dim sql As String
  Dim tabName As String
  Dim hucSize As String
  Dim strSoFar As String
  Dim areaRec As Recordset
  Dim areaRec2 As Recordset

  If NationalDB Then
    'Set Areas to be used as part of sql when creating report table in database
    If Left(MyP.UnitArea, 1) = "H" Then
      Areas = " AND Len(Trim(" & TableName & "Data.Location))=" & MyP.Length
    Else
      Areas = ""
    End If
    'create location array with all of the selected areas if National DB
    ReDim LocnArray(2, lstStates.SelCount - 1)
    k = 0
    For i = 0 To lstStates.ListCount - 1
      If lstStates.Selected(i) Then
        strSoFar = CStr(lstStates.ItemData(i))
        If Len(strSoFar) < 2 Then strSoFar = "0" & strSoFar
        LocnArray(0, k) = strSoFar
        strSoFar = ""
        selArea = Trim(lstStates.List(i)) & "  "
        j = 1
        While j <= Len(Trim(selArea))
        strSoFar = strSoFar & " " & UCase(Left(StrSplit(Mid(selArea, j), " ", ""), 1)) & _
                    LCase(Mid(StrSplit(Mid(selArea, j), " ", ""), 2))
          j = j + InStr(1, Mid(selArea, j), " ")
        Wend
        LocnArray(1, k) = Trim(strSoFar)
        LocnArray(2, k) = LocnArray(1, k) & " - " & LocnArray(0, k)
        k = k + 1
      End If
    Next i
  ElseIf (Left(lstArea.RightItem(0), 3) = "000" And _
          (MyP.UserOpt = 5 Or MyP.UserOpt = 10)) Or _
          lstArea.LeftCount = 0 Then 'all unit areas selected, include compare years (opt=10)
    If Left(MyP.UnitArea, 1) = "H" Then
      hucSize = " And Len(Trim(huc_cd))=" & MyP.Length
    Else
      hucSize = ""
    End If
    Areas = " And Len(Trim(" & TableName & "Data.Location))=" & MyP.Length
    tabName = TableName
    If MyP.UserOpt = 9 Then j = 2 Else j = 1
    For i = 1 To j
      If i = 1 Then
        'Use SQL language to create recordset of unit area codes and names
        ' from the "county", "huc", "aquifer", or "state" table in General.mdb.
        ' Certain counties and aquifers were added or eliminated recently so
        ' "begin" and "end" attributes were added to exclude such areas from
        ' the query for years when they did not exist.
        ' For example:
        '   SELECT [County].County_cd, [County].County_nm From [County]
        '   WHERE [County].state_cd='10'
        '     AND (([county].begin<=0 OR IsNull([county].begin))
        '     AND ([county].end>=0 OR IsNull([county].end)))
        '   ORDER BY County_cd;
        If MyP.Length = 3 Or MyP.Length = 10 Then
          sql = "SELECT [" & tabName & "]." & tabName & "_cd, [" & _
              tabName & "]." & tabName & "_nm " & "From [" & tabName & _
              "] WHERE [" & tabName & "].state_cd='" & MyP.stateCode & "'" & _
              " AND (([" & tabName & "].begin<=" & MyP.Year1Opt & " OR IsNull([" & tabName & "].begin))" & _
              " AND ([" & tabName & "].end>=" & MyP.Year1Opt & " OR IsNull([" & tabName & "].end)))" & _
              hucSize & " ORDER BY " & tabName & "_cd;"
        Else
          sql = "SELECT [" & tabName & "]." & tabName & "_cd, [" & _
              tabName & "]." & tabName & "_nm " & "From [" & tabName & _
              "] WHERE [" & tabName & "].state_cd='" & MyP.stateCode & "'" & _
              hucSize & " ORDER BY " & tabName & "_cd;"
        End If
        Set areaRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
        areaRec.MoveLast
        areaRec.MoveFirst
      ElseIf i = 2 Then  'comparing by areas
        If Left(MyP.UnitArea2, 1) = "H" Then
          hucSize = " And Len(Trim(" & TableName2 & "_cd))=" & MyP.length2
        Else
          hucSize = ""
        End If
        tabName = TableName2
        If tabName <> "" Then
          If MyP.length2 = 3 Or MyP.length2 = 10 Then
            sql = "SELECT [" & tabName & "]." & tabName & "_cd, [" _
                & tabName & "]." & tabName & "_nm " & "From [" & tabName & _
                "] WHERE [" & tabName & "].state_cd='" & MyP.stateCode & "'" & _
                " AND (([" & tabName & "].begin<=" & MyP.Year1Opt & " OR IsNull([" & tabName & "].begin))" & _
                " AND ([" & tabName & "].end>=" & MyP.Year1Opt & " OR IsNull([" & tabName & "].end)))" & _
                hucSize & " ORDER BY " & tabName & "_cd;"
          Else
            sql = "SELECT [" & tabName & "]." & tabName & "_cd, [" _
                & tabName & "]." & tabName & "_nm " & "From [" & tabName & _
                "] WHERE [" & tabName & "].state_cd='" & MyP.stateCode & "'" & _
                hucSize & " ORDER BY " & tabName & "_cd;"
          End If
          Set areaRec2 = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
          areaRec2.MoveLast
          areaRec2.MoveFirst
        End If
      End If
    Next i
    If MyP.UserOpt = 9 And tabName <> "" Then  'comparing by areas
      If areaRec.RecordCount > areaRec2.RecordCount Then
        j = areaRec.RecordCount - 1
      Else
        j = areaRec2.RecordCount - 1
      End If
      i = 5
    Else
      i = 2
      j = areaRec.RecordCount - 1
    End If
    'Create location array with all of the selected areas
    ReDim LocnArray(i, j)
    For i = 0 To areaRec.RecordCount - 1
      LocnArray(0, i) = Trim(areaRec(TableName & "_cd"))
      If Not IsNull(Trim(areaRec(TableName & "_nm"))) Then
        LocnArray(1, i) = Trim(areaRec(TableName & "_nm"))
      Else
        LocnArray(1, i) = ""
      End If
      LocnArray(2, i) = LocnArray(1, i) & " - " & LocnArray(0, i)
      areaRec.MoveNext
    Next i
    If MyP.UserOpt = 9 And tabName <> "" Then  'comparing by areas
      For i = 0 To areaRec2.RecordCount - 1
        LocnArray(3, i) = areaRec2(TableName2 & "_cd")
        LocnArray(4, i) = Trim(areaRec2(TableName2 & "_nm"))
        LocnArray(5, i) = LocnArray(4, i) & " - " & LocnArray(3, i)
        areaRec2.MoveNext
      Next i
    End If
  Else
    For i = 0 To lstArea.RightCount - 1
      'Set Areas to be used as part of sql when creating report table in database
      If i = 0 Then
        Areas = " AND ([" & TableName & "Data].Location = '" & Left(lstArea.RightItem(i), MyP.Length) & "'"
      Else
        Areas = Areas & " Or [" & TableName & "Data].Location = '" & Left(lstArea.RightItem(i), MyP.Length) & "'"
      End If
      If i = lstArea.RightCount - 1 Then Areas = Areas & ")"
    Next i
    If lstArea.RightCount > 0 Then ReDim LocnArray(2, lstArea.RightCount - 1)
    For i = 0 To lstArea.RightCount - 1
      j = InStr(1, lstArea.RightItem(i), "-")
      LocnArray(0, i) = Left(lstArea.RightItem(i), j - 2)
      LocnArray(1, i) = Mid(lstArea.RightItem(i), j + 2)
      LocnArray(2, i) = Trim(lstArea.RightItem(i))
    Next i
  End If

End Sub

Private Sub cmdExit_Click()
Attribute cmdExit_Click.VB_Description = "Closes the database then unloads main form, closing AWUDS"
' ##SUMMARY Closes the database then unloads main form, closing AWUDS
' ##REMARKS cmdExit located on lower portion of main form, always in view.
  Dim compactDB As Database
  
  On Error GoTo x
  
  If Len(Dir(AwudsDataPath & MyP.stateCode & ".mdb")) > 0 Then MyP.stateDB.Close
x:
  Unload Me
End Sub

Private Sub cmdHelp_Click()
Attribute cmdHelp_Click.VB_Description = "Raises new window containing CHM help file."
' ##SUMMARY Raises new window containing CHM help file.
' ##REMARKS cmdHelp located on lower portion of main form, always in view.
  Dim helpFilename As String
  
  On Error GoTo ErrTrap
    
  helpFilename = AwudsDataPath & "AWUDS.chm"
  If Len(helpFilename) = 0 Then GoTo TryAnotherName
  If Len(Dir(helpFilename)) > 0 Then 'CHM file in data directory
    'All 3 following calls work, but last 2 may be less prone to crashing than 1st.
'    OpenFile helpFilename, cdlgFileSel
'    Shell "hh.exe " & helpFilename, vbNormalFocus
    WinExec "hh.exe " & helpFilename, 1
  Else
TryAnotherName:
    helpFilename = ExePath & "AWUDS.chm"
    If Len(Dir(helpFilename)) > 0 Then 'CHM file in EXE directory
      'All 3 following calls work, but last 2 may be less prone to crashing than 1st.
'      OpenFile helpFilename, Me.cdlgFileSel
'      Shell "hh.exe " & helpFilename, vbNormalFocus
      WinExec "hh.exe " & helpFilename, 1
    Else 'can not find CHM file
      MsgBox "Help file not available"
    End If
  End If
ErrTrap:
End Sub

Private Sub cmdImport_Click()
Attribute cmdImport_Click.VB_Description = "Begins the import process for a selected Excel spreadsheet."
' ##SUMMARY Begins the import process for a selected Excel spreadsheet.
' ##REMARKS Calls XLImport in module&nbsp;ImpExp.bas.
' ##HISTORY 5/21/2007, prhummel Use file filter/extension determined from version of Excel

'  If LCase(Right(txtCurFile, 4)) <> ".xls" And NewImpFile = True Then
  If LCase(Right(txtCurFile, 4)) <> Right(XLFileExt, 4) And NewImpFile = True Then
    MyMsgBox.Show _
        "The import file must be an Excel(" & XLFileExt & ") file with the proper format." & vbCrLf _
        & "See 'Import' in the on-line help for more information.", _
        "Import file type verification", "+-&OK"
    Exit Sub
  Else
    If Not IsNumeric(txtYear.value) Then
        MyMsgBox.Show _
            "The year of import must be between 1800 and 2100." & vbCrLf & _
            "You have entered a non-numeric value", "Invalid Year", "+-&OK"
        Exit Sub
    End If
    MyP.Year1Opt = txtYear.value
    If Len(Dir(txtCurFile.Text)) > 0 Then
      If MyP.Year1Opt > 1800 And MyP.Year1Opt < 2101 Then
        XLImport txtCurFile.Text
      Else
        MyMsgBox.Show _
            "The year of import must be between 1800 and 2100.", _
            "Invalid Year", "+-&OK"
        Exit Sub
      End If
    Else
      MyMsgBox.Show _
          "Could not locate " & txtCurFile.Text & "." & vbCrLf _
          & "Click on the Browse button and select an existing file.", _
          "Can not find import file", "+-&OK"
      Exit Sub
    End If
    fraImport.Visible = False
    MainOpt(1) = False
    MainOpt(1).Font.Bold = False
    MyP.Year1Opt = 0
    txtYear.value = -999
  End If
End Sub

Private Sub cmdNationalDB_Click()
Attribute cmdNationalDB_Click.VB_Description = "Alternates between selection of National or individual state database."
' ##SUMMARY Alternates between selection of National or individual state database.
' ##REMARKS National database is created at runtime by aggregating all state databases.
  Dim i As Long
  
  If NationalDB = False Then
    If Len(Dir(AwudsDataPath & "Nation.mdb")) = 0 Then
      MsgBox "There is no 'Nation.mdb' file in your Data Directory (" & AwudsDataPath & ")." & vbCrLf & vbCrLf & _
          "You must either download or copy the current version of the 'Nation.mdb' Access database" & vbCrLf & _
          "into the Data Directory before working with the National Database options of AWUDS." _
          , vbCritical
      Exit Sub
    End If
    NationalDB = True
    MainOpt(0).Enabled = False
    MainOpt(1).Enabled = False
    'export of national DB now allowed, PRH 11/2005
    'MainOpt(10).Enabled = False
    MainOpt(11).Enabled = False
    Ok2DoMore
    MyP.StateStuff "United States", "Nation"
  Else
    lstStates.ListIndex = -1
    NationalDB = False
    MyP.State = ""
    If UserAccess = "WRITE" Then
      MainOpt(0).Enabled = True
      MainOpt(1).Enabled = True
      MainOpt(10).Enabled = True
      MainOpt(11).Enabled = True
    End If
  End If

  ClearOpts
  If Not NationalDB Then
    tabMain.TabEnabled(2) = False
    tabMain.TabEnabled(3) = False
    tabMain.TabEnabled(4) = False
  End If
  MyP.UserOpt = 0
  fraImport.Visible = False
  fraNewYear.Visible = False
  For i = 0 To MainOpt.Count - 1
    MainOpt(i) = False
  Next i
  EmboldenMe MainOpt, -1
  
  ListStates
  Ok2DoMore
End Sub

Private Sub cmdRetrieve_Click()
Attribute cmdRetrieve_Click.VB_Description = "Retrieves frequently selected set of area units (counties, HUCs, or aquifers) for certain year(s) previously saved by user."
' ##SUMMARY Retrieves frequently selected set of area units (counties, HUCs, or aquifers) _
          for certain year(s) previously saved by user.
' ##REMARKS Set of user selections are stored in an ascii file; default directory is 'AWUDSReports'.
  Dim fileTitle As String
  Dim textLine As String
  Dim dataStrng As String
  Dim InFile As Integer
  Dim i As Long

  On Error GoTo x

  With cdlgFileSel
    .DialogTitle = "Select name of saved selection group"
    .filename = ReportPath & "Selection Group.txt"
    .Filter = "(*.txt)|*.txt|All Files|*.*"
    .FilterIndex = 1
    .CancelError = True
    .ShowOpen
    fileTitle = .filename
  End With
  
  InFile = FreeFile
  Open fileTitle For Input As InFile

  'erase all current selections
  tabMain.Tab = 1
  ClearOpts
  tabMain.Tab = 2
  
  Do While Not EOF(InFile)  'Loop until end of file.
    Line Input #InFile, textLine
    i = InStr(1, textLine, " ")
    dataStrng = Mid(textLine, i + 1)
    Select Case Left(textLine, 1)
      Case 3:
        rdoAreaUnit(CInt(dataStrng)) = True
      Case 4:
        rdoAreaUnit2(CInt(dataStrng)) = True
      Case 5:
        rdoID(CInt(dataStrng)) = True
      Case 6:
        For i = 0 To lstArea.LeftCount - 1
          If Left(lstArea.LeftItem(i), MyP.Length) = dataStrng Then
            lstArea.MoveRight (i)
            Exit For
          End If
        Next i
      Case 7:
        For i = 0 To lstYears.ListCount
          If lstYears.List(i) = dataStrng Then
            lstYears.Selected(i) = True
            Exit For
          End If
        Next i
        If Not (MyP.UserOpt = 10 And lstYears.SelCount = 1) _
            And lstArea.RightCount > 0 Then cmdExeOpts_Click
      Case 8:
        For i = 0 To lstDataCats.LeftCount - 1
          If lstDataCats.LeftItemData(i) = dataStrng Then
            lstDataCats.MoveRight (i)
            Exit For
          End If
        Next i
    End Select
  Loop
  
  Close InFile
x:
End Sub

Private Sub CmdReturn_Click()
Attribute CmdReturn_Click.VB_Description = "Returns from 5th to 4th tab on main form without saving edits made to data in selected category."
' ##SUMMARY Returns from 5th to 4th tab on main form without saving edits made to data in _
          selected category.
' ##REMARKS Button titled 'Cancel' displayed on the 5th tab of the main form.
  Me.MousePointer = vbHourglass
  tabMain.TabEnabled(4) = False
  tabMain.Tab = 3
  Me.MousePointer = vbDefault
End Sub

Private Sub cmdSaveGroup_Click()
Attribute cmdSaveGroup_Click.VB_Description = "Saves current set of user-selected area units (counties, HUCs, or aquifers) and year(s) for later retrieval."
' ##SUMMARY Saves current set of user-selected area units (counties, HUCs, or aquifers) _
          and year(s) for later retrieval.
' ##REMARKS Set of user selections are stored in an ascii file; default directory is 'AWUDSReports'.
  Dim fileTitle As String
  Dim OutFile As Integer
  Dim i As Long

  On Error GoTo x

  With cdlgFileSel
    .DialogTitle = "Select name for selection group as:"
    .filename = ReportPath & "Selection Group.txt"
    .Filter = "(*.txt)|*.txt|All Files|*.*"
    .FilterIndex = 1
    .CancelError = True
    .ShowSave
    fileTitle = .filename
  End With
  
  OutFile = FreeFile
  Open fileTitle For Output As OutFile
  Print #OutFile, "#Saved AWUDS selection group"
  For i = 0 To 3
    If rdoAreaUnit(i) Then
      Print #OutFile, "3UnitArea1 " & i
    End If
  Next i
  For i = 0 To 3
    If rdoAreaUnit2(i) Then
      Print #OutFile, "4UnitArea2 " & i
    End If
  Next i
  For i = 0 To 2
    If rdoID(i) Then
      Print #OutFile, "5AreaID " & i
    End If
  Next i
  For i = 0 To lstArea.RightCount - 1
    Print #OutFile, "6Areas " & Left(lstArea.RightItem(i), MyP.Length)
  Next i
  For i = 0 To lstYears.ListCount - 1
    If lstYears.Selected(i) Then Print #OutFile, "7Year(s) " & lstYears.List(i)
  Next i
  For i = 0 To lstDataCats.RightCount - 1
    Print #OutFile, "8Categories " & lstDataCats.RightItemData(i)
  Next i
  Close OutFile
x:
End Sub

Private Sub Form_Load()
' ##SUMMARY Loads main form.&nbsp; Reads domain name and sets user access.
' ##REMARKS If domain name unknown, displays text box for user entry.
  Dim lastState As String
  Dim DomainName As String
  Dim v As AtcoValidateUser
    
  Set MyMsgBox.Icon = Me.Icon
  tabMain.Tab = 0
  
  Me.Show
  
  Set v = New AtcoValidateUser
  DomainName = GetSetting("AWUDS", "Defaults", "Domain", "unknown")
  UserAccess = v.GetUserAccess()
'  UserAccess = "WRITE"
  If UserAccess = "" Then
    MsgBox "Enter the domain name of this computer."
    fraDomainName.Visible = True
  End If
  
  If UserAccess <> "WRITE" Then
    MainOpt(0).Enabled = False
    MainOpt(1).Enabled = False
    MainOpt(11).Enabled = False
  End If
  
  'determine Excel version and set properties for saved spreadsheets
  SetExcelProps
  
  'Write in Data path on first Tab
  txtDataPath.Text = AwudsDataPath
  
  Ok2DoMore
  ListStates
  If MyP.State = "" Then
    MsgBox "Please select a state in order to proceed."
  End If
 
End Sub

Private Sub ListStates()
Attribute ListStates.VB_Description = "Populates list box of states on 1st tab."
' ##SUMMARY Populates list box of states on 1st tab.
' ##REMARKS Automatically reselects last state to be selected, or, if NationalDB _
          selected, selects all states.
  Dim stRec As Recordset
  Dim stateCode As String
  Dim stName As String
  Dim lastState As String
  Dim i As Long
  
  If Not NationalDB Then
    lstStates.Clear
    cmdNationalDB.Caption = "Access National Database"
    'Populate the list box with the list of states
    lastState = GetSetting("AWUDS", "Defaults", "LastState", "unknown")
    Set stRec = MyP.GenDB.OpenRecordset("state", dbOpenSnapshot)
    stRec.MoveFirst
    While Not (stRec.EOF)
      stateCode = stRec("state_cd")
      If Len(Dir(AwudsDataPath & stateCode & ".mdb")) > 0 Then
        stName = stRec("state_nm")
        lstStates.AddItem stName
        lstStates.ItemData(lstStates.ListCount - 1) = stateCode
        If stName = UCase(lastState) Then
          lstStates.ListIndex = lstStates.ListCount - 1
          lstStates.Selected(lstStates.ListCount - 1) = True
        Else
          lstStates.Selected(lstStates.ListCount - 1) = False
        End If
      End If
      stRec.MoveNext
    Wend
    If lstStates.SelCount > 0 Then
      MyP.State = lastState
    Else
      MyP.State = ""
      SaveSetting "AWUDS", "Defaults", "LastState", ""
    End If
    stRec.Close
  Else
    cmdNationalDB.Caption = "Select Single State"
    For i = 0 To lstStates.ListCount - 1
      lstStates.Selected(i) = True
    Next i
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
' ##SUMMARY Unloads status monitor.
  AtcoLaunch1.SendMonitorMessage "(EXIT)"
End Sub

Private Sub lblInstructs_Change()
Attribute lblInstructs_Change.VB_Description = "Changes ToolTipText to display relevant instructions, depending upon location of cursor."
' ##SUMMARY Changes ToolTipText to display relevant instructions, depending upon location _
          of cursor.
  tabMain.ToolTipText = ReplaceString(lblInstructs.Caption, vbCrLf, "  ")
End Sub

Private Sub lstArea_Change()
Attribute lstArea_Change.VB_Description = "Ensures that only one HUC/county/aquifer is selected&nbsp;when editing data, and that multiple years of data are available when&nbsp;performing "
' ##SUMMARY Ensures that only one HUC/county/aquifer is selected&nbsp;when editing data, _
          and that multiple years of data are available when&nbsp;performing "Compare _
          Data for 2 Years" report.
' ##HISTORY 5/25/2007, prhummel Moved LoneArea from Global to local Static _
                       since it is only used here
  Dim i As Long
  Static LoneArea As String
  Static changing As Boolean
  
  If Not changing Then
    changing = True
    
    If MyP.UserOpt = 1 Then
      If lstArea.RightCount > 1 Then
        If LoneArea = Left(lstArea.RightItem(0), MyP.Length) Then i = 0 Else i = 1
        lstArea.MoveLeft (i)
      End If
      If lstArea.RightCount = 1 Then
        LoneArea = Left(lstArea.RightItem(0), MyP.Length)
      End If
    ElseIf MyP.UserOpt = 10 Then
      'State total selected, remove all others
      If lstArea.RightCount > 1 Then
        If Left(LoneArea, 3) = "000" Then
          'State total was on right, remove it
          lstArea.MoveLeft (0)
        ElseIf Left(lstArea.RightItem(0), 3) = "000" Then 'adding State total, remove all others
          While lstArea.RightCount > 1
            lstArea.MoveLeft (1)
          Wend
        End If
      End If
      LoneArea = Left(lstArea.RightItem(0), MyP.Length)
    End If
    If (MyP.Length = 3 Or MyP.Length = 10) And MyP.UserOpt <> 9 And lstYears.SelCount > 0 And _
        Left(lstArea.RightItem(0), 3) <> "000" Then CheckAreas
    changing = False
    If MyP.UserOpt = 10 And lstYears.ListCount < 2 And lstArea.RightCount > 0 Then
      MyMsgBox.Show _
          "There are not 2 available years of data to compare for" & vbCrLf & _
          Trim(lstArea.RightItem(lstArea.RightCount - 1)) & ".  Select another Area Unit.", _
          "Option Selection Problem", "+-&OK"
      lstArea.MoveLeft (lstArea.RightCount - 1)
    End If
    Ok2DoMore
  End If
End Sub

Private Sub lstStates_Click()
Attribute lstStates_Click.VB_Description = "Sets properties for newly selected state."
' ##SUMMARY Sets properties for newly selected state.
' ##REMARKS Sets 'LastState' entry in registry to current selection on 1st tab of main _
          form.
  Dim response As Long
  Dim i As Long
  Dim j As Long
  Dim dbPath As String
  Dim stCode As String
  Dim sql As String
  Dim selArea As String
  Dim strSoFar As String
  
  ClearOpts
  MyP.StateDBClose
  
  If NationalDB = False Then
    If lstStates.SelCount > 1 Then
      'clear previous selection
      i = CLng(MyP.stateCode)
      For j = 0 To lstStates.ListCount
        If lstStates.ItemData(j) = i Then
          lstStates.Selected(j) = False
          Exit For
        End If
      Next j
    End If
    If lstStates.SelCount = 0 Then
      MyP.State = ""
    Else
      stCode = lstStates.ItemData(lstStates.ListIndex)
      If Len(stCode) < 2 Then
        stCode = "0" & stCode
      End If
      dbPath = AwudsDataPath & stCode & ".mdb"
      If Len(Dir(dbPath)) = 0 Then
        On Error Resume Next
        MyMsgBox.Show "There is not a database for " & lstStates.List(lstStates.ListIndex) _
            & " in the '" & Mid(dbPath, 2, 5) & "' directory." & vbCrLf & _
            "The file would be titled '" & Right(dbPath, 6) & "'." & _
            vbCrLf & "Please select another state." _
            , "User Action Verification", "+-&OK"
      Else
        MyP.StateStuff lstStates.List(lstStates.ListIndex), stCode
        selArea = lstStates.List(lstStates.ListIndex) & "  "
        j = 1
        While j <= Len(Trim(selArea))
          strSoFar = strSoFar & " " & UCase(Left(StrSplit(Mid(selArea, j), " ", ""), 1)) & _
                      LCase(Mid(StrSplit(Mid(selArea, j), " ", ""), 2))
          j = j + InStr(1, Mid(selArea, j), " ")
        Wend
        MyP.State = Trim(strSoFar)
      End If
    End If
    Ok2DoMore
    SaveSetting "AWUDS", "Defaults", "LastState", MyP.State
  End If
  Me.Caption = "Aggregate Water-Use Data System (AWUDS) - " & MyP.State
End Sub

Private Sub lstYears_Click()
Attribute lstYears_Click.VB_Description = "Ensures that selected HUCs/counties/aquifers exist for selected year (or 2 years in the case of 'Compare Data for 2 Years' report)."
' ##SUMMARY Ensures that selected HUCs/counties/aquifers exist for selected year (or 2 _
          years in the case of 'Compare Data for 2 Years' report).
' ##REMARKS Selects appropriate Data Dictionary based on year(s) selection.
  Dim i As Long
  Dim j As Long
  Dim dictYear As Long
  Dim year As Long
  Dim yearStr As String
  Dim dummyRec As Recordset
  
  If lstYears.SelCount < 2 Then
    MyP.Year2Opt = 0
    TwoYears = False
  End If
    
  If lstYears.List(0) = "none" Then Exit Sub
  MyP.Year1Opt = 0
  If MyP.UserOpt = 10 Then
    If lstYears.SelCount > 2 Then
      lstYears.Selected(OldYr) = False
    ElseIf lstYears.SelCount < 2 Then
      MyP.Year2Opt = 0
    End If
  Else
    If lstYears.SelCount > 1 Then
      lstYears.Selected(OldYr) = False
    End If
  End If
  For i = 0 To lstYears.ListCount - 1
    If lstYears.Selected(i) = True Then
      MyP.Year1Opt = lstYears.List(i)
      If (MyP.Length = 3 Or MyP.Length = 10) And MyP.UserOpt <> 9 And _
          Not NationalDB And Left(lstArea.RightItem(0), 3) <> "000" Then CheckAreas
      If MyP.UserOpt = 10 Then
        For j = i + 1 To lstYears.ListCount - 1
          If lstYears.Selected(j) = True Then
            MyP.Year2Opt = MyP.Year1Opt
            MyP.Year1Opt = lstYears.List(j)
            If (MyP.Length = 3 Or MyP.Length = 10) And _
               Left(lstArea.RightItem(0), 3) <> "000" Then CheckAreas
            Exit For
          End If
        Next j
        Exit For
      End If
    End If
  Next i
  
  'find the most recent Data Dictionary on file
  For j = 1995 To 2050 Step 5
    On Error GoTo x
    Set dummyRec = MyP.stateDB.OpenRecordset(j & "Fields1", dbOpenSnapshot)
    dictYear = j
    dummyRec.Close
  Next j
x:
  If MyP.Year1Opt >= dictYear Then
    MyP.YearFields = dictYear & "Fields1"
  Else
    If MyP.Year1Opt < 1996 And MyP.Year2Opt < 1996 Then
      MyP.YearFields = 1995 & "Fields1"
    Else
      year = dictYear
      If MyP.UserOpt = 10 Then
        If MyP.Year1Opt > MyP.Year2Opt And MyP.Year2Opt > 0 Then
          year = MyP.Year1Opt
          MyP.Year1Opt = MyP.Year2Opt
          MyP.Year2Opt = year
        Else
          If lstYears.SelCount = 2 Then
            year = MyP.Year2Opt
          Else
            year = MyP.Year1Opt
          End If
        End If
      End If
      year = Int(year / 5) * 5
      yearStr = CStr(year)
      MyP.YearFields = yearStr & "Fields1"
    End If
  End If
  
  Ok2DoMore
  If MyP.UserOpt = 10 Then
    yearStr = MyP.Year1Opt & " and " & MyP.Year2Opt
  Else
    yearStr = MyP.Year1Opt
  End If
  If Not (MyP.UserOpt = 10 And MyP.Year2Opt = 0) Then _
      Me.Caption = "Aggregate Water-Use Data System (AWUDS) - " & MyP.State & ", " & yearStr
  OldYr = lstYears.ListIndex
End Sub

Private Sub CheckAreas()
Attribute CheckAreas.VB_Description = "Ensures that all&nbsp;selected counties exist for selected year."
' ##SUMMARY Ensures that all&nbsp;selected counties/aquifers exist for selected year.
' ##REMARKS If a county/aquifer does not exist for the selected year then it _
          will be removed from the list of selected counties/aquifers.

  Dim i As Long
  Dim j As Long
  Dim sql As String
  Dim areaRec As Recordset
  Dim selectAll As Boolean
  
  If lstArea.LeftCount = 0 Then selectAll = True

  'Use SQL language to create recordset of unit area codes and names from
  ' the "county", "huc", "aquifer" tables for the selected state and year.
  ' Certain counties and aquifers were added or eliminated recently so
  ' "begin" and "end" attributes were added to exclude such areas from
  ' the query for years when they did not exist.
  ' For example:
  '  SELECT Trim(county_cd) As Code, Trim(county_nm) As Name FROM [county]
  '  WHERE state_cd='10' AND (([county].begin<=1995 OR IsNull([county].begin))
  '    AND ([county].end>=1995 OR IsNull([county].end)))
  '  ORDER BY Trim(county_cd);
  sql = "SELECT Trim(" & TableName & "_cd) As Code, Trim(" & TableName & "_nm) As Name" & _
      " FROM [" & TableName & "]" & _
      " WHERE state_cd='" & MyP.stateCode & "'" & _
      " AND (([" & TableName & "].begin<=" & MyP.Year1Opt & " OR IsNull([" & TableName & "].begin))" & _
      " AND ([" & TableName & "].end>=" & MyP.Year1Opt & " OR IsNull([" & TableName & "].end)))" & _
      " ORDER BY Trim(" & TableName & "_cd);"
  Set areaRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  j = 0
  For i = 0 To lstArea.RightCount - 1
    If Left(lstArea.RightItem(i - j), 4) <> "All " Then
      areaRec.FindFirst "Code='" & Left(lstArea.RightItem(i - j), MyP.Length) & "'"
      If areaRec.NoMatch Then
        If Not selectAll Then
          MsgBox lstArea.RightItem(i - j) & " County did not exist during the year " & _
              MyP.Year1Opt & vbCrLf & "and it will be removed from the selected counties."
        End If
        lstArea.MoveLeft (i - j)
        j = j + 1
      End If
    End If
  Next i
  areaRec.Close
End Sub

Private Sub MainOpt_Click(index As Integer)
Attribute MainOpt_Click.VB_Description = "Sets properties of controls on 3rd tab of&nbsp;main form&nbsp;depending upon which operation selected on 2nd tab."
' ##SUMMARY Sets properties of controls on 3rd tab of&nbsp;main form&nbsp;depending upon _
          which operation selected on 2nd tab.
' ##PARAM index I Integer indicating which option has been selected from the _
          'Operation' tab of the main form.
' ##REMARKS Populates location array for National Db, if selected.
  Dim operType As String
  Dim ampLocn As Long
  Dim i As Long

  If NationalDB Then
    SetAreas
    If MainOpt(8) Then
      rdoAreaUnit(0).Caption = "County"
      rdoAreaUnit(1).Visible = True
      rdoAreaUnit(2).Visible = False
      rdoAreaUnit2(2).Visible = False
      rdoAreaUnit(3).Visible = False
    ElseIf MainOpt(10) Then 'County and Aquifer available for National Export, PRH 11/2005
      rdoAreaUnit(0).Caption = "County"
      rdoAreaUnit(1).Visible = False
      rdoAreaUnit(2).Visible = False
      rdoAreaUnit2(2).Visible = False
      rdoAreaUnit(3).Visible = True
    Else
      rdoAreaUnit(0).Caption = "State"
      rdoAreaUnit(1).Visible = True
      rdoAreaUnit(2).Visible = True
      rdoAreaUnit2(2).Visible = True
      rdoAreaUnit(3).Visible = True
    End If
  Else
    rdoAreaUnit(1).Visible = True
    rdoAreaUnit(2).Visible = True
    rdoAreaUnit(0).Caption = "County"
  End If

  EmboldenMe MainOpt, index
  MyP.UserOpt = index + 1
  MyP.UnitArea2 = ""
  If index > 1 And index < 10 And index <> 8 Then
    fraAreaUnitID.Visible = True
  Else
    fraAreaUnitID.Visible = False
  End If
  If index = 1 Then
    fraImport.Visible = True
    fraNewYear.Visible = False
    txtYear.value = ""
    txtCurFile.Text = " no file selected"
  ElseIf index = 11 Then
    fraImport.Visible = False
    fraNewYear.Visible = True
    
    fraDO.Visible = False
    fraPSPop.Visible = False
    fraIR.Visible = False
    lblPSTotal.Visible = False
    txtPSTotal.Visible = False
    lblDOTotalGW.Visible = False
    txtDOTotalGW.Visible = False
    lblDOTotalSW.Visible = False
    txtDOTotalSW.Visible = False
    lblNewYear.Visible = False
    txtNewYear.Visible = False
    For i = 0 To 1
      rdoNewAreaUnit(i) = False
      rdoDO(i) = False
      rdoPSPop(i) = False
      rdoIR(i) = False
    Next i
    rdoPSPop(2) = False
    rdoNewAreaUnit(2) = False
    EmboldenMe rdoNewAreaUnit, -1
    EmboldenMe rdoDO, -1
    EmboldenMe rdoPSPop, -1
    EmboldenMe rdoIR, -1
    txtPSTotal.value = ""
    txtDOTotalGW.value = ""
    txtDOTotalSW.value = ""
    txtNewYear.value = ""
    MyP.Year1Opt = 0
    MyP.UnitArea = ""
  Else
    cmdExeOpts.Caption = "Finalize Selections"
    ClearOpts
    If index = 8 Then
      fraAreaUnitID.Caption = "Area Unit IDs"
      fraAreaUnits2.Visible = True
      lstArea.Enabled = False
      lblDate.Caption = "Common Year(s) of Data"
      TwoAreas = True
    ElseIf NationalDB Then
      fraAreaUnitID.Caption = "State IDs"
      lstArea.Enabled = False
      fraAreaUnits2.Visible = False
      lblDate.Caption = "Available Year(s) of Data"
    ElseIf index = 5 Or index = 6 Or index = 7 Then
      fraAreaUnitID.Caption = "Area Unit IDs"
      fraAreaUnits2.Visible = False
      lstArea.Enabled = True
      lblDate.Caption = "Available Year(s) of Data"
    Else
      fraAreaUnitID.Caption = "Area Unit IDs"
      fraAreaUnits2.Visible = False
      lstArea.Enabled = True
      lblDate.Caption = "Available Year(s) of Data"
    End If
  End If
  
  'Can't do 'Calculated Tables' report with aquifers
  If index = 5 Or index = 6 Then
    rdoAreaUnit(3).Enabled = False
  Else
    rdoAreaUnit(3).Enabled = True
  End If
    
  If index = 0 Then
    cmdCatOpt.Caption = "Edit Data"
    operType = ""
  ElseIf index = 1 Then
    lblInstructs = "Select import file and type in year of data. Click 'Execute' button when selections complete."
  ElseIf index = 11 Then
    cmdAddYear.Enabled = False
    lblInstructs = "Choose from data storage options and enter new year value. Click 'Execute' when done."
  ElseIf index > 1 And index < 7 Then
    cmdCatOpt.Caption = "Produce Report"
    operType = "Report "
  ElseIf index > 6 And index < 10 Then
    cmdCatOpt.Caption = "Perform Quality Control Measure"
    operType = "Q-C Measure "
  ElseIf index = 10 Then
    cmdCatOpt.Caption = "Export Data"
    operType = ""
  End If
  If index = 0 Then
    lstDataCats.ButtonVisible(3) = False
    lstDataCats.ButtonVisible(4) = False
    lstArea.ButtonVisible(3) = False
    lstArea.ButtonVisible(4) = False
  Else
    lstDataCats.ButtonVisible(3) = True
    lstDataCats.ButtonVisible(4) = True
    If index = 9 Then
      lstArea.ButtonVisible(3) = False
      lstArea.ButtonVisible(4) = False
    Else
      lstArea.ButtonVisible(3) = True
      lstArea.ButtonVisible(4) = True
    End If
  End If
  ampLocn = InStr(1, MainOpt(index).Caption, "&", vbTextCompare)
  operation = Left(MainOpt(index).Caption, ampLocn - 1)
  operation = " for " & operType & " '" & operation & Mid(MainOpt(index).Caption, ampLocn + 1) & "'"
  Ok2DoMore
End Sub

Private Sub rdoAreaUnit_Click(index As Integer)
Attribute rdoAreaUnit_Click.VB_Description = "Sets unit area selection as either County, HUC-8, HUC-4, or Aquifer."
' ##SUMMARY Sets unit area selection as either County, HUC-8, HUC-4, or Aquifer.
' ##REMARKS Populates area listbox with available unit areas, and dates list box with _
          available years.
' ##PARAM Index I integer indicating which area unit has been selected from the _
          OptionButton array (0-3).
  Dim str As String
  Dim i As Long
  Dim response As Long
  Dim myDB As Database
  Dim myTabDef As TableDef
  
  For i = 0 To rdoAreaUnit2.Count - 1
    rdoAreaUnit2(i).Enabled = True
    rdoAreaUnit2(i) = False
  Next i
  TableName = StrSplit(rdoAreaUnit(index).Caption, " ", "")
  If TableName = "State" Then TableName = "County"
  
  AggregateHUCs = False
  If NationalDB And Not (rdoAreaUnit(0) Or rdoAreaUnit(3)) Then TableName = "HUC8"
  Select Case index
    Case 0: MyP.Length = 3
    Case 1: MyP.Length = 8
    Case 2: MyP.Length = 4
            'If HUC-4s are being analyzed, determine if they should be
            'from original records or from aggregations of HUC-8s
            If Not (MyP.UserOpt = 1 Or MyP.UserOpt = 9 Or MyP.UserOpt = 11) Then
              If TableName = "HUC8" Then
                AggregateHUCs = True
              Else
                response = MyMsgBox.Show("Are the HUC-4 data stored as original records or" & _
                    vbCrLf & "should they be aggregated from HUC-8 records?", _
                    "User Action Verification", "+&Original", "-&Aggregate")
              End If
              If response = 2 Then AggregateHUCs = True
            End If
    Case 3: MyP.Length = 10
  End Select
  
  ClearOpts
  
  If MyP.UserOpt = 9 Or NationalDB = True Then
    lstArea.Enabled = False
    lstArea.RightItem(0) = "All " & rdoAreaUnit(index).Caption & " areas selected"
    lstArea.RightItem(1) = ""
    If NationalDB Then
      Set myDB = OpenDatabase(AwudsDataPath & "Nation.mdb", , True)
      str = "Do you wish to construct a new national database " & _
            "that calculates " & vbCrLf
      If MainOpt(8) Or rdoAreaUnit(0) Then
        Set myTabDef = myDB.TableDefs(TableName & "Data")
        str = str & "state totals by summing the " & TableName & " areas in each state?"
      ElseIf rdoAreaUnit(3) Then
        Set myTabDef = myDB.TableDefs("AquiferData")
        str = str & "aquifer totals across the U.S.?"
      Else
        Set myTabDef = myDB.TableDefs("HUC8Data")
        str = str & "HUC totals across the U.S.?"
      End If
      str = str & vbCrLf & "The last such database was created on " & myTabDef.DateCreated & "." & _
            vbCrLf & "The process might take several minutes or longer depending on the quantity of data."
      response = MyMsgBox.Show(str, "National Database Creation", "+&Yes", "-&No")
      Set myTabDef = Nothing
      If response = 1 Then CreateNationalDB 1
    End If
  End If
  
  Me.MousePointer = vbHourglass
  rdoAreaUnit2(index).Enabled = False
  MyP.UnitArea = rdoAreaUnit(index).Caption
  MyP.UnitArea2 = ""
  MyP.AreaTable = TableName & "Data"
  SetAreas
  If Not (MyP.UserOpt = 9 Or NationalDB = True) Then
    fillArea
  End If
  
  Ok2DoMore
  If Not (MyP.UserOpt = 9 And MyP.UnitArea2 = "") Then FillDate
  Me.MousePointer = vbDefault
End Sub

Private Sub rdoAreaUnit2_Click(index As Integer)
Attribute rdoAreaUnit2_Click.VB_Description = "Sets second unit area selection for the "
' ##PARAM Index I integer indicating which area unit has been selected from the _
          OptionButton array (0-3).
' ##SUMMARY Sets second unit area selection for the "Compare State Totals by Area" report.
' ##REMARKS Populates dates list box with years for which both unit areas have data.
  Dim i As Long
  Dim response As Long
  Dim myDB As Database
  Dim myTabDef As TableDef
  
  'assign length of 2nd unit area string ID.
  Select Case index
    Case 0: MyP.length2 = 3
    Case 1: MyP.length2 = 8
    Case 2: MyP.length2 = 4
    Case 3: MyP.length2 = 10
  End Select
  
  Me.MousePointer = vbHourglass
  If index = 2 Then
    'flip-flop unit areas because HUC-4, if selected, must be first area unit
    For i = 0 To rdoAreaUnit.UBound
      If rdoAreaUnit(i) = True Then Exit For
    Next i
    MyP.length2 = 4
    rdoAreaUnit(2) = True
    rdoAreaUnit2(i) = True
    lstArea.RightItem(0) = "All " & MyP.UnitArea & " areas selected"
  Else
    TableName2 = StrSplit(rdoAreaUnit2(index).Caption, " ", "")
    If NationalDB And Not (rdoAreaUnit2(0) Or rdoAreaUnit2(3)) Then TableName2 = "HUC8"
    If NationalDB And MyP.Length <> 4 Then
      Set myDB = OpenDatabase(AwudsDataPath & "Nation.mdb", , True)
      Set myTabDef = myDB.TableDefs(TableName2 & "Data")
        response = MyMsgBox.Show("Do you wish to construct a new national database " & _
            "that calculates " & vbCrLf & "state totals by summing the " & TableName2 & _
            " areas in each state?" & vbCrLf & "The last such database was created on " & _
            myTabDef.DateCreated & vbCrLf & "This process may take 10 minutes or more.", _
            "National Database Creation", "+&Yes", "-&No")
      If response = 1 Then CreateNationalDB 2
    End If
    MyP.UnitArea2 = rdoAreaUnit2(index).Caption
    If Len(MyP.UnitArea) > 0 And Len(MyP.UnitArea2) > 0 Then
      lstArea.RightItem(1) = "All " & MyP.UnitArea2 & " areas selected"
      lstArea.RightItem(0) = "All " & MyP.UnitArea & " areas selected"
      FillDate
    End If
    Ok2DoMore
  End If
  Me.MousePointer = vbDefault
End Sub

Private Sub CreateNationalDB(Counter As Long)
Attribute CreateNationalDB.VB_Description = "Aggregates all data for each State, HUC-8,&nbsp;or aquifer&nbsp;in the U.S. into a single set of water-use records."
' ##SUMMARY Aggregates all data for each State, HUC-8,&nbsp;or aquifer&nbsp;in the U.S. _
          into a single set of water-use records.
' ##PARAM Counter I Integer indicating whether&nbsp;creating database for 1st unit area _
          type or appending 2nd type for 'Compare State Totals by Area' report.
' ##RETURNS Reads data from state DB's and creates new table in Nation.mdb.
  Dim Length As Long
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim opt As Long
  Dim numStates As Long
  Dim start As Long
  Dim lastYear As Long
  Dim numDataYears As Long
  Dim sql As String
  Dim dbPath As String
  Dim tabName As String
  Dim filler As String
  Dim myDB As Database
  Dim nationDB As Database
  Dim myTabDef As TableDef
  Dim myIndx As index
  Dim qdfReport As QueryDef
  Dim qdfAppend As QueryDef
  Dim myRec As Recordset

  On Error GoTo ErrTrap

  AtcoLaunch1.SendMonitorMessage "(OPEN AWUDS)"
  AtcoLaunch1.SendMonitorMessage "(BUTTOFF DETAILS)"
  AtcoLaunch1.SendMonitorMessage "(BUTTON CANCEL)"
  AtcoLaunch1.SendMonitorMessage "(BUTTON PAUSE)"
  AtcoLaunch1.SendMonitorMessage "(MSG1 Creating NationalDB)"
  AtcoLaunch1.SendMonitorMessage "(PROGRESS 0)"

  SetAreas
  NewTable = ""
  Set nationDB = OpenDatabase(AwudsDataPath & "Nation.mdb", , False)
  If Counter = 1 Then  'creating new table for County/HUC/Aquifer in Nation.mdb
    Length = MyP.Length
    On Error Resume Next
    If Not MainOpt(8) Then 'not comparing totals by area
      Select Case Length
        Case 4, 8:
          Length = 8
          NewTable = "HUC8Data"
      End Select
    End If
    tabName = TableName & "Data"
    start = 0
    numStates = UBound(LocnArray, 2)
  Else  'second unit area being added to table for Compare by Areas report
    Length = MyP.length2
    tabName = TableName2 & "Data"
    start = 1
    numStates = UBound(LocnArray, 2)
  End If
  If NewTable = "" Then NewTable = tabName
  
  On Error Resume Next
  nationDB.Execute "DROP TABLE [" & NewTable & "];"
  On Error GoTo 0
  'The filler is added to the state code to make the location field equal
  'in length to the unit area codes (i.e., state code + "C" = 3 digits and
  'county codes are three digits long)
  If MainOpt(8) Then  'comparing tables by area
    Select Case Length
      Case 3: filler = "C"
      Case 4: filler = "H4"
      Case 8: filler = "_HUC-8"
      Case 10: filler = "Aquifer_"
    End Select
  ElseIf rdoAreaUnit(0) Then
    filler = "C"
  End If

  'Create data table
  lastYear = 2000
  While lastYear < CLng(Format(Date, "yyyy"))
    lastYear = lastYear + 1
  Wend
  numDataYears = 3 + lastYear - 1999
  For i = 1985 To lastYear 'loop for year
    If i = 1985 Or i = 1990 Or i = 1995 Or i >= 2000 Then
      For j = 0 To numStates  'loop for # of states
        Select Case NextPipeCharacter(AtcoLaunch1.ComputeRead)
          Case "P"
            While NextPipeCharacter(AtcoLaunch1.ComputeRead) <> "R"
              DoEvents
            Wend
          Case "C"
            ImportDone = True
            MyMsgBox.Show "Creation of the new National Database has been interrupted." & _
                vbCrLf & "The " & tabName & " table is currently unsuitable for use.", _
                "National DB interrupted", "+-&OK"
            GoTo ErrTrap
        End Select
        'open DB for current state
        dbPath = AwudsDataPath & LocnArray(0, j) & ".mdb"
        Set myDB = OpenDatabase(dbPath, False, False, "MS Access; pwd=B7Q6C9B752")
        If i = 1985 And j = 0 Then
          'Use SQL language to create table "CountyData", "HUCData", or "AquiferData" in Nation.mdb
          ' containing the first year of data (1985) from the initial state/aquifer/HUC.
          ' The data table and Data Dictionary tables are joined to access attributes from both tables.
          ' The criteria "trim([Field1].Formula)=''" ensures that the query only includes user-entered
          ' fields; i.e., not products of formulas referencing other fields.
          ' "Excluded" property is coded to omit certain fields from certain reports.
          ' For example:
          '  SELECT '01C' AS Location, [CountyData].Date, [CountyData].FieldID,
          '  Sum([CountyData].Value) AS [Value], [CountyData].QualFlg, '02' AS State
          '  INTO [CountyData] IN 'C:\VBExperimental\AwudsBuild\Data\Nation.mdb'
          '  FROM [CountyData] INNER JOIN AllFields ON [CountyData].FieldID = AllFields.ID
          '  Where Len([CountyData].Location) = 3
          '  GROUP BY [CountyData].Date, [CountyData].FieldID, [CountyData].QualFlg, [AllFields].Formula
          '  HAVING [CountyData].Date=1985 AND Len(Trim([AllFields].Formula))=0;
          Select Case NewTable
            Case "HUC8Data":
              sql = "SELECT [HUCData].*, '" & LocnArray(0, j) & "' AS [State] " & _
                  "INTO [HUC8Data] IN '" & AwudsDataPath & "Nation.mdb' " & _
                  "FROM [HUCData] INNER JOIN AllFields ON [HUCData].FieldID = [AllFields].ID " & _
                  "WHERE Len([HUCData].Location)=" & Length & _
                  " AND [HUCData].Date=" & i & " AND Len(Trim([AllFields].Formula))=0" & _
                  " ORDER BY Date, Location, FieldID;"
            Case "AquiferData":
              sql = "SELECT [AquiferData].*, '" & LocnArray(0, j) & "' AS [State] " & _
                  "INTO [AquiferData] IN '" & AwudsDataPath & "Nation.mdb' " & _
                  "FROM [AquiferData] INNER JOIN AllFields ON [AquiferData].FieldID = [AllFields].ID " & _
                  "WHERE [AquiferData].Date=" & i & " AND Len(Trim([AllFields].Formula))=0" & _
                  " ORDER BY Date, Location, FieldID;"
            Case Else:
              sql = "SELECT '" & LocnArray(0, j) & filler & "' AS Location, [" & _
                  tabName & "].Date, [" & tabName & "].FieldID, Sum([" & tabName & "].Value) AS [Value], [" & _
                  tabName & "].QualFlg, '" & LocnArray(0, j) & "' AS [State] " & _
                  "INTO [" & NewTable & "] IN '" & AwudsDataPath & "Nation.mdb' " & _
                  "FROM [" & tabName & "] INNER JOIN AllFields ON [" & tabName & "].FieldID = AllFields.ID " & _
                  "WHERE Len([" & tabName & "].Location)=" & Length & _
                  " GROUP BY [" & tabName & "].Date, [" & tabName & "].FieldID, [" & tabName & "].QualFlg, " & _
                  "[AllFields].Formula " & _
                  "HAVING [" & tabName & "].Date=" & i & " AND Len(Trim([AllFields].Formula))=0;"
          End Select
          Set qdfReport = myDB.CreateQueryDef("", sql)
          qdfReport.Execute
          qdfReport.Close
          Set qdfReport = Nothing
          Set myTabDef = nationDB.TableDefs(NewTable)
          With myTabDef
            For k = 1 To 3
              Select Case k
                Case 1: sql = "FieldID"
                Case 2: sql = "Location"
                Case 3: sql = "Date"
              End Select
              Set myIndx = .CreateIndex("FldID" & k)
              With myIndx
                .Unique = False
                .Fields = sql
              End With
              .Indexes.Append myIndx
              Set myIndx = Nothing
            Next k
          End With
          Set myTabDef = Nothing
        Else  '
          'Append data to "CountyData", "HUCData", or "AquiferData" in Nation.mdb for all areas
          ' and years subsequent to the creation of table above.
          ' For example:
          '  INSERT INTO [CountyData] IN 'C:\VBExperimental\AwudsBuild\Data\Nation.mdb'
          '  SELECT '02C' AS Location, [CountyData].Date, [CountyData].FieldID, Sum([CountyData].Value) AS [Value],
          '  [CountyData].QualFlg, '02' AS State
          '  FROM [CountyData] INNER JOIN AllFields ON [CountyData].FieldID = [AllFields].ID
          '  Where Len([CountyData].Location) = 3
          '  GROUP BY [CountyData].Date, [CountyData].FieldID, [CountyData].QualFlg, [AllFields].Formula
          '  HAVING [CountyData].Date=1985 AND Len(Trim([AllFields].Formula))=0;
          Select Case NewTable
            Case "HUC8Data":
              sql = "INSERT INTO [HUC8Data] IN '" & AwudsDataPath & "Nation.mdb' " & _
                    "SELECT [HUCData].*, '" & LocnArray(0, j) & "' AS [State] FROM [HUCData] " & _
                    "INNER JOIN AllFields ON [HUCData].FieldID = [AllFields].ID " & _
                    "WHERE Len([HUCData].Location)=" & Length & _
                    " AND [HUCData].Date=" & i & " AND Len(Trim([AllFields].Formula))=0" & _
                    " ORDER BY Date, Location, FieldID;"
            Case "AquiferData":
              sql = "INSERT INTO [AquiferData] IN '" & AwudsDataPath & "Nation.mdb' " & _
                  "SELECT [AquiferData].*, '" & LocnArray(0, j) & "' AS [State] FROM [AquiferData] " & _
                  "INNER JOIN AllFields ON [AquiferData].FieldID = [AllFields].ID " & _
                  "WHERE [AquiferData].Date=" & i & " AND Len(Trim([AllFields].Formula))=0" & _
                  " ORDER BY Date, Location, FieldID;"
            Case Else:
              sql = "INSERT INTO [" & tabName & "] IN '" & AwudsDataPath & "Nation.mdb' " & _
                    "SELECT '" & LocnArray(0, j) & filler & "' AS Location, [" & tabName & "].Date, [" & tabName & "].FieldID, " & _
                    "Sum([" & tabName & "].Value) AS [Value], [" & tabName & "].QualFlg, '" & LocnArray(0, j) & "' AS [State] " & _
                    "FROM [" & tabName & "] INNER JOIN AllFields ON [" & tabName & "].FieldID = [AllFields].ID " & _
                    "WHERE Len([" & tabName & "].Location)=" & Length & _
                    " GROUP BY [" & tabName & "].Date, [" & tabName & "].FieldID, [" & tabName & "].QualFlg, " & _
                    "[AllFields].Formula " & _
                    "HAVING [" & tabName & "].Date=" & i & " AND Len(Trim([AllFields].Formula))=0;"
            End Select
          Set qdfAppend = myDB.CreateQueryDef("", sql)
          qdfAppend.Execute
          qdfAppend.Close
          Set qdfAppend = Nothing
        End If
nextState:
        If i < 2000 Then
          AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (((j + 1) / ((numStates + 1) * numDataYears) + ((i - 1985) / 5) / numDataYears) * 100) & ")"
        Else
          AtcoLaunch1.SendMonitorMessage "(PROGRESS " & _
            (((j + 1) / ((numStates + 1) * numDataYears) + (3 + i - 2000) / numDataYears) * 100) & ")"
        End If
        'change data values for 2000 as necessitated by data storage options
        myDB.Close
        If i >= 2000 And NewTable <> "AquiferData" Then
          OrderData NewTable, nationDB, i, LocnArray(0, j)
        End If
      Next j
      If NewTable = "HUC8Data" Or NewTable = "AquiferData" Then CombineAreasInState nationDB, i, Counter
    End If
  Next i

  nationDB.Close
ErrTrap:
  If Err.Number <> 999 And Err.Number <> 0 Then
    MsgBox "An error occured while constructing the " & NewTable & _
        " table in the National Database." & vbCrLf & _
        "Check the Nation.mdb file in '" & AwudsDataPath & _
        "' to make sure the table was not deleted," & vbCrLf & _
        "then rebuild the national database for " & MyP.UnitArea & "areas.", _
        vbCritical, "National Database Error"
    Err.Clear
  End If
  AtcoLaunch1.SendMonitorMessage "(CLOSE)"
End Sub

Private Sub OrderData(NewTable As String, nationDB As Database, _
                      CurYear As Long, CurState As String)
' ##SUMMARY This sub edits the National database to standardize state data storage _
          options, enabling state totals to be summed uniformly across the U.S.
' ##PARAM NewTable I String with the name of the table in Nation.mdb to be edited.
' ##PARAM NumAreas I Integer with the number of unit areas in the table to be edited.
' ##PARAM NationDB I Object representing the National database.
' ##PARAM CurYear I Integer indicating the year of data to be edited.
' ##PARAM CurState I Integer indicating the FIPS code of the state to be edited.
  Dim myDB As Database
  Dim dataRec As Recordset
  Dim tmpVal As Double
  Dim irFlds(7) As Double
  Dim area As String
  Dim sql As String
  Dim j As Long
  Dim k As Long
  Dim fldID As Long
  Dim numFlds As Long
  Dim numAreas As Long
  Dim opt As Long
  Dim isHUC As Boolean
  
  'open DB for current state and count # of areas
  Set myDB = OpenDatabase(AwudsDataPath & CurState & ".mdb", False, False, "MS Access; pwd=B7Q6C9B752")
  If NewTable = "HUC8Data" Then
    isHUC = True
    sql = "SELECT DISTINCT Location FROM [HUCData] WHERE Date=" & CurYear & ";"
  Else
    sql = "SELECT DISTINCT Location FROM [" & NewTable & _
          "] WHERE Date=" & CurYear & ";"
  End If
  Set dataRec = myDB.OpenRecordset(sql, dbOpenSnapshot)
  If dataRec.RecordCount > 0 Then
    dataRec.MoveLast
    numAreas = dataRec.RecordCount
  Else
    Exit Sub
  End If
  dataRec.Close
  myDB.Close
  
  sql = "SELECT * FROM [" & NewTable & _
        "] WHERE State='" & CurState & "' AND Date=" & CurYear & " ORDER BY Location, FieldID;"
  Set dataRec = nationDB.OpenRecordset(sql, dbOpenDynaset)
  With dataRec
    .MoveFirst
    If !QualFlg < 7 Then
      opt = !QualFlg
    ElseIf !QualFlg = 7 Then
      opt = 1
    ElseIf !QualFlg = 8 Then
      opt = 5
    End If
    While Not .EOF
      If !FieldID = 1 Then area = !Location
      Select Case !FieldID
        'can not store PS-PopServed or DO-Withdrawals for Nation by HUCs
        Case 2
          If isHUC And CurYear <> 2005 Then
            .Delete
            .MoveNext
            .Delete
          Else
            j = .AbsolutePosition
            If Not IsNull(!value) Then tmpVal = !value
            .Delete
            .MoveNext
            If Not IsNull(!value) Then tmpVal = tmpVal + !value
            .Delete
            .AddNew
            !Date = CurYear
            !Location = area
            !FieldID = 4
            If Not IsEmpty(tmpVal) Then !value = tmpVal
            If CurYear = 2005 Then !QualFlg = 8 Else !QualFlg = 5
            .Update
            .AbsolutePosition = j - 1
          End If
        Case 4
          If CurYear = 2005 Then
            GoTo NormalField
          Else
            If isHUC And CurYear <> 2005 Then
              .Delete
            Else
              .Edit
              If opt = 3 Or opt = 4 Then 'divide sum of state totals by number of areas
                If !value > 0 Then !value = !value / numAreas Else !value = 0
              End If
              If CurYear = 2005 Then !QualFlg = 8 Else !QualFlg = 5
              .Update
            End If
          End If
        Case 40, 43
          If CurYear = 2005 Then
            GoTo NormalField
          Else
            If isHUC Then 'can't store this or Field 43 for HUCs in National DB
              .Delete
            Else
              .Edit
              !QualFlg = 5
              If opt = 2 Or opt = 4 Or opt = 6 Then
                If !value > 0 Then !value = !value / numAreas Else !value = 0
              End If
              .Update
            End If
          End If
        Case Else
NormalField:
          .Edit
          If CurYear = 2005 Then
            !QualFlg = 8
          ElseIf isHUC Then
            !QualFlg = 4
          Else
            !QualFlg = 5
          End If
          .Update
      End Select
      .MoveNext
    Wend
    .Close
  End With
End Sub

Private Sub CombineAreasInState(NationalDB As Database, year As Long, Counter As Long)
Attribute CombineAreasInState.VB_Description = "This sub aggregates data in the National database&nbsp;for HUC-8s that have area in more than one state."
' ##SUMMARY This sub aggregates data in the National database&nbsp;for HUC-8s that have _
          area in more than one state.
' ##REMARKS The national database contains a unique list of counties per _
          state. HUC-8 and aquifer area units, however, do not abide by _
          state boundaries and often occur in multiple states. Hence, their _
          partial values may be stored in multiple state databases and need _
          to be combined into a single value. The module CreateNationalDB _
          writes all data from each state into the national database table. _
          For HUC-8 and aquifer areas that cross one or more state boundaries, _
          multiple area entries are put in the national database originally. _
          CombineAreasInState sums data across state boundaries then deletes the _
          repetitive values so that a single entry remains in the national _
          database representing the total for each unique HUC-8 or aquifer area. _
' ##PARAM NationalDB I Object representing the National database.
' ##PARAM Year I Integer indicating the year of data to be aggregated.
' ##PARAM Counter I Integer: 2 when combining 2nd unit areas for 'Compare by Area' _
          report; 1 otherwise.
  Dim myRec As Recordset
  Dim sql As String
  Dim areaCode As String
  Dim tmpVal As Double
  Dim tmpID As Long
  Dim qFlg As Long
  Dim i As Long  'counter for # of states, above one, partially occupied by a HUC
  Dim j As Long
  Dim haveVal As Boolean
  
  On Error GoTo x
  
  'create ordered recordset of all HUC data for given year
  sql = "SELECT * FROM [" & NewTable & "] " & _
        "WHERE Date=" & year & _
        " ORDER BY Location, FieldID;"
  Set myRec = NationalDB.OpenRecordset(sql, dbOpenDynaset)
  With myRec
    'run thru recordset looking for multiple entries of same HUC-8
    While Not .EOF
      tmpID = !FieldID
      .MoveNext
      i = 0
      'determine how many states the HUC-8 was in: (i+1) = number of states
      While !FieldID = tmpID
        i = i + 1
        .MoveNext
      Wend
      If i > 0 Then  'have multiple entries for same HUC-8
        'move back to first record for this HUC-8
        For j = 0 To i
          .MovePrevious
        Next j
        areaCode = !Location
        'following loop is thru all records for current HUC-8 and given year
        While !Location = areaCode
          tmpVal = 0
          tmpID = !FieldID
          haveVal = False
          If Not IsNull(!value) Then
            tmpVal = !value
            haveVal = True
          End If
          .MoveNext
          'following loop sums then deletes common data elements
          While !FieldID = tmpID
            If Not IsNull(!value) Then
              tmpVal = tmpVal + !value
              haveVal = True
            End If
            .Delete
            .MoveNext
          Wend
          'edit value for lone remaining data element
          .MovePrevious
          .Edit
          If haveVal Then !value = tmpVal Else !value = Null
          .Update
          .MoveNext
        Wend
      Else
        sql = !Location
        .FindNext "Location<>'" & sql & "'"
      End If
    Wend
  End With
x:
End Sub

Private Sub EmboldenMe(o As Object, index As Integer)
Attribute EmboldenMe.VB_Description = "This sub emboldens text labels on the main form to indicate the selected option in an object array."
' ##SUMMARY This sub emboldens text labels on the main form to indicate the selected _
          option in an object array.
' ##PARAM O I Object containing the array of options.
' ##PARAM Index I Integer indicating the selected item in the array of options.
  Dim objF As Font
  Dim i As Long
  
  For i = 0 To o.Count - 1
    Set objF = o(i).Font
    If i = index Then
      objF.Bold = True 'Embolden new selection
    Else
      objF.Bold = False 'disEmbolden new selection
    End If
  Next i

End Sub

Private Sub rdoDataDict_Click(index As Integer)

  'Make sure user has selected proper Data Dictionary for 2000 or 2005 data
  If index <> 0 And txtNewYear.value = 2000 Then
    'trying to assign 2005 Data Dict to 2000 data
    rdoDataDict(0) = True
    MsgBox "Year 2000 data must use the 2000 Data dictionary."
    Exit Sub
  ElseIf index <> 1 And txtNewYear.value = 2005 Then
    rdoDataDict(1) = True
    MsgBox "Year 2005 data must use the 2005 Data dictionary."
    Exit Sub
  End If
  
  If Len(MyP.UnitArea) > 0 Then
    If Not MyP.UnitArea = "Aquifer" Then fraPSPop.Visible = True
    fraIR.Visible = True
    lblNewYear.Visible = True
    txtNewYear.Visible = True
  End If
  If index = 0 Then
    If MyP.Year1Opt = 2005 And MyP.UnitArea <> "Aquifer" Then 'undo automatic selection
      rdoDO(0) = False
      rdoDO(0).Font.Bold = False
    End If
    If MyP.UnitArea = "County" Or MyP.UnitArea = "HUC - 8" Then fraDO.Visible = True
    rdoPSPop(2).Visible = True
    If rdoPSPop(2) Then
      lblPSTotal.Visible = True
      txtPSTotal.Visible = True
    Else
      lblPSTotal.Visible = False
      txtPSTotal.Visible = False
    End If
    If rdoDO(1) Then
      lblDOTotalGW.Visible = True
      txtDOTotalGW.Visible = True
      lblDOTotalSW.Visible = True
      txtDOTotalSW.Visible = True
    Else
      lblDOTotalGW.Visible = False
      txtDOTotalGW.Visible = False
      lblDOTotalSW.Visible = False
      txtDOTotalSW.Visible = False
    End If
  Else
    'make DO selection
    rdoDO(0) = True
    fraDO.Visible = False
    lblDOTotalGW.Visible = False
    txtDOTotalGW.Visible = False
    lblDOTotalSW.Visible = False
    txtDOTotalSW.Visible = False
    'eliminate statewide PS-Pop option
    rdoPSPop(2) = False
    rdoPSPop(2).Font.Bold = False
    rdoPSPop(2).Visible = False
    lblPSTotal.Visible = False
    txtPSTotal.Visible = False
  End If
  MyP.Year1Opt = rdoDataDict(index).Caption
End Sub

Private Sub rdoID_Click(index As Integer)
Attribute rdoID_Click.VB_Description = "Sets option of labeling HUCs/Counties/Aquifers by name, code, or both."
' ##SUMMARY Sets option of labeling HUCs/Counties/Aquifers by name, code, or both.
' ##REMARKS Index values: 0 = name, 1 = code, 2 = both.&nbsp;With both option, code is _
          stored in 1st column of spreadsheet and name in 2nd column.
' ##PARAM Index I Integer indicating the selected item in the array of OptionButtons.
  If index = 0 Then
    AreaID = 1
  ElseIf index = 1 Then
    AreaID = 0
  Else
    AreaID = 2
  End If
  Ok2DoMore
End Sub

Private Sub rdoNewAreaUnit_Click(index As Integer)
Attribute rdoNewAreaUnit_Click.VB_Description = "Sets unit area type (County, HUC-8, or Aquifer) for new year of data being created."
' ##SUMMARY Sets unit area type (County, HUC-8, or Aquifer) for new year of data being _
          created.
' ##PARAM Index I Integer indicating the selected item in the array of OptionButtons.
  Dim i As Long
  
  If index = 2 Or Not (rdoDataDict(0) Or rdoDataDict(1)) Then
    fraDO.Visible = False
    fraPSPop.Visible = False
    If index = 2 Then
      rdoDO(0) = True
      rdoPSPop(0) = True
      If MyP.Year1Opt = 0 Then
        fraIR.Visible = False
        lblNewYear.Visible = False
        txtNewYear.Visible = False
      Else
        fraIR.Visible = True
        lblNewYear.Visible = True
        txtNewYear.Visible = True
      End If
    End If
  Else
    If HasDataForThisYear > 0 Then
      MsgBox "There is already " & MyP.Year1Opt & " " & rdoNewAreaUnit(index) & " data in the state database for " & MyP.State
    Else
      If rdoDataDict(0) Then
        If MyP.UnitArea = "Aquifer" Then 'undo automatic DO selection
          rdoDO(0) = False
          EmboldenMe rdoDO, -1
        End If
        fraDO.Visible = True
      Else
        fraDO.Visible = False
      End If
      If MyP.UnitArea = "Aquifer" Then 'undo automatic PS-Pop selection
        rdoPSPop(0) = False
        EmboldenMe rdoPSPop, -1
      End If
      fraPSPop.Visible = True
      fraIR.Visible = True
      lblNewYear.Visible = True
      txtNewYear.Visible = True
    End If
  End If
  MyP.UnitArea = rdoNewAreaUnit(index).Caption
  EmboldenMe rdoNewAreaUnit, index
  OKtoAddYear
End Sub

Private Sub rdoDO_Click(index As Integer)
Attribute rdoDO_Click.VB_Description = "Sets Domestic withdrawals data storage option for new year of data being created."
' ##SUMMARY Sets Domestic withdrawals data storage option for new year of data being _
          created.
' ##REMARKS Index values: 0 = "by Unit Area", 1 = "State Total".
' ##PARAM Index I Integer indicating the selected item in the array of OptionButtons.
  EmboldenMe rdoDO, index
  If index = 0 Then
    lblDOTotalGW.Visible = False
    txtDOTotalGW.Visible = False
    lblDOTotalSW.Visible = False
    txtDOTotalSW.Visible = False
  Else
    lblDOTotalGW.Visible = True
    txtDOTotalGW.Visible = True
    lblDOTotalSW.Visible = True
    txtDOTotalSW.Visible = True
  End If
  OKtoAddYear
End Sub

Private Sub rdoPSPop_Click(index As Integer)
Attribute rdoPSPop_Click.VB_Description = "Sets Public Supply population served data storage option for new year of data being created."
' ##SUMMARY Sets Public Supply population served data storage option for new year of data _
          being created.
' ##REMARKS Index values: 0 = "by Unit Area for both GW and SW", 1 = "by Unit Area for _
          combined GW/SW", 2 = "State Total for combined GW/SW".
' ##PARAM Index I Integer indicating the selected item in the array of OptionButtons.
  EmboldenMe rdoPSPop, index
  If index = 0 Or index = 1 Then
    lblPSTotal.Visible = False
    txtPSTotal.Visible = False
  Else
    lblPSTotal.Visible = True
    txtPSTotal.Visible = True
  End If
  OKtoAddYear
End Sub

Private Sub rdoIR_Click(index As Integer)
Attribute rdoIR_Click.VB_Description = "Sets Irrigation data storage option for new year of data being created."
' ##SUMMARY Sets Irrigation data storage option for new year of data being created.
' ##RETURNS Index values: 0 = divided into crops and golf, 1 = overall.
' ##PARAM Index I Integer indicating the selected item in the array of OptionButtons.
  EmboldenMe rdoIR, index
  OKtoAddYear
End Sub

Private Sub txtDOTotalGW_Change()
Attribute txtDOTotalGW_Change.VB_Description = "Performs QA check on user-entered value for state total of Domestic GW withdrawals for new year of data being created."
' ##SUMMARY Performs QA check on user-entered value for state total of Domestic GW _
          withdrawals for new year of data being created.
  txtDOTotalGW.value = DataEntry(txtDOTotalGW.value)
End Sub

Private Sub txtDOTotalSW_Change()
Attribute txtDOTotalSW_Change.VB_Description = "Performs QA check on user-entered value for state total of Domestic SW withdrawals for new year of data being created."
' ##SUMMARY Performs QA check on user-entered value for state total of Domestic SW _
          withdrawals for new year of data being created.
  txtDOTotalSW.value = DataEntry(txtDOTotalSW.value)
End Sub

Private Sub txtPSTotal_Change()
Attribute txtPSTotal_Change.VB_Description = "Performs QA check on user-entered value for state total of Public Supply population served for new year of data being created."
' ##SUMMARY Performs QA check on user-entered value for state total of Public Supply _
          population served for new year of data being created.
  txtPSTotal.value = DataEntry(txtPSTotal.value)
End Sub

Private Sub txtNewYear_Change()
Attribute txtNewYear_Change.VB_Description = "Performs QA check on user-entered value for year of new data being created."
' ##SUMMARY Performs QA check on user-entered value for year of new data being created.
  Dim i As Long

  For i = 1 To Len(txtNewYear.value)
    If Not IsNumeric(Mid(txtNewYear.value, i, 1)) Then
      txtNewYear.value = Left(txtNewYear.value, i - 1) & Mid(txtNewYear.value, i + 1)
    End If
  Next i
  If Len(txtNewYear.value) = 4 And IsNumeric(txtNewYear.value) Then
    If txtNewYear.value = "2000" And Not rdoDataDict(0) Then
      rdoDataDict(0) = True
      MsgBox "Year 2000 data must use the 2000 Data dictionary."
    ElseIf txtNewYear.value = "2005" And Not rdoDataDict(1) Then
      rdoDataDict(1) = True
      MsgBox "Year 2005 data must use the 2005 Data dictionary."
    End If
    If HasDataForThisYear > 0 Then
      MsgBox "There is already " & txtNewYear.value & " " & MyP.UnitArea & " data in the state database for " & MyP.State
    Else
      MyP.Year1Opt = txtNewYear.value
    End If
  End If
  OKtoAddYear
End Sub

Private Function HasDataForThisYear() As Long
' ##SUMMARY Checks to see if data exists for this year for other unit area type. _
          If so, the same data storage options will be imposed on the new data set.
  Dim dataRec As Recordset
  Dim sql As String
    
  If (rdoNewAreaUnit(0) Or rdoNewAreaUnit(1)) And Len(Trim(txtNewYear.value)) = 4 Then
    'Use SQL language to create recordset with this year's data for other
    ' unit area (i.e., create 2005 HUC recordset if creating new year of 2005
    ' county data).  If the recordset is not empty, use the same QualFlg.
    ' A restiction instituted in 12/05 forbids the user from mixing
    ' Data Dictionaries for HUC and County data in the same year.
    ' For example:
    '  SELECT * From CountyData
    '  Where (CountyData.Date = 2005);
    If rdoNewAreaUnit(0) Then
      sql = "SELECT * FROM [CountyData]" & _
            " WHERE (CountyData.Date=" & txtNewYear.value & ");"
    ElseIf rdoNewAreaUnit(1) Then
      sql = "SELECT * FROM [HUCData]" & _
            " WHERE (HUCData.Date=" & txtNewYear.value & ")" & _
            " ORDER BY Location DESC;"
    ElseIf rdoNewAreaUnit(2) Then
      sql = "SELECT * FROM [AquiferData]" & _
            " WHERE (AquiferData.Date=" & txtNewYear.value & ")" & _
            " ORDER BY Location DESC;"
    End If
    Set dataRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    If dataRec.RecordCount > 0 Then
      HasDataForThisYear = dataRec("QualFlg")
      Select Case HasDataForThisYear
        Case 1, 7:
          rdoPSPop(0) = True
          rdoPSPop_Click 0
          rdoDO(0) = True
          rdoDO_Click 0
        Case 2:
          rdoPSPop(0) = True
          rdoPSPop_Click 0
          rdoDO(1) = True
          rdoDO_Click 1
        Case 3:
          rdoPSPop(2) = True
          rdoPSPop_Click 2
          rdoDO(0) = True
          rdoDO_Click 0
        Case 4:
          rdoPSPop(2) = True
          rdoPSPop_Click 2
          rdoDO(1) = True
          rdoDO_Click 1
        Case 5, 8:
          rdoPSPop(1) = True
          rdoPSPop_Click 1
          rdoDO(0) = True
          rdoDO_Click 0
        Case 6:
          rdoPSPop(1) = True
          rdoPSPop_Click 1
          rdoDO(1) = True
          rdoDO_Click 1
      End Select
      'Check for irrigation selection
      dataRec.FindFirst "FieldID=195"
      If dataRec.NoMatch Then
        rdoIR(0) = True
        rdoIR_Click 0
      Else
        rdoIR(1) = True
        rdoIR_Click 1
      End If
    Else
      HasDataForThisYear = -1
    End If
  End If
End Function

Private Sub OKtoAddYear()
Attribute OKtoAddYear.VB_Description = "This sub enables command button that creates blank template for new year of water-use data&nbsp;if user has made necessary selections."
' ##SUMMARY This sub enables command button that creates blank template for new year of _
          water-use data&nbsp;<EM>if</EM> user has made necessary selections.
  If ((rdoNewAreaUnit(0) Or rdoNewAreaUnit(1) Or rdoNewAreaUnit(2)) And _
      (rdoDO(0) Or rdoDO(1)) And _
      (rdoPSPop(0) Or rdoPSPop(1) Or rdoPSPop(2)) And _
      (rdoIR(0) Or rdoIR(1))) And _
      Len(txtNewYear.value) = 4 Then
    cmdAddYear.Enabled = True
  Else
    cmdAddYear.Enabled = False
  End If
  
End Sub

Private Sub cmdAddYear_Click()
Attribute cmdAddYear_Click.VB_Description = "Creates blank set of records in state DB to store new year of water-use data based upon which data storage options were selected."
' ##SUMMARY Creates blank set of records in state DB to store new year of water-use data _
          based upon which data storage options were selected.
  Dim newDataRec As Recordset
  Dim fldRec As Recordset
  Dim areaRec As Recordset
  Dim sql As String
  Dim sqlAddOn As String
  Dim locn As String
  Dim irr As String
  Dim fieldTab As String
  Dim i As Long
  Dim j As Long
  Dim year As Long
  Dim qualFlag As Long
  Dim dummyRec As Recordset
  
  For i = 0 To 2
    If rdoNewAreaUnit(i) Then Exit For
  Next i
  Select Case i
    Case 0:
      MyP.Length = 3
      MyP.UnitArea = "County"
    Case 1:
      MyP.Length = 8
      MyP.UnitArea = "HUC - 8"
    Case 2:
      MyP.Length = 10
      MyP.UnitArea = "Aquifer"
  End Select
  
  If txtNewYear.value < 1850 Or txtNewYear.value > 2100 Then
    MyMsgBox.Show "The value entered for Year must be between 1850-2100", "Bad Year", "+&OK"
    Exit Sub
  ElseIf IsNumeric(txtDOTotalGW.value) And (txtDOTotalGW.value < 0 Or txtDOTotalGW.value > 99999.99) Or _
         IsNumeric(txtDOTotalSW.value) And (txtDOTotalSW.value < 0 Or txtDOTotalSW.value > 99999.99) Or _
         IsNumeric(txtPSTotal.value) And (txtPSTotal.value < 0 Or txtPSTotal.value > 99999.99) Then
    MyMsgBox.Show "All data values must be in the range of 0 to 99999.99", "Bad Data Value", "+&OK"
    Exit Sub
  End If
  Me.MousePointer = vbHourglass
  
  If rdoNewAreaUnit(2) Then  'new aquifer data records
    MyP.YearFields = "2000FieldsA"
    fieldTab = "FieldA"
    If rdoDataDict(0) Then
      qualFlag = 1
    Else
      qualFlag = 7
    End If
  Else
    If rdoDO(0) And rdoPSPop(0) Then
      MyP.YearFields = "2000Fields1"  'DO and PS by county, GW/SW
      fieldTab = "Field1"
      If rdoDataDict(0) Then
        qualFlag = 1
      Else
        qualFlag = 7
      End If
    ElseIf rdoDO(1) And rdoPSPop(0) Then  'DO by state, GW/SW.  PS by county, GW/SW
      MyP.YearFields = "2000Fields2"
      fieldTab = "Field2"
      qualFlag = 2
    ElseIf rdoDO(0) And rdoPSPop(2) Then  'DO by county GW/SW.  PS by state, total
      MyP.YearFields = "2000Fields3"
      fieldTab = "Field3"
      qualFlag = 3
    ElseIf rdoDO(1) And rdoPSPop(2) Then 'DO by state, GW/SW.  PS by state, total
      MyP.YearFields = "2000Fields4"
      fieldTab = "Field4"
      qualFlag = 4
    ElseIf rdoDO(0) And rdoPSPop(1) Then 'DO by county.  PS by county, total
      MyP.YearFields = "2000Fields5"
      fieldTab = "Field5"
      If rdoDataDict(0) Then
        qualFlag = 5
      Else
        qualFlag = 8
      End If
    ElseIf rdoDO(1) And rdoPSPop(1) Then 'DO by state.  PS by county, total
      MyP.YearFields = "2000Fields6"
      fieldTab = "Field6"
      qualFlag = 6
    End If
  End If
  If rdoIR(0) Then
    irr = "Not([" & fieldTab & "].ID > 194 AND [" & fieldTab & "].ID < 213)"
  Else
    irr = "Not([" & fieldTab & "].ID > 281 AND [" & fieldTab & "].ID < 302)"
  End If
  If MyP.Length <> 8 Then  'do not include RE fields
    sqlAddOn = " AND Not([" & fieldTab & "].ID = 232 OR [" & fieldTab & "].ID = 233)"
  End If
  'Create recordset to which records will be added
  TableName = LCase(StrSplit(MyP.UnitArea, " ", ""))
  sql = TableName & "Data"
  Set newDataRec = MyP.stateDB.OpenRecordset(sql, dbOpenDynaset)
  If newDataRec.AbsolutePosition > -1 Then newDataRec.MoveLast
  'Use SQL language to create recordset of null values containing all fields to be
  ' added for new year of data. Since Irrigation can be a single category
  ' or divided into crops and golf courses, one of those options is
  ' excluded from the query in the WHERE statement. The criteria
  ' "trim([Field1].Formula)=''" ensures that the query only includes original
  ' data fields; i.e., not products of formulas referencing other fields.
  ' For example:
  '  SELECT [Field1].ID FROM [2000Fields1]
  '  RIGHT JOIN Field1 ON [2000Fields1].FieldID = Field1.ID
  '  Where Not ([Field1].id > 194 And [Field1].id < 213)
  '    AND Not([Field1].ID = 232 OR [Field1].ID = 233)
  '    AND Trim([Field1].Formula)=''
  '  ORDER BY [Field1].ID;
  sql = "SELECT [" & fieldTab & "].ID FROM [" & MyP.YearFields & _
        "] RIGHT JOIN " & fieldTab & " ON [" & MyP.YearFields & "].FieldID = " & fieldTab & ".ID" & _
        " Where " & irr & sqlAddOn & " AND Trim([" & fieldTab & "].Formula)=''" & _
        " ORDER BY [" & fieldTab & "].ID;"
  Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  fldRec.MoveFirst
  'Use SQL language to create recordset with all area codes/names for unit area
  ' type (county,huc,aquifer) in this state. Certain counties and aquifers may
  ' been added after the year of new data, or were eliminated beforehand,
  ' so "begin" and "end" attributes were added to exclude such areas from the
  ' query for years when they did not exist.
  ' For example:
  '  SELECT county_cd, county_nm FROM [county]
  '  WHERE state_cd='09' AND (([county].begin<=2010 OR IsNull([county].begin))
  '                      AND ([county].end>=2010 OR IsNull([county].end)));
  If MyP.Length = 3 Or MyP.Length = 10 Then
    sql = "SELECT " & TableName & "_cd, " & TableName & "_nm FROM [" & TableName & "]" & _
          " WHERE state_cd='" & MyP.stateCode & _
          "' AND (([" & TableName & "].begin<=" & MyP.Year1Opt & " OR IsNull([" & TableName & "].begin))" & _
          " AND ([" & TableName & "].end>=" & MyP.Year1Opt & " OR IsNull([" & TableName & "].end)));"
  Else
    sql = "SELECT " & TableName & "_cd, " & TableName & "_nm " & _
          "From [" & TableName & _
          "] WHERE state_cd='" & MyP.stateCode & _
          "' AND Len(Trim(" & TableName & "_cd))=" & MyP.Length & ";"
  End If
  Set areaRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  areaRec.MoveLast
  areaRec.MoveFirst
  year = MyP.Year1Opt
  FillDate
  For i = 0 To lstYears.ListCount - 1
    If txtNewYear.value = lstYears.List(i) Then
      MyMsgBox.Show "Data already exists for this year." & vbCrLf & _
          "Export data to view and edit.", _
          "User Action Verification", "+&OK"
      GoTo x
    End If
  Next i
  With newDataRec
    'open the ATCo timebar
    AtcoLaunch1.SendMonitorMessage "(OPEN AWUDS)"
    AtcoLaunch1.SendMonitorMessage "(BUTTOFF DETAILS)"
    AtcoLaunch1.SendMonitorMessage "(BUTTOFF CANCEL)"
    AtcoLaunch1.SendMonitorMessage "(BUTTOFF PAUSE)"
    AtcoLaunch1.SendMonitorMessage "(MSG1 Creating data records for " & _
        MyP.UnitArea & " areas in " & MyP.State & ", " & txtNewYear.value & ")"
    AtcoLaunch1.SendMonitorMessage "(PROGRESS 0)"
    For i = 1 To areaRec.RecordCount
      AtcoLaunch1.SendMonitorMessage "(PROGRESS " & i * 100 / areaRec.RecordCount & ")"
      locn = Trim(areaRec(TableName & "_cd"))
      For j = 1 To fldRec.RecordCount
        .AddNew
        !Date = year
        !Location = locn
        !FieldID = fldRec("ID")
        Select Case !FieldID
          Case 4:
            If IsNumeric(txtPSTotal.value) And rdoPSPop(2) Then
              !value = txtPSTotal.value
            Else
              !value = Null
            End If
          Case 40:
            If IsNumeric(txtDOTotalGW.value) And (qualFlag = 2 Or qualFlag = 4 Or qualFlag = 6) Then
              !value = txtDOTotalGW.value
            Else
              !value = Null
            End If
          Case 43:
            If IsNumeric(txtDOTotalSW.value) And (qualFlag = 2 Or qualFlag = 4 Or qualFlag = 6) Then
              !value = txtDOTotalSW.value
            Else
              !value = Null
            End If
          Case Else:
            !value = Null
        End Select
        !QualFlg = qualFlag
        .Update
        fldRec.MoveNext
      Next j
      fldRec.MoveFirst
      areaRec.MoveNext
    Next i
  End With
x:
  newDataRec.Close
  txtNewYear.value = ""
  For i = 0 To 2
    rdoNewAreaUnit(i) = False
  Next i
  For i = 0 To 1
    rdoDataDict(i) = False
    rdoDataDict(i).Font.Bold = False
    rdoDO(i).Enabled = True
    rdoPSPop(i).Enabled = True
    rdoIR(i).Enabled = True
  Next i
  rdoPSPop(2).Enabled = True
  EmboldenMe rdoNewAreaUnit, -1
  EmboldenMe MainOpt, -1
  MyP.Length = 0
  MyP.Year1Opt = 0
  MyP.UnitArea = ""
  fraNewYear.Visible = False
  lstYears.Clear
  lblInstructs = "Select an Operation"
  AtcoLaunch1.SendMonitorMessage "(CLOSE)"
  Me.MousePointer = vbDefault
  MainOpt(11) = False
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
Attribute tabMain_Click.VB_Description = "Sets appropriate control properties when user selects new tab."
' ##SUMMARY Sets appropriate control properties when user selects new tab.
' ##PARAM PreviousTab I Integer indicating the previous tab receiving focus on the main _
          form.
  Dim response As Long
  Dim byArea As String
  If PreviousTab = 4 And tabMain.TabEnabled(2) = True Then  'leaving Values tab
    If DataSelOK = True Then
      If cmdEdit.Enabled = True Then  'values have been edited
        response = MyMsgBox.Show("Would you like to save your edits?", _
            "User Action Verification", "+&Yes", "-&No")
        If response = 1 Then
          EditOK = True
          cmdEdit_Click
        End If
      End If
    End If
    'clear Categories list
    If lstDataCats.RightCount > 0 Then
      lstDataCats.ClearRight
    End If
    If lstDataCats.LeftCount > 0 Then
      lstDataCats.ClearLeft
    End If
    FillCats
  End If
   
  If tabMain.Tab = 0 Then
    fraWhat2Do.Caption = "Instructions"
    If NationalDB Then
      lblInstructs = "You may continue to select and deselect individual states"
    Else
      lblInstructs = "Specify State to analyze"
    End If
  ElseIf tabMain.Tab = 1 Then
    fraWhat2Do.Caption = "Instructions"
    fraDataOpts.Caption = "Data Operations for " & MyP.State
    If Not (MyP.UserOpt = 2 Or MyP.UserOpt = 12) Then lblInstructs = "Select an Operation"
  ElseIf tabMain.Tab = 2 Then
    byArea = " by " & MyP.UnitArea
    If Len(MyP.UnitArea) > 0 And Len(MyP.UnitArea2) > 0 Then
      byArea = " by " & MyP.UnitArea & ", " & MyP.UnitArea2
    ElseIf Len(MyP.UnitArea) > 0 Then
      byArea = " by " & MyP.UnitArea
    Else
      byArea = ""
    End If
    fraWhat2Do.Caption = "Instructions" & operation & byArea
    If MyP.UserOpt <> 2 Then fraImport.Visible = False
    If MyP.UserOpt <> 12 Then fraNewYear.Visible = False
    If DataSelOK = True Then
      If MyP.UserOpt = 2 Then
        lblInstructs = "Click 'Execute' when ready."
      ElseIf MyP.UserOpt = 12 Then
        lblInstructs = "Click 'Execute' when ready."
      Else
        lblInstructs = "Click 'Finalize Selections' when ready."
      End If
    ElseIf MyP.YearValid Then
      If (fraAreaUnitID.Visible = False Or (rdoID(0) Or rdoID(1) Or rdoID(2))) Then
        If MyP.UserOpt = 1 Then
          lblInstructs = "Select one " & MyP.UnitArea & " from the Available list"
        Else
          lblInstructs = "Select one or more " & MyP.UnitArea & _
                         " from the Available list"
        End If
      Else
        lblInstructs = "Choose from one of the three Area Unit IDs"
      End If
    ElseIf MyP.numAreas > 0 = True Then
      If (fraAreaUnitID.Visible = False Or (rdoID(0) Or rdoID(1) Or rdoID(2))) Then
        If MyP.UserOpt = 10 Then
          If lstYears.SelCount = 1 Then
            lblInstructs = "Choose one more of the available years of data"
          Else
            lblInstructs = "Choose two of the available years of data"
          End If
        Else
          If MyP.UserOpt = 9 And Len(MyP.UnitArea) > 0 And Len(MyP.UnitArea2) = 0 Then
            lblInstructs = "Choose the second unit area for comparison"
          Else
            lblInstructs = "Choose one of the available years of data"
          End If
        End If
      Else
        lblInstructs = "Choose from one of the three Area Unit IDs"
      End If
    ElseIf Len(MyP.UnitArea) > 0 Then
      If (fraAreaUnitID.Visible = False Or (rdoID(0) Or rdoID(1) Or rdoID(2))) Then
        If Left(lstArea.LeftItem(0), 2) = "No" Then
          lblInstructs = "Try a different Area Unit"
        Else
          lblInstructs = "Select from the list of available areas"
        End If
        If lstArea.LeftCount > 1 And MyP.UserOpt <> 1 Then
          lblInstructs = lblInstructs & "." & vbCrLf & "You may select multiple " _
              & "areas from the Available " & MyP.UnitArea & " list."
        End If
      Else
        lblInstructs = "Choose from one of the three Area Unit IDs"
      End If
    ElseIf Left(lstArea.LeftItem(0), 2) = "No" Then
      lblInstructs = "Try a different Area Unit"
    ElseIf NationalDB And (MyP.UserOpt = 9 Or MyP.UserOpt = 11) Then 'PRH 11/2005
      lblInstructs = "Select one of the two Area Units."
    Else
      lblInstructs = "Select one of the four Area Units."
    End If
  ElseIf tabMain.Tab = 3 Then
    fraWhat2Do.Caption = "Instructions" & operation
    If lstDataCats.RightCount > 0 Then
      If MyP.UserOpt = 1 Then
        lblInstructs = "Click '" & cmdCatOpt.Caption & "' to edit selected category"
      ElseIf MyP.UserOpt > 2 And MyP.UserOpt < 8 Then
        lblInstructs = "Click '" & cmdCatOpt.Caption & "' to produce the " _
            & Mid(operation, 6)
      ElseIf MyP.UserOpt > 7 And MyP.UserOpt < 11 Then
        lblInstructs = "Click '" & cmdCatOpt.Caption & "' to execute " _
            & Mid(operation, 19)
      ElseIf MyP.UserOpt = 11 Then
        lblInstructs = "Click '" & cmdCatOpt.Caption & "' to export selected data"
      End If
      If MyP.UserOpt = 3 Or MyP.UserOpt = 10 Or MyP.UserOpt = 11 Then
        lblInstructs = lblInstructs & "." & vbCrLf & "You may select " _
            & "multiple data categories if desired."
      End If
    ElseIf MyP.UserOpt = 1 Then
      lblInstructs = "Selected one category"
      If Asterisk Then
        lblInstructs = lblInstructs & "." & vbCrLf & "An * before the " & _
            "listing indicates there is no data at this location for that category."
      End If
    ElseIf MyP.UserOpt = 3 Or MyP.UserOpt = 10 Or MyP.UserOpt = 11 Then
      lblInstructs = "Select one or more data categories"
      If Asterisk Then
        lblInstructs = lblInstructs & "." & vbCrLf & "An * before the " & _
            "listing indicates there is no data at any location for that category."
      End If
    End If
  ElseIf tabMain.Tab = 4 Then
    fraWhat2Do.Caption = "Instructions" & operation
    EditOK = False
    lblInstructs = "Edit data fields then press 'Save' to finalize the changes" & _
        vbCrLf & "Fields highlighted in red are required data elements." & _
        vbCrLf & "Fields highlighted in blue are required, but null values are allowed."
  End If

End Sub

Private Sub txtCurFile_GotFocus()
Attribute txtCurFile_GotFocus.VB_Description = "Records when user selects new file for import."
' ##SUMMARY Records when user selects new file for import.
' ##REMARKS Check is made after TextBox loses focus to ensure that an Excel file was _
          selected.
  NewImpFile = True
End Sub

Private Sub txtCurFile_LostFocus()
Attribute txtCurFile_LostFocus.VB_Description = "Makes sure file selected for import is an Excel workbook."
' ##SUMMARY Makes sure file selected for import is an Excel workbook.
' ##HISTORY 5/21/2007, prhummel Use file filter/extension determined from version of Excel

'  If LCase(Right(txtCurFile, 4)) <> ".xls" And NewImpFile = True Then
  If LCase(Right(txtCurFile, 4)) <> Right(XLFileExt, 4) And NewImpFile = True Then
    MyMsgBox.Show _
        "The import file must be an Excel(" & XLFileExt & ") file with the proper format." & vbCrLf _
        & "See 'Import' in the on-line help for more information.", _
        "Import file type verification", "+-&OK"
    NewImpFile = False
  End If
End Sub

Private Sub txtCurFile_Change()
Attribute txtCurFile_Change.VB_Description = "Opens selected Excel file and reads which year of data is being imported."
' ##SUMMARY Opens selected Excel file and reads which year of data is being imported.
' ##REMARKS Writes year to interface for user inspection.
  Dim excelApp As Excel.Application
  Dim rangeTemp As Excel.Range
  Dim xlRange As Excel.Range
  Dim i As Integer
  Dim firstLine As Long
  Dim excelBook As Excel.Workbook
  Dim excelSheet As Excel.Worksheet
  
  If Len(Dir(txtCurFile)) > 0 Then
    On Error GoTo x
    Set excelApp = New Excel.Application
    Set excelBook = Workbooks.Open(txtCurFile)
    Set excelSheet = Worksheets(1)
    excelSheet.Activate
    With excelSheet
      Set xlRange = Range(Cells(1, 1), Cells(10, 20))
      For i = 1900 To 2050
        Set rangeTemp = xlRange.Find(" " & i & " ", , xlValues, , xlByRows, xlNext)
        If Not rangeTemp Is Nothing Then
          txtYear.value = i
          Exit For
        End If
      Next i
    End With
    excelBook.Close False
x:
    excelApp.Quit
    Set excelApp = Nothing
  End If
  If IsNumeric(txtYear.value) Then
    If (txtYear.value > 1899 And txtYear.value < 2050) _
        And Len(Dir(txtCurFile)) > 0 Then
      cmdImport.Enabled = True
    Else
      cmdImport.Enabled = False
    End If
  End If
End Sub

Private Sub txtDataFlds_GotFocus(index As Integer)
' ##SUMMARY Selects/highlights text in user-selected textbox.
' ##PARAM Index I Integer indicating the selected item in the array of TextBoxes.
  Dim i As Long
  
    txtDataFlds(index).SelStart = 0
    txtDataFlds(index).SelLength = Len(txtDataFlds(index))
End Sub

Private Sub txtDataFlds_Change(index As Integer)
Attribute txtDataFlds_Change.VB_Description = "Enables "
' ##SUMMARY Enables "Save" button if data values for current category have been changed.
' ##PARAM Index I Integer indicating the selected item in the array of TextBoxes.
  Dim i As Long
  If tabMain.Tab = 4 And DeleteOK = False Then
    For i = 0 To txtDataFlds.Count - 1
      cmdEdit.Enabled = False
      If PreEditVal(i) <> txtDataFlds(i) Then
        cmdEdit.Enabled = True
        Exit For
      End If
    Next i
  End If
  If PreEditVal(index) <> txtDataFlds(index) Then
    txtDataFlds(index) = DataEntry(txtDataFlds(index), index)
  End If
End Sub

Function DataEntry(Entry As String, Optional index As Integer) As String
Attribute DataEntry.VB_Description = "Performs QA check on prospective value to be entered in database."
' ##SUMMARY Performs QA check on prospective value to be entered in database.
' ##PARAM Entry I String user-entered value.
' ##PARAM index I Integer array index of selected field on data entry tab.
' ##RETURNS User-entered value, with leading '-' sign or most recently typed _
   non-numeric character removed if need be ('HY-OfPow' is the only field that _
   can have a negative value).
   
  Dim i As Integer       'length of user-entered value
  Dim maxVal As Single   'maximum allowable user-entered value
  Dim oneDec As Boolean  'true if entry contains a decimal place
  
  Entry = Trim(Entry)
  If Len(Entry) > 0 Then
    'Check for non-numeric symbols; not allowed
    For i = 1 To Len(Entry)
      If Mid(Entry, i, 1) = "." Then
        If oneDec Then
          MsgBox "Only one decimal place is allowed per entry." & vbCrLf & _
            "The entry will be reset to its previous value.", , "Improper Syntax"
          DataEntry = PreEditVal(index)
          Exit Function
        Else
          oneDec = True
        End If
      ElseIf Not IsNumeric(Mid(Entry, i, 1)) And Mid(Entry, i, 1) <> "." Then
        If Not (i = 1 And Mid(Entry, i, 1) = "-") Then
          MsgBox "The character " & Mid(Entry, i, 1) & " is not allowed; it will be removed", , "Improper Syntax"
          DataEntry = Left(Entry, i - 1) & Mid(Entry, i + 1)
          Exit Function
        End If
      End If
    Next i
    If Not IsNumeric(Entry) Then
      If Not (Entry = "-" Or Entry = "." Or (Left(Entry, 1) = "." And IsNumeric(Mid(Entry, 2)))) Then
        MyMsgBox.Show _
            "You have entered a non-numeric value." & vbCrLf & _
            "Only positive numbers may be entered", _
            "Bad data value", "+-&OK"
        i = Len(Entry)
        Entry = Left(Entry, i - 1)
      End If
    ElseIf Entry < 0 Then
      If lblDataFlds(index) <> "HY-OfPow" Then
        i = Len(Entry)
        MyMsgBox.Show _
            "You have entered a negative value for " & lblDataFlds(index) & "." & vbCrLf & _
            "Only positive numbers may be entered", _
            "Bad data value", "+-&OK"
        Entry = Right(Entry, i - 1)
      ElseIf Entry < -999.99 Then
        MyMsgBox.Show _
            "You have entered a value outside the allowable range for HY-OfPow." & vbCrLf & _
            "The allowable range is (-999.99 to 99999.99)." & vbCrLf & _
            "The entry will be reset to its previous value.", _
            "Bad data value", "+-&OK"
        txtDataFlds(index) = PreEditVal(index)
        Entry = PreEditVal(index)
      End If
    ElseIf Entry > 99999.999 Then
      Entry = Right(Entry, Len(Entry) - 1)
      If LCase(Right(lblDataFlds(index), 6)) = "-facil" Or _
         LCase(Right(lblDataFlds(index), 3)) = "fac" Then
        maxVal = 99999
      ElseIf InStr(1, LCase(lblDataFlds(index)), "pop") > 0 Then
        maxVal = 99999.999
      Else
        maxVal = 99999.99
      End If
      MyMsgBox.Show _
          "You have entered too large of a value." & vbCrLf & _
          "The maximum allowable data value is " & maxVal & "." & vbCrLf & _
          "The entry will be reset to its previous value.", _
          "Bad data value", "+-&OK"
       txtDataFlds(index) = PreEditVal(index)
       Entry = PreEditVal(index)
    End If
    'Check to ensure decimal place limit not exceeded.
    i = InStr(1, Entry, ".")
    If i > 0 Then
      i = Len(Entry) - i
      If LCase(Right(lblDataFlds(index), 6)) = "-facil" Or _
         LCase(Right(lblDataFlds(index), 3)) = "fac" Then
        MsgBox "'number of facilities' must be an integer."
        Entry = Round(Entry, 0)
      ElseIf i > 2 Then
        If InStr(1, LCase(lblDataFlds(index)), "pop") > 0 Then
          If i > 3 Then MsgBox "You have entered " & i & " decimal places for the '" & _
              lblDataFlds(index) & "' field." & vbCrLf & "Only 3 are allowed, so the value will be rounded."
          Entry = Format(Round(Entry, 3), "0.000")
        Else
          MsgBox "You have entered " & i & " decimal places for the '" & _
              lblDataFlds(index) & "' field." & vbCrLf & "Only 2 are allowed, so the value will be rounded."
          Entry = Format(Round(Entry, 2), "0.00")
        End If
      End If
    End If
  End If
  DataEntry = Entry
End Function

Private Sub ClearOpts()
Attribute ClearOpts.VB_Description = "Clears certain user selections depending upon which tab is in focus."
' ##SUMMARY Clears certain user selections depending upon which tab is in focus.
  Dim i As Long
  
  If tabMain.Tab < 2 Then
    For i = 0 To rdoAreaUnit.Count - 1
      rdoAreaUnit(i) = False
      rdoAreaUnit2(i) = False
      rdoAreaUnit(i).Enabled = True
      rdoAreaUnit2(i).Enabled = True
    Next i
    For i = 0 To rdoID.Count - 1
      rdoID(i) = False
    Next i
  End If
  If lstArea.RightCount > 0 Then
    lstArea.ClearRight
  End If
  If lstArea.LeftCount > 0 Then
    lstArea.ClearLeft
  End If
  If lstYears.ListCount > 0 Then
    lstYears.Clear
  End If
  lstDataCats.Enabled = True
  If lstDataCats.RightCount > 0 Then
    lstDataCats.ClearRight
  End If
  If lstDataCats.LeftCount > 0 Then
    lstDataCats.ClearLeft
  End If
  MyP.Year1Opt = 0
  MyP.Year2Opt = 0
  MyP.UnitArea = ""
  For i = 0 To 1
    rdoDataDict(i) = False
  Next i
  If tabMain.Tab = 0 Then
    MyP.UserOpt = 0
    fraImport.Visible = False
    fraNewYear.Visible = False
    For i = 0 To MainOpt.Count - 1
      MainOpt(i) = False
    Next i
    EmboldenMe MainOpt, -1
  End If
End Sub

Private Sub Ok2DoMore()
Attribute Ok2DoMore.VB_Description = "Automatically selects next tab if user has made necessary selections on current tab, and enables appropriate controls given user selections."
' ##SUMMARY Automatically selects next tab if user has made necessary selections on _
          current tab, and enables appropriate controls given user selections.
  Dim userOk As Boolean
  
  MyP.numAreas = lstArea.RightCount
  userOk = (MyP.StateValid Or NationalDB) And fraDomainName.Visible = False
  DataSelOK = MyP.numAreas > 0 And _
      (rdoID(0) = True Or rdoID(1) = True Or rdoID(2) = True _
       Or MainOpt(0) Or MainOpt(8) Or MainOpt(10)) _
      And MyP.YearValid And (Len(MyP.YearFields) > 0)
  If tabMain.Tab <> 3 And tabMain.Tab <> 4 Then
    If MyP.UserOptValid = True Then
      tabMain.Tab = 2
    ElseIf userOk = True Then
      tabMain.Tab = 1
    Else
      tabMain.Tab = 0
    End If
  End If
  
  If lstArea.RightCount > 0 Then
    cmdSaveGroup.Enabled = True
  Else
    cmdSaveGroup.Enabled = False
  End If
  
  tabMain_Click (tabMain.Tab)
 
  tabMain.TabEnabled(1) = userOk
  tabMain.TabEnabled(2) = tabMain.TabEnabled(1) And MyP.UserOptValid
  cmdRetrieve.Enabled = tabMain.TabEnabled(2)
  cmdExeOpts.Enabled = tabMain.TabEnabled(2) And DataSelOK
  tabMain.TabEnabled(3) = (tabMain.Tab = 3)
  If lstDataCats.RightCount > 0 Then cmdCatOpt.Enabled = True Else cmdCatOpt.Enabled = False
  tabMain.TabEnabled(4) = (tabMain.Tab = 4)
  
End Sub

Private Sub fillArea()
Attribute fillArea.VB_Description = "Populates area ListBox on 3rd tab with available HUCs/counties/aquifers."
' ##SUMMARY Populates area ListBox on 3rd tab with available HUCs/counties/aquifers.
  Dim listName As String
  Dim sql As String
  Dim unitAreaRec As Recordset
  Dim i As Long
  
  On Error GoTo x
  
  'Use SQL language to create recordset of unit area codes and names
  ' from the "county", "huc", "aquifer" tables for the selected state.
  ' For example:
  '   SELECT Trim(County_cd) As Code, Trim(County_nm) As Name FROM [County]
  '   WHERE state_cd='10' AND len(trim(County_cd))=3
  '   ORDER BY Trim(County_cd);
  sql = "SELECT Trim(" & TableName & "_cd) As Code, Trim(" & TableName & "_nm) As Name" & _
      " FROM [" & TableName & _
      "] WHERE state_cd='" & MyP.stateCode & "'" & _
      " AND len(trim(" & TableName & "_cd))=" & MyP.Length & _
      " ORDER BY Trim(" & TableName & "_cd);"
  Set unitAreaRec = MyP.stateDB.OpenRecordset(sql, dbOpenForwardOnly)

  If (MyP.UserOpt = 1 And Not rdoAreaUnit(3)) Or MyP.UserOpt = 5 Or MyP.UserOpt = 10 Then
    For i = 1 To MyP.Length
      listName = listName & "0"
    Next i
    lstArea.LeftItem(0) = listName & " - STATE VALUES"
  End If
  While Not unitAreaRec.EOF
    listName = unitAreaRec("Code") & " - " & unitAreaRec("Name")
    If MainOpt(8) = False Then  'not Comp2Years report
      lstArea.LeftItem(lstArea.LeftCount) = listName
    Else
      lstArea.RightItem(lstArea.RightCount) = listName
    End If
    unitAreaRec.MoveNext
  Wend
  unitAreaRec.Close
  If lstArea.LeftCount = 0 And lstArea.RightCount = 0 Then
x:
    lstArea.LeftItem(lstArea.LeftCount) = "No data available for this unit area."
  End If
End Sub

Private Sub FillDate()
Attribute FillDate.VB_Description = "Populates year listbox on 3rd tab with all years for which data are available, depending upon area selections."
' ##SUMMARY Populates year listbox on 3rd tab with all years for which data are available, _
          depending upon area selections.
  Dim sql As String
  Dim findDate As String
  Dim i As Long
  Dim j As Long
  Dim Length As Long
  Dim length2 As Long
  Dim dateRec As Recordset
  Dim dateRec2 As Recordset
  Dim ready As Boolean

  On Error GoTo y
  lstYears.Clear
  Length = MyP.Length
  length2 = MyP.length2
  If AggregateHUCs = True Then
    If Length = 4 Then Length = 8
  End If
  'Open recordset with list of distinct dates for available data
  sql = "Select DISTINCT Date from [" & TableName & "Data]" & _
        " WHERE Len(Trim(Location))=" & Length & _
        " ORDER BY Date DESC;"
  Set dateRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  If MyP.UserOpt = 9 Then  'compare areas report
    If Len(MyP.UnitArea) > 0 And Len(MyP.UnitArea2) > 0 Then
      sql = "Select DISTINCT Date from [" & TableName2 & "Data]" & _
            " WHERE Len(Trim(Location))=" & length2 & _
            " ORDER BY Date DESC;"
      Set dateRec2 = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
      ready = True
    End If
  Else
    ready = True
  End If
  If ready = True Then
    While Not dateRec.EOF
      lstYears.AddItem dateRec("Date")
      dateRec.MoveNext
    Wend
    j = 0
    If MyP.UserOpt = 9 Then
      For i = 0 To lstYears.ListCount - 1
        While Not dateRec2.EOF
          If lstYears.List(j) = dateRec2("Date") Then GoTo x
          dateRec2.MoveNext
        Wend
        lstYears.RemoveItem (j)
        j = j - 1
        dateRec2.MoveFirst
x:
        j = j + 1
      Next i
    End If
    dateRec.Close
y:
    If lstYears.ListCount = 0 Then lstYears.List(0) = "none"
  End If
End Sub

Private Sub FillCats()
Attribute FillCats.VB_Description = "Fills the listbox of available categories on the 4th tab."
' ##SUMMARY Fills the listbox of available categories on the 4th tab.
' ##REMARKS The available list is dependent upon which Data Dictionary is selected as the _
          reference and which data storage options are selected.
  Dim CatRec As Recordset, catRec2 As Recordset, haveDataRec As Recordset
  Dim year As Long
  Dim i As Long
  Dim addOn As String
  Dim sql As String
  Dim dontAdd As Boolean
  Dim noData As Boolean
  
  TwoYears = False
  WhichCats CatRec
  Asterisk = False
  
  'Get the categories for the dataset resulting from user selections
  If MyP.UserOpt = 1 And Left(lstArea.RightItem(0), 3) = "000" Then
    sql = "SELECT QualFlg FROM [LastEdit]"
    Set haveDataRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    Select Case haveDataRec("QualFlg")
      Case 0, 1, 5, 7, 8: MyMsgBox.Show _
                      "There are no state values for this year of data.", _
                      "Data Inadequacy Notification", "+-&OK"
      Case 2: lstDataCats.LeftItem(0) = "Domestic"
              lstDataCats.LeftItemData(0) = 4
      Case 3: lstDataCats.LeftItem(0) = "Public Supply"
              lstDataCats.LeftItemData(0) = 2
      Case 4: lstDataCats.LeftItem(0) = "Public Supply"
              lstDataCats.LeftItemData(0) = 2
              lstDataCats.LeftItem(1) = "Domestic"
              lstDataCats.LeftItemData(1) = 4
      Case 6: lstDataCats.LeftItem(0) = "Domestic"
              lstDataCats.LeftItemData(0) = 4
    End Select
  Else
    If MyP.UserOpt = 10 Then
      TwoYears = True
      WhichCats catRec2
      Cats = ""
      While Not CatRec.EOF
        While Not catRec2.EOF
          If IRinTwo <> IRinTwo2 And (CatRec("ID") = 16 Or CatRec("ID") = 17) Then
            'Irrigation storage options don't match - have to do total
            If Len(Cats) = 0 Then
              Cats = " And ([" & CatTable & "].ID=16"
            Else
              Cats = Cats & " Or [" & CatTable & "].ID=16"
            End If
            'Move over Irr-crop and Irr-golf if necessary
            While CatRec("ID") < 20 And Not CatRec.EOF
              CatRec.MoveNext
            Wend
            CatRec.MovePrevious
            catRec2.MoveLast
          ElseIf CatRec("ID") = catRec2("ID") Then
            If Len(Cats) = 0 Then
              Cats = " And ([" & CatTable & "].ID=" & CatRec("ID")
            Else
              Cats = Cats & " Or [" & CatTable & "].ID=" & CatRec("ID")
            End If
            catRec2.MoveLast
          End If
          catRec2.MoveNext
        Wend
        catRec2.MoveFirst
        CatRec.MoveNext
      Wend
      Cats = Cats & ")"
      Set catRec2 = Nothing
      If Len(Cats) > 0 Then sql = " WHERE " & Mid(Cats, 6) Else sql = ""
      'Use SQL language to create recordset with all categories to be used
      ' in "Compare Data for 2 Years" report
      ' For example:
      '  SELECT * From [Category2]
      '  WHERE ([Category2].ID=1 Or [Category2].ID=2 Or [Category2].ID=3 Or ...
      '    ... Or [Category2].ID=21 Or [Category2].ID=22 Or [Category2].ID=23)
      '  ORDER BY [Category2].ID;
      sql = sql & " ORDER BY [" & CatTable & "].ID;"
      Set CatRec = MyP.stateDB.OpenRecordset("SELECT * From [" & CatTable & "]" & sql, dbOpenSnapshot)
    End If
    'Use SQL language to create recordset with data values and associated
    ' category of each datum.  The recordset will be used to determine
    ' which categories actually have data entered for them (i.e., not all
    ' data fields in category are null).  Data values are joined with data
    ' dictionary to link datum to their respective categories.
    ' For example:
    '  SELECT [Category2].ID, [CountyData].Value
    '  FROM (([Category2] INNER JOIN [Field1] ON [Category2].ID = [Field1].CategoryID)
    '    INNER JOIN [CountyData] ON [Field1].ID = [CountyData].FieldID)
    '    INNER JOIN [2000Fields1] ON [Field1].ID = [2000Fields1].FieldID
    '  Where ([CountyData].Date = 2000 Or [CountyData].Date = 0)
    '    And IsNull([CountyData].Value) = False
    '  ORDER BY [Category2].ID ASC
    If MyP.UserOpt <> 1 And MyP.UserOpt <> 9 Then
      sql = "SELECT [" & CatTable & "].ID, [" & MyP.AreaTable & "].Value " & _
            "FROM (([" & CatTable & "] INNER JOIN [" & FieldTable & "] ON [" & CatTable & "].ID = [" & FieldTable & "].CategoryID) " & _
            "INNER JOIN [" & MyP.AreaTable & "] ON [" & FieldTable & "].ID = [" & MyP.AreaTable & "].FieldID) " & _
            "INNER JOIN [" & MyP.YearFields & "] ON [" & FieldTable & "].ID = [" & MyP.YearFields & "].FieldID " & _
            "WHERE ([" & MyP.AreaTable & "].Date = " & MyP.Year1Opt & _
                   " Or [" & MyP.AreaTable & "].date = " & MyP.Year2Opt & _
                   ") And IsNull([" & MyP.AreaTable & "].Value) = False" & Areas & _
            " ORDER BY [" & CatTable & "].ID ASC"
    Else
      sql = "SELECT [" & CatTable & "].ID, [" & MyP.AreaTable & "].Value " & _
            "FROM (([" & CatTable & "] INNER JOIN [" & FieldTable & "] ON [" & CatTable & "].ID = [" & FieldTable & "].CategoryID) " & _
            "INNER JOIN [" & MyP.AreaTable & "] ON [" & FieldTable & "].ID = [" & MyP.AreaTable & "].FieldID) " & _
            "INNER JOIN [" & MyP.YearFields & "] ON [" & FieldTable & "].ID = [" & MyP.YearFields & "].FieldID " & _
            "WHERE [" & MyP.AreaTable & "].Date = " & MyP.Year1Opt & _
                   " And IsNull([" & MyP.AreaTable & "].Value) = False" & _
                   " And [" & MyP.AreaTable & "].Location = '" & LocnArray(0, 0) & _
            "' ORDER BY [" & CatTable & "].ID ASC"
    End If
    Set haveDataRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    'Determine categories without data and put an asterisk (*) next to them
    While Not CatRec.EOF
      noData = False
      haveDataRec.FindFirst "ID=" & CatRec("ID")
      If haveDataRec.NoMatch And MyP.DataOpt > 0 Then noData = True
      If MyP.UserOpt = 1 Or MyP.UserOpt = 3 Or MyP.UserOpt = 4 Or MyP.UserOpt > 9 Then
        dontAdd = False
        If ((MyP.UserOpt = 1 Or MyP.UserOpt = 9 Or MyP.UserOpt = 11) And (CatRec("Original"))) = -1 _
            Or (Left(TableName, 3) <> "HUC" And CatRec("Name") = "RE") _
            Or (MyP.UserOpt = 3 And CatRec("Name") = "TP") _
            Or (MyP.UserOpt = 4 And CatRec("Name") = "TP") _
            Then dontAdd = True
        If NationalDB And MyP.UserOpt = 9 And CatRec("Name") = "IT" Then dontAdd = False
        If dontAdd = False Then
          If noData = False Or CatRec("Original") = -1 Then
            lstDataCats.LeftItem(lstDataCats.LeftCount) = CatRec("Description")
          Else
            lstDataCats.LeftItem(lstDataCats.LeftCount) = "*" & CatRec("Description")
            Asterisk = True
          End If
          lstDataCats.LeftItemData(lstDataCats.LeftCount - 1) = CatRec("ID")
        End If
      Else
        If (Left(MyP.AreaTable, 3) <> "HUC" And CatRec("Name") = "RE") _
            Or (MyP.UserOpt = 9 And (CatRec("Original") = -1 Or CatRec("Name") = "TP")) _
            Or (MyP.UserOpt = 5 And (CatRec("Original") = -1 And Not (NationalDB And CatRec("Name") = "IT"))) Then
            'Or (noData And (catRec("ID") = 17 Or catRec("ID") = 18 Or catRec("ID") = 19)) Then
          dontAdd = True
        Else
          dontAdd = False
        End If
        If NationalDB And MyP.UserOpt = 9 And CatRec("Name") = "IT" Then dontAdd = False
        If dontAdd = False Then
          If noData = False Or CatRec("Original") = -1 Then
            lstDataCats.RightItem(lstDataCats.RightCount) = CatRec("Description")
          Else
            lstDataCats.RightItem(lstDataCats.RightCount) = "*" & CatRec("Description")
          End If
          lstDataCats.RightItemData(lstDataCats.RightCount - 1) = CatRec("ID")
          lstDataCats.Enabled = False
          lstDataCats.ToolTipText = "All categories with data on file automatically selected"
          cmdCatOpt.Enabled = True
        End If
      End If
      CatRec.MoveNext
    Wend
    haveDataRec.Close
    addOn = ""
    For i = 0 To lstDataCats.LeftCount - 1
      If Left(lstDataCats.LeftItem(i), 1) <> "*" Then GoTo y
    Next i
    addOn = "*"
y:
    If MyP.UserOpt = 3 Then
      If Not rdoAreaUnit(3) Then
        lstDataCats.LeftItem(lstDataCats.LeftCount) = addOn & "Totals, SW by categories"
        lstDataCats.LeftItemData(lstDataCats.LeftCount - 1) = 50
      End If
      lstDataCats.LeftItem(lstDataCats.LeftCount) = addOn & "Totals, GW by categories"
      lstDataCats.LeftItemData(lstDataCats.LeftCount - 1) = 51
      If Not rdoAreaUnit(3) Then
        lstDataCats.LeftItem(lstDataCats.LeftCount) = addOn & "Totals, Overall by categories"
        lstDataCats.LeftItemData(lstDataCats.LeftCount - 1) = 52
      End If
    End If
    If MyP.UserOpt = 1 Then
      lblInstructs = "Select the data category of choice"
    Else
      lblInstructs = "Select the data categories of choice"
    End If
    CatRec.Close
  End If
x:
End Sub

Private Sub WhichCats(CatRec As Recordset)
Attribute WhichCats.VB_Description = "Checks which data storage options are used for this state and assigns field table, category table, and Excel header file accordingly."
' ##SUMMARY Checks which data storage options are used for this state and assigns field _
          table, category table, and Excel header file accordingly.
' ##PARAM catRec O Recordset containing the categories to be listed on the 4th tab of the _
          main form.
  Dim sql As String
  Dim addOn As String
  Dim opt As String
  Dim myRec As Recordset
  
  If MyP.UserOpt = 1 Then
    opt = "LastEdit"
  ElseIf MyP.UserOpt = 11 Then
    opt = "LastExport"
  Else
    opt = "LastReport"
  End If
  Set myRec = MyP.stateDB.OpenRecordset(opt, dbOpenSnapshot)
  If TwoYears Then myRec.MoveLast
  If fraAreaUnits2.Visible Then 'comparing totals by area
    myRec.MoveLast
    'Ensure compatibility b/t data storage options
    If MyP.DataOpt > 4 Or MyP.DataOpt2 > 4 Then
      If (MyP.DataOpt2 > 4 And (MyP.DataOpt = 2 Or MyP.DataOpt = 4 Or MyP.DataOpt = 6)) Or _
         (MyP.DataOpt > 4 And (MyP.DataOpt2 = 2 Or MyP.DataOpt2 = 4 Or MyP.DataOpt2 = 6)) Then
        opt = 6
      Else
        opt = 5
      End If
    ElseIf MyP.DataOpt = 4 Or MyP.DataOpt2 = 4 Then
      opt = 4
    ElseIf MyP.DataOpt = 3 Or MyP.DataOpt2 = 3 Then
      opt = 3
    ElseIf MyP.DataOpt = 2 Or MyP.DataOpt2 = 2 Then
      opt = 2
    ElseIf MyP.DataOpt > 0 Or MyP.DataOpt2 > 0 Then
      opt = 1
    Else
      opt = 0
    End If
  Else
    If (NationalDB And myRec("QualFlg") <> 0) Then
      opt = 5
    Else
      If myRec("QualFlg") < 7 Then
        opt = myRec("QualFlg")
      ElseIf myRec("QualFlg") = 7 Then
        opt = 1
      ElseIf myRec("QualFlg") = 8 Then
        opt = 5
      End If
    End If
  End If
  If opt = 0 Then
    MyP.YearFields = "1995Fields1"
    CatTable = "Category1"
  ElseIf rdoAreaUnit(3) Then  'aquifer
    MyP.YearFields = "2000FieldsA"
    CatTable = "Category2"
  Else
    MyP.YearFields = "2000Fields" & opt
    CatTable = "Category2"
  End If
  If rdoAreaUnit(3) Or (rdoAreaUnit2(3) And MyP.UserOpt = 9) Then  'aquifer
    FieldTable = "FieldA"
    HeaderFile = "HeaderA.xls"
  Else
    FieldTable = "Field" & opt
    HeaderFile = "Header" & opt & ".xls"
  End If
  myRec.FindFirst "FieldID=282"  'Crop Irrigation field
  'Sort out Irrigation options
  If MyP.UserOpt = 9 Then 'Comparing Areas
    If myRec.NoMatch Or Len(Trim(myRec("Location"))) <> MyP.Length Then IRinTwo = False Else IRinTwo = True
    myRec.MoveLast
    myRec.FindPrevious "FieldID=282"
    If Len(Trim(myRec("Location"))) = MyP.length2 And MyP.DataOpt2 > 0 Then
      IRinTwo2 = True
    Else
      IRinTwo2 = False
    End If
  ElseIf MyP.UserOpt = 10 Then 'comparing 2 years
    If myRec.NoMatch Or myRec("Date") <> MyP.Year1Opt Then IRinTwo = False Else IRinTwo = True
    myRec.MoveLast
    myRec.FindPrevious "FieldID=282"
    If myRec("Date") = MyP.Year2Opt And MyP.DataOpt2 > 0 Then
      IRinTwo2 = True
    Else
      IRinTwo2 = False
    End If
  Else
    If myRec.NoMatch Then IRinTwo = False Else IRinTwo = True
  End If
  'Figure our which Irrigation categories to list
  If (MyP.DataOpt = 0 And MyP.DataOpt2 = 0) Or _
     (Not NationalDB And Not IRinTwo And Not IRinTwo2) Or _
     (Not NationalDB And IRinTwo <> IRinTwo2 And (MyP.UserOpt = 9 Or MyP.UserOpt = 10) Or _
     (NationalDB And MyP.UserOpt = 11)) Then 'this clause of If to display IR for national export, PRH 11/2005
    addOn = "Where " & CatTable & ".ID<>16 And " & CatTable & ".ID<>18 And " & CatTable & ".ID<>19 "
  ElseIf IRinTwo And (IRinTwo2 Or (MyP.UserOpt <> 9 And MyP.UserOpt <> 10)) And Not NationalDB Then
    addOn = "Where " & CatTable & ".ID<>17 "
  Else
    addOn = "Where " & CatTable & ".ID<>17 And " & CatTable & ".ID<>18 And " & CatTable & ".ID<>19 "
  End If
  'Use SQL language to create recordset with all possible categories from one of the
  ' "Category_" tables in Categories.mdb.  Since Irrigation can be a single category (17)
  ' or divided into crops and golf courses (18 and 19), one of those options is excluded
  ' from the query.
  ' For example:
  '   SELECT DISTINCT [Category2].*
  '   FROM ([Category2] INNER JOIN [Field1] ON [Category2].ID = [Field1].CategoryID)
  '     INNER JOIN [2000Fields1] ON [Field1].ID = [2000Fields1].FieldID
  '   Where Category2.id <> 17
  '   ORDER BY [Category2].ID ASC
  sql = "SELECT DISTINCT [" & CatTable & "].* " & _
        "FROM ([" & CatTable & "] INNER JOIN [" & FieldTable & "] ON [" & CatTable & "].ID = [" & FieldTable & "].CategoryID) " & _
        "INNER JOIN [" & MyP.YearFields & "] ON [" & FieldTable & "].ID = [" & MyP.YearFields & "].FieldID " & _
        addOn & "ORDER BY [" & CatTable & "].ID ASC"
  Set CatRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  myRec.Close
End Sub

Private Sub cmdCatOpt_Click()
Attribute cmdCatOpt_Click.VB_Description = "Instantiates selected operation (report production or editing session)."
' ##SUMMARY Instantiates selected operation (report production or editing session).
' ##REMARKS For reports, raises dialog box for output file then calls report module.
' ##HISTORY 5/21/2007, prhummel Use file filter/extension determined from version of Excel
  Dim i As Long
  Dim instrCount As Long
  Dim sql As String
  Dim fileTitle As String
  Dim CatRec As Recordset
  Dim areaRec As Recordset
  Dim startTime As Date
  
  On Error GoTo x
  
  Me.MousePointer = vbHourglass
  CreateTable
  i = 1
  
  If MyP.UserOpt = 1 Then
    cmdEdit.Enabled = False
    LoadControlArrays
    tabMain.TabEnabled(4) = True
    tabMain.Tab = 4
    Me.MousePointer = vbDefault
  ElseIf MyP.UserOpt > 2 Then
    With cdlgFileSel
      If MyP.UserOpt = 11 Then
        fileTitle = ReportPath & "Export_" & MyP.Year1Opt & _
            ", " & MyP.State & "_" & MyP.UnitArea
'        While Len(Dir(fileTitle & ".xls")) > 0
        While Len(Dir(fileTitle & XLFileExt)) > 0
          i = i + 1
          If i > 9 Then
            fileTitle = Left(fileTitle, Len(fileTitle) - 3)
          ElseIf i > 2 Then
            fileTitle = Left(fileTitle, Len(fileTitle) - 2)
          End If
          fileTitle = fileTitle & "-" & i
        Wend
        .DialogTitle = "Select name for export file"
        .filename = fileTitle
      Else
        .DialogTitle = "Assign name of report file"
        instrCount = InStr(operation, "'")
        RepTitle = Mid(operation, instrCount + 1, (Len(operation) - instrCount - 1))
        If MyP.UserOpt = 10 Then
          If Left(lstArea.RightItem(0), 3) = "000" Then 'State total selected
            fileTitle = ReportPath & RepTitle & " - " & MyP.State & " - State, " & _
                        MyP.Year1Opt & " vs " & MyP.Year2Opt
          Else
            fileTitle = ReportPath & RepTitle & " - " & MyP.State & " - " & MyP.UnitArea & _
                        ", " & MyP.Year1Opt & " vs " & MyP.Year2Opt
          End If
        ElseIf MyP.UserOpt = 9 Then
          fileTitle = ReportPath & RepTitle & " - " & MyP.State & " - " & MyP.UnitArea & "_" & MyP.UnitArea2 & ", " & MyP.Year1Opt
        Else
          fileTitle = ReportPath & RepTitle & " - " & MyP.State & " - " & MyP.UnitArea & ", " & MyP.Year1Opt
        End If
'        While Len(Dir(fileTitle & ".xls")) > 0
        While Len(Dir(fileTitle & XLFileExt)) > 0
          i = i + 1
          If i > 2 Then fileTitle = Left(fileTitle, Len(fileTitle) - 2)
          fileTitle = fileTitle & "-" & i
        Wend
        .filename = fileTitle
      End If
      fileTitle = .filename
      If Len(Dir(.filename)) > 0 Then Kill .filename
      .Filter = XLFileFilter
      .FilterIndex = 1
      .CancelError = True
      .ShowSave
      fileTitle = .filename
      ReportPath = PathNameOnly(fileTitle) & "\"
      SaveSetting "AWUDS", "Defaults", "ReportPath", ReportPath
      If MyP.UserOpt = 2 Then txtCurFile.Text = fileTitle
    End With
    
    Me.MousePointer = vbHourglass
    If Len(Cats) > 0 Then sql = " WHERE " & Mid(Cats, 6) Else sql = ""
    sql = sql & " ORDER BY [" & CatTable & "].ID;"
    Set CatRec = MyP.stateDB.OpenRecordset("SELECT * From [" & CatTable & "]" & sql, dbOpenSnapshot)
    CatRec.MoveLast
    CatRec.MoveFirst
    startTime = Time
    
    If TableName = "HUC8" Or (TableName = "Aquifer" And NationalDB) Then
      If TableName = "HUC8" Then
        If rdoAreaUnit(1) Then i = 8 Else i = 4
        sql = "SELECT DISTINCT [huc_cd] as [Code], [huc_nm] as [Name] FROM [huc] " & _
              "WHERE len(trim(huc_cd))=" & i & _
              " ORDER BY huc_cd;"
      Else
        sql = "SELECT DISTINCT [aquifer_cd] as [Code], [aquifer_nm] as [Name] " & _
              "FROM [aquifer] ORDER BY aquifer_cd;"
      End If
      Set areaRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
      With areaRec
        .MoveLast
        .MoveFirst
        ReDim LocnArray(2, .RecordCount - 1)
        For i = 0 To .RecordCount - 1
          LocnArray(0, i) = Trim(!code)
          LocnArray(1, i) = Trim(!Name)
          LocnArray(2, i) = LocnArray(1, i) & " - " & LocnArray(0, i)
          .MoveNext
        Next i
        .Close
      End With
    End If
    
    If NationalDB And MyP.UserOpt = 11 Then 'create national export file, PRH 11/2005
      NationalExport fileTitle, CatRec
    Else
      If MyP.UserOpt <> 9 And MyP.UserOpt <> 11 Then 'don't populate arrays for compare and export
        PopulateFieldArrays TableName
        PopulateDataArray
      End If
      Me.MousePointer = vbDefault
      
      Select Case MyP.UserOpt
        Case 3: ByCatReport fileTitle, CatRec
        Case 4: ByAreaReport fileTitle, CatRec
        Case 5: EnteredReport fileTitle, CatRec
        Case 6: CalcReport fileTitle, CatRec
        Case 7: FacilityReport fileTitle
        Case 8: QAReport fileTitle, CatRec
        Case 9: CompAreasReport fileTitle, CatRec
        Case 10: CompYearsReport fileTitle, CatRec
        Case 11: CreateNewXLBook fileTitle, CatRec, LocnArray
      End Select
    End If
  End If
x:
  Me.MousePointer = vbDefault
End Sub

Public Sub CreateTable(Optional TwoAreas As Boolean)
Attribute CreateTable.VB_Description = "Creates table in Categories.mdb titled "
' ##SUMMARY Creates table in Categories.mdb titled "LastEdit", "LastReport", _
          "LastImport", or "LastExport".
' ##REMARKS Edit/Import tables store values being overwritten, while&nbsp;Report/Export _
          store values used to create respective Excel spreadsheet.
' ##PARAM TwoAreas I Boolean indicating whether user has selected _
          'Compare State Totals by Area' report option.
' ##HISTORY PR 18442, 4/25/2007, prhummel Different versions of the year table were being _
            used during import, depending on whether or not data categories had been selected. _
            Updated conditional to always use "Fields1" table when importing.
           
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim Length As Long
  Dim year As Long
  Dim qdfReport As QueryDef
  Dim qdfAppend As QueryDef
  Dim catName As String
  Dim sql As String
  Dim tabName As String
  Dim yearTab As String
  Dim joinType As String
  Dim dataRec As Recordset
  Dim myTabDef As TableDef
  Dim myIndx As index
  Dim myDB As Database
  Dim wrkJet As Workspace
  
  Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
  
  UnloadControlArrays
  If tabMain.Tab = 3 Then
    If Left(lstDataCats.RightItem(0), 1) = "*" Then
      sql = Mid(lstDataCats.RightItem(0), 2)
    Else
      sql = lstDataCats.RightItem(0)
    End If
    fraDataFlds.Caption = "Data Fields in " & _
        LocnArray(2, 0) & " for " & sql & ", " & MyP.Year1Opt
  End If
  If (lstDataCats.LeftCount = 0 And lstDataCats.RightCount = 0) Or MyP.UserOpt = 2 Then
    'use Fields1 when importing (prh, 4/2007)
    If Left(MyP.YearFields, 1) = "2" Then
      yearTab = "2000Fields1"
    Else
      yearTab = "1995Fields1"
    End If
  Else
    If Left(MyP.YearFields, 1) = "2" Then
      yearTab = MyP.YearFields
    Else
      yearTab = MyP.YearFields
    End If
  End If
  
  File = "LastReport"
  If TwoAreas = False Then
    Length = MyP.Length
    tabName = TableName & "Data"
  Else
    Length = MyP.length2
    tabName = TableName2 & "Data"
  End If
  
  'construct the SQL with the selected years and areas
  If MyP.UserOpt = 9 Then  'comparing totals by area
    If TwoAreas = True Then
      If lstYears.SelCount = 1 Then
        Years = "([" & tabName & "].Date=" & MyP.Year1Opt & ")"
      Else
        Years = "([" & tabName & "].Date=" & MyP.Year1Opt & " Or [" & tabName & "].Date=" & MyP.Year2Opt & ")"
      End If
      If lstArea.LeftCount = 0 Then
        If Left(MyP.UnitArea2, 1) = "H" Then
          Areas = " AND Len(Trim(" & tabName & ".Location))=" & Length
        Else
          Areas = ""
        End If
      Else
        For i = 0 To lstArea.RightCount - 1
          If i = 0 Then
            Areas = " AND (" & tabName & ".Location='" & Left(lstArea.RightItem(i), Length) & "'"
          Else
            Areas = Areas & " Or " & tabName & ".Location='" & Left(lstArea.RightItem(i), Length) & "'"
          End If
          If i = lstArea.RightCount - 1 Then Areas = Areas & ")"
        Next i
      End If
    Else
      Years = "([" & TableName & "Data].Date=" & MyP.Year1Opt & ")"
      Areas = " AND Len(Trim(" & TableName & "Data.Location))=" & MyP.Length
    End If
  End If
  
  For j = 0 To lstDataCats.RightCount - 1
    If j = 0 Then
      Cats = " And (" & CatTable & ".ID=" & lstDataCats.RightItemData(j)
    Else
      Cats = Cats & " Or " & CatTable & ".ID=" & lstDataCats.RightItemData(j)
    End If
  Next j
  Cats = Cats & ")"
  'change HUC-8s to HUC-4s if necessary
  If (MyP.UserOpt <> 9 And MyP.UnitArea = "HUC - 4") Then
    If AggregateHUCs = True Then  'change areas to be queried from HUC-4s to HUC-8s
      If lstArea.LeftCount = 0 Then
        Areas = " AND Len(Trim([" & tabName & "].Location))=8"
      Else
        Areas = " AND (Left([HUCData].Location, 4) = '" & LocnArray(0, 0) & "'"
        For i = 1 To UBound(LocnArray, 2)
          Areas = Areas & " Or Left([HUCData].Location, 4) = '" & LocnArray(0, i) & "'"
        Next i
        Areas = Areas & ")"
      End If
    End If
  End If
  'provide only original data if editing and erase old table if it exists
  If MyP.UserOpt = 1 Then
    File = "LastEdit"
  ElseIf MyP.UserOpt = 2 Then
    File = "LastImport"
  ElseIf MyP.UserOpt = 11 Then
    File = "LastExport"
  Else
    File = "LastReport"
  End If
  On Error Resume Next
  If TwoAreas = False Then
    'Kill temp table in Categories.mdb and the link to it from state DB
    MyP.stateDB.Execute "DROP TABLE [" & File & "];"
    MyP.CatDB.Execute "DROP TABLE [" & File & "];"
  End If
  On Error GoTo 0

  'Omit certain fields under certain conditions
  If MyP.UserOpt = 3 Then OmitFlds = " And ([" & yearTab & "].Excluded<>3)" Else OmitFlds = ""
  If MyP.UserOpt = 1 And Left(MyP.YearFields, 4) = 1995 Then
    OmitFlds = OmitFlds & " AND (AllFields.ID<>4 AND AllFields.ID<>46)"
  End If
  
  'Create data table(s)
  If MyP.UserOpt = 9 Then joinType = "RIGHT " Else joinType = "INNER "
  'check if editing state values and make appropriate changes to SQL
  If Not (MyP.Length = 10 Or MyP.UserOpt = 2) Then 'area <> aquifer and not importing
    If MyP.UserOpt = 1 And CLng(LocnArray(0, 0)) = 0 Then  'editing state values
      Areas = " AND (AllFields.ID=4 OR AllFields.ID=40 OR AllFields.ID=43 OR AllFields.ID=46)"
      MyP.YearFields = "2000Fields1"
      yearTab = "2000Fields1"
    End If
  End If
  If Not TwoAreas Then
    'Use SQL language to create table "LastReport", "LastEdit", "LastImport", or "LastExport"
    ' in Categories.mdb containing data to be used in report, edit, import, or export operation.
    ' The data table and Data Dictionary tables are joined to access attributes from both tables.
    ' The criteria "trim([Field1].Formula)=''" ensures that the query only includes user-entered
    ' fields; i.e., not products of formulas referencing other fields.
    ' "Excluded" property is coded to omit certain fields from certain reports.
    ' For example:
    '  SELECT trim([CountyData].Location) as Location, [AllFields].CategoryID, [CountyData].FieldID,
    '    [CountyData].value, [CountyData].QualFlg, [CountyData].Date
    '    INTO LastReport IN 'C:\VBExperimental\AwudsBuild\Data\Categories.mdb'
    '  FROM (AllFields INNER JOIN [CountyData] ON [AllFields].ID = [CountyData].FieldID)
    '    INNER JOIN [1995Fields1] ON [CountyData].FieldID = [1995Fields1].FieldID
    '  WHERE (CountyData.Date=1995) AND ([CountyData].Location = '003' Or [CountyData].Location = '005')
    '    And ([1995Fields1].Excluded<>3) AND Len(Trim([AllFields].Formula))=0
    '  ORDER BY [CountyData].Date, [CountyData].Location, AllFields.ID
    If Right(CurDir, 4) = "Data" Or Right(CurDir, 3) = "Doc" Then ChDir ".."
    sql = "SELECT trim([" & tabName & "].Location) as Location, [AllFields].CategoryID, [" & tabName & "].FieldID, [" _
        & tabName & "].Value, [" & tabName & "].QualFlg, [" & tabName & "].Date INTO " & File & _
        " IN '" & AwudsDataPath & "Categories.mdb' "
    If (MyP.UserOpt = 10 And TwoYears) Or MyP.UserOpt = 9 Then  'perform query w/o join to YearFields
      sql = sql & _
          "FROM AllFields INNER JOIN [" & tabName & "] ON [AllFields].ID = [" & tabName & "].FieldID " & _
          "WHERE " & Years & Areas & OmitFlds & " AND Len(Trim([AllFields].Formula))=0 " & _
          "ORDER BY [" & tabName & "].Date, [" & tabName & "].Location, AllFields.ID"
    Else  'perform query w/ join to YearFields
      sql = sql & _
          "FROM (AllFields INNER JOIN [" & tabName & "] ON [AllFields].ID = [" & tabName & "].FieldID) " & _
          "INNER JOIN [" & yearTab & "] ON [" & tabName & "].FieldID = [" & yearTab & "].FieldID " & _
          "WHERE " & Years & Areas & OmitFlds & " AND Len(Trim([AllFields].Formula))=0 " & _
          "ORDER BY [" & tabName & "].Date, [" & tabName & "].Location, AllFields.ID"
    End If
    Set qdfReport = MyP.stateDB.CreateQueryDef("", sql)
    qdfReport.Execute
    qdfReport.Close
    Set qdfReport = Nothing
    'Check on data storage option
    If AggregateHUCs Then i = 8 Else i = MyP.Length
    If i = 2 Then i = 3  'in case of National DB for County
    sql = "SELECT [" & tabName & "].* FROM [" & tabName & "] WHERE " & Years & _
         " AND (Len([" & tabName & "].Location)=" & i & ") ORDER BY [" & tabName & "].Date;"
    Set dataRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot, False)
    If dataRec("QualFlg") < 7 Then
      MyP.DataOpt = dataRec("QualFlg")
      If dataRec("QualFlg") = 0 Then
        MyP.DataDict = 1995
      Else
        MyP.DataDict = 2000
      End If
    Else
      MyP.DataDict = 2005
      If dataRec("QualFlg") = 7 Then
        MyP.DataOpt = 1
      ElseIf dataRec("QualFlg") = 8 Then
        MyP.DataOpt = 5
      End If
    End If
    If MyP.UserOpt = 10 Then
      dataRec.MoveLast
      If dataRec("QualFlg") < 7 Then
        MyP.DataOpt2 = dataRec("QualFlg")
      ElseIf dataRec("QualFlg") = 7 Then
        MyP.DataOpt2 = 1
      ElseIf dataRec("QualFlg") = 8 Then
        MyP.DataOpt2 = 5
      End If
    End If
    dataRec.Close
  Else
    'Add 2nd year of data for "Compare Data for 2 Years" report or
    ' 2nd unit area (county,huc,aquifer) for "Compare State Totals by Area" report.
    ' Data is addended to "LastReport" table created above.
    ' For example:
    '  SELECT trim([CountyData].Location) as Location, [AllFields].CategoryID, [CountyData].FieldID,
    '    [CountyData].value , [CountyData].QualFlg, [CountyData].Date
    '  INTO LastReport IN 'C:\VBExperimental\AwudsBuild\Data\Categories.mdb'
    '  FROM AllFields INNER JOIN [CountyData] ON [AllFields].ID = [CountyData].FieldID
    '  Where ([CountyData].Date = 1995) And Len(Trim(CountyData.Location)) = 3 And Len(Trim([AllFields].formula)) = 0
    '  ORDER BY [CountyData].Date, [CountyData].Location, AllFields.ID
    sql = "INSERT INTO [LastReport] IN '" & AwudsDataPath & "Categories.mdb' " & _
        "SELECT trim([" & tabName & "].Location) as Location, [AllFields].CategoryID, [" & tabName & "].FieldID, [" & _
        tabName & "].Value, [" & tabName & "].QualFlg, [" & tabName & "].Date " & _
        "FROM [AllFields] INNER JOIN [" & tabName & "] ON [AllFields].ID = [" & tabName & "].FieldID " & _
        "WHERE " & Years & Areas & OmitFlds & " AND Len(Trim([AllFields].Formula))=0 " & _
        "ORDER BY [" & tabName & "].Date, [" & tabName & "].Location, [AllFields].ID"
    Set qdfAppend = MyP.stateDB.CreateQueryDef("", sql)
    qdfAppend.Execute
    qdfAppend.Close
    Set qdfReport = Nothing
    sql = "SELECT " & tabName & ".* FROM " & tabName & " WHERE " & Years & Areas
    Set dataRec = MyP.stateDB.OpenRecordset(sql, dbOpenForwardOnly, False)
    If dataRec("QualFlg") < 7 Then
      MyP.DataOpt2 = dataRec("QualFlg")
    ElseIf dataRec("QualFlg") = 7 Then
      MyP.DataOpt2 = 1
    ElseIf dataRec("QualFlg") = 8 Then
      MyP.DataOpt2 = 5
    End If
    dataRec.Close
  End If
  
  If MyP.UserOpt = 9 And TwoAreas = False Then  'recursive call to this routine to add 2nd area
    TwoAreas = True
    CreateTable (True)
  Else
    'Reestablish link from active state DB to new table
    MyP.CatDBClose
    With MyP.stateDB
      Set myTabDef = .CreateTableDef(File)
      myTabDef.Connect = ";DATABASE=" & AwudsDataPath & "Categories.mdb"
      myTabDef.SourceTableName = File
      .TableDefs.Append myTabDef
    End With
  End If
  
  'Aggregate HUC-8s into HUC-4s if necessary
  If AggregateHUCs And Not TwoAreas Then
    'Use SQL language to create temporary table to house HUC-4 values aggregated from
    ' the HUC-8 values in "LastReport" table in Categories.mdb. Only the left 4 digits
    ' of the Location field are retained, essentially reducing the HUC-8 identifier to
    ' a HUC-4. Also, the values for each field are summed across each HUC-4, thereby
    ' aggregating the HUC-8 data into HUC-4. After the "Temp" table is constructed,
    ' it replaces the "LastReport" table that was used to create it.
    ' For example:
    '  SELECT left([Location],4) AS Locn, [CategoryID], [FieldID],
    '    Sum([LastReport].[Value]) AS [Value], [QualFlg], [Date]
    '  INTO Temp IN 'C:\VBExperimental\AwudsBuild\Data\Categories.mdb' From [LastReport]
    '  GROUP BY left([Location],4), [LastReport].[CategoryID], [LastReport].[FieldID],
    '  [LastReport].[QualFlg], [LastReport].[Date];
    sql = "SELECT left([Location],4) AS Locn, [CategoryID], [FieldID], " & _
          "Sum([LastReport].[Value]) AS [Value], [QualFlg], [Date] " & _
          "INTO Temp IN '" & AwudsDataPath & "Categories.mdb' From [LastReport] " & _
          "GROUP BY left([Location],4), [LastReport].[CategoryID], [LastReport].[FieldID], " & _
          "[LastReport].[QualFlg], [LastReport].[Date];"
    Set qdfReport = MyP.stateDB.CreateQueryDef("", sql)
    qdfReport.Execute
    qdfReport.Close
    Set qdfReport = Nothing
    MyP.StateDBClose
    Set myDB = OpenDatabase(AwudsDataPath & "Categories.mdb")
    Set myTabDef = myDB.TableDefs("Temp")
    myTabDef.Fields(0).Name = "Location"
    myDB.Execute "DROP TABLE [LastReport];"
    myTabDef.Name = "LastReport"
    myDB.Close
  End If
  If tabMain.Tab = 3 And (NationalDB And MyP.UserOpt <> 11) And Not (TableName = "HUC8" Or TableName = "Aquifer") Then
    MyP.Length = 2
  End If
  wrkJet.Close

End Sub

Private Sub LoadControlArrays()
Attribute LoadControlArrays.VB_Description = "Populates arrays of labels and text boxes that store names, values, and units of data being edited on 5th tab of main form."
' ##SUMMARY Populates arrays of labels and text boxes that store names, values, and units _
          of data being edited on 5th tab of main form.
  Dim editRec As Recordset
  Dim fldRec As Recordset
  Dim i As Long
  Dim sql As String
  Dim fldSQL As String
  Dim highlight As String
  
  sql = "SELECT Required FROM [state] WHERE state_cd='" & MyP.stateCode & "'"
  Set editRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  ReqSt = editRec("Required")
  editRec.Close
  Set editRec = MyP.stateDB.OpenRecordset("LastEdit", dbOpenSnapshot)
  fldSQL = "CategoryID=" & lstDataCats.RightItemData(0)
  If Left(lstArea.RightItem(0), 3) = "000" Then
    fldSQL = fldSQL & " AND ([" & FieldTable & "].ID=4 OR [" _
        & FieldTable & "].ID=40 OR [" & FieldTable & "].ID=43)"
    txtDataFlds(0).BackColor = &HC0C0FF 'red
  Else
    txtDataFlds(0).BackColor = &HFFFFFF 'white
  End If
  'Use SQL language to create recordset with all user-entered fields for selected
  ' category that may be edited.  Different "Field" tables that comprise the
  ' data dictionary are joined. The criteria "trim([Field1].Formula)=''" ensures
  ' that the query only includes user-entered fields; i.e., not products of formulas
  ' referencing other fields.
  ' For example:
  '  SELECT * FROM [Field1]
  '  INNER JOIN [2000Fields1] ON [Field1].ID = [2000Fields1].FieldID
  '  WHERE CategoryID=11 And trim([Field1].Formula)=''
  '  ORDER BY [Field1].ID;
  sql = "SELECT * FROM [" & FieldTable & _
        "] INNER JOIN [" & MyP.YearFields & "] ON [" & FieldTable & "].ID = [" & MyP.YearFields & "].FieldID " & _
        "WHERE " & fldSQL & _
        " And trim([" & FieldTable & "].Formula)='' " & _
        "ORDER BY [" & FieldTable & "].ID;"
  Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  fldRec.MoveLast
  fldRec.MoveFirst
  ReDim PreEditVal(fldRec.RecordCount - 1)
  For i = 0 To fldRec.RecordCount - 1
    editRec.FindFirst "FieldID=" & fldRec("ID")
    lblDataFlds(i) = fldRec("Name")
    If Not IsNull(editRec("Value")) Then
      txtDataFlds(i) = editRec("Value")
      PreEditVal(i) = editRec("Value")
    Else
      txtDataFlds(i) = ""
      PreEditVal(i) = ""
    End If
    highlight = MyP.Required(fldRec("ID"), ReqSt)
    If highlight = "red" Then
      txtDataFlds(i).BackColor = &HC0C0FF
    ElseIf highlight = "blue" Then
      txtDataFlds(i).BackColor = &HFFFFC0
    Else
      txtDataFlds(i).BackColor = &HFFFFFF
    End If
    lblDataUnits(i) = fldRec("Description")
    If i < fldRec.RecordCount - 1 Then
      Load txtDataFlds(i + 1)
      txtDataFlds(i + 1).Top = (i + 2) * 360
      txtDataFlds(i + 1).Visible = True
      Load lblDataFlds(i + 1)
      lblDataFlds(i + 1).Top = (i + 2) * 360
      lblDataFlds(i + 1).Visible = True
      Load lblDataUnits(i + 1)
      lblDataUnits(i + 1).Top = (i + 2) * 360
      lblDataUnits(i + 1).Visible = True
    End If
    fldRec.MoveNext
  Next i
  editRec.Close
  fldRec.Close
End Sub

Private Sub UnloadControlArrays()
Attribute UnloadControlArrays.VB_Description = "Unloads arrays of labels and text boxes used during preceding data editing session."
' ##SUMMARY Unloads arrays of labels and text boxes used during preceding data editing _
          session.
  Dim i As Long
  For i = 1 To txtDataFlds.UBound()
    Unload txtDataFlds(i)
    Unload lblDataFlds(i)
    Unload lblDataUnits(i)
  Next i
End Sub

Private Sub lstDataCats_Change()
Attribute lstDataCats_Change.VB_Description = "Ensures only one category can be selected when editing data."
' ##SUMMARY Ensures only one category can be selected when editing data.
  Dim i As Long
  Static changing As Boolean
  Static sort As Boolean
  
  If Not changing Then
    changing = True
    If MyP.UserOpt = 1 Then
      If lstDataCats.RightCount > 1 Then
        If Z = lstDataCats.RightItemData(0) Then i = 0 Else i = 1
        lstDataCats.MoveLeft (i)
      End If
      If lstDataCats.RightCount = 1 Then
        Z = lstDataCats.RightItemData(0)
      End If
    End If
    changing = False
  End If
  Ok2DoMore
End Sub

Private Sub ByCatReport(RepPath As String, CatRec As Recordset)
Attribute ByCatReport.VB_Description = "Creates "
' ##SUMMARY Creates "Basic Tables by Category" report as Excel output file.
' ##PARAM RepPath I String full pathname where report will be saved.
' ##PARAM catRec I Recordset containing user-selected categories.
' ##REMARKS Report contains one category per worksheet.
  Dim xlRange As Excel.Range
  Dim rangeTemp As Excel.Range
  Dim fldRec As Recordset
  Dim dataRec As Recordset
  Dim stateTotalRec As Recordset
  Dim rowCnt As Long
  Dim fldRecCnt As Integer
  Dim areaCnt As Integer
  Dim catCnt As Integer
  Dim sql As String
  Dim areaName As String
  Dim str As String
  Dim i As Long
  Dim col As Long
  Dim lastCol As Long
  Dim numFlds As Long
  Dim cat As Long
  Dim fld As Long
  Dim opt As Long
  Dim totPop As Double
  Dim Val As Double
  
  On Error GoTo x
  
  InitReport RepPath
  catCnt = 0
  Set dataRec = MyP.stateDB.OpenRecordset("LastReport")
  If dataRec("QualFlg") < 7 Then
    opt = dataRec("QualFlg")
  ElseIf dataRec("QualFlg") = 7 Then
    opt = 1
  ElseIf dataRec("QualFlg") = 8 Then
    opt = 5
  End If
  dataRec.Close
  If AreaID = 2 Then AreaID = 0
  'Overwrite total population recordset if aggregating HUC-4s
  If AggregateHUCs Then
    sql = "SELECT Left([Location],4) AS Locn, Sum([HUCData].Value) AS [Value]" & _
        " From [HUCData]" & _
        " GROUP BY Left([Location],4), [HUCData].FieldID, [HUCData].Date" & _
        " HAVING [HUCData].FieldID=1 AND [HUCData].Date=" & MyP.Year1Opt & ";"
    Set TotalPopRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    str = "Locn"
  Else
    str = "Location"
  End If
  'Remove extra worksheets; reduce total to one
  While Worksheets.Count > 1
    With ActiveWorkbook
      Application.DisplayAlerts = False
      .Worksheets(.Worksheets.Count).Delete
      Application.DisplayAlerts = True
    End With
  Wend
  Set XLSheet = Worksheets(1)
  While Not CatRec.EOF  'one worksheet per category
    AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (CatRec.AbsolutePosition * 100) / _
        (CatRec.RecordCount) & ")"
    rowCnt = 1
    catCnt = catCnt + 1
    If catCnt > 1 Then
      'Add worksheet for each additional category then activate new sheet
      With ActiveWorkbook
        Worksheets.Add.Move after:=Worksheets(Worksheets.Count)
        Set XLSheet = Worksheets(catCnt)
      End With
    End If
    With XLSheet
      .Activate
      Columns.ColumnWidth = 5
      Columns(1).HorizontalAlignment = xlHAlignLeft
      cat = CatRec("ID")
      GetHeader cat
      'Paste body of header then write & format header title at top of each page if needed
      .Activate
      Set xlRange = Cells(rowCnt, 1)
      XLSheet.Paste xlRange
      For i = 1 To HeaderRows
        Cells(rowCnt, 1).value = ""
        rowCnt = rowCnt + 1
      Next i
      .Rows(rowCnt - 1).HorizontalAlignment = xlHAlignLeft
      rowCnt = rowCnt - HeaderRows
      If rdoAreaUnit(1) Then
        areaName = "hydrologic cataloging unit"
      ElseIf rdoAreaUnit(2) Then
        areaName = "water-resources region"
      ElseIf NationalDB Then
        areaName = "state"
      Else
        areaName = MyP.UnitArea
      End If
      With xlRange
        If cat = 2 Or cat = 4 And MyP.DataOpt <> 3 And MyP.DataOpt <> 4 Then
          Rows(rowCnt).Insert Shift:=xlDown
        End If
        .HorizontalAlignment = xlHAlignLeft
        Cells(rowCnt, 1).value = RepTitle & ", " & CatRec("Description") & _
            ", by " & areaName & " - " & MyP.State & ", " & MyP.Year1Opt
        Range(Cells(rowCnt, 1), Cells(rowCnt, 5)).Font.Bold = True
        If CatRec("ID") = 2 Or CatRec("ID") = 4 And MyP.DataOpt <> 3 And MyP.DataOpt <> 4 Then
          rowCnt = rowCnt + 1
          If CatRec("ID") = 2 Then
            Cells(rowCnt, 1).value = "[Population Served values have been rounded.]"
          ElseIf CatRec("ID") = 4 Then
            Cells(rowCnt, 1).value = "[Supplied Population values have been rounded.]"
          End If
        End If
      End With
      'Use SQL language to open recordset with all fields in current category.
      ' Get all attributes in table Field1 (ID, Name, Formula, CategoryID, and Description)
      ' that are used in the 2000 Data Dictionary. Different "Field" tables that comprise
      ' the data dictionary are joined.
      ' For example:
      '  SELECT [Field1].* FROM ([Category2]
      '  INNER JOIN [Field1] ON [Category2].ID = [Field1].CategoryID)
      '    INNER JOIN [2000Fields1] ON [Field1].ID = [2000Fields1].FieldID
      '  Where ([Category2].id = 2) And ([2000Fields1].Excluded <> 3)
      '  ORDER BY [Field1].CategoryID, [Field1].ID
      If cat < 50 Then
        sql = "SELECT [" & FieldTable & "].* " & _
            "FROM ([" & CatTable & "] INNER JOIN [" & FieldTable & "] ON [" & CatTable & "].ID = [" & FieldTable & "].CategoryID) " & _
            "INNER JOIN [" & MyP.YearFields & "] ON [" & FieldTable & "].ID = [" & MyP.YearFields & "].FieldID " & _
            "WHERE ([" & CatTable & "].ID=" & cat & ")" & OmitFlds & _
            " ORDER BY [" & FieldTable & "].CategoryID, [" & FieldTable & "].ID"
      Else  'need to total water use by GW, SW, or overall for each category
        If cat = 50 Then
          sql = "SELECT [" & FieldTable & "].* FROM [" & MyP.YearFields & _
              "] INNER JOIN [" & FieldTable & "] ON [" & MyP.YearFields & "].FieldID = [" & FieldTable & "].ID " & _
              "WHERE ((Right([" & FieldTable & "].Name,5)='WSWFr') OR (Right([" & FieldTable & "].Name,5)='WSWSa'))"
        ElseIf cat = 51 Then
          sql = "SELECT [" & FieldTable & "].* FROM [" & MyP.YearFields & _
              "] INNER JOIN [" & FieldTable & "] ON [" & MyP.YearFields & "].FieldID = [" & FieldTable & "].ID " & _
              "WHERE ((Right([" & FieldTable & "].Name,5)='WGWFr') OR (Right([" & FieldTable & "].Name,5)='WGWSa'))"
        ElseIf cat = 52 Then
          sql = "SELECT [" & FieldTable & "].* FROM [" & MyP.YearFields & _
              "] INNER JOIN [" & FieldTable & "] ON [" & MyP.YearFields & "].FieldID = [" & FieldTable & "].ID " & _
              "WHERE ((Right([" & FieldTable & "].Name,5)='WFrTo') OR (Right([" & FieldTable & "].Name,5)='WSaTo'))"
        End If
        If IRinTwo Then
          sql = sql & " AND Not ([" & FieldTable & "].CategoryID = 17"
        Else
          sql = sql & " AND Not ([" & FieldTable & "].CategoryID = 16"
        End If
        If Not rdoAreaUnit(3) Then
          'not aquifer; exclude Thermoelectric Power
          sql = sql & " OR [" & FieldTable & "].CategoryID = 10 OR [" & FieldTable & "].CategoryID = 11"
        End If
        sql = sql & " OR [" & FieldTable & "].CategoryID = 7 OR [" & FieldTable & "].CategoryID = 8" & _
                    " OR [" & FieldTable & "].CategoryID = 9 OR [" & FieldTable & "].CategoryID = 13" & _
                    " OR [" & FieldTable & "].CategoryID = 18 OR [" & FieldTable & "].CategoryID = 19)" & _
                    " ORDER BY [" & FieldTable & "].CategoryID, [" & FieldTable & "].ID;"
      End If
      Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
      fldRec.MoveLast
      fldRec.MoveFirst

      'Construct table beneath header for current category
      rowCnt = rowCnt + HeaderRows
      'Write unit area in heading
      With Cells(rowCnt - 1, 1)
        .HorizontalAlignment = xlHAlignLeft
        .Font.Bold = True
        .value = MyP.UnitArea
      End With
      i = UBound(LocnArray, 2) + 1
      For areaCnt = 0 To i
        Select Case NextPipeCharacter(AtcoLaunch1.ComputeRead)
          Case "P"
            While NextPipeCharacter(AtcoLaunch1.ComputeRead) <> "R"
              DoEvents
            Wend
          Case "C"
            ImportDone = True
              MyMsgBox.Show "The Basic Tables by Category report was cancelled.", _
                  "Report interrupted", "+-&OK"
            Err.Raise 999
        End Select
        AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (CatRec.AbsolutePosition + areaCnt _
              / i) * 100 / CatRec.RecordCount & ")"
        numFlds = fldRec.RecordCount - 1
        'Adjust the number of fields as necessary for certain categories
        If cat = 23 Or cat = 20 Then  'cat = Totals Overall or Hydroelectric
          numFlds = numFlds + 1
          If cat = 23 Then
            If opt <> 2 And opt <> 4 And opt <> 6 Then numFlds = numFlds + 1
          End If
        End If
        If MyP.UnitArea <> "Aquifer" Then
          If cat = 2 And (opt = 3 Or opt = 4) Then 'First column is State Total
            numFlds = numFlds + 1
          ElseIf cat = 4 And (opt = 2 Or opt = 4 Or opt = 6) Then '2 columns are State Totals
            numFlds = numFlds + 3
          End If
        End If
        
        If areaCnt < i Then  'Filling in one one unit area at a time
          'Fill in name of unit area
          If AreaID = 0 Then
            Cells(rowCnt, 1).value = "'" & LocnArray(AreaID, areaCnt)
          Else
            Cells(rowCnt, 1).value = LocnArray(AreaID, areaCnt)
          End If
          'Fill in body of table with data
          For fldRecCnt = 0 To numFlds
            If MyP.UnitArea <> "Aquifer" Then
              If cat = 2 And fldRecCnt = 0 And (opt = 3 Or opt = 4) Then
                'First column is State Total
                Cells(rowCnt, fldRecCnt + 2).value = "--"
                fldRecCnt = fldRecCnt + 1
              End If
              If cat = 4 Then
                If (fldRecCnt = 1 And (opt = 2 Or opt = 6)) Or _
                   (fldRecCnt = 0 And opt = 4) Then
                  'First or second column is State Total
                  Cells(rowCnt, fldRecCnt + 2).value = "--"
                  Cells(rowCnt, fldRecCnt + 3).value = "--"
                  Cells(rowCnt, fldRecCnt + 4).value = "--"
                  fldRecCnt = fldRecCnt + 3
                End If
              End If
            End If
            
            fld = fldRec("ID")
            Val = EvalArray(fld, areaCnt)
            If cat = 20 And fldRecCnt = 1 Then
              Cells(rowCnt, fldRecCnt + 2).value = Round((Cells(rowCnt, 2).value * 1.12), 2)
              fldRecCnt = fldRecCnt + 1
            ElseIf fld = 243 And numFlds > 8 Then 'insert per capita use for TO
              If totPop > -2 And Cells(rowCnt, fldRecCnt - 1).value <> "" Then
                If totPop > 0 Then
                  Cells(rowCnt, fldRecCnt + 2).value = _
                      Round((Cells(rowCnt, fldRecCnt - 1).value * 1000 / totPop), 2)
                Else
                  Cells(rowCnt, fldRecCnt).value = 0
                End If
              End If
              fldRecCnt = fldRecCnt + 1
            End If
            If Not NoRec Then
              If cat > 19 And cat < 50 Then
                If cat = 23 Then  'Totals Overall
                  If fldRecCnt = 0 Then  'write total population
                    If rdoAreaUnit(2) Or rdoAreaUnit(3) Then
                      sql = "Left(" & str & ", " & MyP.Length & ")='" & LocnArray(0, areaCnt) & "'"
                    Else
                      sql = "Left(" & str & ", " & MyP.Length & ")=" & LocnArray(0, areaCnt)
                    End If
                    TotalPopRec.FindFirst sql
                    If TotalPopRec.NoMatch Or IsNull(TotalPopRec("Value")) Then
                      totPop = -1
                    Else
                      totPop = TotalPopRec("Value")
                      Cells(rowCnt, fldRecCnt + 2).value = Round(totPop, 3)
                    End If
                    fldRecCnt = fldRecCnt + 1
                  End If
                End If
                Cells(rowCnt, fldRecCnt + 2).value = Round(Val, 2)
                If fld = 236 Then
                  If rdoAreaUnit(3) Then  'aquifer per capita for TO
                    fldRecCnt = fldRecCnt + 1
                    If totPop > -2 And Cells(rowCnt, fldRecCnt + 1).value <> "" Then
                      If totPop > 0 Then
                        Cells(rowCnt, fldRecCnt + 2).value = _
                          Round((Cells(rowCnt, fldRecCnt + 1).value * 1000 / totPop), 2)
                      Else
                        Cells(rowCnt, fldRecCnt + 2).value = 0
                      End If
                    End If
                  End If
                End If
              Else
                Cells(rowCnt, fldRecCnt + 2).value = Round(Val, 2)
              End If
              fldRec.MoveNext
            Else  'data record does not exist for this field
              If cat = 23 And fldRecCnt = 0 Then  'need to insert Total Population for area
                If rdoAreaUnit(2) Or rdoAreaUnit(3) Then
                  sql = "Left(" & str & ", " & MyP.Length & ")='" & LocnArray(0, areaCnt) & "'"
                Else
                  sql = "Left(" & str & ", " & MyP.Length & ")=" & LocnArray(0, areaCnt)
                End If
                TotalPopRec.FindFirst sql
                If TotalPopRec.NoMatch Or IsNull(TotalPopRec("Value")) Then
                  totPop = 0
                Else
                  totPop = TotalPopRec("Value")
                  Cells(rowCnt, fldRecCnt + 2).value = Round(totPop, 3)
                End If
                fldRecCnt = fldRecCnt + 1
              End If
              fldRec.MoveNext
              If (fld = 213) Or (rdoAreaUnit(3) And (fld = 234 Or fld = 236)) Then
                fldRecCnt = fldRecCnt + 1
              End If
            End If
          Next fldRecCnt
          fldRec.MoveFirst
          rowCnt = rowCnt + 1
        Else 'Total each field for all unit areas
          For fldRecCnt = 2 To numFlds + 2
            'make special summing/averaging considerations for certain fields
            If cat = 2 Then  'Public Supply
              If fldRecCnt = 2 And (opt = 3 Or opt = 4) Then
                'Fill in State total in first column for PS
                sql = "SELECT Value FROM [" & MyP.AreaTable & "] " & _
                      "WHERE FieldID=4 And Date=" & MyP.Year1Opt & _
                      " And Len(trim(location))=" & MyP.Length
                Set stateTotalRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
                If stateTotalRec.RecordCount > 0 Then Cells(rowCnt, 2).value = Round(stateTotalRec("Value"), 2)
                fldRecCnt = fldRecCnt + 1
              End If
              If fldRecCnt = 20 Then
                If opt < 3 Then 'PS-pop by unit area for County or HUC
                  If Cells(rowCnt, 4).value <> "" And Cells(rowCnt, fldRecCnt - 7).value <> "" Then
                    If Cells(rowCnt, 4) <> 0 Then
                      Cells(rowCnt, fldRecCnt).value = _
                          Round((Cells(rowCnt, fldRecCnt - 7).value * 1000 / Cells(rowCnt, 4).value), 2)
                    Else
                      Cells(rowCnt, fldRecCnt).value = 0
                    End If
                  End If
                  GoTo y
                End If
              ElseIf fldRecCnt = 6 And MyP.Length = 10 Then 'PS-pop by unit area for Aquifer
                If Cells(rowCnt, 2).value <> "" And Cells(rowCnt, 5).value <> "" Then
                  If Cells(rowCnt, 2) <> 0 Then
                    Cells(rowCnt, fldRecCnt).value = _
                        Round((Cells(rowCnt, 5).value * 1000 / Cells(rowCnt, 2).value), 2)
                  Else
                    Cells(rowCnt, fldRecCnt).value = 0
                  End If
                End If
                GoTo y
              ElseIf fldRecCnt = 18 Then
                If opt = 5 Or opt = 6 Then 'PS-pop by unit area, GW/SW combined
                  If Cells(rowCnt, 2).value <> "" And Cells(rowCnt, fldRecCnt - 7).value <> "" Then
                    If Cells(rowCnt, 2) <> 0 Then
                      Cells(rowCnt, fldRecCnt).value = _
                          Round((Cells(rowCnt, fldRecCnt - 7).value * 1000 / Cells(rowCnt, 2).value), 2)
                    Else
                      Cells(rowCnt, fldRecCnt).value = 0
                    End If
                  End If
                  GoTo y
                End If
              End If
            ElseIf cat = 4 Then  'Domestic
              If (fldRecCnt = 3 And (opt = 2 Or opt = 6)) Or _
                 (fldRecCnt = 2 And opt = 4) Then
                'Fill in State totals in columns 1-3 for DO
                sql = "SELECT FieldID, Value FROM [" & MyP.AreaTable & "] " & _
                      "WHERE (FieldID=40 or FieldID=43) And Date=" & MyP.Year1Opt & _
                      " And Len(trim(location))=" & MyP.Length & _
                      " ORDER BY FieldID;"
                Set stateTotalRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
                If stateTotalRec.RecordCount > 0 Then Cells(rowCnt, fldRecCnt).value = Round(stateTotalRec("Value"), 2)
                fldRecCnt = fldRecCnt + 1
                If stateTotalRec.RecordCount > 0 Then
                  stateTotalRec.MoveLast
                  Cells(rowCnt, fldRecCnt).value = Round(stateTotalRec("Value"), 2)
                End If
                fldRecCnt = fldRecCnt + 1
                Cells(rowCnt, fldRecCnt).value = _
                    Round(Cells(rowCnt, fldRecCnt - 2).value + Cells(rowCnt, fldRecCnt - 1).value, 2)
                fldRecCnt = fldRecCnt + 1
              End If
              If (fldRecCnt = 12 And opt = 0) Or _
                 (fldRecCnt = 6 And (opt = 1 Or opt = 5)) Then 'total per capita use for DO
                If Cells(rowCnt, 2).value <> "" And Cells(rowCnt, fldRecCnt - 1).value <> "" Then
                  If Cells(rowCnt, 2).value <> 0 Then
                    Cells(rowCnt, fldRecCnt).value = _
                        Round((Cells(rowCnt, fldRecCnt - 1).value * 1000 / Cells(rowCnt, 2).value), 2)
                  Else
                    Cells(rowCnt, fldRecCnt).value = 0
                  End If
                End If
                GoTo y
              ElseIf (fldRecCnt = 15 And opt = 0) Or _
                  (fldRecCnt = 9 And (opt = 1 Or opt = 5)) Or _
                  (fldRecCnt = 5 And (opt = 2 Or opt = 6)) Then 'Public-Supplied per capita for DO
                If Cells(rowCnt, fldRecCnt - 2).value <> "" And Cells(rowCnt, fldRecCnt - 1).value <> "" Then
                  If Cells(rowCnt, fldRecCnt - 2).value <> 0 Then
                    Cells(rowCnt, fldRecCnt).value = Round((Cells(rowCnt, fldRecCnt - 1).value _
                        * 1000 / Cells(rowCnt, fldRecCnt - 2).value), 2)
                  Else
                    Cells(rowCnt, fldRecCnt).value = 0
                  End If
                End If
                GoTo y
              End If
            ElseIf cat = 23 Then  'Totals overall
              If fldRecCnt = 12 Or (fldRecCnt = 6 And MyP.UnitArea = "Aquifer") Then
                If MyP.UnitArea = "Aquifer" Then fld = 5 Else fld = 9
                If Cells(rowCnt, fld).value <> "" And Cells(rowCnt, 2).value <> "" Then
                  If Cells(rowCnt, 2).value <> 0 Then
                    Cells(rowCnt, fldRecCnt).value = _
                        Round((Cells(rowCnt, fld).value * 1000 / Cells(rowCnt, 2).value), 2)
                  Else
                    Cells(rowCnt, fldRecCnt).value = 0
                  End If
                End If
                GoTo y
              End If
            End If
            'perform regular summing functions for rest of fields
            If Application.WorksheetFunction.CountIf _
                (Range(Cells(HeaderRows + 1, fldRecCnt), Cells(rowCnt - 1, fldRecCnt)), "") = 0 Then
              Cells(rowCnt, fldRecCnt).value = Round(Application.WorksheetFunction. _
                  Sum(Range(Cells(rowCnt - i, fldRecCnt), Cells(rowCnt, fldRecCnt))), 3)
            End If
y:
          Next fldRecCnt
          With Range(Cells(rowCnt, 1), Cells(rowCnt, fldRecCnt - 1)).Borders(xlTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
          End With
        End If
      Next areaCnt
      'Write in "Total:" only if totals have been summed
      For fldRecCnt = 2 To numFlds + 2
        If Len(Cells(rowCnt, fldRecCnt).value) > 0 Then
          Cells(rowCnt, 1).value = "Total:"
          Cells(rowCnt, 1).HorizontalAlignment = xlHAlignRight
          Exit For
        End If
      Next
      'Set number format default to 2 decimals
      Range(Cells(rowCnt - i, 2), _
            Cells(rowCnt, numFlds + 2)).NumberFormat = "#,###,##0.00"
      If cat = 23 Then Columns(2).NumberFormat = "#.000"
'      Set xlRange = .UsedRange
'      Set rangeTemp = xlRange.Find("opulation", , , , xlByColumns, xlNext, False)
'      If Not rangeTemp Is Nothing Then col = rangeTemp.Column
'      lastCol = -1
'      While Not rangeTemp Is Nothing
'        If col > lastCol Then
'          If MyP.Length <> 10 And cat = 2 And col = 2 And (opt < 3 Or opt = 7) Then 'PS has both GW and SW pops
'            Range(Cells(rowCnt - i, col), Cells(rowCnt, col + 2)).NumberFormat = "###,##0.000"
'          Else
'            Range(Cells(rowCnt - i, col), Cells(rowCnt, col)).NumberFormat = "###,##0.000"
'          End If
'          lastCol = col
'          Set rangeTemp = xlRange.FindNext(rangeTemp)
'          col = rangeTemp.Column
'        Else
'          Set rangeTemp = Nothing
'        End If
'      Wend
      'Set number format for facilities to 0 decimals
      Set xlRange = .UsedRange
      Set rangeTemp = xlRange.Find("umber of", , , , xlByColumns, xlNext, False)
      If Not rangeTemp Is Nothing Then col = rangeTemp.Column
      lastCol = -1
      While Not rangeTemp Is Nothing
        If col > lastCol Then
          Range(Cells(rowCnt - i, col), Cells(rowCnt, col)).NumberFormat = "###,##0"
          lastCol = col
          Set rangeTemp = xlRange.FindNext(rangeTemp)
          col = rangeTemp.Column
        Else
          Set rangeTemp = Nothing
        End If
      Wend
      'Set number format for per capita fields to 0 decimals
      Set rangeTemp = xlRange.Find("apita", , , , xlByColumns, xlNext, False)
      lastCol = -1
      If Not rangeTemp Is Nothing Then col = rangeTemp.Column
      While Not rangeTemp Is Nothing
        If col > lastCol Then
          Range(Cells(rowCnt - i, col), Cells(rowCnt, col)).NumberFormat = "###,##0"
          lastCol = col
          Set rangeTemp = xlRange.FindNext(rangeTemp)
          col = rangeTemp.Column
        Else
          Set rangeTemp = Nothing
        End If
      Wend
      fldRec.Close
      'Set column widths
      If rdoAreaUnit.item(0) Then
        If rdoID.item(1) Then .Columns(1).ColumnWidth = 4 Else .Columns(1).ColumnWidth = 12
      ElseIf rdoAreaUnit.item(1) Then
        If rdoID.item(1) Then .Columns(1).ColumnWidth = 8 Else .Columns(1).ColumnWidth = 16
      ElseIf rdoAreaUnit.item(2) Then
        If rdoID.item(1) Then .Columns(1).ColumnWidth = 5 Else .Columns(1).ColumnWidth = 16
      ElseIf rdoAreaUnit.item(3) Then
        If rdoID.item(1) Then .Columns(1).ColumnWidth = 11 Else .Columns(1).ColumnWidth = 30
      End If
      With Range(Columns(2), Columns(numFlds + 2))
        .AutoFit
      End With
      
      If rdoID(2) Then
        .Columns(2).Insert
        If cat = 2 Or cat = 4 And MyP.DataOpt <> 3 And MyP.DataOpt <> 4 Then
          HeaderRows = HeaderRows + 1
        End If
        With Range(Cells(HeaderRows, 1), Cells(HeaderRows, 2))
          .Merge
          .HorizontalAlignment = xlHAlignCenter
        End With
        For areaCnt = 1 To i
          Cells(areaCnt + HeaderRows, 2).value = LocnArray(1, areaCnt - 1)
        Next areaCnt
        With Range(Cells(areaCnt + HeaderRows, 1), Cells(areaCnt + HeaderRows, 2))
          .Merge
          .HorizontalAlignment = xlHAlignRight
        End With
      End If
      XLSheet.Name = CatRec("Name")
      CatRec.MoveNext
    End With
    XLBook.Save
    Clipboard.Clear
    Set xlRange = XLSheet.UsedRange
    xlRange.Font.size = 7
    xlRange.Font.Name = "Times New Roman"
  Wend
  CatRec.Close
  TotalPopRec.Close
  XLBook.Worksheets(1).Select
x:
  EndReport
End Sub

Private Sub ByAreaReport(RepPath As String, CatRec As Recordset)
Attribute ByAreaReport.VB_Description = "Creates "
' ##SUMMARY Creates "Basic Tables by Area" report as Excel output file.
' ##PARAM RepPath I String full pathname where report will be saved.
' ##PARAM catRec I Recordset containing user-selected categories.
' ##REMARKS Report is on single worksheet with columnar data and 2 categories per row.
  Dim xlRange As Excel.Range
  Dim fldRec As Recordset
  Dim rowCnt As Long
  Dim Val As Double
  Dim sql As String
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim areaCount As Long
  Dim fldRecCnt As Long
  Dim addCols As Long
  Dim maxAreas As Long
  Dim catCounter As Long
  Dim Length As Long
  Dim thisCat As Long
  Dim thisSheet As Long
  Dim areaName As String
  Dim skip As Boolean
  Dim fldsInCat() As Integer
  
  On Error GoTo x

  InitReport RepPath
  maxAreas = 9999
  thisSheet = 0
  ReDim fldsInCat(lstDataCats.RightCount, 30)
  
  For areaCount = 0 To UBound(LocnArray, 2)
    If areaCount = (maxAreas + 1) * thisSheet Then
      rowCnt = 1
      thisSheet = thisSheet + 1
      Set XLSheet = ActiveWorkbook.Worksheets(thisSheet)
      XLSheet.Activate
    End If
    With XLSheet
      If areaCount = (maxAreas + 1) * (thisSheet - 1) Then
        Columns(1).ColumnWidth = 35
        Columns(2).ColumnWidth = 6
        Columns(3).ColumnWidth = 5
        Columns(4).ColumnWidth = 35
        Columns(5).ColumnWidth = 6
        Columns(2).NumberFormat = "#,###,##0.00"
        Columns(5).NumberFormat = "#,###,##0.00"
      End If
      AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (areaCount * 100) / _
          (UBound(LocnArray, 2) + 1) & ")"
          
      skip = False
      If rdoAreaUnit(0) Then
        If NationalDB Then
          areaName = LocnArray(AreaID, areaCount)
        Else
          areaName = LocnArray(AreaID, areaCount) & " County"
        End If
      ElseIf rdoAreaUnit(1) Then
        areaName = "hydrologic cataloging unit " & LocnArray(AreaID, areaCount)
      ElseIf rdoAreaUnit(2) Then
        areaName = "water-resources region " & LocnArray(AreaID, areaCount)
      Else
        areaName = "aquifer " & LocnArray(AreaID, areaCount)
      End If
      'create and format header for this unit area
      Cells(rowCnt, 1).value = "Estimated use of water for " & _
          areaName & ", " & MyP.State & ", " & MyP.Year1Opt
      rowCnt = rowCnt + 1
      Cells(rowCnt, 1).value = "[Data units in million gallons per day " & _
          "(mgd) unless otherwise noted]"
      With Range(Cells(rowCnt, 1), Cells(rowCnt - 1, 1))
        .Font.Bold = True
        .HorizontalAlignment = xlHAlignLeft
      End With
      rowCnt = rowCnt + 1
      'get header and paste onto report, if necessary
      While Not CatRec.EOF
        Select Case NextPipeCharacter(AtcoLaunch1.ComputeRead)
          Case "P"
            While NextPipeCharacter(AtcoLaunch1.ComputeRead) <> "R"
              DoEvents
            Wend
          Case "C"
            ImportDone = True
              MyMsgBox.Show "The Basic Tables by Area report was cancelled.", _
                  "Report interrupted", "+-&OK"
            Err.Raise 999
        End Select
        thisCat = CatRec("ID")
        catCounter = CatRec.AbsolutePosition
        AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (areaCount + _
            (catCounter / CatRec.RecordCount)) * 100 / (UBound(LocnArray, 2) + 1) & ")"
        If skip = False Then
          rowCnt = rowCnt + 1
          addCols = 0
        Else
          addCols = 3
        End If
        If areaCount = 0 Then
          'Iterate once through each category to establish structure of field tables
          With Range(Cells(rowCnt, 1), Cells(rowCnt, 2 + addCols)).Borders(xlBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
          End With
          'write name of category into header
          Cells(rowCnt, 1 + addCols).value = CatRec("Description") & " Category"
          'Use SQL language to create recordset with all fields from current category
          ' that will be used in report. Different "Field" tables that comprise the data dictionary
          ' are joined.
          ' For example:
          '  SELECT [Field1].* FROM ([Category2]
          '  INNER JOIN [Field1] ON [Category2].ID = [Field1].CategoryID)
          '    INNER JOIN [2000Fields1] ON [Field1].ID = [2000Fields1].FieldID
          '  WHERE ([Category2].ID=6)
          '  ORDER BY [Field1].ID
          sql = "SELECT [" & FieldTable & "].* " & _
              "FROM ([" & CatTable & "] INNER JOIN [" & FieldTable & "] ON [" & CatTable & "].ID = [" & FieldTable & "].CategoryID) " & _
              "INNER JOIN [" & MyP.YearFields & "] ON [" & FieldTable & "].ID = [" & MyP.YearFields & "].FieldID " & _
              "WHERE ([" & CatTable & "].ID=" & thisCat & ")" & _
              " ORDER BY [" & FieldTable & "].ID"
          Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
          fldRec.MoveLast
          fldRec.MoveFirst
          fldsInCat(catCounter, 0) = fldRec.RecordCount
        End If
        'Fill in body of table with field descriptions and data
        k = 0
        For fldRecCnt = 0 To fldsInCat(catCounter, 0) - 1
          rowCnt = rowCnt + 1
          k = k + 1
          If areaCount = 0 Then
            Cells(rowCnt, 1 + addCols).value = fldRec("Description")
            fldsInCat(catCounter, k) = fldRec("ID")
          End If
          Val = EvalArray(fldsInCat(catCounter, k), areaCount)
          If Not NoRec Then
            Cells(rowCnt, 2 + addCols).value = Round(Val, 2)
          Else
            Cells(rowCnt, 2 + addCols).value = Null
          End If
          If areaCount = 0 Then
            If InStr(1, LCase(fldRec("Description")), "umber of") > 0 Then
              Cells(rowCnt, 2 + addCols).NumberFormat = "###,##0"
            ElseIf InStr(1, LCase(fldRec("Description")), "apita") > 0 Then
              Cells(rowCnt, 2 + addCols).NumberFormat = "###,##0"
'            ElseIf InStr(1, LCase(fldRec("Description")), "opulation") > 0 Then
'              Cells(rowCnt, 2 + addCols).NumberFormat = "###,##0.000"
            End If
            fldRec.MoveNext
          End If
        Next fldRecCnt
        If skip = False Then skip = True Else skip = False
        If skip = True Then 'just put in first column of data
          rowCnt = rowCnt - fldsInCat(catCounter, 0)
          Length = fldsInCat(catCounter, 0)
        ElseIf Length > fldsInCat(catCounter, 0) Then
          rowCnt = rowCnt + Length - fldsInCat(catCounter, 0) + 1
        Else
          rowCnt = rowCnt + 1
        End If
        If areaCount = 0 Then fldRec.Close
        CatRec.MoveNext
      Wend
      If skip Then rowCnt = rowCnt + fldRecCnt
      
      'Paste data labels for additional areas as necessary
      If areaCount = 0 Then
        Set xlRange = XLSheet.UsedRange
        i = xlRange.Rows.Count
        Set xlRange = Range(Cells(4, 1), Cells(i, 5))
        xlRange.Copy
        j = 0
        For i = 1 To UBound(LocnArray, 2)
          k = (rowCnt + 5) * (i - j) - (i - j - 1) * 4
          ActiveSheet.Paste ActiveSheet.Range(Cells(k, 1), Cells(k, 1))
          If k > 65000 Then  'start a new worksheet
            If thisSheet = 1 Then maxAreas = i  'number of areas that fit on worksheet
            ActiveSheet.Name = LocnArray(0, j) & "-" & LocnArray(0, i)
            thisSheet = thisSheet + 1
            ActiveWorkbook.Worksheets(thisSheet).Activate
            j = i + 1
          End If
        Next i
        thisSheet = 1
        XLSheet.Activate
      End If
      Clipboard.Clear
      rowCnt = rowCnt + 2
      CatRec.MoveFirst
      ActiveWorkbook.Save
    End With
  Next areaCount
  Set XLSheet = Nothing

  'Name and format used sheets
  j = Right(ActiveWorkbook.ActiveSheet.Name, 1)
  For i = 1 To j
    ActiveWorkbook.Worksheets(i).Activate
    Set xlRange = ActiveSheet.UsedRange
    xlRange.Replace ", in Mgal/d", ""
    xlRange.Font.size = 8
    xlRange.Font.Name = "Times New Roman"
    If i = 1 Then
      ActiveWorkbook.ActiveSheet.Name = "By " & MyP.UnitArea
    Else
      ActiveWorkbook.ActiveSheet.Name = "By " & MyP.UnitArea & " - " & i
    End If
  Next i
  'Get rid of extra sheets
  While j < ActiveWorkbook.Worksheets.Count
    ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count).Delete
  Wend

  CatRec.Close
x:
  AllFldRec.Close
  TotalPopRec.Close
  EndReport
End Sub

Private Sub EnteredReport(RepPath As String, CatRec As Recordset)
Attribute EnteredReport.VB_Description = "Creates "
' ##SUMMARY Creates "Entered Data Elements" report as Excel output file.
' ##PARAM RepPath I String full pathname where report will be saved.
' ##PARAM CatRec I Recordset containing user-selected categories.
' ##REMARKS Report in single worksheet in tabular form with 1 table per unit area.
' ##HISTORY PR 18442, 4/25/2007, prhummel Updated specialized code for output of _
            DO State totals of GW and SW Withdrawals.  Removed hard-coded row numbers _
            and replaced them with row numbers found by call to new function _
            FindRow, which returns the row number containing the correct label.
  Dim xlRange As Excel.Range
  Dim fldRec As Recordset
  Dim rowCnt As Long
  Dim totPop As Variant
  Dim ReqSt As Long
  Dim i As Long
  Dim j As Long
  Dim lRow As Long
  Dim areaCount As Long
  Dim fldRecCnt As Long
  Dim headopt As Long
  Dim lastColumn As Long
  Dim lastRow As Long
  Dim rowsInTable As Long
  Dim colsInTable As Long
  Dim sql As String
  Dim sqlAddOn As String
  Dim areaName As String
  Dim areasType As String
  Dim header As String
  Dim Border As String
  Dim Val As Double
  Dim totalVals() As Double
  Dim needFootnote As Boolean
  Dim writeTable As Boolean
  
  On Error GoTo x

  If NationalDB Then
    ReqSt = 0
  Else
    'Determine which group of special fields are required for this state
    sql = "SELECT Required FROM state WHERE state_cd='" & MyP.stateCode & "'"
    Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    ReqSt = fldRec("Required")
    fldRec.Close
  End If
  
  'Set area type label
  If rdoAreaUnit(0) Then
    areasType = "Counties"
  ElseIf rdoAreaUnit(1) Then
    areasType = "HUC-8's"
  ElseIf rdoAreaUnit(2) Then
    areasType = "HUC-4's"
  ElseIf rdoAreaUnit(3) Then
    areasType = "Aquifers"
  End If

  InitReport RepPath
  rowCnt = 1
  CatRec.MoveFirst
  'Create recordset of only original data fields
  If Left(MyP.UnitArea, 1) <> "H" Then sqlAddOn = " And [" & CatTable & "].ID<>22"
  If IRinTwo Or (NationalDB And MyP.DataOpt > 0) Then 'do not include Irrigation category 17
    sqlAddOn = sqlAddOn & " And Not([" & FieldTable & "].ID>194 And [" & FieldTable & "].ID<213)"
  End If
  If (Not IRinTwo) Or NationalDB Then 'do not include Irrigation categories 1
    sqlAddOn = sqlAddOn & " And Not([" & FieldTable & "].ID>281 And [" & FieldTable & "].ID<302)"
  End If
  'Use SQL language to create recordset with all user-entered fields from Data Dictionary
  ' that will be used in report. Different "Field" tables that comprise the data dictionary
  ' are joined. Varying irrigation options are accounted for via "sqlAddOn" variable used in
  ' WHERE statement.
  ' For example:
  '  SELECT [Field0].ID, [Field0].CategoryID FROM ([Category1]
  '    INNER JOIN [Field0] ON [Category1].ID = [Field0].CategoryID)
  '    INNER JOIN [1995Fields1] ON [Field0].ID = [1995Fields1].FieldID
  '  WHERE [Field0].Formula='' AND [Field0].ID<>1 And [Category1].ID<>22
  '    And Not([Field0].ID>281 And [Field0].ID<302)
  '  ORDER BY CategoryID, [Field0].ID;
  sql = "SELECT [" & FieldTable & "].ID, [" & FieldTable & "].CategoryID FROM " _
      & "([" & CatTable & "] INNER JOIN [" & FieldTable & "] ON [" & CatTable & "].ID = [" & FieldTable & "].CategoryID)" _
      & " INNER JOIN [" & MyP.YearFields & "] ON [" & FieldTable & "].ID = [" & MyP.YearFields & "].FieldID" _
      & " WHERE ([" & FieldTable & "].Formula=''"
  If NationalDB Then
    sql = sql & " Or ([" & CatTable & "].ID=16" _
        & " AND [" & FieldTable & "].ID<>304 AND [" & FieldTable & "].ID<>310)"
  End If
  sql = sql & ") AND [" & FieldTable & "].ID<>1" & sqlAddOn _
        & " ORDER BY CategoryID, [" & FieldTable & "].ID;"
  Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  fldRec.MoveLast
  fldRec.MoveFirst
  'Dimension the totals array. 2nd dimension: 1=value,2=row,3=col,4=MissingMandatoryVal
  ReDim totalVals(1 To fldRec.RecordCount + 1, 1 To 4)
  With XLSheet
    .Columns.ColumnWidth = 6
    If Left(TableName, 1) = "H" Then headopt = 3 Else headopt = 4
    If Left(MyP.YearFields, 1) = "2" And ((Not IRinTwo) Or NationalDB) Then headopt = headopt - 2
    'Create a table for each area unit
    Set xlRange = Range(Columns(1), Columns(25))
    xlRange.Font.size = 8
    xlRange.Font.Name = "Times New Roman"
    For areaCount = 0 To UBound(LocnArray, 2)
      Select Case NextPipeCharacter(AtcoLaunch1.ComputeRead)
        Case "P"
          While NextPipeCharacter(AtcoLaunch1.ComputeRead) <> "R"
            DoEvents
          Wend
        Case "C"
          ImportDone = True
            MyMsgBox.Show "The Entered Data Elements report was cancelled.", _
                "Report interrupted", "+-&OK"
          Err.Raise 999
      End Select
      
      areaName = LocnArray(0, areaCount)
      'Check to see if this unit area was selected by the user
      writeTable = False
      If NationalDB Then
        writeTable = True
      Else
        For i = 0 To lstArea.RightCount - 1
          If areaName = Left(lstArea.RightItem(i), MyP.Length) Then
            writeTable = True
            Exit For
          End If
        Next i
      End If
      
      needFootnote = False
      'Create header for this area unit
      If rdoAreaUnit(3) Then
        TotalPopRec.FindFirst "Left(Location, " & MyP.Length & ")='" & areaName & "'"
      Else
        TotalPopRec.FindFirst "Left(Location, " & MyP.Length & ")=" & areaName
      End If
      If IsNull(TotalPopRec("Value")) Then
        totPop = "(no value in database)"
        NoRec = True
      Else
        totPop = TotalPopRec("Value")
        totalVals(1, 1) = totalVals(1, 1) + totPop
        NoRec = False
      End If
      If NationalDB Then
        If TableName = "HUC8" Then
          areaName = "HUC"
        ElseIf TableName = "Aquifer" Then
          areaName = "Aquifer"
        Else
          areaName = "State"
        End If
      Else
        If rdoAreaUnit(2) Then
          areaName = "Water-resources Region"
        Else
          areaName = MyP.UnitArea
        End If
      End If
            
      If NationalDB Then
        header = RepTitle & " report for " & LocnArray(AreaID, areaCount) & _
            ", " & MyP.Year1Opt
      Else
        header = RepTitle & " report for " & LocnArray(AreaID, areaCount) & _
            " - " & MyP.State & ", " & MyP.Year1Opt
      End If
      Cells(rowCnt, 1) = header
      Cells(rowCnt, 1).Font.Bold = True
      Cells(rowCnt + 2, 1) = areaName & " Code: " & LocnArray(0, areaCount)
      Cells(rowCnt + 2, 2) = areaName & " Name:"
      Cells(rowCnt + 2, 3) = LocnArray(1, areaCount)
      Cells(rowCnt + 2, 5) = "Total Population:"
      Cells(rowCnt + 2, 5).HorizontalAlignment = xlRight
      If IsNumeric(totPop) Then
        Cells(rowCnt + 2, 6) = Format(totPop, "#.000") & " thousand"
      Else
        Cells(rowCnt + 2, 6) = totPop
      End If
      Border = MyP.Required(1, ReqSt)
      'Make text red & bold if total population field is null
      If NoRec Then
        If Border = "red" And Not NationalDB Then
          Cells(rowCnt + 2, 6).Font.Color = RGB(500, 0, 0)
          Cells(rowCnt + 2, 6).Font.Bold = True
          If totalVals(1, 4) = 1 Then totalVals(1, 4) = 3
        End If
        If totalVals(1, 4) = 1 Then totalVals(1, 4) = 2
      Else 'Track whether null values are part of total
        If areaCount = 0 Then 'first datum of those being summed
          totalVals(1, 4) = 1
        ElseIf totalVals(1, 4) = 0 Then 'previous data missing
          If Border = "red" And Not NationalDB Then
            totalVals(1, 4) = 3
          Else
            totalVals(1, 4) = 2
          End If
        End If
      End If
      Cells(rowCnt + 3, 1) = "Data units in million gallons per day (mgd) unless otherwise noted"
      If Not NationalDB Then
        Cells(rowCnt + 4, 1) = "Mandatory data elements containing null values are shown as bold and red: NR"
        Cells(rowCnt + 5, 1) = "NR - indicates that the value for this data element is null"
      Else
        Cells(rowCnt + 4, 1) = "NR - indicates that the value for this data element is null"
      End If
      Cells(rowCnt + 4, 1).Select
      rowCnt = rowCnt + 6
      'Get header with table fields and paste onto report
      GetHeader headopt
      .Activate
      .Paste XLSheet.Range(Cells(rowCnt, 1), Cells(rowCnt, 1))
      
      Cells(rowCnt, 1).value = ""
      rowCnt = rowCnt + 2
      Cells(rowCnt, 2).Select
      'Size the table
      If areaCount = 0 Then
        'Count # of rows in table
        rowsInTable = rowCnt
        While Cells(rowsInTable, 1) <> ""
          rowsInTable = rowsInTable + 1
        Wend
        rowsInTable = rowsInTable - rowCnt
        'Count # of columns in table
        colsInTable = 2
        While Cells(rowCnt - 1, colsInTable) <> ""
          colsInTable = colsInTable + 1
        Wend
        colsInTable = colsInTable - 1
      End If
      Range(Cells(rowCnt - 6, 1), Cells(rowCnt + rowsInTable + 2, 20)).Select
      With Selection
        .Font.size = 8
        .Font.Name = "Times New Roman"
      End With
      Cells(rowCnt - 2, 1).Select
      With ActiveCell.Characters(start:=75, Length:=2).Font
        .ColorIndex = 3
        .Bold = True
      End With
      
      With Range(Cells(rowCnt, 2), Cells(rowCnt + rowsInTable - 1, colsInTable))
        .Interior.ColorIndex = 15
        .NumberFormat = "0.00"
      End With
      If Left(MyP.AreaTable, 1) <> "A" Then
        If Left(MyP.AreaTable, 1) = "H" Then i = 3 Else i = 1
        Rows(rowCnt + rowsInTable - i).NumberFormat = "0"
      End If
      For i = rowCnt To rowCnt + 2
        If InStr(1, Cells(i, 1), "op ") > 0 Then
          Rows(i).NumberFormat = "0.000"
        End If
      Next i
      'Fill in body of table with data
      With Range(Cells(rowCnt - 1, 2), Cells(rowCnt + HeaderRows - 3, CatRec.RecordCount + 3))
        .HorizontalAlignment = xlHAlignLeft
        Set xlRange = .Find(-1, , , , xlByColumns, xlNext, False)
        If Not xlRange Is Nothing Then
          'Loop through all data fields
          Do
            j = fldRec.AbsolutePosition + 2
            lastColumn = xlRange.Column
            lastRow = xlRange.Row
            'Rearrange layout of instream/out-of-stream data to fit table
            If (lastColumn = 12 And headopt < 3) Or _
               (lastColumn = 13 And headopt > 2) Then
              If fldRec("ID") = 214 Or fldRec("ID") = 215 Then
                Cells(lastRow, lastColumn) = ""
                lastColumn = lastColumn + 1
              ElseIf (fldRec("ID") = 218) Or (fldRec("ID") = 221) Then
                Cells(lastRow, lastColumn) = ""
                lastColumn = lastColumn + 1
                lastRow = lastRow - 1
              End If
            End If
            'Rearrange layout of wastewater data to fit table
            If fldRec("ID") = 229 Then
              Cells(lastRow, lastColumn) = ""
              lastRow = lastRow - 2
            End If
            'Match field name with data record and fill cells
            Val = Round(EvalArray(fldRec("ID"), areaCount), 3)
            If areaCount = 0 Then
              totalVals(j, 2) = lastRow
              totalVals(j, 3) = lastColumn
            End If
            Border = MyP.Required(fldRec("ID"), ReqSt)
            Cells(lastRow, lastColumn).Interior.ColorIndex = 0
            If NoRec Then
              Cells(lastRow, lastColumn).value = "NR"
              'Mark running total as containing null value(s)
              If Border = "red" And Not NationalDB Then 'required, no nulls
                Cells(lastRow, lastColumn).Font.Color = RGB(500, 0, 0)
                Cells(lastRow, lastColumn).Font.Bold = True
                If totalVals(j, 4) = 1 Then totalVals(j, 4) = 3
              Else
                Cells(lastRow, lastColumn).BorderAround , Weight:=xlHairline
              End If
              needFootnote = True
              If totalVals(j, 4) = 1 Then totalVals(j, 4) = 2
            Else 'have datum
              Cells(lastRow, lastColumn).value = Val
              totalVals(j, 1) = totalVals(j, 1) + Val
              'Check running total for previous null value(s)
              If areaCount = 0 Then 'first datum of those being summed
                totalVals(j, 4) = 1
              ElseIf totalVals(j, 4) = 0 Then 'first x consecutive data missing
                If Border = "red" And Not NationalDB Then
                  totalVals(j, 4) = 3
                Else
                  totalVals(j, 4) = 2
                End If
              End If
            End If
            fldRec.MoveNext
            Set xlRange = .FindNext(xlRange)
            AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (areaCount + lastColumn / _
                (CatRec.RecordCount + 1)) * 100 / (UBound(LocnArray, 2) + 1) & ")"
          Loop While Not fldRec.EOF
        End If
      End With
      If NationalDB And MyP.DataDict = 2000 And Left(TableName, 1) = "H" Then
        'can not calculate HUC totals for PS-PopServed or DO-Withdrawals
        Rows(rowCnt + 1).Delete
        Range("D" & rowCnt + 1 & ":D" & rowCnt + 2).Select
        Selection.Clear
        Selection.Interior.ColorIndex = 15
        rowCnt = rowCnt - 1
      End If
      rowCnt = rowCnt + HeaderRows - 1
      If Not writeTable Then
        rowCnt = rowCnt - rowsInTable - 10 '9
        Range(Cells(rowCnt, 1), Cells(rowCnt + rowsInTable + 9, colsInTable)).Clear
        Range(Cells(rowCnt, 1), Cells(rowCnt + rowsInTable + 9, colsInTable)).Font.size = 8
        Range(Cells(rowCnt, 1), Cells(rowCnt + rowsInTable + 9, colsInTable)).Font.Name = "Times New Roman"
      End If
      Rows(rowCnt - 1).Select
      If ActiveCell.Row > 1 And ActiveWindow.SelectedSheets.HPageBreaks.Count < 1000 Then
        ActiveWindow.SelectedSheets.HPageBreaks.Add Before:=ActiveCell
      End If
      fldRec.MoveFirst
    Next areaCount
    If NationalDB Then
      .Name = "States"
    Else
      .Name = areasType
    End If
      .Columns(1).ColumnWidth = 18
      .Columns(2).ColumnWidth = 7
      .Columns(3).ColumnWidth = 6
      Range(Columns(4), Columns(lastColumn)).AutoFit
      .Columns(5).ColumnWidth = 6
      .Columns(6).ColumnWidth = 10
      If IRinTwo Then
        .Columns(15).ColumnWidth = 8
      Else
        .Columns(14).ColumnWidth = 8
      End If
    CatRec.Close
    Cells(1, 1).Select
    With .PageSetup
      .LeftMargin = Application.InchesToPoints(0.2)
      .RightMargin = Application.InchesToPoints(0.2)
      .Orientation = xlLandscape
      .FitToPagesWide = 1
    End With
  End With
  
  If Left(lstArea.RightItem(0), 3) = "000" Or NationalDB Then
    'Now create a sheet with the state/national totals
    totPop = 999999
    rowCnt = 1
    Set XLSheet = ActiveWorkbook.Worksheets(2)
    XLSheet.Activate
    With XLSheet
      If NationalDB Then
        header = RepTitle & " report for Nation (" & MyP.UnitArea & " totals) - United States, " & MyP.Year1Opt
        .Name = "Nation"
      Else
        header = RepTitle & " report for " & areasType & " - " & MyP.State & ", " & MyP.Year1Opt
        .Name = "State"
      End If
      Cells(1, 1) = header
      Cells(1, 1).Font.Bold = True
      Cells(3, 1) = "Total Population:"
      'Make text red & bold if total population fields null
      If (totalVals(1, 4) = 0 Or totalVals(1, 4) = 3) And Not NationalDB Then
        Cells(3, 2).Font.Color = RGB(500, 0, 0)
        Cells(3, 2).Font.Bold = True
        If totalVals(1, 4) = 0 Then
          Cells(3, 2) = "(no values in database)"
        Else
          Cells(3, 2) = Format(totalVals(1, 1), "0.000") & " thousand"
        End If
      Else
        Cells(3, 2) = Format(totalVals(1, 1), "0.000") & " thousand"
      End If
      Cells(4, 1) = "Data units in million gallons per day (mgd) unless otherwise noted"
      If Not NationalDB Then
        Cells(5, 1) = "Totals of mandatory data elements with one or more contributing " & MyP.UnitArea & " having a null value are shown as bold and red"
        Cells(6, 1) = "NR - indicates that all values for this data element are null"
      Else
        Cells(5, 1) = "NR - indicates that all values for this data element are null"
      End If
      'Get header with table fields and paste onto report
      GetHeader headopt
      .Activate
      .Paste XLSheet.Range(Cells(7, 1), Cells(7, 1))
      Cells(7, 1).value = ""
      'Clear Table of markers and shade body gray
      Range(Cells(9, 2), Cells(9 + rowsInTable, colsInTable)).Clear
      With Range(Cells(9, 2), Cells(9 + rowsInTable - 1, colsInTable))
        .Interior.ColorIndex = 15
        .NumberFormat = "0.00"
      End With
      Set xlRange = .UsedRange
      xlRange.Font.size = 8
      xlRange.Font.Name = "Times New Roman"
      Cells(5, 1).Select
      If NationalDB And MyP.UnitArea = "State" Then i = 104 Else i = 105
      With ActiveCell.Characters(start:=i, Length:=13).Font
        .ColorIndex = 3
        .Bold = True
      End With
      'Fill in body of table with data
      For i = LBound(totalVals) + 1 To UBound(totalVals)
        Cells(totalVals(i, 2), totalVals(i, 3)) = totalVals(i, 1)
        Cells(totalVals(i, 2), totalVals(i, 3)).Interior.ColorIndex = 0
        'Check if field is mandatory and total has missing values
        If (totalVals(i, 4) = 3 Or (totalVals(i, 4) = 0) _
            And MyP.Required(fldRec("ID"), ReqSt) = "red") _
            And Not NationalDB Then
          Cells(totalVals(i, 2), totalVals(i, 3)).Font.Color = RGB(500, 0, 0)
          Cells(totalVals(i, 2), totalVals(i, 3)).Font.Bold = True
        End If
        If totalVals(i, 4) = 0 Then Cells(totalVals(i, 2), totalVals(i, 3)) = "NR"
        fldRec.MoveNext
      Next i
      If MyP.DataOpt = 2 Or MyP.DataOpt = 4 Or MyP.DataOpt = 6 Then
        'Write state total values for DO - withdrawals for GW and SW
        sql = "SELECT * FROM [" & MyP.AreaTable & "]" & _
              " WHERE (FieldID=40 or FieldID=43) AND Date=" & MyP.Year1Opt & _
              " ORDER BY Location, FieldID;"
        Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
        'find "GW Withdrawal" row, prh 4/2007
        lRow = FindRow("GW Withdrawals, fresh", rowsInTable)
        If lRow > 0 Then
          If fldRec.RecordCount = 0 And NationalDB Then  'total DO withdrawals not kept for HUC's with 2000 DD
            'Can't sum national DO withdrawal totals
            Range("D" & lRow & ":D" & lRow).Select
            Selection.Clear
            Selection.Interior.ColorIndex = 15
            'repeat for "SW withdrawal", prh 4/2007
            lRow = FindRow("SW Withdrawals, fresh", rowsInTable)
            If lRow > 0 Then
              Range("D" & lRow & ":D" & lRow).Select
              Selection.Clear
              Selection.Interior.ColorIndex = 15
            End If
          Else
            Range("D" & lRow & ":D" & lRow).Select
            Selection.Interior.ColorIndex = xlNone
            Val = fldRec("Value")
            If IsNull(fldRec("Value")) Then
              Cells(lRow, 4) = "NR"
              If Not NationalDB Then
                Cells(lRow, 4).Font.Color = RGB(500, 0, 0)
                Cells(lRow, 4).Font.Bold = True
              End If
            Else
              Cells(lRow, 4) = Val
            End If
            'repeat for "SW withdrawal", prh 4/2007
            lRow = FindRow("SW Withdrawals, fresh", rowsInTable)
            If lRow > 0 Then
              Range("D" & lRow & ":D" & lRow).Select
              Selection.Interior.ColorIndex = xlNone
              fldRec.MoveNext
              Val = fldRec("Value")
              If IsNull(fldRec("Value")) Then
                Cells(lRow, 4) = "NR"
                If Not NationalDB Then
                  Cells(lRow, 4).Font.Color = RGB(500, 0, 0)
                  Cells(lRow, 4).Font.Bold = True
                End If
              Else
                Cells(lRow, 4) = Val
              End If
            End If
          End If
        Else 'couldn't find proper label, prh 4/2007
          Err.Raise 998, , "Unable to find GW/SW Withdrawal label for output of DO State total"
        End If
      End If
      If MyP.DataOpt = 3 Or MyP.DataOpt = 4 Then
        sql = "SELECT * FROM [" & MyP.AreaTable & "]" & _
              " WHERE FieldID=4 AND Date=" & MyP.Year1Opt & _
              " ORDER BY Location, FieldID;"
        Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
        If fldRec.RecordCount = 0 And NationalDB Then 'PS-TotPopServed not kept for HUC's with 2000 DD
          'Can't sum national Total Population Served totals
          Rows(10).Delete
          rowsInTable = rowsInTable - 1
        Else
          'Insert state total values for PS - population served for GW and SW
          Rows("10:10").Select
          Selection.Insert Shift:=xlDown
          rowsInTable = rowsInTable + 1
          Range(Cells(10, 2), Cells(10, lastColumn)).Select
          Selection.Interior.ColorIndex = 15
          Range(Cells(10, 2), Cells(10, 2)).Select
          Selection.Interior.ColorIndex = xlNone
          Val = fldRec("Value")
          Cells(10, 1) = "Population Served (thousands)"
          If IsNull(fldRec("Value")) Then
            Cells(10, 2) = "NR"
            If Not NationalDB Then
              Cells(10, 2).Font.Color = RGB(500, 0, 0)
              Cells(10, 2).Font.Bold = True
            End If
          Else
            Cells(10, 2) = Val
          End If
        End If
      End If
      If Left(MyP.AreaTable, 1) <> "A" Then
        If Left(MyP.AreaTable, 1) = "H" Then i = 6 Else i = 8
        Rows(i + rowsInTable).NumberFormat = "0"
      End If
      For i = 10 To 11
        If InStr(1, Cells(i, 1), "op ") > 0 Then
          Rows(i).NumberFormat = "0.000"
        End If
      Next i
      .Columns(1).ColumnWidth = 18
      .Columns(2).ColumnWidth = 7
      .Columns(3).ColumnWidth = 6
      Range(Columns(4), Columns(lastColumn)).AutoFit
      .Columns(5).ColumnWidth = 6
      .Columns(6).ColumnWidth = 10
      If IRinTwo Then
        .Columns(15).ColumnWidth = 8
      Else
        .Columns(14).ColumnWidth = 8
      End If
      With .PageSetup
        .LeftMargin = Application.InchesToPoints(0.2)
        .RightMargin = Application.InchesToPoints(0.2)
        .Orientation = xlLandscape
        .FitToPagesWide = 1
      End With
      Cells(1, 1).Select
    End With
  End If
x:
  If Err.Number > 0 And Err.Number <> 999 Then
    MsgBox "The 'Entered Data Elements' report experienced an error:" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Error"
  End If
  Clipboard.Clear
  AllFldRec.Close
  TotalPopRec.Close
  EndReport
End Sub

Private Function FindRow(RowLabel As String, maxRows As Long)
' ##SUMMARY Finds row containing a specific label (RowLabel) on output spreadsheet
' ##PARAM RowLabel I Label in column 1 to look for
' ##PARAM maxRows I Max number of rows on output spreadsheet
' ##RETURNS Row number containing input label (RowLabel)
  Dim lRow As Long
  Dim i As Long
        
  lRow = 0
  i = 0
  With XLSheet
    While lRow = 0 And i < maxRows
      i = i + 1
      If Cells(i, 1) = RowLabel Then
        lRow = i
      End If
    Wend
  End With
  FindRow = lRow

End Function

Private Sub CalcReport(RepPath As String, CatRec As Recordset)
Attribute CalcReport.VB_Description = "Creates "
' ##SUMMARY Creates "Calculated Tables" report as Excel output file.
' ##PARAM RepPath I String full pathname where report will be saved.
' ##PARAM catRec I Recordset containing user-selected categories.
' ##REMARKS Report is in single table in worksheet.
  Dim xlRange As Excel.Range
  Dim fldRec As Recordset
  Dim rowCnt As Long
  Dim denom As Double
  Dim temp2 As Double
  Dim sumDenom(5) As Double
  Dim sumNum(5) As Double
  Dim i As Long
  Dim j As Long
  Dim hyOff As Long
  Dim irTot As Long
  Dim irWTotl As Long
  Dim irCLoss As Long
  Dim sql As String
  Dim sqlAddOn As String
  Dim str As String
  Dim areaName As String
  
  On Error GoTo x

  InitReport RepPath
  rowCnt = 1
  If AreaID = 2 Then AreaID = 0
  
  With XLSheet
    .Columns.ColumnWidth = 7
    .Columns(1).HorizontalAlignment = xlHAlignLeft
    'Create calculations table
    For i = 0 To 4
      sumDenom(i) = 0
    Next i
    'Create header
    If MyP.Length = 4 Or MyP.Length = 8 Then
      GetHeader 1
    Else
      GetHeader 2
    End If
    .Activate
    Set xlRange = XLSheet.Range(Cells(rowCnt, 1), Cells(rowCnt, 1))
    XLSheet.Paste xlRange
    xlRange.HorizontalAlignment = xlHAlignLeft
    If NationalDB Then
      str = "State"
    Else
      If rdoAreaUnit(1) Then
        str = "hydrologic cataloging unit"
      ElseIf rdoAreaUnit(2) Then
        str = "water-resources region"
      Else
        str = MyP.UnitArea
      End If
    End If
    Cells(rowCnt, 1).value = "Calculated Tables, by " & str & _
        " - " & MyP.State & ", " & MyP.Year1Opt
    Cells(rowCnt, 1).Font.Bold = True
    rowCnt = rowCnt + HeaderRows - 1
    Cells(rowCnt, 1).value = MyP.UnitArea
    rowCnt = rowCnt + 1
    'Create recordset with all relevant data fields
    If Left(MyP.UnitArea, 1) = "H" Then _
        sqlAddOn = " Or CategoryID = 22" Else sqlAddOn = ""
    If Left(MyP.YearFields, 1) = "2" Then
      hyOff = 214
      irTot = 310
      irCLoss = 306
      irWTotl = 304
    Else
      hyOff = 216
      irTot = 211
      irCLoss = 207
      irWTotl = 201
    End If
    'Use SQL language to create recordset with all fields that will be used
    ' in calculations written to report.
    ' For example:
    '  SELECT * FROM [Field1]
    '  Where (id = 84 Or id = 88 Or id = 219 Or id = 213 Or id = 214 Or id = 306 Or id = 310 Or id = 304)
    '  ORDER BY CategoryID, ID
    sql = "SELECT * FROM [" & FieldTable & "] " & _
        "WHERE (ID = 84 Or ID = 88 " & _
        "Or ID = 219 Or ID = 213 " & _
        "Or ID = " & hyOff & " Or ID = " & irCLoss & _
        " Or ID = " & irTot & " Or ID = " & irWTotl & _
        sqlAddOn & ") ORDER BY CategoryID, ID"
    Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    'Fill in body of table
    For i = 1 To UBound(LocnArray, 2) + 1
      AtcoLaunch1.SendMonitorMessage "(PROGRESS " & i * 100 / (UBound(LocnArray, 2) + 1) & ")"
      'Fill in the names of the unit areas
      If AreaID = 0 Then
        areaName = "'" & LocnArray(AreaID, i - 1)
      Else
        areaName = LocnArray(AreaID, i - 1)
      End If
      Cells(rowCnt, 1).value = areaName
      'Fill the body of the table with data
      temp2 = EvalArray(fldRec("ID"), i - 1)
      fldRec.MoveNext
      denom = EvalArray(fldRec("ID"), i - 1)
      If denom = 0 Then
        Cells(rowCnt, 2).value = "N/A"
      Else
        Cells(rowCnt, 2).value = Round(temp2 / denom, 2)
      End If
      sumNum(0) = sumNum(0) + temp2
      sumDenom(0) = sumDenom(0) + denom
      fldRec.MoveNext
      denom = EvalArray(fldRec("ID"), i - 1)
      fldRec.MoveNext
      temp2 = EvalArray(fldRec("ID"), i - 1)
      If denom = 0 Then
        Cells(rowCnt, 4).value = 0
      Else
        Cells(rowCnt, 4).value = Round(temp2 * 100 / denom, 2)
      End If
      sumNum(2) = sumNum(2) + temp2 * 100
      sumDenom(2) = sumDenom(2) + denom
      fldRec.MoveNext
      temp2 = denom
      denom = EvalArray(fldRec("ID"), i - 1)
      If denom = 0 Then
        Cells(rowCnt, 5).value = 0
      Else
        Cells(rowCnt, 5).value = Round(temp2 * 13.452 / denom, 2)
      End If
      sumNum(3) = sumNum(3) + temp2 * 13.452
      sumDenom(3) = sumDenom(3) + denom
      fldRec.MoveNext
      temp2 = EvalArray(fldRec("ID"), i - 1)
      fldRec.MoveNext
      temp2 = temp2 + EvalArray(fldRec("ID"), i - 1)
      fldRec.MoveNext
      denom = EvalArray(fldRec("ID"), i - 1)
      If denom = 0 Then
        Cells(rowCnt, 3).value = "N/A"
      Else
        Cells(rowCnt, 3).value = Round(temp2 / denom, 2)
      End If
      sumNum(1) = sumNum(1) + temp2
      sumDenom(1) = sumDenom(1) + denom
      fldRec.MoveNext
      If MyP.Length = 4 Or MyP.Length = 8 Then
        denom = EvalArray(fldRec("ID"), i - 1)
        fldRec.MoveNext
        temp2 = EvalArray(fldRec("ID"), i - 1)
        If denom = 0 Then
          Cells(rowCnt, 6).value = 0
        Else
          Cells(rowCnt, 6).value = Round(temp2 * 1000 / denom, 2)
        End If
        sumNum(4) = sumNum(4) + temp2 * 1000
        sumDenom(4) = sumDenom(4) + denom
      End If
      rowCnt = rowCnt + 1
      fldRec.MoveFirst
    Next i
    Cells(rowCnt, 1).Characters.Text = "Weighted Averages:"
    Cells(rowCnt, 1).HorizontalAlignment = xlHAlignRight
    If Left(MyP.UnitArea, 1) = "H" Then j = 6 Else j = 5
    For i = 2 To j
      If sumDenom(i - 2) = 0 Then
        Cells(rowCnt, i).value = 0
      Else
        Cells(rowCnt, i).value = Round(sumNum(i - 2) / sumDenom(i - 2), 2)
      End If
      With Range(Cells(rowCnt, 1), Cells(rowCnt, j)).Borders(xlTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
      End With
    Next i
    With Range(Cells(4, 2), Cells(rowCnt, 6))
      .NumberFormat = "##0.00"
      .HorizontalAlignment = xlHAlignRight
    End With
    .Rows(rowCnt).HorizontalAlignment = xlHAlignRight
    rowCnt = rowCnt + 5
    Clipboard.Clear
    If rdoID(2) Then
      .Columns(2).Insert
      With Range(Cells(HeaderRows, 1), Cells(HeaderRows, 2))
        .Merge
        .HorizontalAlignment = xlHAlignCenter
      End With
      For i = 1 To UBound(LocnArray, 2) + 1
        Cells(i + HeaderRows, 2).value = LocnArray(1, i - 1)
      Next i
      With Range(Cells(i + HeaderRows, 1), Cells(i + HeaderRows, 2))
        .Merge
        .HorizontalAlignment = xlHAlignRight
      End With
    End If
    .Columns(1).ColumnWidth = 14
    .Name = "Calculated Fields_" & MyP.UnitArea
    .Range(.Columns(2), .Columns(6)).AutoFit
  End With
  CatRec.Close
x:
  Set xlRange = XLSheet.UsedRange
  With xlRange
    .Font.size = 8
    .Font.Name = "Times New Roman"
  End With
  AllFldRec.Close
  TotalPopRec.Close
  EndReport
End Sub

Private Sub FacilityReport(RepPath As String)
Attribute FacilityReport.VB_Description = "Creates "
' ##SUMMARY Creates "Facility Tables" report as Excel output file.
' ##PARAM RepPath I String full pathname where report will be saved.
' ##REMARKS Report is in single table in worksheet.
  Dim rowCnt As Long
  Dim dictYear As Long
  Dim numFlds As Long
  Dim thisArea As Long
  Dim i As Long
  Dim marker As Variant
  Dim Val As Long
  Dim sql As String
  Dim str As String
  Dim fldRec As Recordset
  Dim xlRange As Excel.Range

  On Error GoTo x

  InitReport RepPath
  rowCnt = 1
  If AreaID = 2 Then AreaID = 0
  
  With XLSheet
    .Columns.ColumnWidth = 7
    .Columns(1).HorizontalAlignment = xlHAlignLeft
    'Create header
    dictYear = CLng(Left(MyP.YearFields, 4))
    If dictYear < 2000 Then
      GetHeader 1
      numFlds = 10
    Else
      GetHeader 2
      numFlds = 9
    End If
    .Activate
    Set xlRange = XLSheet.Range(Cells(rowCnt, 1), Cells(rowCnt, 1))
    XLSheet.Paste xlRange
    xlRange.HorizontalAlignment = xlHAlignLeft
    If NationalDB Then
      str = "State"
    Else
      If rdoAreaUnit(1) Then
        str = "hydrologic cataloging unit"
      ElseIf rdoAreaUnit(2) Then
        str = "water-resources region"
      Else
        str = MyP.UnitArea
      End If
    End If
    Cells(rowCnt, 1).value = "Total Facilities by Category, by " _
        & str & " - " & MyP.State & ", " & MyP.Year1Opt
    Range(Cells(rowCnt, 1), Cells(rowCnt, 5)).Font.Bold = True
    rowCnt = rowCnt + HeaderRows
    Cells(rowCnt - 1, 1).value = MyP.UnitArea
    'Fill in body of table
    For thisArea = 0 To UBound(LocnArray, 2)
      'Use SQL language to create recordset with all fields that will be used in report.
      ' Different "Field" tables that comprise the data dictionary are joined.
      ' For example:
      '  SELECT [Field0].ID FROM ([Category1]
      '  INNER JOIN [Field0] ON [Category1].ID = [Field0].CategoryID)
      '    INNER JOIN [1995Fields1] ON [Field0].ID = [1995Fields1].FieldID
      '  WHERE (Right([Field0].Name, 3)='Fac' OR Mid([Field0].Name, 4, 3)='Fac')
      '  ORDER BY [Category1].ID, [Field0].ID
      sql = "SELECT [" & FieldTable & "].ID FROM " & _
            "([" & CatTable & "] INNER JOIN [" & FieldTable & "] ON [" & CatTable & "].ID = [" & FieldTable & "].CategoryID) " & _
            "INNER JOIN [" & MyP.YearFields & "] ON [" & FieldTable & "].ID = [" & MyP.YearFields & "].FieldID " & _
            "WHERE (Right([" & FieldTable & "].Name, 3)='Fac' OR Mid([" & FieldTable & "].Name, 4, 3)='Fac') " & _
            "ORDER BY [" & CatTable & "].ID, [" & FieldTable & "].ID"
      Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
      'write in name of area
      If AreaID = 0 Then
        Cells(rowCnt, 1).value = "'" & LocnArray(AreaID, thisArea)
      Else
        Cells(rowCnt, 1).value = LocnArray(AreaID, thisArea)
      End If
      For i = 2 To numFlds + 1
        Val = EvalArray(fldRec("ID"), thisArea)
        If Not NoRec Then Cells(rowCnt, i).value = Val
        fldRec.MoveNext
      Next i
      rowCnt = rowCnt + 1
      AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (thisArea + 1) * 100 / (UBound(LocnArray, 2) + 1) & ")"
    Next thisArea
    .Name = "Facility Report - " & MyP.UnitArea
  End With
  With Cells(rowCnt, 1)
    .value = "Totals:"
    .HorizontalAlignment = xlHAlignRight
  End With
  For i = 2 To numFlds + 1
    Cells(rowCnt, i).value = Application.WorksheetFunction. _
        Sum(Range(Cells(HeaderRows + 1, i), Cells(rowCnt - 1, i)))
  Next i
  With Range(Cells(rowCnt, 1), Cells(rowCnt, i - 1)).Borders(xlTop)
    .LineStyle = xlContinuous
    .Weight = xlThin
  End With
  If rdoID(2) Then
    XLSheet.Columns(2).Insert
    With Range(Cells(HeaderRows, 1), Cells(HeaderRows, 2))
      .Merge
      .HorizontalAlignment = xlHAlignCenter
    End With
    For i = 1 To UBound(LocnArray, 2) + 1
      Cells(i + HeaderRows, 2).value = LocnArray(1, i - 1)
    Next i
    With Range(Cells(i + HeaderRows, 1), Cells(i + HeaderRows, 2))
      .Merge
      .HorizontalAlignment = xlHAlignRight
    End With
  End If
x:
  Set xlRange = XLSheet.UsedRange
  xlRange.Font.size = 8
  xlRange.Font.Name = "Times New Roman"
  AllFldRec.Close
  TotalPopRec.Close
  EndReport
End Sub

Private Sub QAReport(RepPath As String, CatRec As Recordset)
Attribute QAReport.VB_Description = "Performs "
' ##SUMMARY Performs "Quality-Assurance Program" QA check with results written to Excel output file.
' ##PARAM RepPath I String full pathname where report will be saved.
' ##PARAM catRec I Recordset containing user-selected categories.
' ##REMARKS Results of QA check are listed one category at a time on single worksheet. _
          When QA infraction is found, 1 new row is inserted into worksheet _
          containing information regarding infraction.
  Dim xlRange As Excel.Range
  Dim dataRec As Recordset
  Dim fldsInCatRec As Recordset
  Dim rowCnt As Long 'tracks which row messages are written on worksheet.
  Dim frUses As Double
  Dim saUses As Double
  Dim totUses As Double
  Dim frSources As Double
  Dim saSources As Double
  Dim totSources As Double
  Dim sql As String
  Dim i As Long
  Dim areaCount As Long
  Dim a As Long 'tracks number of rows inserted into worksheet for Public Supply category.
  Dim b As Long
  Dim c As Long
  Dim d As Long
  Dim e As Long
  Dim f As Long
  Dim g As Long
  Dim qFlag As Long
  Dim allAreas As String
  Dim tempArea As String
  
  On Error GoTo x

  InitReport RepPath
  
  With XLSheet
    'create header for report
    Cells(1, 1).value = "Quality-Assurance checks for " & _
          MyP.State & ", " & MyP.Year1Opt
    Cells(2, 1).value = "Date of this report:  " & _
          Date
    If lstArea.LeftCount = 0 Then
      Cells(4, 1).value = "All " & MyP.UnitArea & _
            " areas are selected for this report."
    Else
      For i = 1 To lstArea.RightCount
        If i = 1 Then
          allAreas = LocnArray(0, 0)
        Else
          allAreas = allAreas & ", " & LocnArray(0, i - 1)
        End If
      Next i
      Cells(4, 1).value = MyP.UnitArea & " areas selected for this report:  " & allAreas
    End If
    With Range(Cells(1, 1), Cells(4, 1))
      .Font.Bold = True
      .HorizontalAlignment = xlHAlignLeft
    End With
    
    GetHeader 1
    .Activate
    XLSheet.Paste Cells(5, 1)
    Cells(5, 1).value = ""
    rowCnt = 8
    'Get rid of the space holders in the header
    For i = 0 To 5
      Cells(rowCnt + 3 * i, 1).value = ""
    Next i
    For areaCount = 0 To UBound(LocnArray, 2)
      sql = "SELECT * FROM [LastReport] " & _
          "WHERE Left(Location, " & MyP.Length & ")='" & LocnArray(0, areaCount) & "'"
      Set dataRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
      If dataRec.EOF Then GoTo x
      If dataRec("QualFlg") < 7 Then
        qFlag = dataRec("QualFlg")
      ElseIf dataRec("QualFlg") = 7 Then
        qFlag = 1
      ElseIf dataRec("QualFlg") = 8 Then
        qFlag = 5
      End If
      rowCnt = 7 + a
      'Perform population check
      AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (1 / 7 + areaCount) * 100 / _
                                        (UBound(LocnArray, 2) + 1) & ")"
      tempArea = LocnArray(AreaID, areaCount)
      If EvalArray(1, areaCount) < EvalArray(4, areaCount) Then
        Cells(rowCnt, 1).value = "The population served by Public Supply in " & _
            tempArea & " is greater than the total population for that area."
        InsertRow rowCnt, a
      End If
      rowCnt = rowCnt + 3 + b
      If Len(Cells(rowCnt - 1, 1)) = 0 Then
        rowCnt = rowCnt - 1
      End If
      'Perform population vs withdrawal & use checks
      AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (2 / 7 + areaCount) * 100 / _
                                        (UBound(LocnArray, 2) + 1) & ")"
      If b = 0 And Left(Cells(rowCnt, 1), 4) <> "-->>" Then
        rowCnt = rowCnt - 1
      End If
      If EvalArray(39, areaCount) <= 0 Then
        For i = 40 To 44
          If EvalArray(i, areaCount) > 0 Then
            AllFldRec.FindFirst "ID=" & i
            Cells(rowCnt, 1).value = "The self-supplied population of " & _
                tempArea & " is zero or null, yet the Domestic " & _
                "category field " & AllFldRec("Name") & " has a positive value."
            InsertRow rowCnt, b
          End If
        Next i
      End If
      If EvalArray(39, areaCount) > 0 And EvalArray(46, areaCount) = 0 Then
        Cells(rowCnt, 1).value = "The self-supplied population of " & _
            tempArea & " is reported as " & EvalArray(39, areaCount) & ", " & _
            "yet there are no Domestic withdrawals in that area."
        InsertRow rowCnt, b
      End If
      If EvalArray(4, areaCount) <= 0 Then
        If EvalArray(51, areaCount) > 0 Then
          AllFldRec.FindFirst "ID=" & 51
          Cells(rowCnt, 1).value = "The Public Supply population of " & tempArea & _
              " is zero or null, yet there are deliveries from Public Supply to Domestic."
          InsertRow rowCnt, b
        End If
      ElseIf EvalArray(4, areaCount) > 0 Then
        If EvalArray(51, areaCount) = 0 Then
          AllFldRec.FindFirst "ID=" & 51
          Cells(rowCnt, 1).value = "The Public Supply population of " & _
              tempArea & " is reported as " & EvalArray(4, areaCount) & ", " & _
              "yet there are no deliveries from Public Supply to Domestic."
          InsertRow rowCnt, b
        End If
      End If
      rowCnt = rowCnt + 3 + c
      If Len(Cells(rowCnt - 1, 1)) = 0 Then
        rowCnt = rowCnt - 1
      End If
      'Perform consuptive use check
      AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (3 / 7 + areaCount) * 100 / _
                                        (UBound(LocnArray, 2) + 1) & ")"
      If c = 0 And Left(Cells(rowCnt, 1), 4) <> "-->>" Then
        rowCnt = rowCnt - 1
      End If
      For i = 1 To CatRec.RecordCount
        sql = "SELECT ID, Name FROM [" & FieldTable & _
            "] WHERE CategoryID=" & CatRec("ID") & " ORDER BY ID;"
        Set fldsInCatRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
        fldsInCatRec.FindFirst "Name='" & CatRec("Name") & "-CUsFr'"
        On Error Resume Next
        If Not fldsInCatRec.NoMatch Then
          frUses = EvalArray(fldsInCatRec("ID"), areaCount)
          fldsInCatRec.MoveNext
          If Right(fldsInCatRec("Name"), 5) = "CUsSa" Then
            saUses = EvalArray(fldsInCatRec("ID"), areaCount)
          Else
            saUses = 0
          End If
          fldsInCatRec.FindFirst "Name='" & CatRec("Name") & "-WFrTo'"
          frSources = EvalArray(fldsInCatRec("ID"), areaCount)
          fldsInCatRec.MoveNext
          If Right(fldsInCatRec("Name"), 5) = "WSaTo" Then
            saSources = EvalArray(fldsInCatRec("ID"), areaCount)
          Else
            saSources = 0
          End If
          fldsInCatRec.FindFirst "Name='" & CatRec("Name") & "-PSDel'"
          If Not fldsInCatRec.NoMatch Then _
              frSources = frSources + EvalArray(fldsInCatRec("ID"), areaCount)
          fldsInCatRec.FindFirst "Name='" & CatRec("Name") & "-RecWW'"
          If Not fldsInCatRec.NoMatch Then _
              frSources = frSources + EvalArray(fldsInCatRec("ID"), areaCount)
          If frUses - frSources > 0.005 Then
            Cells(rowCnt, 1).value = "The fresh consumptive use for " & CatRec("Description") _
                & " in " & tempArea & _
                " is greater than the sum of the freshwater sources.  Sources may include total " & _
                "withdrawals, public supply deliveries, and reclaimed wastewater."
            InsertRow rowCnt, c
          End If
          If saUses - saSources > 0.005 Then
            Cells(rowCnt, 1).value = "The saline consumptive use for " & CatRec("Description") _
                & " in " & tempArea & " is greater than the sum of the saline withdrawals."
            InsertRow rowCnt, c
          End If
        End If
        CatRec.MoveNext
        fldsInCatRec.Close
      Next i
      CatRec.MoveFirst
      rowCnt = rowCnt + 3 + d
      If Len(Cells(rowCnt - 1, 1)) = 0 Then
        rowCnt = rowCnt - 1
      End If
      'Perform irrigation use check
      AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (4 / 7 + areaCount) * 100 / _
                                        (UBound(LocnArray, 2) + 1) & ")"
      If d = 0 And Left(Cells(rowCnt, 1), 4) <> "-->>" Then
        rowCnt = rowCnt - 1
      End If
      If IRinTwo Then
        totSources = EvalArray(302, areaCount) + EvalArray(303, areaCount) _
                + EvalArray(311, areaCount)
        totUses = EvalArray(305, areaCount) + EvalArray(306, areaCount)
      Else
        totSources = EvalArray(195, areaCount) + EvalArray(198, areaCount) _
                + EvalArray(212, areaCount)
        totUses = EvalArray(204, areaCount) + EvalArray(207, areaCount)
      End If
      If totUses - totSources > 0.005 Then
        Cells(rowCnt, 1).value = "The sum of total conveyence loss and " & _
            "total consuptive use is greater than the sum of the totsources in " & _
            tempArea & _
            ".  Sources include total withdrawals and reclaimed wastewater."
            InsertRow rowCnt, d
      End If
      rowCnt = rowCnt + 3 + e
      If Len(Cells(rowCnt - 1, 1)) = 0 Then
        rowCnt = rowCnt - 1
      End If
      'Perform reservoir evaporation check if unit area is a HUC
      If MyP.Length = 4 Or MyP.Length = 8 Then
        If e = 0 And Left(Cells(rowCnt, 1), 4) <> "-->>" Then
          rowCnt = rowCnt - 1
        End If
        AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (5 / 7 + areaCount) * 100 / _
                                        (UBound(LocnArray, 2) + 1) & ")"
        If EvalArray(232, areaCount) <= 0 _
            And EvalArray(233, areaCount) > 0 Then
          Cells(rowCnt, 1).value = "The total reservoir surface area in " & _
              MyP.UnitArea & " " & tempArea & _
              " is zero, yet the reservoir evaporation field has a positive value."
          InsertRow rowCnt, e
        ElseIf EvalArray(233, areaCount) <= 0 _
            And EvalArray(232, areaCount) > 0 Then
          Cells(rowCnt, 1).value = "There is positive total reservoir surface area in " & _
              tempArea & _
              ", yet there is no value for reservoir evaporation."
          InsertRow rowCnt, e
        End If
      End If
      rowCnt = rowCnt + 3 + f
      If Len(Cells(rowCnt - 1, 1)) = 0 Then
        rowCnt = rowCnt - 1
      End If
      'Perform hydroelectric power check
      AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (6 / 7 + areaCount) * 100 / _
                                        (UBound(LocnArray, 2) + 1) & ")"
      If f = 0 And Left(Cells(rowCnt, 1), 4) <> "-->>" Then
        rowCnt = rowCnt - 1
      End If
      If (EvalArray(214, areaCount) + EvalArray(215, areaCount) = 0) And _
          EvalArray(218, areaCount) > 0 Then
        Cells(rowCnt, 1).value = "Offstream power generation has a positive value in " & _
            tempArea & _
            ", yet there are no offstream water withdrawals."
        InsertRow rowCnt, f
      ElseIf (EvalArray(214, areaCount) + EvalArray(215, areaCount)) > 0 _
          And EvalArray(218, areaCount) <= 0 Then
        Cells(rowCnt, 1).value = "There are offstream water withdrawals in " & _
            tempArea & _
            ", yet there is no offstream power produced."
        InsertRow rowCnt, f
      End If
      If EvalArray(213, areaCount) <= 0 And _
          EvalArray(217, areaCount) > 0 Then
        Cells(rowCnt, 1).value = "Instream power generation has a positive value in " & _
            tempArea & _
            ", yet the instream water use field has a zero value."
        InsertRow rowCnt, f
      ElseIf EvalArray(213, areaCount) > 0 _
          And EvalArray(217, areaCount) <= 0 Then
        Cells(rowCnt, 1).value = "The instream water use field has a positive value in " & _
            tempArea & _
            ", yet there is no instream power produced."
        InsertRow rowCnt, f
      End If
      rowCnt = rowCnt + 3 + g
      If Len(Cells(rowCnt - 1, 1)) = 0 Then
        rowCnt = rowCnt - 1
      End If
      'Perform facilities check
      AtcoLaunch1.SendMonitorMessage "(PROGRESS " & (1 + areaCount) * 100 / _
                                       (UBound(LocnArray, 2) + 1) & ")"
      If g = 0 And Left(Cells(rowCnt, 1), 4) <> "-->>" Then
        rowCnt = rowCnt - 1
      End If
      If EvalArray(108, areaCount) <= 0 _
          And EvalArray(106, areaCount) > 0 Then
        Cells(rowCnt, 1).value = "There is fossil-fuel power generation in " & _
            tempArea & _
            ", yet there are no facilities entered for this area."
        InsertRow rowCnt, g
      End If
      If EvalArray(126, areaCount) <= 0 _
          And EvalArray(124, areaCount) > 0 Then
        Cells(rowCnt, 1).value = "There is geothermal power generation in " & _
            tempArea & _
            ", yet there are no facilities entered for this area."
        InsertRow rowCnt, g
      End If
      If EvalArray(144, areaCount) <= 0 _
          And EvalArray(142, areaCount) > 0 Then
        Cells(rowCnt, 1).value = "There is nuclear power generation in " & _
            tempArea & _
            ", yet there are no facilities entered for this area."
        InsertRow rowCnt, g
      End If
      If EvalArray(264, areaCount) <= 0 And _
          EvalArray(262, areaCount) > 0 Then
        Cells(rowCnt, 1).value = "There is power generation by facilities with " & _
            "once-through cooling systems in " & tempArea & _
            ", yet there are no such facilities entered for this area."
        InsertRow rowCnt, g
      End If
      If EvalArray(281, areaCount) <= 0 _
          And EvalArray(279, areaCount) > 0 Then
        Cells(rowCnt, 1).value = "There is power generation by facilities with " & _
            "closed-loop cooling systems in " & tempArea & _
            ", yet there are no such facilities entered for this area."
        InsertRow rowCnt, g
      End If
      If EvalArray(220, areaCount) <= 0 _
          And EvalArray(217, areaCount) > 0 Then
        Cells(rowCnt, 1).value = "Instream Hydroelectric power generation " & _
            "has a positive value in " & tempArea & _
            ", yet there are no facilities entered for this area."
        InsertRow rowCnt, g
      End If
      If EvalArray(221, areaCount) <= 0 _
          And EvalArray(218, areaCount) > 0 Then
        Cells(rowCnt, 1).value = "Offstream Hydroelectric power generation " & _
            "has a positive value in " & tempArea & _
            ", yet there are no facilities entered for this area."
        InsertRow rowCnt, g
      End If
      If EvalArray(227, areaCount) <= 0 _
          And EvalArray(226, areaCount) > 0 Then
        Cells(rowCnt, 1).value = "There are public wastewater returns entered for " & _
            tempArea & _
            ", yet there are no public wastewater facilities entered for this area."
        InsertRow rowCnt, g
      End If
x:
    Next areaCount
    Set xlRange = .UsedRange
    xlRange.Font.size = 8
    xlRange.Font.Name = "Times New Roman"
    Clipboard.Clear
    .Name = "QA Report_" & MyP.UnitArea
  End With
  CatRec.Close
  AllFldRec.Close
  TotalPopRec.Close
  EndReport
End Sub

Private Sub InsertRow(rowCnt As Long, Charctr As Long)
Attribute InsertRow.VB_Description = "Inserts row into "
' ##SUMMARY Inserts row into "Quality-Assurance Program" Excel output file.
' ##PARAM rowCnt I Integer indicating location of new row in worksheet.
' ##PARAM Charctr M Long used to track number of rows inserted into worksheet for each _
          respective category.
  rowCnt = rowCnt + 1
  If Charctr > 0 Then
    XLSheet.Rows(rowCnt).Select
    XLSheet.Rows(rowCnt).Insert
    XLSheet.Rows(rowCnt).HorizontalAlignment = xlHAlignLeft
  End If
  Charctr = Charctr + 1
End Sub

Private Sub CompAreasReport(RepPath As String, CatRec As Recordset)
Attribute CompAreasReport.VB_Description = "Performs "
' ##SUMMARY Performs "Compare State Totals by Area" QA check with results written to _
          Excel output file.
' ##PARAM RepPath I String full pathname where report will be saved.
' ##PARAM catRec I Recordset containing user-selected categories.
' ##REMARKS Results of QA check are listed in single table in worksheet.
  Dim xlRange As Excel.Range
  Dim sumDataRec As Recordset
  Dim sumDataRec2 As Recordset
  Dim fldRec As Recordset
  Dim fldRec2 As Recordset
  Dim rowCnt As Long
  Dim totalPop As Single
  Dim temp1 As Double
  Dim temp2 As Double
  Dim fldID1 As Long
  Dim fldID2 As Long
  Dim numAreas1 As Long
  Dim numAreas2 As Long
  Dim i As Long
  Dim fldRecCnt As Long
  Dim numFldRecs As Long
  Dim lastColumn As Long
  Dim lastRow As Long
  Dim headopt As Long
  Dim sql As String
  Dim ir1 As String
  Dim ir2 As String
  Dim irrFld As String
  Dim areaName As String
  Dim header As String
  Dim message As String
  Dim sameIR As Boolean
  
  On Error GoTo x
  
  InitReport RepPath
  rowCnt = 1
  CatRec.MoveFirst
  
  'Check to see if the same data storage options are used for both areas
  'Set Headerfile accordingly
  If MyP.Length = 2 Then
    For i = 0 To 3
      If rdoAreaUnit(i) Then Exit For
    Next i
    Select Case i 'first unit area can only be either County of HUC-8
      Case 0: MyP.Length = 3
      Case 1: MyP.Length = 8
    End Select
  ElseIf MyP.length2 = 2 Then
    MyP.length2 = 3
  End If
  'Find out number of areas for both unit areas in order to divide
  'sum of state totals if QualFlg > 1
  sql = "SELECT DISTINCT Location FROM [LastReport] " & _
        "WHERE Len(Trim(Location))=" & MyP.Length
  Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  fldRec.MoveLast
  numAreas1 = fldRec.RecordCount
  fldRec.Close
  sql = "SELECT DISTINCT Location FROM [LastReport] " & _
        "WHERE Len(Trim(Location))=" & MyP.length2
  Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  fldRec.MoveLast
  numAreas2 = fldRec.RecordCount
  fldRec.Close
  
  'Check to see if the same IR storage options are used for both unit areas
  sql = "SELECT * FROM [LastReport] WHERE len(Location)=" & MyP.Length
  Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  fldRec.FindFirst "FieldID=195"
  If fldRec.NoMatch Then sameIR = True Else sameIR = False
  sql = "SELECT * FROM [LastReport] WHERE len(Location)=" & MyP.length2
  Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  fldRec.FindFirst "FieldID=195"
  If (fldRec.NoMatch And sameIR) Or (fldRec.NoMatch = False And sameIR = False) Then
    sameIR = True
  Else
    sameIR = False
  End If
  'IR1 and IR2 establish which IR fields to include
  If MyP.DataOpt = 0 Then  '95 Data Dictionary
    If rdoAreaUnit2(3) Then 'Aquifer is 2nd unit area; need IR totals
      headopt = 0
      ir1 = " OR ID=302) AND Not ([" & FieldTable & "].CategoryID>16 AND [" & FieldTable & "].CategoryID<20)"
      ir2 = ir1
    Else
      headopt = 2
      ir1 = ")"
      ir2 = ir1
    End If
  ElseIf NationalDB Then  'IR is "Total"
    headopt = 1
    ir1 = " OR [" & FieldTable & "].CategoryID=16) " & _
        "AND Not ([" & FieldTable & "].CategoryID>16 AND [" & FieldTable & "].CategoryID<20)"
    ir2 = ir1
    'Add commentary re fields not included in comparison
    message = "When aggregating state data for the National DB," & vbCrLf & _
        "all data is aggregated on the grossest possible level." & vbCrLf & _
        "Specifically, irrigation is taken as a total (not broken out into Golf and Crop);" & vbCrLf & _
        "Domestic freshwater withdrawals are not broken out into GW and SW; and " & vbCrLf & _
        "Public Supply - total population served is not broken out into GW and SW." & vbCrLf & vbCrLf & _
        "Hence, comparisons between different unit areas on the national level " & vbCrLf & _
        "will not include those more detailed fields."
  ElseIf sameIR And IRinTwo Then  'IR split for both areas
    headopt = 2
    ir1 = ") AND " & _
        "Not([" & FieldTable & "].ID>194 AND [" & FieldTable & "].ID<213)"
    ir2 = ir1
  ElseIf sameIR And Not IRinTwo Then  'IR not split for either area
    headopt = 1
    ir1 = ") AND " & _
        "Not([" & FieldTable & "].ID>281 AND [" & FieldTable & "].ID<302)"
    ir2 = ir1
  ElseIf fldRec.NoMatch Then  'IR in second unit area is split, but not in first
    headopt = 1
    ir1 = ") AND Not([" & FieldTable & "].ID>281 AND [" & FieldTable & "].ID<302)"
    ir2 = " OR [" & FieldTable & "].CategoryID=16) AND [" & FieldTable & "].ID<>304 " & _
                "AND [" & FieldTable & "].ID<>310 AND Not([" & FieldTable & _
                "].CategoryID>16 AND [" & FieldTable & "].CategoryID<20)"
    'Add commentary re fields not included in comparison
    message = MyP.UnitArea & " areas in " & MyP.State & " keep irrigation as a total," & vbCrLf & _
        "while " & MyP.UnitArea2 & " areas divide irrigation into golf and crop use." & vbCrLf & _
        "By necessity, the " & MyP.UnitArea2 & " irrigation data will be totaled before" & vbCrLf & _
        "being compared to the " & MyP.UnitArea & " data."
  Else  'IR in first unit area is split, but not in second
    headopt = 1
    ir1 = " OR [" & FieldTable & "].CategoryID=16) AND [" & FieldTable & "].ID<>304 " & _
                "AND [" & FieldTable & "].ID<>310 AND Not([" & FieldTable & _
                "].CategoryID>16 AND [" & FieldTable & "].CategoryID<20)"
    ir2 = ") AND Not([" & FieldTable & "].ID>281 AND [" & FieldTable & "].ID<302)"
    'Add commentary re fields not included in comparison
    message = MyP.UnitArea2 & " areas in " & MyP.State & " keep irrigation as a total," & vbCrLf & _
        "while " & MyP.UnitArea & " areas divide irrigation into golf and crop use." & vbCrLf & _
        "By necessity, the " & MyP.UnitArea & " irrigation data will be totaled before" & vbCrLf & _
        "being compared to the " & MyP.UnitArea2 & " data."
  End If
  'Adjust header choice if aquifer is one unit area and other unit area
  'does not keep PS-GWPop
  If (MyP.Length = 10 And MyP.DataOpt2 > 2) Or (MyP.length2 = 10 And MyP.DataOpt > 2) Then
    headopt = headopt + 2
    If Not NationalDB Then
      'Add commentary re fields not included in comparison
      If Len(message) > 0 Then message = message & vbCrLf & vbCrLf & "Also, "
      message = message & "You are comparing aquifer data against "
      If MyP.DataOpt2 > 2 Then message = message & MyP.UnitArea2 Else message = message & MyP.UnitArea
      message = message & " data" & vbCrLf & "that does not break out Public Supply - " & _
      "total population served into GW and SW." & vbCrLf & "As a result, that GW datum " & _
      "can not be compared."
    End If
  ElseIf MyP.DataOpt <> MyP.DataOpt2 And Not NationalDB Then
    'Add commentary re fields not include  in comparison
    If MyP.DataOpt > 2 Then
      If Len(message) > 0 Then
        message = message & vbCrLf & vbCrLf & "Also, for " & MyP.UnitArea & " areas in " & MyP.State & ", "
      Else
        message = "For " & MyP.UnitArea & " areas in " & MyP.State & ", "
      End If
      message = message & "Public Supply - total population served" & vbCrLf & _
            "is not broken out into GW and SW; therefore, this datum is compared as a total."
    End If
    If MyP.DataOpt2 > 2 Then
      If Len(message) > 0 Then
        message = message & vbCrLf & vbCrLf & "Also, for " & MyP.UnitArea2 & " areas in " & MyP.State & ", "
      Else
        message = "For " & MyP.UnitArea2 & " areas in " & MyP.State & ", "
      End If
      message = message & "Public Supply - total population served" & vbCrLf & _
            "is not broken out into GW and SW; therefore, this datum is compared as a total."
    End If
  End If
  If Len(message) > 0 Then MsgBox message, , "Fields Not Used in Comparison"
  'Use SQL language to create recordsets with all fields of user-entered data
  ' that will be compared in report. There are 2 queries; one for each unit area.
  ' "Total Population" and "Reservoir Evaporation" categories are excluded,
  ' and considerations for the Irrigation storage options are made via the
  ' ir1 and ir2 variables.
  ' For example:
  '  SELECT ID, Name, CategoryID FROM [Field1]
  '  WHERE (Formula='') AND Not([Field1].ID>194 AND [Field1].ID<213)
  '  AND CategoryID<>1 AND CategoryID<>22
  '  ORDER BY CategoryID, ID;
  sql = "SELECT ID, Name, CategoryID FROM [" & FieldTable & _
      "] WHERE (Formula=''" & ir1 & _
      " AND CategoryID<>1 AND CategoryID<>22" & _
      " ORDER BY CategoryID, ID;"
  Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  fldRec.MoveLast
  fldRec.MoveFirst
  sql = "SELECT ID, Name, CategoryID FROM [" & FieldTable & _
      "] WHERE (Formula=''" & ir2 & _
      " AND CategoryID<>1 AND CategoryID<>22" & _
      " ORDER BY CategoryID, ID;"
  Set fldRec2 = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
  fldRec2.MoveLast
  fldRec2.MoveFirst
  numFldRecs = fldRec.RecordCount

  'Assemble report
  With XLSheet
    .Columns.ColumnWidth = 5.8
    'Create header for this area unit
    For i = 1 To 4
      Select Case i
        Case 1: header = RepTitle & " report for " & MyP.UnitArea & " and " & _
                         MyP.UnitArea2 & " - " & MyP.State & ", " & MyP.Year1Opt
        Case 2: header = ""
        Case 3: header = "[Data units in million gallons per day " & _
                         "(mgd) unless otherwise noted]"
        Case 4: header = "Differences are the state totals by " & MyP.UnitArea & _
                         " minus the state totals by " & MyP.UnitArea2
      End Select
      Cells(rowCnt, 1).value = header
      rowCnt = rowCnt + 1
    Next i
    rowCnt = rowCnt + 1
    'calc difference in total population of unit areas
    sql = "SELECT Sum(Value) AS Population FROM [" & MyP.AreaTable & "] " & _
          "WHERE FieldID=1 And Date=" & MyP.Year1Opt & _
          " And Len(trim(location))=" & MyP.Length
    Set TotalPopRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    If Not IsNull(TotalPopRec("Population")) Then totalPop = TotalPopRec("Population")
    sql = "SELECT Sum(Value) AS Population FROM [" & TableName2 & "Data] " & _
          "WHERE FieldID=1 And Date=" & MyP.Year1Opt & _
          " And Len(trim(location))=" & MyP.length2
    Set TotalPopRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    If Not IsNull(TotalPopRec("Population")) Then totalPop = totalPop - TotalPopRec("Population")
    Cells(rowCnt, 1).value = "Total population:  " & Format(Round(totalPop, 3), "#,##0.000")
    rowCnt = rowCnt + 2
    'Get header and paste onto report if necessary
    GetHeader headopt
    .Activate
    .Paste XLSheet.Range(Cells(rowCnt, 1), Cells(rowCnt, 1))
    Clipboard.Clear
    rowCnt = rowCnt + 1
    Cells(rowCnt - 1, 1).value = ""
    If (MyP.Length = 10 Or MyP.length2 = 10) And headopt > 2 Then
      fldRec.MoveNext
      fldRec2.MoveNext
      numFldRecs = numFldRecs - 1
    End If
    Cells(rowCnt, 2).Select
    i = 1
    'Fill in body of table with data
    With Range(Cells(rowCnt, 2), _
               Cells(rowCnt + HeaderRows - 2, CatRec.RecordCount + 3))
      Set xlRange = .Find(-1, , , , xlByColumns, xlNext, False)
      For fldRecCnt = 0 To numFldRecs - 1
        If Not xlRange Is Nothing Then
          If Not (fldID1 = 217 Or fldID1 = 220) Then
            lastColumn = xlRange.Column
            lastRow = xlRange.Row
          Else
            lastColumn = lastColumn + 1
          End If
'          If (lastColumn = 13 And headopt = 2) Or _
'             (lastColumn = 12 And headopt = 1) Then
'            'rearranging order to account for division of hydroelectric
'            'into instream and offstream
'            If i = 2 Or (i = 3 And MyP.DataOpt = 0) Then
'              Cells(lastRow, lastColumn) = ""
'              lastColumn = lastColumn + 1
'            ElseIf (MyP.DataOpt = 0 And (i = 5 Or i = 7)) _
'                Or (MyP.DataOpt > 0 And (i = 4 Or i = 6)) Then
'              Cells(lastRow, lastColumn) = ""
'              lastColumn = lastColumn + 1
'              lastRow = lastRow - 1
'            End If
'            i = i + 1
          If (lastColumn = 15 And headopt = 2) Or _
             (lastColumn = 14 And headopt = 1) Then
            'rearranging order to flip-flop Public WW Facilities and WW Returns
            If i > 3 Then i = 1  'reset variable from hydroelectric columns
            If i = 3 Then
              Cells(lastRow, lastColumn) = ""
              lastRow = lastRow - 2
            End If
            i = i + 1
          End If
          'Determine sum of current field for first unit area as baseline value
          fldID1 = fldRec("ID")
          fldID2 = fldRec2("ID")
          If fldID1 > 301 Then  'IR is sum of agriculture and golf
            Select Case fldID1
              Case 302: irrFld = " OR FieldID=195"
              Case 303: irrFld = " OR FieldID=198"
              Case 305: irrFld = " OR FieldID=204"
              Case 306: irrFld = " OR FieldID=207"
              Case 307: irrFld = " OR FieldID=208"
              Case 308: irrFld = " OR FieldID=209"
              Case 309: irrFld = " OR FieldID=210"
              Case 311: irrFld = " OR FieldID=212"
              Case Else: irrFld = ""
            End Select
            sql = "SELECT Sum(Value) AS SumVal FROM [LastReport] " & _
                  "WHERE (FieldID=" & fldID1 - 10 & " OR FieldID=" & fldID1 - 20 & irrFld & _
                  ") And Len(trim(location))=" & MyP.Length
          ElseIf fldID1 = 4 And MyP.DataOpt < 3 Then 'PS - Population Served not a state total
            sql = "SELECT Sum(Value) AS SumVal FROM [LastReport] " & _
                  "WHERE (FieldID=2 Or FieldID=3) " & _
                  "And Len(trim(location))=" & MyP.Length
          ElseIf MyP.UnitArea2 = "Aquifer" And MyP.DataOpt = 0 _
              And (fldID1 = 248 Or fldID1 = 249) Then
            'Comparing Aquifer w/ 1995 data; need to take Thermoelectric total
            sql = "SELECT Sum(Value) AS SumVal FROM [LastReport] " & _
                  "WHERE FieldID=" & fldID1 - 174 & _
                  " And Len(trim(location))=" & MyP.Length
          Else
            sql = "SELECT Sum(Value) AS SumVal FROM [LastReport] " & _
                  "WHERE FieldID=" & fldID1 & _
                  " And Len(trim(location))=" & MyP.Length
          End If
          Set sumDataRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
          'Determine sum of current field for second unit area as comparison value
          If fldID2 > 301 Then  'IR is sum of agriculture and golf
            Select Case fldID2
              Case 302: irrFld = " OR FieldID=195"
              Case 303: irrFld = " OR FieldID=198"
              Case 305: irrFld = " OR FieldID=204"
              Case 306: irrFld = " OR FieldID=207"
              Case 307: irrFld = " OR FieldID=208"
              Case 308: irrFld = " OR FieldID=209"
              Case 309: irrFld = " OR FieldID=210"
              Case 311: irrFld = " OR FieldID=212"
              Case Else: irrFld = ""
            End Select
            sql = "SELECT Sum(Value) AS SumVal FROM [LastReport] " & _
                  "WHERE (FieldID=" & fldID2 - 10 & " OR FieldID=" & fldID2 - 20 & irrFld & _
                  ") And Len(trim(location))=" & MyP.length2
          ElseIf fldID2 = 4 And MyP.DataOpt2 < 3 Then
            sql = "SELECT Sum(Value) AS SumVal FROM [LastReport] " & _
                  "WHERE (FieldID=2 Or FieldID=3) " & _
                  "And Len(trim(location))=" & MyP.length2
          ElseIf MyP.UnitArea2 = "Aquifer" And MyP.DataOpt = 0 _
              And (fldID1 = 248 Or fldID1 = 249) Then
            'Comparing Aquifer w/ 1995 data; need to take Thermoelectric total
            sql = "SELECT Sum(Value) AS SumVal FROM [LastReport] " & _
                  "WHERE FieldID=" & fldID1 - 174 & _
                  " And Len(trim(location))=" & MyP.length2
            If fldID1 = 249 Then 'done with PO fields, skip PC
              fldRecCnt = fldRecCnt + 2
              fldRec.MoveNext
              fldRec.MoveNext
              fldRec2.MoveNext
              fldRec2.MoveNext
            End If
          Else
            sql = "SELECT Sum(Value) AS SumVal FROM [LastReport] " & _
                  "WHERE FieldID=" & fldID2 & _
                  " And Len(trim(location))=" & MyP.length2
          End If
          Set sumDataRec2 = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
          If IsNull(sumDataRec("SumVal")) Then temp1 = 0 Else temp1 = sumDataRec("SumVal")
          If IsNull(sumDataRec2("SumVal")) Then temp2 = 0 Else temp2 = sumDataRec2("SumVal")
          'Make adjustments as necessary for data kept as state totals
          If fldID1 = 40 Or fldID2 = 43 Then  'DO-withdrawals
            If MyP.DataOpt = 2 Or MyP.DataOpt = 4 Or MyP.DataOpt = 6 Then  'DO-withdrawals stored by state for first unit area
              If numAreas1 <> 0 Then temp1 = Round((temp1 / numAreas1), 2) Else temp1 = 0
            End If
            If MyP.DataOpt2 = 2 Or MyP.DataOpt2 = 4 Or MyP.DataOpt2 = 6 Then  'DO-withdrawals stored by state for second unit area
              If numAreas2 <> 0 Then temp2 = Round((temp2 / numAreas2), 2) Else temp2 = 0
            End If
          End If
          If fldID1 = 4 Then  'PS-pop served may be stored as state total for one or both areas
            If numAreas1 <> 0 Then
              If MyP.DataOpt = 3 Or MyP.DataOpt = 4 Then temp1 = Round((temp1 / numAreas1), 2)
            Else
              temp1 = 0
            End If
            If numAreas2 <> 0 Then
              If MyP.DataOpt2 = 3 Or MyP.DataOpt2 = 4 Then temp2 = temp2 / numAreas2
            Else
              temp2 = 0
            End If
          End If
          If Left(Cells(lastRow, 1), 3) = "Pop" Then
            Cells(lastRow, lastColumn).value = Round(temp1 - temp2, 3)
          Else
            Cells(lastRow, lastColumn).value = Round(temp1 - temp2, 2)
          End If
          If Not (fldID1 = 217 Or fldID1 = 220) Then
            Set xlRange = .FindNext(xlRange)
            While xlRange <> -1 And xlRange.Column = lastColumn
              Set xlRange = .FindNext(xlRange)
            Wend
          End If
          AtcoLaunch1.SendMonitorMessage "(PROGRESS " & ((fldRecCnt + 1) * 100 / numFldRecs) & ")"
        End If
        If Not fldRec.EOF Then
          fldRec.MoveNext
          fldRec2.MoveNext
        End If
      Next fldRecCnt
    End With
    rowCnt = rowCnt + 5
    sumDataRec.Close
    sumDataRec2.Close
    CatRec.Close
    Set xlRange = .UsedRange
    xlRange.Font.size = 8
    xlRange.Font.Name = "Times New Roman"
    Columns(1).ColumnWidth = 22
    Range(Columns(2), Columns(lastColumn)).AutoFit
    ActiveWorkbook.ActiveSheet.Name = MyP.UnitArea & " vs " & MyP.UnitArea2 & " - " & MyP.Year1Opt
  End With
x:
  AllFldRec.Close
  TotalPopRec.Close
  EndReport
End Sub

Private Sub CompYearsReport(RepPath As String, CatRec As Recordset)
Attribute CompYearsReport.VB_Description = "Performs "
' ##SUMMARY Performs "Compare Data for 2 Years" QA check with results written to Excel _
          output file.
' ##PARAM RepPath I String full pathname where report will be saved.
' ##PARAM catRec I Recordset containing user-selected categories.
' ##REMARKS Results are listed one table per category for each unit area in single table _
          in worksheet.
' ##HISTORY 5/31/2007, prhummel Updated module to produce State Total report (PR 13110)
  Dim fldRec As Recordset
  Dim sql As String
  Dim yearFields2 As String
  Dim i As Long
  Dim j As Long
  Dim sheetNum As Long
  Dim rowCnt As Long
  Dim rowCntHolder As Long
  Dim areaCnt As Long
  Dim catCnt As Long
  Dim fldRecCnt As Long
  Dim sigDigits As Long
  Dim xlRange As Excel.Range
  Dim divByZero As Boolean
  Dim nullCell As Boolean
  Dim message As String
  Dim tmpStr As String
  Dim lC1 As Double, lC2 As Double
  Dim TotVal1() As Double
  Dim TotVal2() As Double
  Dim curBnd As Long
  Dim curInd As Long
  Dim StateTotal As Boolean

  On Error GoTo x

  If Left(lstArea.RightItem(0), 3) = "000" Then 'doing report for State Totals
    StateTotal = True
    curBnd = 0
    curInd = 0
  Else
    StateTotal = False
  End If
  InitReport RepPath
  ActiveWorkbook.ActiveSheet.Name = MyP.Year1Opt & " vs " & MyP.Year2Opt & " - " & MyP.UnitArea
  rowCnt = 1
  sheetNum = 1
  j = 0
  CatRec.MoveLast
  CatRec.MoveFirst
  'Determine which field table is used for each of 2 years
  Set fldRec = MyP.stateDB.OpenRecordset("LastReport", dbOpenSnapshot)
  If fldRec("QualFlg") = 0 Then
    MyP.YearFields = "1995Fields1"
  ElseIf rdoAreaUnit(3) Then
    MyP.YearFields = "2000FieldsA"
  Else
    If fldRec("QualFlg") < 7 Then
      MyP.YearFields = "2000Fields" & fldRec("QualFlg")
    ElseIf fldRec("QualFlg") = 7 Then
      MyP.YearFields = "2000Fields" & 1
    ElseIf fldRec("QualFlg") = 8 Then
      MyP.YearFields = "2000Fields" & 5
    End If
  End If
  fldRec.MoveLast
  If fldRec("QualFlg") = 0 Then
    yearFields2 = "1995Fields1"
  ElseIf rdoAreaUnit(3) Then
    yearFields2 = "2000FieldsA"
  Else
    If fldRec("QualFlg") < 7 Then
      yearFields2 = "2000Fields" & fldRec("QualFlg")
    ElseIf fldRec("QualFlg") = 7 Then
      yearFields2 = "2000Fields" & 1
    ElseIf fldRec("QualFlg") = 8 Then
      yearFields2 = "2000Fields" & 5
    End If
  End If
  fldRec.Close
newSheet:
  j = j + 1
  With XLSheet
    If rowCnt < 10 Then
      Cells.Select
      Selection.Font.size = 8
      .Columns.ColumnWidth = 5.8
      Range(Columns(4), Columns(6)).NumberFormat = "#,###,##0.00"
      'Create header for report
      With Cells(rowCnt, 3)
        tmpStr = "Comparison of Water Use Data for " & _
            MyP.State & " between " & MyP.Year1Opt & " and " & MyP.Year2Opt
        If StateTotal Then _
          tmpStr = tmpStr & " -- State values calculated from " & MyP.UnitArea & " data"
        .value = tmpStr
        .HorizontalAlignment = xlHAlignLeft
        .Font.Bold = True
      End With
    End If
    rowCnt = rowCnt + 1
    If StateTotal Then
      Cells(rowCnt, 3).value = "(values in bold have one or more nulls in the source data)"
      Cells(rowCnt, 3).HorizontalAlignment = xlHAlignLeft
      rowCnt = rowCnt + 1
    End If
    'Add commentary re fields not included in comparison
    tmpStr = ""
    If IRinTwo And Not IRinTwo2 Then
      message = MyP.UnitArea2 & " areas in " & MyP.State & " keep irrigation as a total," & vbCrLf & _
          "while " & MyP.UnitArea & " areas divide irrigation into golf and crop use." & vbCrLf & _
          "By necessity, the " & MyP.UnitArea & " irrigation data will be totaled before" & vbCrLf & _
          "being compared to the " & MyP.UnitArea2 & " data."
    ElseIf IRinTwo2 And Not IRinTwo Then
      message = MyP.UnitArea & " areas in " & MyP.State & " keep irrigation as a total," & vbCrLf & _
          "while " & MyP.UnitArea2 & " areas divide irrigation into golf and crop use." & vbCrLf & _
          "By necessity, the " & MyP.UnitArea & " irrigation data will be totaled before" & vbCrLf & _
          "being compared to the " & MyP.UnitArea2 & " data."
    End If
    If MyP.DataOpt <> MyP.DataOpt2 Then
      If MyP.DataOpt = 2 And MyP.DataOpt2 <> 4 Or MyP.DataOpt2 <> 6 Then
        tmpStr = tmpStr & "Domestic withdrawals are kept as a state total;" & vbCrLf & _
        "therefore, they can not be compared to those from " & MyP.UnitArea2 & " areas."
      ElseIf MyP.DataOpt = 3 And MyP.DataOpt2 <> 4 Then
        tmpStr = tmpStr & "Public Supply - total population served" & vbCrLf & _
        "is kept as a state total; therefore, this datum can not be compared to " & MyP.UnitArea2 & " areas."
      ElseIf MyP.DataOpt = 4 Then
        If (MyP.DataOpt2 = 2 Or MyP.DataOpt2 = 6) Then
          tmpStr = tmpStr & "Public Supply - total population served" & vbCrLf & _
          "is kept as a state total; therefore, this datum can not be compared to " & MyP.UnitArea2 & " areas."
        ElseIf MyP.DataOpt2 = 3 Then
          tmpStr = tmpStr & "Domestic withdrawals are kept as a state total;" & vbCrLf & _
          "therefore, they can not be compared to those from " & MyP.UnitArea2 & " areas."
        Else
          tmpStr = tmpStr & "Domestic withdrawals are kept as a state total;" & vbCrLf & _
          "therefore, they can not be compared to those from " & MyP.UnitArea2 & " areas." & _
          "Public Supply - total population served is also kept as a state total;" & vbCrLf & _
          "therefore, this datum can not be compared to " & MyP.UnitArea2 & " areas either."
        End If
      ElseIf MyP.DataOpt = 5 And (MyP.DataOpt2 <> 3 And MyP.DataOpt2 <> 6) Then
        tmpStr = tmpStr & "Public Supply - total population served" & vbCrLf & _
        "is not broken out into GW and SW; therefore, this datum is compared as a total."
      ElseIf MyP.DataOpt = 6 Then
        If MyP.DataOpt2 = 2 Then
          tmpStr = tmpStr & "Public Supply - total population served" & vbCrLf & _
          "is not broken out into GW and SW; therefore, this datum is compared as a total."
        ElseIf (MyP.DataOpt2 = 3 Or MyP.DataOpt2 = 5) Then
          tmpStr = tmpStr & "Domestic withdrawals are kept as a state total;" & vbCrLf & _
          "therefore, they can not be compared to " & MyP.UnitArea2 & " areas."
        ElseIf (MyP.DataOpt2 = 1) Then
          tmpStr = tmpStr & "Domestic withdrawals are kept as a state total;" & vbCrLf & _
          "therefore, they can not be compared to " & MyP.UnitArea2 & " areas." & vbCrLf & _
          "Public Supply - total population served" & vbCrLf & _
          "is not broken out into GW and SW; therefore, this datum is compared as a total."
        End If
      End If
      If Len(tmpStr) > 0 Then
        If Len(message) > 0 Then
          message = message & vbCrLf & vbCrLf & "Also, for " & MyP.UnitArea & " areas in " & MyP.State & ", "
        Else
          message = "For " & MyP.UnitArea & " areas in " & MyP.State & ", "
        End If
        message = message & tmpStr
      End If
      'Now check 2nd set of data storage option
      If MyP.DataOpt2 = 2 And MyP.DataOpt <> 4 Or MyP.DataOpt <> 6 Then
        tmpStr = tmpStr & "Domestic withdrawals are kept as a state total;" & vbCrLf & _
        "therefore, they can not be compared to those from " & MyP.UnitArea2 & " areas."
      ElseIf MyP.DataOpt2 = 3 And MyP.DataOpt <> 4 Then
        tmpStr = tmpStr & "Public Supply - total population served" & vbCrLf & _
        "is kept as a state total; therefore, this datum can not be compared to " & MyP.UnitArea2 & " areas."
      ElseIf MyP.DataOpt2 = 4 Then
        If (MyP.DataOpt = 2 Or MyP.DataOpt = 6) Then
          tmpStr = tmpStr & "Public Supply - total population served" & vbCrLf & _
          "is kept as a state total; therefore, this datum can not be compared to " & MyP.UnitArea2 & " areas."
        ElseIf MyP.DataOpt = 3 Then
          tmpStr = tmpStr & "Domestic withdrawals are kept as a state total;" & vbCrLf & _
          "therefore, they can not be compared to those from " & MyP.UnitArea2 & " areas."
        Else
          tmpStr = tmpStr & "Domestic withdrawals are kept as a state total;" & vbCrLf & _
          "therefore, they can not be compared to those from " & MyP.UnitArea2 & " areas." & _
          "Public Supply - total population served is also kept as a state total;" & vbCrLf & _
          "therefore, this datum can not be compared to " & MyP.UnitArea2 & " areas either."
        End If
      ElseIf MyP.DataOpt2 = 5 And (MyP.DataOpt <> 3 And MyP.DataOpt <> 6) Then
        tmpStr = tmpStr & "Public Supply - total population served" & vbCrLf & _
        "is not broken out into GW and SW; therefore, this datum is compared as a total."
      ElseIf MyP.DataOpt2 = 6 Then
        If MyP.DataOpt = 2 Then
          tmpStr = tmpStr & "Public Supply - total population served" & vbCrLf & _
          "is not broken out into GW and SW; therefore, this datum is compared as a total."
        ElseIf (MyP.DataOpt = 3 Or MyP.DataOpt = 5) Then
          tmpStr = tmpStr & "Domestic withdrawals are kept as a state total;" & vbCrLf & _
          "therefore, they can not be compared to " & MyP.UnitArea2 & " areas."
        ElseIf MyP.DataOpt = 1 Then
          tmpStr = tmpStr & "Domestic withdrawals are kept as a state total;" & vbCrLf & _
          "therefore, they can not be compared to " & MyP.UnitArea2 & " areas." & vbCrLf & _
          "Public Supply - total population served" & vbCrLf & _
          "is not broken out into GW and SW; therefore, this datum is compared as a total."
        End If
      End If
      If Len(tmpStr) > 0 Then
        If Len(message) > 0 Then
          message = message & vbCrLf & vbCrLf & "Also, for " & MyP.UnitArea & " areas in " & MyP.State & ", "
        Else
          message = "For " & MyP.UnitArea & " areas in " & MyP.State & ", "
        End If
        message = message & tmpStr
      End If
    End If
    'Loop thru areas
    For areaCnt = 0 To UBound(LocnArray, 2)
      'Check to see if returning here after pasting framework of 1st area repeatedly
      If j > UBound(LocnArray, 2) And UBound(LocnArray, 2) <> 0 Then
        areaCnt = areaCnt + 1
        j = 0
      End If
      'Check to see if we need to move to the next sheet to fill more data
      If rowCnt > 65200 Then
        j = i
        Cells(1, 1).Select
        sheetNum = sheetNum + 1
        ActiveWorkbook.Sheets(sheetNum).Select
        Set XLSheet = ActiveWorkbook.Sheets(sheetNum)
        rowCnt = 2
      End If
      If StateTotal Then
        rowCnt = 3 'only one table produced, reset row counter
        curInd = 0 'reset state total array index
      End If
      CatRec.MoveFirst
      rowCnt = rowCnt + 1
      'Divide fields of report into categories
      For catCnt = 0 To CatRec.RecordCount - 1
        Select Case NextPipeCharacter(AtcoLaunch1.ComputeRead)
          Case "P"
            While NextPipeCharacter(AtcoLaunch1.ComputeRead) <> "R"
              DoEvents
            Wend
          Case "C"
            ImportDone = True
              MyMsgBox.Show "The Compare Data for 2 Years report was cancelled.", _
                  "Report interrupted", "+-&OK"
            Err.Raise 999
        End Select
        AtcoLaunch1.SendMonitorMessage "(PROGRESS " & _
            (((catCnt / CatRec.RecordCount) + areaCnt) * 100 / (UBound(LocnArray, 2) + 1)) & ")"
        divByZero = False
        'Create header
        If areaCnt = 0 Then
          Cells(rowCnt, 1).value = MyP.UnitArea
          Cells(rowCnt, 2).value = "Category"
          Cells(rowCnt, 3).value = "Field"
          Cells(rowCnt, 4).value = MyP.Year1Opt
          Cells(rowCnt, 5).value = MyP.Year2Opt
          Cells(rowCnt, 6).value = "Change, in units"
          Cells(rowCnt, 7).value = "Change, in percent"
          Rows(rowCnt).NumberFormat = "###0"
          'Format the header
          With Range(Cells(rowCnt, 1), Cells(rowCnt, 7))
            .NumberFormat = "####"
            .Borders(xlTop).LineStyle = xlContinuous
            .Borders(xlTop).Weight = xlThin
          End With
          With Range(Cells(rowCnt, 1), Cells(rowCnt, 7)).Borders(xlBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
          End With
        End If
        rowCnt = rowCnt + 1
        'Use SQL language to create recordset of Data Dictionary fields within current
        ' category that will be compared in report.  Different "Field" tables that
        ' comprise the data dictionary are joined.
        ' For example:
        '  SELECT [AllFields].ID, [AllFields].Description, [1995Fields1].Excluded FROM ([1995Fields1]
        '  INNER JOIN [2000Fields1] ON [1995Fields1].FieldID = [2000Fields1].FieldID)
        '    INNER JOIN [AllFields] ON [2000Fields1].FieldID = [AllFields].ID
        '  Where [AllFields].CategoryID = 4
        '  ORDER BY [AllFields].ID
        If CatRec("ID") = 16 Then  'need special query for Total Irrigation
          sql = "SELECT [AllFields].ID, [AllFields].Description, [" & yearFields2 & "].Excluded" & _
              " FROM [" & yearFields2 & _
              "] INNER JOIN [AllFields] ON [" & yearFields2 & "].FieldID = [AllFields].ID" & _
              " Where [AllFields].CategoryID=" & CatRec("ID") & _
              " ORDER BY [AllFields].ID"
        ElseIf MyP.YearFields = yearFields2 Then
          sql = "SELECT [AllFields].ID, [AllFields].Description, [" & MyP.YearFields & "].Excluded" & _
              " FROM [" & MyP.YearFields & _
              "] INNER JOIN [AllFields] ON [" & MyP.YearFields & "].FieldID = [AllFields].ID" & _
              " Where [AllFields].CategoryID=" & CatRec("ID") & _
              " ORDER BY [Allfields].ID;"
        Else
          sql = "SELECT [AllFields].ID, [AllFields].Description, [" & MyP.YearFields & "].Excluded" & _
              " FROM ([" & MyP.YearFields & "] INNER JOIN [" & yearFields2 & _
              "] ON [" & MyP.YearFields & "].FieldID = [" & yearFields2 & "].FieldID)" & _
              " INNER JOIN [AllFields] ON [" & yearFields2 & "].FieldID = [AllFields].ID" & _
              " Where [AllFields].CategoryID=" & CatRec("ID") & _
              " ORDER BY [AllFields].ID"
        End If
        Set fldRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
        fldRec.MoveLast
        fldRec.MoveFirst
        If areaCnt = 0 Then 'size state total arrays 1st time through
          ReDim Preserve TotVal1(1, curBnd + fldRec.RecordCount)
          ReDim Preserve TotVal2(1, curBnd + fldRec.RecordCount)
          curBnd = UBound(TotVal1, 2)
        End If
        'Loop through fields in current category
        For fldRecCnt = 0 To fldRec.RecordCount - 1
          If StateTotal Then
            Cells(rowCnt, 1).value = MyP.State
          ElseIf rdoID(1) Then
            Cells(rowCnt, 1).value = "'" & LocnArray(AreaID, areaCnt)
          Else
            Cells(rowCnt, 1).value = LocnArray(AreaID, areaCnt)
          End If
          'Find values of present field for both years and compare
          If areaCnt = 0 Then
            Cells(rowCnt, 2).value = CatRec("Description")
            Cells(rowCnt, 3).value = fldRec("Description")
            If InStr(1, LCase(fldRec("Description")), "umber of") > 0 Then
              sigDigits = 0
              Range(Cells(rowCnt, 4), Cells(rowCnt, 6)).NumberFormat = "###0"
            ElseIf InStr(1, LCase(fldRec("Description")), "opulat") Then
              sigDigits = 2
              Range(Cells(rowCnt, 4), Cells(rowCnt, 6)).NumberFormat = "###0.000"
            Else
              sigDigits = 2
              Range(Cells(rowCnt, 4), Cells(rowCnt, 6)).NumberFormat = "###0.00"
            End If
          End If
          nullCell = False
          TwoYears = False
          'Get values
          lC1 = EvalArray(fldRec("ID"), areaCnt)
          If NoRec Then
            If Not StateTotal Then Cells(rowCnt, 4).value = Null
            nullCell = True
            If TotVal1(1, curInd) = 1 Then TotVal1(1, curInd) = 2
          Else
            If Not StateTotal Then Cells(rowCnt, 4).value = Round(lC1, 3)
            If TotVal1(1, curInd) = 0 Then TotVal1(1, curInd) = 1
            TotVal1(0, curInd) = TotVal1(0, curInd) + lC1
          End If
          TwoYears = True
          lC2 = EvalArray(fldRec("ID"), areaCnt)
          If NoRec Then
            If Not StateTotal Then Cells(rowCnt, 5).value = Null
            nullCell = True
            If TotVal2(1, curInd) = 1 Then TotVal2(1, curInd) = 2
          Else
            If Not StateTotal Then Cells(rowCnt, 5).value = Round(lC2, 3)
            If TotVal2(1, curInd) = 0 Then TotVal2(1, curInd) = 1
            TotVal2(0, curInd) = TotVal2(0, curInd) + lC2
          End If
          If StateTotal And areaCnt = UBound(LocnArray, 2) Then
            'last area in state total compilation
            If TotVal1(1, curInd) = 0 Then 'no data for 1st year
              Cells(rowCnt, 4).value = Null
            Else
              If TotVal1(1, curInd) = 2 Then Cells(rowCnt, 4).Font.Bold = True 'total contains nulls
              Cells(rowCnt, 4).value = Round(TotVal1(0, curInd), 3)
            End If
            If TotVal2(1, curInd) = 0 Then 'no data for 2nd year
              Cells(rowCnt, 5).value = Null
            Else
              If TotVal2(1, curInd) = 2 Then Cells(rowCnt, 5).Font.Bold = True 'total contains nulls
              Cells(rowCnt, 5).value = Round(TotVal2(0, curInd), 3)
            End If
            If TotVal1(1, curInd) = 0 And TotVal2(1, curInd) = 0 Then
              Cells(rowCnt, 6).value = "No Data"
            ElseIf TotVal1(1, curInd) = 0 Then
              Cells(rowCnt, 6).value = "No " & MyP.Year1Opt & " Data"
            ElseIf TotVal2(1, curInd) = 0 Then
              Cells(rowCnt, 6).value = "No " & MyP.Year2Opt & " Data"
            Else 'report difference and %change
              If TotVal1(1, curInd) = 2 Or TotVal2(1, curInd) = 2 Then
                'totals contain null values
                Range(Cells(rowCnt, 6), Cells(rowCnt, 7)).Font.Bold = True
              End If
              Cells(rowCnt, 6).value = Round(TotVal2(0, curInd) - TotVal1(0, curInd), 3)
              If TotVal1(0, curInd) = 0 Then
                If TotVal2(0, curInd) = 0 Then
                  Cells(rowCnt, 7).value = 0
                Else
                  Cells(rowCnt, 7).value = "****"
                  Cells(rowCnt, 7).HorizontalAlignment = xlHAlignCenter
                  divByZero = True
                End If
              Else
                Cells(rowCnt, 7).value = Round((TotVal2(0, curInd) - TotVal1(0, curInd)) * 100 / TotVal1(0, curInd), 1)
              End If
            End If
          ElseIf Not StateTotal Then 'report difference and %change for this area
            If Not nullCell Then
              Cells(rowCnt, 6).value = Round(lC2 - lC1, 3)
              If lC1 = 0 Then
                If lC2 = 0 Then
                  Cells(rowCnt, 7).value = 0
                Else
                  Cells(rowCnt, 7).value = "****"
                  Cells(rowCnt, 7).HorizontalAlignment = xlHAlignCenter
                  divByZero = True
                End If
              Else
                Cells(rowCnt, 7).value = Round((lC2 - lC1) * 100 / lC1, 1)
              End If
            Else
              If NoRec Then
                If Cells(rowCnt, 4).value = "" Then
                  Cells(rowCnt, 6).value = "No Data"
                Else
                  Cells(rowCnt, 6).value = "No " & MyP.Year2Opt & " Data"
                End If
              Else
                Cells(rowCnt, 6).value = "No " & MyP.Year1Opt & " Data"
              End If
              Cells(rowCnt, 7).value = ""
            End If
          End If
          lC1 = 0
          lC2 = 0
          If areaCnt = 0 Then
            If fldRec("Excluded") > 1 Then _
                Range(Cells(rowCnt, 4), Cells(rowCnt, 6)).NumberFormat = "##,##0"
          End If
          rowCnt = rowCnt + 1
          curInd = curInd + 1
          fldRec.MoveNext
        Next fldRecCnt
        fldRec.Close
        CatRec.MoveNext
        If divByZero Then
          Cells(rowCnt, 3).value = "**** indicates an infinite " & _
              "% increase from the previous data value of zero."
          Cells(rowCnt, 3).Font.Superscript = True
        Else
          Cells(rowCnt, 3).value = ""
        End If
        rowCnt = rowCnt + 1
      Next catCnt
      If areaCnt = 0 And UBound(LocnArray, 2) > 0 And Not StateTotal Then
        rowCntHolder = rowCnt
        Set xlRange = Range(Cells(1, 1), Cells(rowCnt - 1, 7))
        xlRange.Copy
        For i = j To UBound(LocnArray, 2)
          If rowCnt + (i - j) * (rowCnt - 1) > 65200 Then
            rowCnt = 1
            'consideration made for writing, not pasting, first HUC on successive sheets
            If sheetNum > 1 Then i = i + 1
            sheetNum = sheetNum + 1
            ActiveWorkbook.ActiveSheet.Name = LocnArray(0, j - 1) & "-" & LocnArray(0, i - 1)
            ActiveWorkbook.Worksheets.Add.Move after:=Worksheets(Worksheets.Count)
            ActiveWorkbook.Sheets(sheetNum).Select
            ActiveWorkbook.ActiveSheet.Name = LocnArray(0, i) & "-" & LocnArray(0, UBound(LocnArray, 2))
            Set XLSheet = ActiveWorkbook.Sheets(sheetNum)
            j = i
            GoTo newSheet
          End If
          'consideration made for writing, not pasting, first HUC on successive sheets
          If Not (sheetNum > 1 And i = UBound(LocnArray, 2) - 1) Then
            Cells(rowCnt + (i - j) * (rowCnt - 1), 1).Select
            .Paste
          End If
        Next i
        j = i
        rowCnt = rowCntHolder
        sheetNum = 1
        'Get back to 2nd HUC/aquifer
        Set XLSheet = ActiveWorkbook.Sheets(1)
        XLSheet.Select
        GoTo newSheet
      End If
      rowCnt = rowCnt + 1
    Next areaCnt
  End With
  CatRec.Close
  For i = 1 To ActiveWorkbook.Sheets.Count
    Set XLSheet = ActiveWorkbook.Sheets(i)
    With XLSheet
      .Select
      Set xlRange = .UsedRange
      xlRange.Select
      Selection.Font.Name = "Times New Roman"
      Selection.Columns(7).NumberFormat = "##0.0"
      For j = 1 To 7
        Columns(j).AutoFit
      Next j
      Cells(1, 3).Select
      Clipboard.Clear
    End With
  Next i
  ActiveWorkbook.Sheets(1).Select
x:
  TwoYears = True
  TotalPopRec.Close
  EndReport
End Sub

Private Sub InitReport(RepPath As String)
Attribute InitReport.VB_Description = "Initiates instance of status monitor and Excel workbook for report creation."
' ##SUMMARY Initiates instance of status monitor and Excel workbook for report creation.
' ##PARAM RepPath I String full pathname where report will be saved.
' ##REMARKS Also determines required fields for selected state, and checks for selected _
          areas with no data, which should never happen unless the database has been _
          corrupted.  It also opens the workbook containing the report headers that _
          will be used in the Excel reports.
  Dim sql As String
  Dim fileTitle As String
  Dim missingAreas As String
  Dim i As Long
  Dim j As Long
  Dim OutFile As Long
  Dim myRec As Recordset

  'open the ATCo status bar
  AtcoLaunch1.SendMonitorMessage "(OPEN AWUDS)"
  AtcoLaunch1.SendMonitorMessage "(BUTTOFF DETAILS)"
  AtcoLaunch1.SendMonitorMessage "(BUTTON CANCEL)"
  AtcoLaunch1.SendMonitorMessage "(BUTTON PAUSE)"
  AtcoLaunch1.SendMonitorMessage "(MSG1 Creating file for " & MainOpt(MyP.UserOpt - 1).Caption _
      & " for " & MyP.State & ")"
  AtcoLaunch1.SendMonitorMessage "(PROGRESS 0)"
  
  If Not NationalDB Then
    'Determine which group of special fields are required for this state
    sql = "SELECT Required FROM [state] WHERE state_cd = '" & MyP.stateCode & "'"
    Set myRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    ReqSt = myRec("Required")
    myRec.Close
  Else
    ReqSt = 0
  End If

  On Error GoTo x

  'Open the Excel workbook and worksheets
  'Create new sheet for report
  Set XLApp = New Excel.Application

  Set XLBook = Excel.Workbooks.Add
  XLBook.Activate
  If rdoAreaUnit(1) And MainOpt(3) Then 'need to add 4 extra sheets to fit all HUC-8's
    XLApp.SheetsInNewWorkbook = 5
    While ActiveWorkbook.Worksheets.Count < 5
      With ActiveWorkbook
        Worksheets.Add.Move after:=Worksheets(Worksheets.Count)
      End With
    Wend
  Else
    If MainOpt(4) Then
      XLApp.SheetsInNewWorkbook = 2
      Worksheets.Add.Move after:=Worksheets(Worksheets.Count)
    Else
      XLApp.SheetsInNewWorkbook = 1
    End If
    While ActiveWorkbook.Worksheets.Count > XLApp.SheetsInNewWorkbook
      With ActiveWorkbook
        Application.DisplayAlerts = False
        Worksheets(Worksheets.Count).Delete
        Application.DisplayAlerts = True
      End With
    Wend
  End If
  XLBook.SaveAs RepPath, XLFileFormatNum
  Set XLSheet = ActiveWorkbook.Worksheets(1)
  Set AllFldRec = MyP.stateDB.OpenRecordset(FieldTable, dbOpenSnapshot)
  If MyP.UserOpt <> 9 Then
    sql = "SELECT * FROM [" & MyP.AreaTable & "]" & _
          " WHERE FieldID=1 And Date=" & MyP.Year1Opt & Areas & _
          " ORDER BY Location;"
    Set TotalPopRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot)
    TotalPopRec.MoveLast
    TotalPopRec.MoveFirst
    If UBound(LocnArray, 2) > TotalPopRec.RecordCount Then
      'There are selected areas with no data. This should never happen unless DB corrupted.
      j = 1
      missingAreas = ""
      For i = 0 To UBound(LocnArray, 2)
        If LocnArray(0, i) <> Left(TotalPopRec("Location"), MyP.Length) Then
          missingAreas = missingAreas & LocnArray(0, i) & vbTab & LocnArray(1, i) & vbCrLf
        Else
          TotalPopRec.MoveNext
          If TotalPopRec.EOF Then TotalPopRec.MovePrevious
        End If
      Next i
      If TableName = "HUC8" Then i = 1 Else i = 3
      If MainOpt(9) Then
        sql = MyP.Year1Opt & " and " & MyP.Year2Opt
      Else
        sql = MyP.Year1Opt
      End If
      OutFile = FreeFile
      fileTitle = ReportPath & "Missing_" & MyP.UnitArea & "s"
      i = 1
      While Len(Dir(fileTitle & ".txt")) > 0
        i = i + 1
        If i > 99 Then
          fileTitle = Left(fileTitle, Len(fileTitle) - 4)
        ElseIf i > 9 Then
          fileTitle = Left(fileTitle, Len(fileTitle) - 3)
        ElseIf i > 2 Then
          fileTitle = Left(fileTitle, Len(fileTitle) - 2)
        End If
        fileTitle = fileTitle & "-" & i
      Wend
      fileTitle = fileTitle & ".txt"
      Open fileTitle For Output As OutFile
      MyMsgBox.Show _
          "There is not data for all of the selected " & MyP.UnitArea & _
          " areas in" & vbCrLf & MyP.State & " for " & sql & _
          ".  The missing areas are listed in the file " & vbCrLf & _
          "'" & fileTitle & "'", _
          "Data Inadequacy Notification", "+-&OK"
      Print #OutFile, missingAreas
      Close OutFile
    End If
  End If
  'Look for and open the header worksheet
  If MyP.UserOpt = 4 Or MyP.UserOpt = 10 Then Exit Sub  'These reports do not use headerfile
LookForFile:
  fileTitle = AwudsDataPath & HeaderFile
  If Len(Dir(fileTitle)) = 0 Then
    fileTitle = ExePath & HeaderFile
    If Len(Dir(fileTitle)) = 0 Then
      fileTitle = GetSetting("AWUDS", "Defaults", "HeaderFile", "")
      If Len(fileTitle) = 0 Then GoTo AskForHeaderFile
      If Len(Dir(fileTitle)) = 0 Then
AskForHeaderFile:
        Select Case MyMsgBox.Show("The file " & HeaderFile & " was not found." & vbCr _
                        & "It was expected to be in '" & AwudsDataPath & "'" & vbCrLf _
                        & "It is available for download along with AWUDS data files.", _
                          "AWUDS Report Header File", "+&Browse for file", "&Retry", "-&Cancel")
          Case 1:
            With frmDialog.cdlg
              .fileTitle = AwudsDataPath & HeaderFile
              .DialogTitle = "Locate Header File"
              .ShowOpen
              fileTitle = .fileTitle
              If fileTitle = "" Then GoTo AskForHeaderFile
              If Len(Dir(fileTitle)) = 0 Then GoTo AskForHeaderFile
            End With
          Case 2: GoTo LookForFile
          Case 3: Exit Sub
        End Select
      End If
    End If
  End If
  Set HeaderBook = XLApp.Workbooks.Open(AwudsDataPath & HeaderFile, , False)
  If MyP.UserOpt = 3 Then i = 2 Else i = 3
  Set HeaderSheet = HeaderBook.Worksheets(MyP.UserOpt - i)
  XLSheet.Activate
  
  Exit Sub
x:
  MyMsgBox.Show "Unable to access the file " & RepPath & "." & vbCrLf _
      & "Make sure this file is not presently open." & vbCrLf & Err.Description, _
      "AWUDS Report Error", "+&OK"
End Sub

' ##SUMMARY Closes Excel (or atleast it should!) and status monitor.
' ##REMARKS Excel dies hard and, despite exhaustive efforts, will not terminate until the _
          current instance of AWUDS is closed.
Private Sub EndReport()
Attribute EndReport.VB_Description = "Closes Excel (or atleast it should!) and status monitor."
 
  On Error Resume Next
  AtcoLaunch1.SendMonitorMessage "(CLOSE)"
  
  Set HeaderSheet = Nothing
  HeaderBook.Activate
  HeaderBook.Close False
  Set HeaderBook = Nothing
  Set XLSheet = Nothing
  XLBook.Activate
  XLBook.Close True
  Set XLBook = Nothing
  XLApp.Application.Quit
  Set XLApp = Nothing
  Excel.Application.Quit
End Sub

Private Sub GetHeader(CatID As Long)
Attribute GetHeader.VB_Description = "Retrieves appropriate report header from appropriate Excel file."
' ##SUMMARY Retrieves appropriate report header from appropriate Excel file.
' ##REMARKS The report headers are tagged with the category id using the _
          syntax "_ID_" where ID is the category id number. For example, _
          the string (tag) "_6_7_8_9_" is used identify the report header _
          to be used for categories 6, 7, 8, and 9. GetHeader will copy the _
          header following the tag. The header is the set of Excel _
          rows containing text up to but not including the first _
          blank row following the tag. The tag is located in column A.
' ##PARAM CatID I Integer identifying the category for which the header will be retrieved.
  Dim headerRange As Excel.Range
  
  HeaderSheet.Activate
  Set headerRange = HeaderSheet.UsedRange.Find _
      ("_" & CatID & "_", , , , xlByRows, xlNext, False)
  HeaderRows = headerRange.CurrentRegion.Rows.Count
  headerRange.CurrentRegion.Copy
End Sub

Private Sub NationalExport(BookName As String, SelCatRec As Recordset)
' ##SUMMARY Creates new workbook file for National database export and saves it to file.
' ##PARAM SelCatRec I Recordset specifies IDs of selected categories and determines _
          number of worksheets in workbook.
' ##PARAM BookName I Full path and filename of report.
' ##RETURNS Excel workbook object.
  Dim NewBook As Excel.Workbook 'new Excel workbook object
  Dim i As Long
  Dim j As Long
  Dim opt As Variant
  Dim stCode As String
  Dim stName As String
  Dim tabName As String
  Dim dataRec As Recordset
  Dim sql As String
  Dim CurRow As Long
  Dim ExpType As String
  
  'Open Status Monitor
  AtcoLaunch1.SendMonitorMessage "(OPEN AWUDS)"
  AtcoLaunch1.SendMonitorMessage "(BUTTON CANCEL)"
  AtcoLaunch1.SendMonitorMessage "(BUTTON PAUSE)"
  
  On Error GoTo x
  
  Set XLApp = New Excel.Application
  Set NewBook = Excel.Workbooks.Add
  Application.SheetsInNewWorkbook = SelCatRec.RecordCount
  NewBook.SaveAs BookName, XLFileFormatNum
  'set the unit area type for the export
  ExpType = LCase(Left(MyP.AreaTable, 2))
  SetNatExpHeaders NewBook, SelCatRec, ExpType
  MyP.StateDBClose 'close national database
  CurRow = 5
  For i = 0 To lstStates.ListCount - 1 'export each state in state list
    stCode = Trim(lstStates.ItemData(i))
    If Len(stCode) < 2 Then
      stCode = "0" & stCode
    End If
    stName = Left(lstStates.List(i), 1) & LCase(Mid(lstStates.List(i), 2))
    If InStr(stName, " ") > 0 Then 'convert first characters after a blank to upper case
      For j = 2 To Len(stName)
        If Mid(stName, j - 1, 1) = " " Then Mid(stName, j, 1) = UCase(Mid(stName, j, 1))
      Next j
    End If
    MyP.StateStuff stName, stCode
    'get quality code for this state's data
    tabName = TableName & "Data"
    sql = "SELECT " & tabName & ".* FROM " & tabName & " WHERE " & Years & Areas & _
        " ORDER BY [" & tabName & "].Date;"
    Set dataRec = MyP.stateDB.OpenRecordset(sql, dbOpenSnapshot, False)
    If dataRec.RecordCount > 0 Then 'data exists for this state and year
      If rdoAreaUnit(3) Then  'aquifer
        opt = "A"
      ElseIf dataRec("QualFlg") < 7 Then
        opt = dataRec("QualFlg")
      ElseIf dataRec("QualFlg") = 7 Then
        opt = 1
      ElseIf dataRec("QualFlg") = 8 Then
        opt = 5
      End If
      MyP.YearFields = "2000Fields" & opt
      dataRec.Close
      'set areas for this state, use SetAreas call with NationalDB=False
      NationalDB = False
      SetAreas
      NationalDB = True
      If Len(Cats) > 0 Then sql = " WHERE " & Mid(Cats, 6) Else sql = ""
      sql = sql & " ORDER BY [" & CatTable & "].ID;"
      Set SelCatRec = MyP.stateDB.OpenRecordset("SELECT * From [" & CatTable & "]" & sql, dbOpenSnapshot)
      SelCatRec.MoveLast
      SelCatRec.MoveFirst
      CreateTable
      FillNationalExport NewBook, SelCatRec, LocnArray, CurRow, ExpType
      If ExpType = "done" Then GoTo x
    End If
    MyP.StateDBClose
  Next i
  ImportDone = True
x:
  'Reopen National Database
  MyP.StateStuff "United States", "Nation"
  
  If Err.Number = 999 Then _
      AtcoLaunch1.SendMonitorMessage "(MSG1 User Canceled Export)"
  If ImportDone = False Then
    MyMsgBox.Show "The import was not successful." & vbCrLf & _
        "Make sure the destination file is not currently open.", _
        "Import not successful", "+-&OK"
  End If
  NewBook.Close True
  Set NewBook = Nothing
  XLApp.Quit
  AtcoLaunch1.SendMonitorMessage "(CLOSE)"

End Sub

Private Sub txtYear_Change()
  If IsNumeric(txtYear.value) Then
    If (txtYear.value > 1799 And txtYear.value < 2101) _
        And Len(Dir(txtCurFile)) > 0 Then
      cmdImport.Enabled = True
    Else
      cmdImport.Enabled = False
    End If
  End If
End Sub

Private Sub SetExcelProps()
' ##SUMMARY Sets Excel file format and extension
  Dim lXLApp As Excel.Application
  
  On Error GoTo ErrProc
  '
  ' Determine which Excel the user is using.
  '
  Set lXLApp = New Excel.Application
  ' Less than 12 is Excel 97-2003
  If Val(lXLApp.Version) < 12 Then
    ' Code -4143 is used to identify Excel 97-2003.
    XLFileFormatNum = -4143
    XLFileFilter = "(*.xls)|*.xls"
    XLFileExt = ".xls"
  Else
    ' Code 51 represents the enumeration for a macro-free
    ' Excel 2007 Workbook (.xlsx).
    XLFileFormatNum = 51
    XLFileFilter = "(*.xlsx)|*.xlsx"
    XLFileExt = ".xlsx"
  End If

  lXLApp.Quit
  Set lXLApp = Nothing
  Exit Sub
ErrProc:
  MsgBox Err.Description

End Sub
