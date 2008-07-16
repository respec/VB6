VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmDialog 
   ClientHeight    =   2556
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2556
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   120
      Top             =   120
      _ExtentX        =   699
      _ExtentY        =   699
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Copyright 2000 by AQUA TERRA Consultants

' ##MODULE_NAME frmDialog
' ##MODULE_DATE June 19, 2000
' ##MODULE_AUTHOR Mark Gray of AQUA TERRA CONSULTANTS
' ##MODULE_SUMMARY This form provides an interface that allows the user to _
          browse for requested files.


