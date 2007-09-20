VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtFName 
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cdlFName 
      Left            =   2880
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      InitDir         =   "c:\nss\ROI-MN"
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtFName 
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblRec 
      Caption         =   "Enter name of REC file:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblRho 
      Caption         =   "Enter name of RHO file:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RhoBin As String, RecBin As String

Private Sub cmdBrowse_Click(Index As Integer)
  cdlFName.ShowOpen
  txtFName(Index).Text = cdlFName.Filename
End Sub

Private Sub cmdConvert_Click()
  If Len(txtFName(1).Text) > 0 Then
    RhoBin = MatrixToBinary(txtFName(1).Text, 1, "MN")
  End If
  If Len(txtFName(2).Text) > 0 Then
    RecBin = MatrixToBinary(txtFName(2).Text, 2, "MN")
  End If
End Sub

Public Function MatrixToBinary(FName As String, MatrixType As Integer, StName As String) As String
  'FName - flat file containing matrix data
  'MatrixType - 1 - RHO, 2 - REC/MCon (coincident years)
  Dim FirstVal As Long, FldLen As Long, ipos As Long, Funit As Long
  Dim Maxnv As Long, nv As Long
  Dim RVals() As Single, TotVal As Single
  Dim IVals() As Integer
  Dim Fstr As String, istr As String, Rstr As String, FType(2) As String

  Fstr = WholeFileString(FName)
  FirstVal = 0
  While FirstVal = 0 And Len(Fstr) > 0
    istr = StrSplit(Fstr, vbCrLf, "")
    If IsNumeric(istr) Then
      If MatrixType = 1 Then 'RHO file should have first value of 1.0
        If CSng(istr) = 1# Then
          FirstVal = 1
          FldLen = Len(istr)
        End If
      Else 'assume first valid in REC file
        FirstVal = 1
      End If
    End If
  Wend
  If FirstVal = 1 Then
    Maxnv = 500000
    If MatrixType = 1 Then
      ReDim Preserve RVals(Maxnv)
      RVals(1) = 1# '1st RHO value is always 1.0
      TotVal = RVals(1)
    Else
      ReDim Preserve IVals(Maxnv)
      IVals(1) = CInt(istr)
      TotVal = IVals(1)
    End If
    nv = 1
    While Len(Fstr) > 0 'process rest of file
      Rstr = StrSplit(Fstr, vbCrLf, "") 'next record
      If MatrixType = 1 Then
        ipos = 1
        While ipos < Len(Rstr)
          nv = nv + 1
          RVals(nv) = Mid(Rstr, ipos, FldLen)
          TotVal = TotVal + RVals(nv)
          ipos = ipos + FldLen
        Wend
      Else
        While Len(Rstr) > 0
          istr = StrSplit(Rstr, " ", "")
          nv = nv + 1
          IVals(nv) = CInt(istr)
          TotVal = TotVal + IVals(nv)
        Wend
      End If
    Wend
    'write out the binary version
    FType(1) = ".rho"
    FType(2) = ".rec"
    Fstr = StName & FType(MatrixType) & ".bin"
    Funit = FreeFile(0)
    Open Fstr For Binary As #Funit
    Put #Funit, , MatrixType
    Put #Funit, , nv
    Put #Funit, , TotVal
    For ipos = 1 To nv
      If MatrixType = 1 Then
        Put #Funit, , RVals(ipos)
      Else
        Put #Funit, , IVals(ipos)
      End If
    Next ipos
    Close #Funit
  Else
    Fstr = ""
  End If
  MatrixToBinary = Fstr

End Function

Private Sub cmdRead_Click()
  Dim Rho As Variant, Rec As Variant
  If Len(RhoBin) > 0 Then
    GetMatrixBinary RhoBin, Rho
  End If
  If Len(RecBin) > 0 Then
    GetMatrixBinary RecBin, Rho
  End If
End Sub

Private Sub GetMatrixBinary(BinFile As String, MatVar As Variant)
  Dim i As Long, j As Long, Fun As Long, Fnv As Long, nv As Long
  Dim FileType As Integer '1 - RHO, 2 - REC
  Dim iVal As Integer
  Dim rval As Single
  Dim FTotVal As Single, TotVal As Single

  Fun = FreeFile(0)
  Open BinFile For Binary As #Fun
  Get #Fun, , FileType
  Get #Fun, , Fnv
  Get #Fun, , FTotVal
  'determine size of matrix array
  i = Fix(Sqr(CDbl(2 * Fnv)))
  ReDim MatVar(i, i)
  If FileType = 1 Then 'RHO values are reals
    For i = 1 To UBound(MatVar)
      For j = 1 To i
        Get #Fun, , rval
        MatVar(i, j) = rval
        nv = nv + 1
        TotVal = TotVal + rval
      Next j
    Next i
  Else 'REC values are integers
    For i = 1 To UBound(MatVar)
      For j = 1 To i
        Get #Fun, , iVal
        MatVar(i, j) = iVal
        nv = nv + 1
        TotVal = TotVal + iVal
      Next j
    Next i
  End If
  If Fnv <> nv Or Abs(FTotVal - TotVal) > 0.001 Then 'didn't read expected number or total of values
    MsgBox "PROBLEM: read different count or total of values than expected from binary matrix file!", vbExclamation
  End If
End Sub
