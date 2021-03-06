Attribute VB_Name = "modDebug"
Option Explicit
'Copyright 2000 by AQUA TERRA Consultants
Option Private Module

Private Type dbuff
  msg As String
  lev As Long
  typ As String
  tim As Date
  mod As String
End Type
Public d(1000) As dbuff, dnull As dbuff

Public p&
Public lev&
Public flsh&
Public instSave As Boolean

Public Function BldDebugRec(d As dbuff)
  Dim s$
  
  s = d.tim & "  " & d.mod
  s = s & SPACE(30 - Len(s)) & d.typ & d.lev & ":" & d.msg
  BldDebugRec = s
End Function

Public Sub DbgMsg(msg$, Optional level& = 7, Optional modul$ = "?", Optional typ$ = "?")
  Dim q As Boolean, l&
  Dim DT As dbuff

  On Error GoTo lpt:
  l = level
  If l <= flsh Then 'save it, dont flush
    DT.msg = msg
    DT.lev = l
    DT.tim = Time
    DT.typ = UCase(typ)
    If p = 0 Then
      q = True
    ElseIf DT.msg = d(p - 1).msg And DT.tim = d(p - 1).tim Then
      q = False
    Else
      q = True
    End If
    If q Then
      If instSave Then
        Print #101, BldDebugRec(DT)
      End If
      d(p) = DT
      p = p + 1
      If p = UBound(d) + 1 Then
        p = 0
      End If
      If frmDebug.Visible Then ReDo True
    End If
  End If
  Exit Sub
lpt:
  Debug.Print msg
End Sub

Public Sub ReDo(RefreshTypes As Boolean)
  Dim t$, j&, s$, x$
  Static InRedo As Boolean
  If InRedo Then Exit Sub Else InRedo = True
  If frmDebug.Visible = False Then Exit Sub
  If RefreshTypes Then frmDebug.ListType.Clear
  t = ""
  x = ""
  For j = p - 1 To 0 Step -1
    GoSub AddToBuff
  Next j
  For j = UBound(d) To p Step -1
    GoSub AddToBuff
  Next j
  If Len(t) > 0 Then
    frmDebug.txtDetails = t
  Else
    frmDebug.txtDetails = ""
  End If
  InRedo = False
  Exit Sub
  
AddToBuff:
  If Len(d(j).msg) > 0 And lev >= d(j).lev Then
   With frmDebug.ListType
    Dim I&, found As Boolean
    I = 0
    found = False
    While I < frmDebug.ListType.ListCount And Not found
      If Left(frmDebug.ListType.List(I), 1) = d(j).typ Then found = True Else I = I + 1
    Wend
    If RefreshTypes And Not found Then
      Dim str As Variant
      str = Switch( _
        d(j).typ = "C", "Calculation", _
        d(j).typ = "E", "Error", _
        d(j).typ = "F", "Focus", _
        d(j).typ = "I", "In Routine", _
        d(j).typ = "K", "Keyboard", _
        d(j).typ = "L", "Load", _
        d(j).typ = "M", "Mouse", _
        d(j).typ = "O", "Out of Routine", _
        d(j).typ = "P", "Property Change", _
        d(j).typ = "W", "Window")
      If IsNull(str) Then
        str = d(j).typ
      Else
        str = d(j).typ & " - " & str
      End If
      .AddItem str
      I = .ListCount - 1
      .Selected(I) = True
    End If
    If frmDebug.ListType.Selected(I) Then
      s = BldDebugRec(d(j))
      If x <> s Then
        t = t & s & vbCrLf
      End If
      x = s
    End If
   End With
  End If
  Return
End Sub
