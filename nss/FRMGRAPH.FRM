VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "comdlg32.ocx"
Begin VB.Form frmGraph 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Graph"
   ClientHeight    =   5295
   ClientLeft      =   1185
   ClientTop       =   3420
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5295
   ScaleWidth      =   7935
   Tag             =   "-1"
   Begin VB.PictureBox scrGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5292
      Left            =   0
      ScaleHeight     =   5265
      ScaleWidth      =   7905
      TabIndex        =   0
      Top             =   0
      Width           =   7932
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   2400
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   327681
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuMeta 
         Caption         =   "&Metafile"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'plotting data buffer
Const BUFMAX% = 6000
Const POSMAX% = 18
Dim bufpos(2, POSMAX) As Integer
Dim wchvr(2, POSMAX) As Integer
Dim yx(BUFMAX) As Single
Dim GRAPH As Object
'plotting stuff
Dim Title As String
Dim GridX As Integer  'grid x - 1:yes
Dim MinX As Single    'minimum x value, world coord
Dim MaxX As Single
Dim AtypeX As Integer 'axis type 0TM,1AR,2LG,3-8PRB
Dim NticX As Integer
Dim TlenX As Single   'X tic length
Dim LabelX As String
Dim GridY As Integer
Dim MinY As Single
Dim MaxY As Single
Dim AtypeY As Integer 'axis type
Dim NticY As Integer
Dim TlenY As Single   'Y tic length
Dim LabelY As String
Dim MrkSiz As Single
Dim ncrv As Integer
Dim nvar As Integer
Dim leglbl(POSMAX) As String * 80

'logical font
Dim lf As LOGFONT
Dim FontToUse%, FontToSave%

Dim ZoomFlg%, ZoomSpec
Dim zx1!, zy1!, zx2!, zy2!
Private Sub ArAxis()

    'put tics and numbers on an arithmetic y-axis
    Dim i!, mystep%, clr&, lbl$

    mystep = (MaxY - MinY) / CSng(NticY)
    'keeps the outline box looking good and updated from zooms
    For i = MinY To MaxY Step mystep
      If (i <> MinY And i <> MaxY) Then
        If (GridY = 1) Then
          'grey grid
          clr = RGB(128, 128, 128)
          'dotted lines
          GRAPH.DrawStyle = 2
          'Grid the Y axis
          GRAPH.Line (CSng(MinX), i)-(CSng(MaxX), i), clr
          GRAPH.Line (1, 1)-(1, 1), clr
          'solid lines
          GRAPH.DrawStyle = 0
        End If
        'need left and right tics
        GRAPH.Line (CSng(MinX), i)-Step(TlenY, 0)
        GRAPH.Line (CSng(MaxX), i)-Step(-TlenY, 0)
      End If
      'labels
      lbl = CStr(i)
      GRAPH.CurrentY = i - GRAPH.TextHeight(lbl) / 2
      'use 1.1 to move labels away from axis a bit
      GRAPH.CurrentX = MinX - 1.1 * GRAPH.TextWidth(lbl)
      GRAPH.Print lbl
    Next i

End Sub

Private Sub AxAxis()

    'put tics and numbers on an arithmetic x-axis
    Dim i!, clr&, lbl$
    Dim mystep!

    mystep = (MaxX - MinX) / CSng(NticX)
    For i = MinX To MaxX Step mystep
      If (i <> MinX And i <> MaxX) Then
        'bottom & top tics
        GRAPH.Line (i, MinY)-Step(0, TlenX)
        GRAPH.Line (i, MaxY)-Step(0, -TlenX)
        If (GridX = 1) Then
          'grey grid
          clr = RGB(128, 128, 128)
          'dotted lines
          GRAPH.DrawStyle = 2
          'Grid the X axis
          GRAPH.Line (i, MinY)-(i, MaxY), clr
          GRAPH.DrawStyle = 0
        End If
      End If
      'labels
      lbl = CStr(i)
      If (i < MaxX) Then
        GRAPH.CurrentX = i - (GRAPH.TextWidth(lbl) / 2)
      Else
        GRAPH.CurrentX = i - GRAPH.TextWidth(lbl)
      End If
      GRAPH.CurrentY = MinY + (GRAPH.TextHeight(lbl) * 0.25)
      GRAPH.Print lbl
    Next i

End Sub

Private Sub LgAxis()

    'put tics and numbers on a logarithmic y-axis
    Dim i%, j%, ip!, lbl$

    For i = MinY To MaxY
      'left and right tics
      If i > MinY And i < MaxY Then
        'large tics for full cycles
        GRAPH.Line (CSng(MinX), i)-Step(TlenY, 0)
        GRAPH.Line (CSng(MaxX), i)-Step(-TlenY, 0)
      End If
      If i < MaxY Then
        'might want intermediate tics
        If MaxY - MinY < 5 Then
          'put small intermediate tics in between cycles
          For j = 2 To 9
            ip = i + Log10(CDbl(j))
            GRAPH.Line (MinX, ip)-Step(0.5 * TlenY, 0)
            GRAPH.Line (MaxX, ip)-Step(-0.5 * TlenY, 0)
          Next j
        End If
      End If
      'labels
      lbl = CStr(10 ^ i)
      GRAPH.CurrentY = i - GRAPH.TextHeight(lbl) / 2
      'use 1.1 to move labels away from axis a bit
      GRAPH.CurrentX = MinX - 1.1 * GRAPH.TextWidth(lbl)
      GRAPH.Print lbl
    Next i

End Sub

Private Sub Form_Activate()

    Set frmGraph = G(Tag)

End Sub

Private Sub Form_Resize()

    If Height > 675 And Width > 100 Then
      'ok to resize (not minimized)
      GRAPH.Height = Height - 675
      GRAPH.Width = Width - 100
      Call ReDrawGraph(1)
    End If

End Sub




Private Sub mnuClose_Click()

'    ZoomSpec = False
    Unload frmGraph
    Tag = -1

End Sub

Private Sub mnuMeta_Click()

    Dim ErrRet%, fname$
    On Error GoTo ErrHandler
    ErrRet = 0
    'get name from user
    CommonDialog1.DialogTitle = "Generate Windows Metafile"
    CommonDialog1.Filter = "Status Files (*.wmf)|*.wmf|All Files|*.*"
    CommonDialog1.FilterIndex = 0
    CommonDialog1.CancelError = True
    CommonDialog1.Action = 1
    fname = CommonDialog1.filename

    ErrRet = 1
    If FileLen(fname) > 0 Then
      'get rid of existing file
      Kill fname
    End If
    ErrRet = 0
BackFromErr3:

    SavePicture frmGraph.Image, fname
    
BackFromErr:
    Exit Sub
ErrHandler:
    If ErrRet = 0 Then
      Resume BackFromErr
    Else
      Resume BackFromErr3
    End If
    Resume BackFromErr

End Sub

Private Sub mnuPrint_Click()

    Call ReDrawGraph(2)

End Sub


Private Sub GridsTics()

    'put tics and numers on axes
On Error GoTo errorhandler9
    Dim i&, ifrst&, ilast&, ypos!, yint!, mystep&, clr&, lbl$
    'Y axis
    If AtypeY = 1 Then
      'arithmetic axis
      Call ArAxis
    Else
      'logarithmic axis
      Call LgAxis
    End If
    'X axis
    If AtypeX = 0 Then
      'time axis, not yet implemented
    ElseIf AtypeX = 1 Then
      'arithmetic axis
      Call AxAxis
    ElseIf AtypeX = 2 Then
      'logarithmic axis
'      Call LxAxis
    ElseIf AtypeX >= 3 Then
      Call PbAxis(AtypeX - 2)
    End If
Exit Sub
errorhandler9:   ' Error handler line label.
 Resume Next ' Resume procedure.

End Sub

Private Sub DrwLabelX()

    'label the X axis
    Dim twid!, lwid!, xpos!
    'where the label is printed
    lwid = GRAPH.TextWidth(LabelX)
    twid = MaxX - MinX
    xpos = MinX + (twid - lwid) / 2#
    GRAPH.CurrentX = xpos
    GRAPH.CurrentY = MinY + 1.25 * GRAPH.TextHeight(LabelX)
    GRAPH.Print LabelX

End Sub

Private Sub DrwLabelY()

    'label the Y axis
    Dim twid!, lwid!, tleft!, twidth!, ypos!
    Dim i%

    'temporarily set horizontal scale to
    'vertical to determine vertical text extent
    tleft = GRAPH.ScaleLeft
    twidth = GRAPH.ScaleWidth
    GRAPH.ScaleLeft = 0
    GRAPH.ScaleWidth = Abs(GRAPH.ScaleHeight)
    'multiply text extent by width/height ratio
    lwid = ScaleWidth / ScaleHeight * GRAPH.TextWidth(LabelY)
    'reset horizontal scale values
    GRAPH.ScaleLeft = tleft
    GRAPH.ScaleWidth = twidth
    'get a font with rotate 270 deg
    'rotates the label
    lf.lfEscapement = 900
    FontToUse = CreateFontIndirect(lf)
    FontToSave = SelectObject(GRAPH.hdc, FontToUse)
    twid = MaxY - MinY
    ypos = MinY + (twid - lwid) / 2#
    GRAPH.CurrentY = ypos
    'use .1 to get y label away from edge a bit
    GRAPH.CurrentX = GRAPH.ScaleLeft + 0.2 * Abs(MinX - GRAPH.ScaleLeft)
    'draw the label
    GRAPH.Print LabelY
    'back to original font
    lf.lfEscapement = 0
    i = SelectObject(GRAPH.hdc, FontToSave)
    i = DeleteObject(FontToUse)

End Sub

Private Sub ReDrawGraph(outflag%)

    'this is the guts of drawing any graph
On Error GoTo ErrorHandler3
    Dim i%, retcod%
    Dim lstr$
    Dim LogPixelsY&, numcol&, clr&
    'set output object
    If outflag = 1 Then
      'output to screen
      Set GRAPH = scrGraph
      'clear the graph
      GRAPH.Cls
    Else
      'output to printer
      CommonDialog1.ShowPrinter
      Set GRAPH = Printer
    End If
    'scale graph
    Call ReScaleGraph(retcod)
    If retcod <> 0 Then GoTo ErrorHandler3
    lstr = "Arial" & Chr$(0)
#If Win16 Then
    lf.lfFaceName = lstr
#Else
    For i = 0 To 5
      lf.lfFaceName(i) = AscB(Mid(lstr, i + 1, 1))
    Next i
#End If
    LogPixelsY = GetDeviceCaps(GRAPH.hdc, 90)
    numcol = GetDeviceCaps(GRAPH.hdc, 24)
    If outflag = 2 And numcol > 2 Then
      'printer has color capability
      outflag = 3
    End If
    lf.lfHeight = -0.15 * LogPixelsY
    lf.lfWeight = 400
    FontToUse = CreateFontIndirect(lf)
    FontToSave = SelectObject(GRAPH.hdc, FontToUse)
    'Debug.Print FontToUse
    'title
    Call DrwTitle
    'Xaxis label
    Call DrwLabelX
    'Yaxis with label
    Call DrwLabelY
    'box around data area, box where plot data goes
    GRAPH.Line (MinX, MinY)-(MaxX, MaxY), clr, B
    'axes numbering/tics
    Call GridsTics
    'draw the curves
    Call ReDrawLine(outflag)
    If outflag = 1 Then
      'show it on screen
      Show
    Else
      'terminate printer output
      GRAPH.EndDoc
    End If
    'back to original font
    i = SelectObject(GRAPH.hdc, FontToSave)
    i = DeleteObject(FontToUse)
Exit Sub
ErrorHandler3:   ' Error handler line label.
Exit Sub

End Sub

Private Sub ReDrawLine(outflag%)

    Dim kx%, ky%
    'draw curves for a plot
On Error GoTo errorhandler10
    Dim i%, cnt%, ipos%, edate%, llen%
    Dim rtmp!, ypleg!, lextnt!, clr&, tclr&
    Dim lstr$

    For i = 1 To ncrv
      'assign variables for this curve
      ky = wchvr(1, i)
      kx = wchvr(2, i)
      'where is first point?
      GRAPH.CurrentX = yx(bufpos(1, kx))
      If AtypeY = 1 Then
        GRAPH.CurrentY = yx(bufpos(1, ky))
      Else
        GRAPH.CurrentY = Log10(CDbl(yx(bufpos(1, ky))))
      End If
      If outflag Mod 2 = 1 Then
        'plot to screen (or color printer)
        If outflag = 1 Then
          'plot to screen, increase width of lines
          GRAPH.DrawWidth = 2
        End If
        'indicate curves w/color
        clr = QBColor(i)
      Else
        'plot to black and white printer
        'indicate curves w/drawing styles
        GRAPH.DrawStyle = (i - 1) Mod 5
      End If
      For cnt = bufpos(1, ky) To bufpos(2, ky)
        ipos = bufpos(1, kx) + cnt - bufpos(1, ky)
        If AtypeY = 1 Then
          GRAPH.Line -(yx(ipos), yx(cnt)), clr
        Else
          GRAPH.Line -(yx(ipos), Log10(CDbl(yx(cnt)))), clr
        End If
'        If (Gid = 1) Then
'          'marker for data points on zoom or fill
'          GRAPH.Circle (yx(cnt), yx(ipos)), MrkSiz, clr
'        End If
      Next cnt
      If ncrv > 1 Then
        'draw legend
        lstr = Trim(leglbl(i))
        ypleg = MaxY - (i - 1 + 0.25) * -GRAPH.TextHeight(lstr) - 1.25 * TlenX
        GRAPH.CurrentY = ypleg
        If AtypeX > 2 And AtypeX < 6 Then
          GRAPH.CurrentX = 0.05 * MaxX
          GRAPH.Line -(0.1 * MaxX, ypleg), clr
          GRAPH.CurrentX = 0.1 * MaxX + GRAPH.TextWidth(" ")
        Else
          GRAPH.CurrentX = 0.5 * MaxX
          GRAPH.Line -(0.55 * MaxX, ypleg), clr
          GRAPH.CurrentX = 0.55 * MaxX + GRAPH.TextWidth(" ")
        End If
        GRAPH.CurrentY = ypleg + 0.4 * -GRAPH.TextHeight(lstr)
        If outflag = 1 Then
          'make legend text smaller for screen
          GRAPH.FontSize = 0.75 * GRAPH.FontSize
        End If
        lextnt = GRAPH.CurrentX + GRAPH.TextWidth(lstr)
        If lextnt > MaxX Then
          'legend label exceeds right y axis
          llen = LenStr(lstr)
          While lextnt > MaxX
            llen = llen - 1
            lstr = Left(lstr, llen)
            lextnt = GRAPH.CurrentX + GRAPH.TextWidth(lstr)
          Wend
        End If
        tclr = GRAPH.ForeColor
        GRAPH.ForeColor = clr
        GRAPH.Print lstr
        If outflag = 1 Then
          'reset text size
          GRAPH.FontSize = 1 / 0.75 * GRAPH.FontSize
        End If
        GRAPH.ForeColor = tclr
      End If
    Next i
Exit Sub
errorhandler10:   ' Error handler line label.
 Resume Next ' Resume procedure.
 
End Sub

Private Sub ReDrawZoomBx()

On Error GoTo ErrorHandler4
    'get rid of current line
    GRAPH.Line (zx1, zy1)-(zx2, zy2), RGB(255, 255, 255), B
Exit Sub
ErrorHandler4:   ' Error handler line label.
 Resume Next ' Resume procedure.

End Sub

Private Sub ReScaleGraph(retcod%)

    'reset scale based on data being plotted
On Error GoTo errorhandler5
    'local version units are world coord
    Dim RmrgX!, TmrgY!, LmrgX!, BmrgY!
    Dim wid!, hei!
    '720 is 1/2 inch in twips
    RmrgX = 720
    LmrgX = 2160
    TmrgY = 720
    BmrgY = 2160
    wid = GRAPH.Width
    hei = GRAPH.Height
    'rescale margins based on new graph scale
    RmrgX = (RmrgX / wid) * (MaxX - MinX)
    TmrgY = (TmrgY / hei) * (MaxY - MinY)
    'tics are 1/2 of margins (for now)
    TlenY = RmrgX / 3
    TlenX = TmrgY / 3
    LmrgX = (LmrgX / wid) * (MaxX - MinX)
    BmrgY = (BmrgY / hei) * (MaxY - MinY)
    'reset scale
    GRAPH.Scale (MinX - LmrgX, MaxY + TmrgY)-(MaxX + RmrgX, MinY - BmrgY)
    retcod = 0
    
Exit Sub
errorhandler5:   ' Error handler line label.
retcod = 1

End Sub

Private Sub DrwTitle()

    'put title on the plot
    Dim twid!, lwid!, xpos!
    'where the title is printed
    lwid = GRAPH.TextWidth(Title)
    twid = MaxX - MinX
    xpos = MinX + (twid - lwid) / 2#
    GRAPH.CurrentX = xpos
    GRAPH.CurrentY = MinY + 2.5 * GRAPH.TextHeight(Title)
    GRAPH.Print Title

End Sub


Public Sub gpdatr(inum%, ipos%, nv%, arra() As Single, retcod%)

    'put data to be plotted in plot buffer
    Dim i%, k%, i1%, i2%

    bufpos(1, inum) = ipos
    bufpos(2, inum) = ipos + nv - 1
    If ipos + nv - 1 > BUFMAX Then
      'not enough space
      retcod = 1
    Else
      'put values in buffer
      For i = 1 To nv
        k = i + ipos - 1
        yx(k) = arra(i)
      Next i
      retcod = 0
    End If

End Sub

Public Sub gplabl(xtype%, ytype%, xlab$, ylab$, titl$, capt$)

    'set axes types and labels and plot title
    AtypeX = xtype
    AtypeY = ytype
    LabelX = xlab
    LabelY = ylab
    Title = titl
    Caption = capt

End Sub

Public Sub gplgnd(icrv%, ilbl$)

    'set the legend for a curve
    leglbl(icrv) = ilbl

End Sub

Public Sub gpncrv(icrv%, ivar%)

    'set the number of curves and variables
    ncrv = icrv
    nvar = ivar

End Sub

Public Sub gpscle(plmn!(), plmx!(), Ntics%())

    'set min/max for left/right y, aux, and x axes
    '(only left y and x functional at present)
    MinY = plmn(0)
    MaxY = plmx(0)
    MinX = plmn(3)
    MaxX = plmx(3)
    'set number of tics (if axis arithmetic)
    NticY = Ntics(0)
    NticX = Ntics(3)

End Sub

Public Sub GraphIt()

    'generate current graph
    If AtypeY = 2 Then
      'adjust y min/max to log coordinate system
      MinY = Fix(1.01 * (Log10(CDbl(MinY))))
      MaxY = Fix(1.01 * (Log10(CDbl(MaxY))))
    End If
    Set GRAPH = scrGraph
    frmGraph.Show

End Sub

Public Sub grinit()

    'initialization routine for generating a plot
    Dim i%

    'find next available plot form
    i = 0
    While i < MXGRPH
      If G(i).Tag = -1 Then
        'this form is available, use it for this plot
        Load G(i)
        G(i).Tag = i
        Set frmGraph = G(i)
        i = MXGRPH
      Else
        i = i + 1
      End If
    Wend

End Sub

Private Sub PbAxis(atype%)
    
    'put tics and numbers on a probability x-axis
    Dim i%, j%, xpos!, lbl$
    Static inifg%, pct!(18), rtp!(18)

    If inifg = 0 Then
      'init percentage and return period values
      Call PbInit(pct(), rtp())
      inifg = 1
    End If
    For i = 0 To 18
      If atype <= 3 Then
        'exceedance plot
        j = i
      Else
        'non-exceedance plot
        j = 18 - i
      End If
      If atype = 2 Or atype = 5 Then
        xpos = gausex(1 / rtp(j))
      Else
        xpos = gausex(pct(j) / 100)
      End If
      If xpos >= MinX And xpos <= MaxX Then
        'within x-axis range, put bottom/top tics
        GRAPH.Line (xpos, CSng(MinY))-Step(0, TlenX), 0
        GRAPH.Line (xpos, CSng(MaxY))-Step(0, -TlenX), 0
        'labels
        If atype = 2 Or atype = 5 Then
          lbl = CStr(rtp(j))
        ElseIf atype = 3 Or atype = 6 Then
          lbl = CStr(0.01 * pct(j))
        Else
          lbl = CStr(pct(j))
        End If
        GRAPH.CurrentX = xpos - (GRAPH.TextWidth(lbl) / 2)
        GRAPH.CurrentY = MinY + (GRAPH.TextHeight(lbl) * 0.25)
        GRAPH.Print lbl
      End If
    Next i

End Sub

Private Sub PbInit(pct!(), rtp!())

    'initialize array constants for probability axis
    pct(0) = 99.9
    pct(1) = 99.8
    pct(2) = 99.5
    pct(3) = 99
    pct(4) = 98
    pct(5) = 95
    pct(6) = 90
    pct(7) = 80
    pct(8) = 70
    pct(9) = 50
    pct(10) = 30
    pct(11) = 20
    pct(12) = 10
    pct(13) = 5
    pct(14) = 2
    pct(15) = 1
    pct(16) = 0.5
    pct(17) = 0.2
    pct(18) = 0.1
    rtp(0) = 1.001
    rtp(1) = 1.002
    rtp(2) = 1.005
    rtp(3) = 1.01
    rtp(4) = 1.02
    rtp(5) = 1.05
    rtp(6) = 1.1
    rtp(7) = 1.25
    rtp(8) = 1.5
    rtp(9) = 2
    rtp(10) = 3
    rtp(11) = 5
    rtp(12) = 10
    rtp(13) = 20
    rtp(14) = 50
    rtp(15) = 100
    rtp(16) = 200
    rtp(17) = 500
    rtp(18) = 1000

End Sub

Private Sub scrGraph_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler6  ' Set up error handler.
    If ZoomSpec = True Then
      If ZoomFlg = -1 Then
'       clear off last zoom
        Call ReDrawZoomBx
      End If
      zx1 = x
      zy1 = y
      ZoomFlg = 1
      GRAPH.DrawMode = 7
      ZoomSpec = False
    End If
Exit Sub
ErrorHandler6:   ' Error handler line label.
 Resume Next ' Resume procedure.

End Sub


Private Sub scrGraph_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler  ' Set up error handler.
    If (ZoomFlg > 0) Then
      If (ZoomFlg = 2) Then
'        erase old box when moves
         Call ReDrawZoomBx
      End If
'     draw new box
      zx2 = x
      zy2 = y
      Call ReDrawZoomBx
      ZoomFlg = 2
    End If
Exit Sub
ErrorHandler:   ' Error handler line label.
 Resume Next ' Resume procedure.

End Sub


Private Sub scrGraph_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler1  ' Set up error handler.

    If (ZoomFlg > 0) Then
        MousePointer = 0
        If (Abs(zx1 - x) > 1 And Abs(zy1 - y) > 2) Then
            If (zx1 > x) Then
                MinX = x
                MaxX = zx1
            Else
                MinX = zx1
                MaxX = x
            End If
            If (zy1 > y) Then
                MinY = y
                MaxY = zy1
            Else
                MinY = zy1
                MaxY = y
            End If
            ZoomFlg = -1
            'draw the zoomed window
            Call ReDrawGraph(1)
        Else
            Const MB_OK = 0, MB_OKCANCEL = 1    ' Define buttons.
            Const MB_YESNOCANCEL = 3, MB_YESNO = 4
            Const MB_ICONSTOP = 16, MB_ICONQUESTION = 32    ' Define Icons.
            Const MB_ICONEXCLAMATION = 48, MB_ICONINFORMATION = 64
            Const MB_DEFBUTTON2 = 256, idyes = 6, IDNO = 7  ' Define other.
            Dim dgdef, msg, response, Title ' Declare variables.
        
            Title = "Default Printer"
            dgdef = MB_OK + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialo
            msg = "Please draw a larger Zoom Box."
            response = MsgBox(msg, dgdef, Title)
            ZoomFlg = 0
       End If
    End If
Exit Sub
ErrorHandler1:   ' Error handler line label.
End
' Resume Next ' Resume procedure.

End Sub



Public Sub gpwcxy(curve%, yvar%, xvar%)

    'store which variables to use for each curve
    wchvr(1, curve) = yvar
    wchvr(2, curve) = xvar

End Sub
