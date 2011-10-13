Attribute VB_Name = "modNSSLib"
Option Explicit
'Copyright 2001 by AQUA TERRA Consultants

Public Function ssMessageBox(Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = "NSS") As VbMsgBoxResult
  Static Suppress As Boolean
  Static LogFileName As String
  Dim OutFile As Integer
  
  Select Case UCase(Prompt)
    Case "SUPPRESSMESSAGES": Suppress = True
    Case "ENABLEMESSAGES":   Suppress = False
    Case "LOGMESSAGES"
      LogFileName = ""
      If Len(Title) > 0 And Title <> "NSS" Then
        On Error GoTo CantLog
        If FileExists(Title) Then
          If (GetAttr(Title) And vbReadOnly) = 0 Then LogFileName = Title
        Else
          MkDirPath FilenameOnly(Title)
          If FileExists(PathNameOnly(Title), True, False) Then
            OutFile = FreeFile(0)
            Open Title For Append As OutFile
            Print #OutFile, "Log starting " & Date & " " & Time
            Close OutFile
            LogFileName = Title
          End If
        End If
      End If
CantLog:
    Case Else
      If Suppress Then
        Debug.Print "Suppressed message: " & Prompt
        ssMessageBox = vbOK
      Else
        ssMessageBox = MsgBox(Prompt, Buttons, Title) ', App.HelpFile)
      End If
      If Len(LogFileName) > 0 Then
        On Error GoTo SkipLogging
        OutFile = FreeFile(0)
        Open LogFileName For Append As OutFile
        Print #OutFile, Date & " " & Time _
                             & " Prompt: " & Prompt _
                             & ", Buttons: " & Buttons _
                             & ", Title: " & Title _
                             & ", Result: " & ssMessageBox
        Close OutFile
SkipLogging:
      End If
  End Select


End Function

Public Function CrippenBueMaxFloodEnvelope(ByVal drainage_area As Double, ByVal flood_region As Long) As Double
  Static k1!(17), k2!(17), k3!(17) 'Crippen and Bue constants
  If k1(1) <> 23200! Then
    k1(1) = 23200!
    k1(2) = 28000!
    k1(3) = 54400!
    k1(4) = 42600!
    k1(5) = 121000!
    k1(6) = 70500!
    k1(7) = 49100!
    k1(8) = 43800!
    k1(9) = 75000!
    k1(10) = 62500!
    k1(11) = 40800!
    k1(12) = 89900!
    k1(13) = 64500!
    k1(14) = 10000!
    k1(15) = 116000!
    k1(16) = 98900!
    k1(17) = 80500!

    k2(1) = 0.895
    k2(2) = 0.77
    k2(3) = 0.924
    k2(4) = 0.938
    k2(5) = 0.838
    k2(6) = 0.937
    k2(7) = 0.883
    k2(8) = 0.954
    k2(9) = 0.849
    k2(10) = 1.116
    k2(11) = 0.919
    k2(12) = 0.935
    k2(13) = 0.873
    k2(14) = 0.71
    k2(15) = 1.059
    k2(16) = 1.029
    k2(17) = 1.024

    k3(1) = -1.082
    k3(2) = -0.897
    k3(3) = -1.373
    k3(4) = -1.327
    k3(5) = -1.354
    k3(6) = -1.297
    k3(7) = -1.352
    k3(8) = -1.357
    k3(9) = -1.368
    k3(10) = -1.371
    k3(11) = -1.352
    k3(12) = -1.304
    k3(13) = -1.338
    k3(14) = -0.844
    k3(15) = -1.572
    k3(16) = -1.341
    k3(17) = -1.461
  End If
  If flood_region >= 1 And flood_region <= 17 Then
    Dim Tmp#
    Tmp = 5# + drainage_area ^ 0.5
    CrippenBueMaxFloodEnvelope = k1(flood_region) _
                               * drainage_area ^ k2(flood_region) _
                               * Tmp ^ k3(flood_region)
  End If
End Function

Public Function nss500(Q() As Double, t() As Single) As Double

  'Perform extrapolation of state-equation flood freq. curves
  'to 500 year flood along Log-Pearson Type III curves.
  'q() is array of flow values for intervals up to 500
  't() is array of intervals corresponding to q() (e.g. 2, 5, 10, 25, 50, 100, 500)

  Dim B0!, B1!, b2!
  Dim bt0!, bt1!, bt2!
  Dim zx!, wx#
  Dim npts&, errflg&
  Dim z() As Single
  Dim w() As Single
  Dim ql()  As Single
  Dim sku!, tempa!
  Const one# = 1#
  Const zero# = 0#
  Dim w2 As Single, w10 As Single, w100 As Single, qx As Single
  Dim i As Integer, k As Integer
  Dim pz As Double

  B0 = 0#
  B1 = 0#
  b2 = 0#


  npts = UBound(Q) - 1
  If npts > 1 Then
    ReDim ql(0 To npts - 1)
    ReDim z(0 To npts - 1)
    ReDim w(0 To npts - 1)
    
    For i = 1 To npts
      ql(i - 1) = Log10(Q(i))
      tempa = one / t(i)
      z(i - 1) = gausex(tempa)
    Next i
    
    If Q(1) = 0 Then
      'use 5-year peak for 2-year peak value of 0
      ql(0) = Log10(Q(2))
    End If
    'Fit quadratic to freq curve on log-probability coordinates
'      bt0 = CSng(b0)
'      bt1 = CSng(b1)
'      bt2 = CSng(b2)
    Call slreg2(ql(), z(), npts, B0, B1, b2) ', bt0, bt1, bt2)
'      b0 = CDbl(bt0)
'      b1 = CDbl(bt1)
'      b2 = CDbl(bt2)

    'Determine Pearson Type III skew 2, 10, 100-year points on curve
    w2 = B0
    pz = gausex(0.1)
    w10 = B0 + pz * (B1 + pz * b2)
    pz = gausex(0.01)
    w100 = B0 + pz * (B1 + pz * b2)
    If Abs(w10 - w2) < 0.0000001 Then
      ssMessageBox "Division by zero in estimating 500-year flow value"
      sku = 0
    Else
      sku = hartrg((w100 - w10) / (w10 - w2))
    End If
    'Transform to Pearson/Wilfrt probability scale
    For i = 0 To npts - 1
      w(i) = wilfrt(sku, z(i), errflg)
    Next i

    'Fit straight line to freq curve in log-Pearson coordinates */
'      bt0 = b0
'      bt1 = b1
    Call slreg(ql(), w(), npts, B0, B1) ' bt1, bt0)
'      b0 = bt0
'      b1 = bt1
    b2 = zero

    'Extrapolate straight line in log-Pearson-III coordinates                */
    zx = gausex(1# / 500#)
    wx = wilfrt(sku, zx, errflg)
    
    qx = B0 + wx * (B1 + wx * b2)
    nss500 = 10# ^ qx

  End If
End Function

Private Sub slreg(y() As Single, x() As Single, n&, ByRef B0!, ByRef B1!)

  'simple linear (straight line) regression of Y on X.  WK 750212.
  Dim i&, sy!, sx!, sxy!, sxx!, varx!
  
  sy = 0#
  sx = 0#
  sxy = 0#
  sxx = 0#
  For i = 0 To n - 1
    sy = sy + y(i)
    sx = sx + x(i)
    sxy = sxy + x(i) * y(i)
    sxx = sxx + x(i) ^ 2
  Next i
  sy = sy / n
  sx = sx / n
  sxy = sxy / n
  sxx = sxx / n
  varx = sxx - sx ^ 2
  B1 = 0#
  If varx > 0 Then
    B1 = (sxy - sx * sy) / varx
  End If
  B0 = sy - B1 * sx

End Sub

Private Sub slreg2(y!(), x!(), n&, ByRef B0!, ByRef B1!, ByRef b2!)

  'SLREG2 - simple quadratic regression of Y on X
  '         Y = B0 + B1*X + B2*X**2
  Dim i&, ybar!, xbar!, xi!, yi!
  Dim sxx#, sx3#, sx4#, sxy#, sxxy#, d#

  sxx = 0#
  sx3 = 0#
  sx4 = 0#
  sxy = 0#
  sxxy = 0#
  ybar = 0#
  xbar = 0#
  For i = 0 To n - 1
    xbar = xbar + x(i)
    ybar = ybar + y(i)
  Next i
  xbar = xbar / n
  ybar = ybar / n
  For i = 0 To n - 1
    xi = x(i) - xbar
    yi = y(i) - ybar
    sxx = sxx + xi ^ 2
    sx3 = sx3 + xi ^ 3
    sx4 = sx4 + xi ^ 4
    sxy = sxy + xi * yi
    sxxy = sxxy + yi * xi ^ 2
  Next i
  d = sx4 - sxx ^ 2 / n
  b2 = (sxxy * sxx - sx3 * sxy) / (sxx * d - sx3 ^ 2)
  B1 = (sxy - sx3 * b2) / sxx
  B0 = -b2 * sxx / n
  B0 = B0 + ybar - B1 * xbar + b2 * xbar ^ 2
  B1 = B1 - b2 * xbar * 2#

End Sub

Private Function hartrg(r!) As Single

  'COMPUTES SKEW COEFF OF PEARSON TYPE III  DISTN,
  'GIVEN THE RATIO (Q.100 - Q.10)/(Q.10 - Q.2),
  'WHERE Q.T IS THE T-YEAR (1.-1/T - PROBABILITY)
  'QUANTILE.   THE EQUATIONS WERE FOUND BY POLYNOMIAL
  'REGRESSION, ETC., OF SKEW VS RATIO, WHERE THE RATIOS
  'WERE LOOKED UP IN HARTER'S TABLES FOR GIVEN SKEWS.
  '       WK 11/80.  FOR WRC BULL 17-B.

  Dim rtmp!, rt1!
  
  If r < 0.243 Then
    If r > 0# Then
      rtmp = -6# + 10# ^ (0.72609 + 0.15397 * Log10(CDbl(r)))
    Else
      rtmp = -4.8
    End If
  ElseIf r > 1.6 Then
    rtmp = 7.1 + 1.6 * (r - 2.4) - 1.4 * ((r - 2.4) ^ 2 + 5.1888) ^ 0.5
  Else
    rt1 = 2.35713 + r * (-0.7387)
    rtmp = -2.51898 + r * (3.82069 + r * (-2.3196 + r * (rt1)))
  End If
  hartrg = rtmp

End Function

Private Function wilfrt(sku!, zeta!, errflg&) As Single

  'WILFRT -- WILSON-HILFERTY REVISED TRANSFORM
  'PURPOSE -- APPROXIMATE TRANSFORMATION OF GAUSSIAN PERCENTAGE POINT
  '   INTO STANDARDIZED PEARSON TYPE III.   THIS VERSION REPRODUCES
  '   CORRECT MEAN, VARIANCE, SKEW AND LOWER BOUND OF STANDARDIZED
  '   PEARSON-III AT SKEWS UP TO 9.0 AT LEAST.  DIFFERENCES BETWEEN
  '   WILFRT PERCENTAGE POINTS AND HARTERS TABLES ARE OF THE ORDER OF
  '   A FEW HUNDREDTHS OF A STD. DEVIATION, EXCEPT IN EXTREME POSITIV
  '   TAIL (95% OR SO) WHERE ERROR IS OF ORDER OF TENTHS IN MAGNI-
  '   TUDE BUT ABOUT 3% IN RELATIVE MAG.
  'USAGE --      X=WILFRT(SKEW,ZETA)*STDDEV+AMEAN
  '   SKEW IS INPUT SKEW, MAY BE ZERO OR NEGATIVE OR POSITIVE.
  '         IF ABS(SKEW) IS GREATER THAN 9.75, 9.75 IS USED.
  '   ZETA IS STANDARD GAUSSIAN VARIATE.   FOR EXAMPLE, GAUSSB(IRAN)
  '         YIELDS RANDOM NOS WHILE GAUSAB(PROB) YIELDS THE
  '         PROB-TH QUANTILE.
  '   STDDEV AND AMEAN ARE DESIRED VALUES OF STD DEVIATION AND
  '         MEAN, IF DIFFERENT FROM ONE AND ZERO.
  'NOTE -- EACH INPUT SKEW VALUE IS COMPARED WITH PREVIOUS INPUT
  '   VALUE. IF DIFFERENT BY MORE THAN 0.0003, TABLE LOOKUP OF NEW
  '   PARAMETERS TAKES PLACE.  THEREFORE, CHAGE THE INPUT SKEW
  '   AS SELDOM AS POSSIBLE.
  'WKIRBY  72-02-25
  'REVISED 73-02-09   TO ACCEPT ZERO SKEW.
  'REF -- W.KIRBY, COMPUTER-ORIENTED WILSON-HILFERTY TRANSFORMATION..
  '    WATER RESOUR RESCH 8(5)1251-4, OCT 72.
  'REV 6/83 WK FOR PRIME ---- SAVE STTMNT ----
  'REV 7/86 BY AML TO OSW CODING CONVENTION
  'rev 9/96 by PRH for VB
    
  Dim ask!, a!, b!, G!, h!, z!, sig!, fmu!
  Const skutol As Single = 0.0003

  'first time thru or new sku (skew)
  ask = Abs(sku)
  If ask >= skutol Then
    'nonzero skew
    Call wilfrs(ask, G, h, a, b, errflg)
    sig = G * 0.1666667
    fmu = 1# - sig * sig
    If sku < 0# Then
      sig = -sig
      a = -a
    End If
    z = fmu + sig * zeta
    If z < h Then
      z = h
    End If
    wilfrt = a * (z * z * z - b)
  Else
    'zero skew
    wilfrt = zeta
  End If

End Function

Private Sub wilfrs(sk!, G!, h!, a!, b!, errflg&)

  'COMPUTES PARAMETERS USED BY WILFRT TRANSFORMATIN
  'USES APPROX FORMULA AND CORRECTION TERMS PREPARED FROM
  'ROUTINE WHMPP (E443-5).  WKIRBY  FEB72
  'PARAMETERS RETURNED TO WILFRT ARE INTENDED TO MAKE
  'WILFRT A STANDARDIZED R.V.  (MEAN=0,STDEV=1) WITH
  'SPECIFIED SKEW AND CORRECT LOWER BOUND
  'REVISED CALC OF CORRECTION TABLE  72-03-03 WK

  Const nroz As Single = 40
  Static table!(1 To 40, 1 To 4)
  Dim flag&, i&, k&
  Dim row!(1 To 4)
  Dim s!, Q!, p!, tog!
  
  If table(1, 1) = 0 Then
    table(1, 1) = 0#
    table(2, 1) = 0.25
    table(3, 1) = 0.5
    table(4, 1) = 0.75
    table(5, 1) = 1#
    table(6, 1) = 1.25
    table(7, 1) = 1.5
    table(8, 1) = 1.75
    table(9, 1) = 2#
    table(10, 1) = 2.25
    table(11, 1) = 2.5
    table(12, 1) = 2.75
    table(13, 1) = 3#
    table(14, 1) = 3.25
    table(15, 1) = 3.5
    table(16, 1) = 3.75
    table(17, 1) = 4#
    table(18, 1) = 4.25
    table(19, 1) = 4.5
    table(20, 1) = 4.75
    table(21, 1) = 5#
    table(22, 1) = 5.25
    table(23, 1) = 5.5
    table(24, 1) = 5.75
    table(25, 1) = 6#
    table(26, 1) = 6.25
    table(27, 1) = 6.5
    table(28, 1) = 6.75
    table(29, 1) = 7#
    table(30, 1) = 7.25
    table(31, 1) = 7.5
    table(32, 1) = 7.75
    table(33, 1) = 8#
    table(34, 1) = 8.25
    table(35, 1) = 8.5
    table(36, 1) = 8.75
    table(37, 1) = 9#
    table(38, 1) = 9.25
    table(39, 1) = 9.5
    table(40, 1) = 9.75
    table(1, 2) = 0#
    table(2, 2) = -0.000144
    table(3, 2) = -0.001137
    table(4, 2) = -0.003762
    table(5, 2) = -0.008674
    table(6, 2) = -0.011555
    table(7, 2) = -0.010076
    table(8, 2) = -0.006049
    table(9, 2) = -0.000921
    table(10, 2) = 0.004189
    table(11, 2) = 0.008515
    table(12, 2) = 0.011584
    table(13, 2) = 0.013139
    table(14, 2) = 0.013122
    table(15, 2) = 0.010945
    table(16, 2) = 0.007546
    table(17, 2) = 0.002767
    table(18, 2) = -0.003181
    table(19, 2) = -0.010089
    table(20, 2) = -0.017528
    table(21, 2) = -0.025476
    table(22, 2) = -0.033609
    table(23, 2) = -0.042434
    table(24, 2) = -0.050525
    table(25, 2) = -0.058192
    table(26, 2) = -0.065221
    table(27, 2) = -0.07141
    table(28, 2) = -0.076638
    table(29, 2) = -0.080655
    table(30, 2) = -0.083349
    table(31, 2) = -0.084584
    table(32, 2) = -0.084203
    table(33, 2) = -0.082089
    table(34, 2) = -0.078126
    table(35, 2) = -0.072165
    table(36, 2) = -0.064188
    table(37, 2) = -0.054059
    table(38, 2) = -0.041633
    table(39, 2) = -0.027005
    table(40, 2) = -0.010188
    table(1, 3) = 0#
    table(2, 3) = 0.004614
    table(3, 3) = 0.009159
    table(4, 3) = 0.013553
    table(5, 3) = 0.017753
    table(6, 3) = 0.021764
    table(7, 3) = 0.025834
    table(8, 3) = 0.030406
    table(9, 3) = 0.03571
    table(10, 3) = 0.04173
    table(11, 3) = 0.048321
    table(12, 3) = 0.055309
    table(13, 3) = 0.062538
    table(14, 3) = 0.069873
    table(15, 3) = 0.077334
    table(16, 3) = 0.084682
    table(17, 3) = 0.091926
    table(18, 3) = 0.099028
    table(19, 3) = 0.105967
    table(20, 3) = 0.112695
    table(21, 3) = 0.119245
    table(22, 3) = 0.106551
    table(23, 3) = 0.095488
    table(24, 3) = 0.085671
    table(25, 3) = 0.07699
    table(26, 3) = 0.06929
    table(27, 3) = 0.062443
    table(28, 3) = 0.056349
    table(29, 3) = 0.050908
    table(30, 3) = 0.046047
    table(31, 3) = 0.041702
    table(32, 3) = 0.037815
    table(33, 3) = 0.034339
    table(34, 3) = 0.031229
    table(35, 3) = 0.028445
    table(36, 3) = 0.025964
    table(37, 3) = 0.023753
    table(38, 3) = 0.021782
    table(39, 3) = 0.020043
    table(40, 3) = 0.018528
    table(1, 4) = 0#
    table(2, 4) = 0#
    table(3, 4) = -0.000001
    table(4, 4) = -0.000004
    table(5, 4) = -0.000021
    table(6, 4) = -0.000075
    table(7, 4) = -0.00019
    table(8, 4) = -0.000326
    table(9, 4) = -0.000317
    table(10, 4) = 0.000116
    table(11, 4) = 0.000434
    table(12, 4) = 0.000116
    table(13, 4) = -0.000464
    table(14, 4) = -0.000981
    table(15, 4) = -0.001165
    table(16, 4) = -0.000743
    table(17, 4) = 0.000435
    table(18, 4) = 0.002479
    table(19, 4) = 0.005462
    table(20, 4) = 0.009353
    table(21, 4) = 0.014206
    table(22, 4) = 0.019964
    table(23, 4) = 0.026829
    table(24, 4) = 0.034307
    table(25, 4) = 0.042495
    table(26, 4) = 0.051293
    table(27, 4) = 0.060593
    table(28, 4) = 0.070324
    table(29, 4) = 0.080332
    table(30, 4) = 0.090532
    table(31, 4) = 0.100831
    table(32, 4) = 0.111114
    table(33, 4) = 0.121283
    table(34, 4) = 0.131245
    table(35, 4) = 0.140853
    table(36, 4) = 0.15012
    table(37, 4) = 0.158901
    table(38, 4) = 0.167085
    table(39, 4) = 0.174721
    table(40, 4) = 0.181994
  End If
  s = sk
  k = 1
  flag = 0
  errflg = 0
  i = 1
  Do
    i = i + 1
    If table(i, 1) > s Then flag = 1
    k = i - 1
  Loop While i < nroz And flag = 0

  If flag = 0 Then
    errflg = 1
    For i = 1 To 4
      row(i) = table(nroz, i)
    Next i
'     replace "row" equivalence
'      s = table(nroz, 1)
'      q = table(nroz, 2)
'      p = table(nroz, 3)
'      tog = table(nroz, 4)
  Else
    p = (s - table(k, 1)) / (table(k + 1, 1) - table(k, 1))
    Q = 1# - p
    For i = 2 To 4
      row(i) = Q * table(k, i) + p * table(k + 1, i)
    Next i
'     replace "row" equivalence
'      q = q * table(k, 2) + p * table(k + 1, 2)
'      p = q * table(k, 3) + p * table(k + 1, 3)
'      tog = q * table(k, 4) + p * table(k + 1, 4)
  End If

  G = s + row(2)
'   replace "row" equivalence
'    g = s + q
  If s > 1# Then G = G - 0.063 * (s - 1#) ^ 1.85
  tog = 2# / s
  Q = tog
  If Q < 0.4 Then Q = 0.4
  a = Q + row(3)
'   replace "row" equivalence
'    a = q + p
  Q = 0.12 * (s - 2.25)
  If Q < 0# Then Q = 0#
  b = 1# + Q * Q + row(4)
'   replace "row" equivalence
'    b = 1# + q * q + tog
  If (b - tog / a) < 0# Then
    'Stop WILFRS
    ssMessageBox "Very serious problem in routine WILFRS.  Contact software distributor", 16
  End If
  h = (b - tog / a) ^ 0.3333333

End Sub

Public Function gausex(exprob!) As Single
    'GAUSSIAN PROBABILITY FUNCTIONS   W.KIRBY  JUNE 71
       'GAUSEX=VALUE EXCEEDED WITH PROB EXPROB

    'GAUSCF MODIFIED 740906 WK -- REPLACED ERF FCN REF BY RATIONAL APPRX N
    'ALSO REMOVED DOUBLE PRECISION FROM GAUSEX AND GAUSAB.
    '76-05-04 WK -- TRAP UNDERFLOWS IN EXP IN GUASCF AND DY.

    'rev 8/96 by PRH for VB
    Const C0! = 2.515517
    Const C1! = 0.802853
    Const C2! = 0.010328
    Const D1! = 1.432788
    Const d2! = 0.189269
    Const d3! = 0.001308
    Dim pr!, rtmp!, p!, t!, numerat!, denom!
    
    p = exprob
    If p >= 1# Then
      'set to minimum
      rtmp = -10#
    ElseIf p <= 0# Then
      'set at maximum
      rtmp = 10#
    Else
      'compute value
      pr = p
      If p > 0.5 Then pr = 1# - pr
      t = (-2# * Log(pr)) ^ 0.5
      numerat = (C0 + t * (C1 + t * C2))
      denom = (1# + t * (D1 + t * (d2 + t * d3)))
      rtmp = t - numerat / denom
      If p > 0.5 Then rtmp = -rtmp
    End If
    gausex = rtmp
End Function

Public Function GetLabelID(StatLabel As String, DB As nssDatabase) As Long
  Dim myRec As Recordset

  Set myRec = DB.DB.OpenRecordset("STATLABEL", dbOpenSnapshot)
  With myRec
    .FindFirst "StatLabel='" & StatLabel & "'"
    If .NoMatch Then .FindFirst "StatisticLabel='" & StatLabel & "'"
    If Not .NoMatch Then
      GetLabelID = .Fields("StatisticLabelID")
    Else
      GetLabelID = -1
    End If
  End With
End Function

