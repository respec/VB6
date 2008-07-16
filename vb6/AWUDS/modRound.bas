Attribute VB_Name = "modRound"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A118FC90372"
Option Explicit
'##MODULE_NAME modRound (RetrievalEngine.dll)
'##MODULE_AUTHOR Todd W. Augenstein, Steve Tessler, and Ken Getz
'##MODULE_DATE 1998
'##MODULE_SUMMARY Rounding routines.
'##MODULE_REMARKS The rounding routines were specifically written to meet the Water Use requirements for rounding. The requirements state that stored values on output can be rounded based on the number of significant figures or can be rounded to a specified number of decimal places.</p><p> SWUDS does not use the NWIS common rounding arrays or NWIS common rounding routines.</p><p> On input SWUDS stores data values as both character and double percision float. The input routines call GetSg to determine the number of significant figures in a number. The number of significant figures is stored with the data values.
'##MODULE_REMARKS The original values are store exactly as the user _
entered them using attributes defined as varchar(15). GetSg is called to get the number of significant figures which is stored in the attributes defined as char(1). The original value is converted to million gallons per day and is stored in attributes defined as float8. On output the original values are never rounded. The stored float8 values are rounded using Round or RoundSg. </p><p> Attributes Used with Rounding Routines: <ul> <li>CN_QNTY_YR_##.cn_qnty_yr_va - varchar(15) N Yearly conveyance quantity value. The yearly conveyance quantity value is expressed in the units of the original measurement.</li> <li>CN_QNTY_YR_##.cn_qnty_yr_mgd_va - float8 N Yearly conveyance quantity value expressed in millions of gallons per day (MGD).</li> <li>CN_QNTY_YR_##.cn_qnty_yr_mgd_sg - char(1) N Yearly conveyance quantity value expressed in millions of gallons per day (MGD) significant figures.</li> <li>CN_QNTY_MO_##.cn_qnty_mo_va - varchar(15) Y Monthly _
 raw conveyance quantity value.</li> _
 <li>CN_QNTY_MO_##.cn_qnty_mo_mgd_va - float8 Y Monthly preferred conveyance quantity value.</li> <li>CN_QNTY_MO_##.cn_qnty_mo_mgd_sg - char(1) Y Monthly preferred conveyance quantity significant figures.</li> <li>CN_QNTY_AGGR_##.cn_qnty_aggr_va - varchar(15) Yearly incomplete transfer of water conveyance quantity value.</li> <li>CN_QNTY_AGGR_##.cn_qnty_aggr_mgd_va - float8 Y Yearly incomplete transfer of water conveyance quantity value.</li> <li>CN_QNTY_AGGR_##.cn_qnty_aggr_mgd_sg - char(1) Y Yearly incomplete transfer of water conveyance quantity significant figures.</li> <li>SITE_QNTY_##.site_qnty_yr_va - varchar(15) N Yearly site quantity value.</li> <li>SITE_QNTY_##.site_qnty_yr_mgd_va - float8 N Site quantity yearly MGD value. A site-based quantity value expressed in millions of gallons per day.</li> <li>SITE_QNTY_##..site_qnty_yr_mgd_sg - char(1) N Site quantity yearly MGD value significant figures.</li> <li>SITE_QNTY_MO_##.site_qnty_mo_va - varchar(15) Y Monthly ra _
w site quantity value.</li> <li> _
SITE_QNTY_MO_##.site_qnty_mo_mgd_va - float8 Y Monthly preferred site quantity value.</li> <li>SITE_QNTY_MO_##.site_qnty_mo_mgd_sg - char(1) Y Monthly preferred site quantity significant figures.</li></ul>
Public Function RoundDc(ByVal dblNumber As Double, _
                        ByVal intDecimals As Integer, _
                        Optional ByVal UseUSGSRounding As Boolean = True, _
                        Optional ByVal PadZeros As Boolean = True) As String
Attribute RoundDc.VB_Description = "Rounds to a specified number of decimal places."
'##SUMMARY Rounds to a specified number of decimal places.
'##DATE April 20, 2004
'##AUTHOR Todd Augenstein
'##PARAM dblNumber (I) Value to be rounded.
'##PARAM intDecimals (I) Number of decimals to round to. <ul> _
 <li>Positive integer 1 through 14, round to number of digits to the left of decimal; i.e. 2 places (51 rounds up to 100) </li> _
 <li>Negative integer -1 through -14, round to number of digits to the right of decimal; i.e. -1 places (0.05 rounds up to 0.1) </li> _
 <li>Zero - zero places, round to single digits e.g. 0 places rounds .9 to 1.
'##PARAM PadZeros (I) Pad with zeros flag. If TRUE then the resulting number _
         will be padded with zeros to the right of the last digit in order _
         to fill intDecimals e.g. if intDecimals = 2 and the number rounds _
         to 1.2 then the number returned will be 1.20.
'##PARAM UseUSGSRounding (I) USGS rounding flag. When set to TRUE then _
         rounding occurs using USGS rounding rules. _
         When set to FALSE rounding will occur by rounding up if _
         the first of the discarded digits is 5 or greater.
'##REMARKS The value is rounded accuratly to 14 significant figures. _
           If more than 14 significant figures are in dblNumber the _
           number returned from this routine will only be significant to _
           14 digits.</p>
'##RETURNS String, the double values as rounded and converted to a string _
           ready to be displayed as text on output.
'##ON_ERROR If something causes a runtime error the routine returns the word "overflow".
'##HISTORY PR 14096 Zeros will be output in the format as was originally _
           entered by the user when the user outputs _
           original values (e.g. 0, 0.0, 0.00 etc.). _
           However, zeros will be output as the value 0 when _
           the user outputs stored rounded or un-rounded values.
'##HISTORY PR 14055 A rounding option was added so that the user has an option _
           to use Survey based rounding. If the number to be rounded is a 5 then _
           if any numbers follow the five then the number rounds up. If _
           there are no numbers following the 5 and the number to the left of _
           the 5 is an even number then the number does not round up. _
           1.5 rounds 2 _
           2.5 rounds 2 _
           2.51 rounds 3 _
           .5 rounds 1 _
           .51 rounds 1


Dim tempString As String       ' Copy of dblNumber as string
Dim strArr(1 To 26) As String  ' Copy of dblNumber as array
Dim n As Integer               ' Loop index
Dim absDblNumber As Double     ' Copy of the dblNumber as absolute
Dim strChar As String          ' A single character of TempString being tested
Dim decimalFound As Boolean    ' TRUE when the decimal character was found.
Dim decimalPosition As Integer ' Position of the decimal if it exists.
Dim lastSigLocation As Integer ' Position of the last significant digit.
Dim testLocation As Integer    ' Position of a character being tested for rounding.
Dim errMsg As String           ' Holds an error message to be displayed to _
                                 the user. Use this in debugging. For the _
                                 runtime code output the word "overflow" _
                                 if an error occurs.
Dim IsNine As Boolean          ' TRUE if the first significant figure was _
                                 a nine.
On Error GoTo errRoundDc
'
' Check for valid arguments.
'
   If intDecimals < -14 Or intDecimals > 14 Then
     errMsg = "Runtime error: Attemped to round value on outside the range of -14 through 14"
     GoTo errRoundDc
   End If
    '
    ' Test for zero value
    '
    ' PR 14096 Zeros will be output as the value 0 when _
               the user outputs stored rounded or un-rounded values.
    '
    '  Email message from John P.Nawyn
    '  U.s.Geological Survey
    '  810 Bear Tavern Road
    '  West Trenton, NJ 08628
    '
    ' "I checked with the District and the consensus
    ' is that the "0" should appear (as default) without
    ' a decimal point.  When you report data in a table,
    ' there is no "0." the decimal point is absent."
    '
    If dblNumber = 0 Then
       RoundDc = "0"
       Exit Function
    End If
'
' Force an exponential number back to a text number.
'
    absDblNumber = Abs(dblNumber)
    '
    ' The number can be no larger than 1000000000000000000000000
    '
    If absDblNumber >= 1E+25 Then GoTo errRoundDc
    
    tempString = Format(absDblNumber, "#0.########################")
    '
    ' The number can not get truncated to zero.
    '
    If CDbl(tempString) = 0 Then GoTo errRoundDc
    
'
' Add a decimal point if it does not exist.
'
    decimalPosition = InStr(1, tempString, ".")
    If decimalPosition = 0 Then
       tempString = tempString & "."
       decimalPosition = Len(tempString)
    End If
    '
    ' The number has too many digits to preform this rounding algorithum.
    '
    If Len(tempString) > 26 Then GoTo errRoundDc
    
'
' Fill a 26 character array with the digits that comprise the number.
' If the start position of the mid function is set to a number
' greater than the number of characters in string,
' it returns an empty string (""); right pad with zeros
'
    
    For n = 1 To 26
       strChar = Mid(tempString, n, 1)
       
       If strChar = "" Then
          strArr(n) = "0"
       Else
          strArr(n) = strChar
       End If
    Next n
    
    '
    ' Get the last significant position as
    ' the number of decimal places the user requested.
    '
    If intDecimals < 0 Then
        lastSigLocation = decimalPosition - intDecimals
    Else
        lastSigLocation = decimalPosition - intDecimals - 1
    End If
    '
    ' Get the TestLocation for rounding.
    '
    testLocation = lastSigLocation + 1
    If lastSigLocation = 0 Then lastSigLocation = 1
    '
    ' If the test location is less than 1 then there is nothing
    ' to round. Rounds to zero.  e.g. round 3 to intDecimals = 2 = zero
    '
    If testLocation < 1 Then
       RoundDc = "0"
       Exit Function
    '
    ' If the testlocation is in position 1 then the first significant figure
    ' is at position 0 which does not exist. If also means the user is
    ' rounding a whole number and we have to handle rounding 51 to say 100.
    ' So append a leading zero so we can round up if we have too.
    '
    ElseIf testLocation = 1 Then
        testLocation = testLocation + 1
        decimalPosition = decimalPosition + 1
        '
        ' Append a leading zero
        '
        For n = 26 To 2 Step -1
           strArr(n) = strArr(n - 1)
        Next n
        strArr(1) = "0"
    End If
    
    '
    ' Skip over any decimal.
    '
    If testLocation = decimalPosition Then testLocation = testLocation + 1
'
' Process rounding.
'
'
    '
    ' The number can not be smaller than a number that can be formatted
    ' with the format #0.######################## when considering how
    ' many significant figures must be processed.
    ' The location of the last significant figure can not exceed the
    ' 23th decimal location to the right of the decimal.
    ' ps: the 23th decimal location is the 25 array element.
    '
    If lastSigLocation > 25 Then GoTo errRoundDc
    
    If Not testLocation > 26 Then
     '
     ' See if the value begins with nine. If the value begins with
     ' a nine there are cases where we have to round up to the next
     ' highest decimal place.
     '
     If strArr(1) = "9" Then
        IsNine = True
     Else
        IsNine = False
     End If
     
     Select Case strArr(testLocation)
            Case "6", "7", "8", "9"
                 '
                 ' If the first if the discarded digits is
                 ' greater than 5 then add 1 to the nth digit.
                 '
                  RoundUp strArr, lastSigLocation

            Case Is = "5"
                 '
                 ' Either do or not do USGS rounding.
                 '
                 If Not UseUSGSRounding Then
                       RoundUp strArr, lastSigLocation
                       GoTo Rounded
                 End If
                 '
                 ' Special USGS rounding rules:
                 '
                 ' If 5 is followed by any non zero digits then
                 ' round up.
                 '
                 For n = testLocation + 1 To 26
                    If strArr(n) <> "0" Then
                       If strArr(n) <> "." Then
                          RoundUp strArr, lastSigLocation
                          GoTo Rounded
                       End If
                    End If
                 Next n
                 '
                 ' If 5 is followed by all zero digits then
                 ' if the LastSigLocation number is odd then round up
                 '
                 Select Case strArr(lastSigLocation)
                        Case "1", "3", "5", "7", "9"
                              
                        RoundUp strArr, lastSigLocation
                 End Select
Rounded:
                 
            Case Else
                 '
                 ' If the first if the discarded digits is
                 ' less than 5 then leave the nth digit unchanged.
                 '
            
            End Select
         End If

'
' Rebuild final number for output.
'
    tempString = ""
    
    '
    ' Replace trailing zeros with blank spaces.
    '
    If Not PadZeros Then
    
       If decimalPosition < lastSigLocation Then
            For n = lastSigLocation To decimalPosition Step -1
               If strArr(n) <> "0" Then
                  If strArr(n) = "." Then strArr(n) = ""
                  GoTo ZerosRemoved
               End If
               
               strArr(n) = ""
            Next n
ZerosRemoved:
       End If
    End If
    
    If decimalPosition < lastSigLocation Then
       
       ' Its a fraction.
              For n = 1 To lastSigLocation
          tempString = tempString & strArr(n)
       Next n
       
    Else
       
       'Its a whole number.
       For n = 1 To lastSigLocation
          tempString = tempString & strArr(n)
       Next n
       
       For n = lastSigLocation + 1 To decimalPosition - 1
          tempString = tempString & "0"
       Next n
    End If
    '
    ' Add an extra 1 if round went to the next place.
    '
     If IsNine Then
        If strArr(1) = "0" Then
           tempString = "1" & tempString
        End If
     End If
    '
    ' Add the negative sign back.
    '
    If CDbl(tempString) = 0 Then
       RoundDc = 0
    
    ElseIf dblNumber < 0 Then
       RoundDc = "-" & tempString
       
       Else
       RoundDc = tempString
    End If
    

   Exit Function
    
errRoundDc:
   If errMsg <> "" Then
      MsgBox errMsg
   End If
   RoundDc = "overflow"
      
End Function

Public Function RoundSg(ByRef dblNumber As Double, _
                        ByRef strSg As String, _
                        Optional ByRef UseUSGSRounding As Boolean = True, _
                        Optional ByRef RetainDecimalPrecision = True) As String
Attribute RoundSg.VB_Description = "Rounds to a specified number of significant figures."
'##SUMMARY Rounds to a specified number of significant figures.
'##DATE April 19, 2004
'##AUTHOR Todd Augenstein
'##PARAM dblNumber (I) Value to be rounded.
'##PARAM strSg (I) Text version of the SWUDS significant _
         figure code (1,2,3,4,5,6,7,8,9,0,:,;,&lt,=,&gt,?,@)
'##PARAM UseUSGSRounding (I) USGS rounding flag. When set to TRUE then _
         rounding occurs using USGS rounding rules. _
         When set to FALSE rounding will occur by rounding up if _
         the first of the discarded digits is 5 or greater.
'##PARAM RetainDecimalPrecision (I) If TRUE then the precision of the _
         decimal place is retained when rounding a value < 0 to the _
         next whole number e.g. 0.95 with 1 significant figure will round to 1.0.
'##REMARKS Reads the significant figures field and SWUDS and uses the _
           number of significant figures to round the input value _
           (dblNumber) too.
'##RETURNS String, the double values as rounded and converted to a string _
           ready to be displayed as text on output.
'##COMMENTS_FROM Source="StandardModules~modRound~Round", Filter="##ON_ERROR"
'##SEEALSO Target="StandardModules~modRound~GetSg", Caption="GetSg"
'##HISTORY PR 14096 Zeros will be output in the format as was originally _
           entered by the user when the user outputs _
           original values (e.g. 0, 0.0, 0.00 etc.). _
           However, zeros will be output as the value 0 when _
           the user outputs stored rounded or un-rounded values.
'##HISTORY PR 14055 RoundSg was rewritten to use the Survey method for _
           rounding as documented _
           in Suggestions to Authors of the Reports of the United States _
           Geogloical Survey, 1991, pages 119-120, Rounding Off Numbers. _
'##ON_ERROR If the number can not be rounded and displayed correctly then _
            the word "overflow" is returned.

Dim tempString As String       ' Copy of dblNumber as string
Dim strArr(1 To 26) As String  ' Copy of dblNumber as array
Dim n As Integer               ' Loop index
Dim absDblNumber As Double     ' Copy of the dblNumber as absolute
Dim strChar As String          ' A single character of TempString being tested
Dim decimalFound As Boolean    ' TRUE when the decimal character was found.
Dim decimalPosition As Integer ' Position of the decimal if it exists.
Dim lastSigLocation As Integer ' Position of the last significant digit.
Dim testLocation As Integer    ' Position of a character being tested for rounding.
Dim firstSigLocation As Integer ' Postion of the first significant digit.
Dim errMsg As String           ' Holds an error message to be displayed to _
                                 the user. Use this in debugging. For the _
                                 runtime code output the word "overflow" _
                                 if an error occurs.
Dim IsNine As Boolean          ' TRUE if the first significant figure was _
                                 a nine.
Dim IsZero As Boolean          ' TRUE if the first digit to left of decimal is a zero.
Dim intSg As Integer           ' Integer version of number of significant figures.
On Error GoTo errRoundSg
    '
    ' Test for zero value
    '
    ' PR 14096 Zeros will be output as the value 0 when _
               the user outputs stored rounded or un-rounded values.
    '
    '  Email message from John P.Nawyn
    '  U.s.Geological Survey
    '  810 Bear Tavern Road
    '  West Trenton, NJ 08628
    '
    ' "I checked with the District and the consensus
    ' is that the "0" should appear (as default) without
    ' a decimal point.  When you report data in a table,
    ' there is no "0." the decimal point is absent."
    '
    If dblNumber = 0 Then
       RoundSg = "0"
       Exit Function
    End If
    '
    ' Get an integer version of the
    ' number of significant figures.
    '
    intSg = Asc(strSg) - 48
    If Not intSg < 17 Then
         errMsg = "Runtime error: bad significant figure value: " & strSg
         GoTo errRoundSg
    End If
'
' A negative sign, positive sign, and the digit 0 to the
' left of any whole number or a single 0 left of the
' the decimal point is not considered significant.
' Find the first nonzero digit (1-9) or the decimal
' point.
'
'
' Force an  exponential number back to a text number.
'
    absDblNumber = Abs(dblNumber)
    '
    ' The number can be no larger than 1000000000000000000000000
    '
    If absDblNumber >= 1E+25 Then GoTo errRoundSg
    
    tempString = Format(absDblNumber, "#0.########################")
    '
    ' The number can not get truncated to zero.
    '
    If CDbl(tempString) = 0 Then GoTo errRoundSg
    
'
' Add a decimal point if it does not exist.
'
    decimalPosition = InStr(1, tempString, ".")
    If decimalPosition = 0 Then
       tempString = tempString & "."
       decimalPosition = Len(tempString)
    End If
    '
    ' The number has too many digits to preform this rounding algorithum.
    '
    If Len(tempString) > 26 Then GoTo errRoundSg
'
' Fill a 26 character array with the digits that comprise the number.
' If the start position of the mid function is set to a number
' greater than the number of characters in string,
' it returns an empty string (""); right pad with zeros
'
    
    For n = 1 To 26
       strChar = Mid(tempString, n, 1)
       
       If strChar = "" Then
          strArr(n) = "0"
       Else
          strArr(n) = strChar
       End If
    Next n
'
' Find the position of the first significant figure.
'
    For n = 1 To 26
       If strArr(n) <> "0" And strArr(n) <> "." Then
          firstSigLocation = n
          GoTo FoundFirstSigLocation
       End If
    Next n
FoundFirstSigLocation:
'
' Get the last sigfig position if the decimal place is
' to the left of the first sigfig.
'
    If decimalPosition < firstSigLocation Then
        lastSigLocation = firstSigLocation + intSg - 1
        
    '
    ' The following must be true DecimalPosition => FirstSigLocation
    ' so see if the decimal position falls after the number of
    ' significant figures.
    '
    ElseIf decimalPosition > intSg Then
        lastSigLocation = firstSigLocation + intSg - 1
    
    '
    ' The decimal must be somewhere between a string of numbers.
    '
    Else
        lastSigLocation = firstSigLocation + intSg
    End If
    
    testLocation = lastSigLocation + 1
    If testLocation = decimalPosition Then testLocation = testLocation + 1
'
' Process rounding.
'
'
    '
    ' The number can not be smaller than a number that can be formatted
    ' with the format #0.######################## when considering how
    ' many significant figures must be processed.
    ' The location of the last significant figure can not exceed the
    ' 23th decimal location to the right of the decimal.
    ' ps: the 23th decimal location is the 25 array element.
    '
    If lastSigLocation > 25 Then GoTo errRoundSg
    
    If Not testLocation > 26 Then
     '
     ' See if the value begins with nine. If the value begins with
     ' a nine there are cases where we have to round up to the next
     ' highest decimal place.
     '
     If strArr(1) = "9" Then
        IsNine = True
     Else
        IsNine = False
     End If
     '
     ' See if the value begins with "0."
     '
     If strArr(1) = "0" And strArr(2) = "." Then
        IsZero = True
     Else
        IsZero = False
     End If
     
     Select Case strArr(testLocation)
            Case "6", "7", "8", "9"
                 '
                 ' If the first if the discarded digits is
                 ' greater than 5 then add 1 to the nth digit.
                 '
                  RoundUp strArr, lastSigLocation

            Case Is = "5"
                 '
                 ' Either do or not do USGS rounding.
                 '
                 If Not UseUSGSRounding Then
                       RoundUp strArr, lastSigLocation
                       GoTo Rounded
                 End If
                 '
                 ' Special USGS rounding rules:
                 '
                 ' If 5 is followed by any non zero digits then
                 ' round up.
                 '
                 For n = testLocation + 1 To 26
                    If strArr(n) <> "0" Then
                       If strArr(n) <> "." Then
                          RoundUp strArr, lastSigLocation
                          GoTo Rounded
                       End If
                    End If
                 Next n
                 '
                 ' If 5 is followed by all zero digits then
                 ' if the LastSigLocation number is odd then round up
                 '
                 Select Case strArr(lastSigLocation)
                        Case "1", "3", "5", "7", "9"
                              
                        RoundUp strArr, lastSigLocation
                 End Select
Rounded:
                 
            Case Else
                 '
                 ' If the first if the discarded digits is
                 ' less than 5 then leave the nth digit unchanged.
                 '
            
            End Select
         End If

'
' Rebuild final number for output.
'
    tempString = ""
    If decimalPosition < lastSigLocation Then
       '
       ' Its a fraction
       '
       ' Check to see if the number rounded up to
       ' a number = or > 1.
       '
       If Not RetainDecimalPrecision Then
          If IsZero Then
              If strArr(1) = "1" Then
                 If intSg = 1 Then
                    lastSigLocation = 1
                 Else
                    lastSigLocation = lastSigLocation - 1
                 End If
              End If
          End If
       End If
       
       
       For n = 1 To lastSigLocation
          tempString = tempString & strArr(n)
       Next n
       
    Else
       'Its a whole number.
       For n = 1 To lastSigLocation
          tempString = tempString & strArr(n)
       Next n
       
       For n = lastSigLocation + 1 To decimalPosition - 1
          tempString = tempString & "0"
       Next n
    End If
    '
    ' Add an extra 1 if round went to the next place.
    '
     If IsNine Then
        If strArr(1) = "0" Then
           tempString = "1" & tempString
        End If
     End If
    '
    ' Add the negative sign back.
    '
    If dblNumber < 0 Then
       RoundSg = "-" & tempString
    Else
       RoundSg = tempString
    End If
    
    Exit Function
    
errRoundSg:
    If errMsg <> "" Then
       MsgBox errMsg, vbCritical
    End If
    RoundSg = "overflow"
End Function
Private Sub RoundUp(ByRef strArr() As String, _
                    ByRef lastSigLocation As Integer)
Attribute RoundUp.VB_Description = "Recursive routine that rounds a number up starting at the specified array location. The routine assumes the starting array position is an ASCII number from 0-9 (ASCII characters 48-57).The routine also assumes the value in strArr needs to be rounded up."
'##SUMMARY Recursive routine that rounds a number up starting _
           at the specified array location. The routine assumes _
           the starting array position is an ASCII number from _
           0-9 (ASCII characters 48-57).The routine also assumes _
           the value in strArr needs to be rounded up.
'##DATE April 19, 2004
'##AUTHOR Todd Augenstein
'##PARAM strArr (M) 26 single character array containing the _
                    number to be rounded. The array contains _
                    a positive, left justified value. For example, 63.5 _
                    is stored as: <br><br> _
                    strArr(1) = "6" <br> _
                    strArr(2) = "3" <br> _
                    strArr(3) = "." <br> _
                    strArr(4) = "5" <br> _
                    strArr(5) = "0" <br> _
                    strArr(6) = "0" <br> _
                    strArr(7) = "0" <br> _
                    strArr(8 through 26) = "0"
'##PARAM LastSigLocation (I) Position in strArr where rounding _
         shoud begin. All digits to position LastSigLocation + 1 _
         are discarded. Rounding starts at postion LastSigLocation.
'##REMARKS Round up will always round up. Rounding occurs by adding _
           1 to position LastSigLocation if postion _
           LastSigLocation => 5-8 and changes the value of _
           postion LastSigLocation to "0" if the value was 9. _
           The complete array is rounded; the _
           routine calls itself in order to round until _
           LastSigLocation = 0. Array strArr contains the modified _
           rounded number. If the number is rounded to the next _
           place (9.5 rounds to 10.0) then strArr(1) will contain _
           the value "0". The calling routine must append the _
           preceeding digit, 1.
           
Dim n As Integer  ' Used to get Ascii collating sequence
Dim i As Integer  ' Index to loop
Dim NextLocation As Integer  ' The next position to round using recursion
     
     '
     ' No other position to round. Pop the stack.
     '
     If lastSigLocation = 0 Then Exit Sub
     
     '
     ' Get the ASCII sequence number for the postion to
     ' round.
     '
     n = Asc(strArr(lastSigLocation))
     '
     ' If the number is less than nine (57 is the Ascii
     ' character for nine) then round up and exit.
     ' The recursion can be stopped because our resulting
     ' number does not cause the next position to the
     ' left in strArr to need to be rounded.
     '
     If n < 57 Then
        strArr(lastSigLocation) = Chr(n + 1)
     Else
        '
        ' If the number was nine then we need to
        ' assign the position the value 0 and then
        ' optionally call this routine to round up
        ' the next number.
        '
        strArr(lastSigLocation) = "0"
        '
        ' If we are at the last position then just
        ' exit; pop the stack.
        '
        If lastSigLocation - 1 <= 0 Then Exit Sub
        '
        ' Else round the next number to the left as long as
        ' the next number is not the decimal point.
        ' Ignore a decimal point. Round by calling this
        ' routine with the next position to the left to
        ' be rounded.
        '
        For i = lastSigLocation - 1 To 1 Step -1
           If strArr(i) <> "." Then
               NextLocation = i
               GoTo getnext
           End If
        Next i
        
        '
        ' No need for further recursion; exit.
        '
        Exit Sub
        
getnext:
        RoundUp strArr(), NextLocation
     End If
     
End Sub

Public Function GetSg(ByVal varNumber As String) As String
Attribute GetSg.VB_Description = "Determines number of significant figures in String"
'##SUMMARY Determines number of significant figures in String
'##DATE 1998
'##AUTHOR Todd Augenstein
'##PARAM varNumber (I) String containing a Number.
'##REMARKS<p>Leading sign + or - are not counted and the decimal is _
not counted as significant digits.</p> _
<p>Text pasted and modified from (10/25/1999):<br> _
http://www.umassd.edu/1Academic/CArtsandSciences/Chemistry/Catalyst/sf.html</p> _
<p>Guidelines for determining the number of significant figures.</p><p> _
1. All nonzero digits are significant.<br> _
2. Zeros between nonzero digits are significant.<br> _
3. Zeros to the left of the first nonzero digit are not significant.<br> _
4. Zeros that fall to the right of a decimal point are significant, _
   but terminal zeroes in very large numbers are usually not _
   significant, as discussed below.<br></p> _
<p>Examples<br> _
<p>2.456     4 significant figures<br> _
1003.2    5<br> _
1.03000   6<br> _
0.0000402 3<br> _
230000    2 - 6 (The SWUDS software assumes 6)</p>
'##RETURNS Blank string, if there was an error, 1 if the varNumber is zero (0 always has 1 sig fig.), or _
number of significant figures in varNumber represented as a single character code: </p><p> _
<table border="1" width="225"> <tr> <td width="107"><p> Number of Significant Figures&nbsp;&nbsp;&nbsp;</p> _
</td> <td width="102"><p> Value of Significant Figure code</p> </td> </tr> <tr> <td width="107"><p> _
0&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; 0</p> </td> </tr> <tr> <td width="107"><p> _
1&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; 1</p> </td> </tr> <tr> <td width="107"><p> _
2&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; 2</p> </td> </tr> <tr> <td width="107"><p> _
3&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; 3</p> </td> </tr> <tr> <td width="107"><p> _
4&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; 4</p> </td> </tr> <tr> <td width="107"><p> _
5&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; 5</p> </td> </tr> <tr> <td width="107"><p> _
6&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; 6</p> </td> </tr> <tr> <td width="107"><p> _
7&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; 7</p> </td> </tr> <tr> <td width="107"><p> _
8&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; 8</p> </td> </tr> <tr> <td width="107"><p> _
9&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; 9</p> </td> </tr> <tr> <td width="107"><p> _
10&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; :</p> </td> </tr> <tr> <td width="107"><p> _
11&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; ;</p> </td> </tr> <tr> <td width="107"><p> _
12&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; &lt;</p> </td> </tr> <tr> <td width="107"><p> _
13&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; =</p> </td> </tr> <tr> <td width="107"><p> _
14&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; ></p> </td> </tr> <tr> <td width="107"><p> _
15&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; ?</p> </td> </tr> <tr> <td width="107"><p> _
16&nbsp;&nbsp;&nbsp;</p> </td> <td width="102"><p>&nbsp; @</p> </td> </tr> </table>
'##ON_ERROR If the input string number is not a valid number, or the _
input precision can not be determined, this module returns _
a null value (e.g. blank string). _
If something causes a runtime error the routine returns _
a null value (e.g. blank string).
'##SEEALSO Target="StandardModules~modRound~RoundSg", Caption="RoundSg"
'##TODO At some future release we could allow the entry of _
 exponential notation, then the software could correctly _
 count the number of significant numbers. But not at 4_1 </p><p> _
 Instead of    User would enter _
 3000000000. = 3.0 x 10**9 (two significant figures) </p><p> _
 Or we could allow a shortcut code for entry of _
 an overriding sig fig</p><p> _
 Instead of    User would enter _
 3000000000. = 3000000000/3 (meaning 3 significant figures)</p><p> _
 The /3 would override any computer calculated sig fig. _
 But not at 4_1, well this ones easy to do.
'##HISTORY PR 14096 In SWUDS zero values are considered exact numbers having an infinite _
           number of significant figures. SWUDS stores zero with 1 significant _
           figure.
           
On Error GoTo errGetSg

Dim NumSigFig As Long          ' Number of sig figs
Dim NumLeadingZeros As Long    ' Number of leading zeros
Dim tempString As String       ' Temp
Dim TestString As String       ' Temp
Dim TestDouble As Double       ' Tests for double
Const CodeList = "123456789:;<=>?@"
'
' See if the entered number is a number
'
   TestDouble = CDbl(varNumber)
'
' If the value is zero return 1.
'
   If TestDouble = 0 Then
      GetSg = 1
      Exit Function
   End If
'
' Zeros to the left of the first nonzero digit are not significant.
' Find first nonzero number
'
'
    NumLeadingZeros = 0
    tempString = Trim(varNumber)
    While True
       TestString = Left(tempString, 1)
       Select Case TestString
          Case ".", "-", "+":
          
          Case "0":
              NumLeadingZeros = NumLeadingZeros + 1
          Case Else:
              GoTo wendEnd1
       End Select
       tempString = Mid(tempString, 2)
       If tempString = "" Then GoTo wendEnd1
    Wend
wendEnd1:
'
' then just add up the number of digits
' in string.
'
    NumSigFig = 0
    While True
       TestString = Left(tempString, 1)
       Select Case TestString
          Case ".", "-", "+":
          
          Case Else:
              NumSigFig = NumSigFig + 1
       End Select
       tempString = Mid(tempString, 2)
       If tempString = "" Then GoTo wendEnd2
    Wend
wendEnd2:
'
    If NumSigFig > 0 And NumSigFig < 17 Then
      GetSg = Mid(CodeList, NumSigFig, 1)
    ElseIf NumLeadingZeros <> 0 Then
      GetSg = "1"
    ElseIf NumSigFig = 0 Then
      GetSg = "0"
    Else
 '     GetSg = "-1"
       GetSg = ""
    End If
    
    Exit Function
    
errGetSg:
    GetSg = ""
    
End Function

Public Function GetLeastDec(ByVal NumDecimals As Integer, _
                            ByVal varNumber As String, _
                            ByVal NumSigFigs As String) As Integer
Attribute GetLeastDec.VB_Description = "Determines the least number of decimal places in a number; the number of digits right of the decimal place."
'##SUMMARY Determines the least number of decimal places in a number; the number of _
           digits right of the decimal place.
'##AUTHOR Todd Augenstein
'##DATE 1998
'##PARAM NumDecimals (I) Number of decimals places to the right of the decimal.
'##PARAM varNumber (I) String containing Number.
'##PARAM NumSigFigs (I) Number of significant figures in the Number (varNumber).
'##REMARKS Determines the number of significant decimal places in varNumber based on the number of _
 significant figures NumSigFigs. Then compares the number of significant decimal places to the _
 number of decimal places entered, NumDecimals. Then the routine returns the smaller of _
 the number of decimal places entered or number of significant decimal places. </p><p> _
 For example, if NumDecimals was 3 and the actual number of significant decimal places was 6 based _
 on varNumber and NumSigFigs then the routine will return 3. </p><p> _
 The routine helps support rounding in algorithms that require addition and subtraction: _
 <p>When adding or subtracting the resulting number should contain _
 the same number of _
 decimal places as in the term with the least number _
 of decimal places.</p> _
 <p>Example:<br> _
 2.487 + 330.4 + 22.59 355.477 --><br> _
 round off to 335.5 (Uncertainty in tenths place)<br>
'##RETURNS Integer, least number of decimal places between the varNumber entered and the NumDecimals to test against.
'##ON_ERROR Returns -1 if there is a runtime error or the varNumber entered was not a number.
'##SEEALSO Target="StandardModules~modRound~GetDec", Caption="GetDec"
'##HISTORY PR 14096 In SWUDS zero values are considered exact numbers having an infinite _
           number of significant figures. The following routine should ignore zero values when determining _
           the number of decimal places to the right of the decimal place. Therefore, _
           a zero value (varNumber) should not effect the number of _
           decimal places (NumDecimals) that a resulting number should have when _
           zero is included in an add or subtraction operation. _
           The routine should not change the number of _
           significant decimal places of the resulting number _
           if zero is being used in a calculation. Thus for zero, NumDecimals is _
           returned, unchanged, as the number of decimal places.

On Error GoTo errGetLeastDec

Dim NumDec As Long             ' Number of decimals
'
' PR 14096 see if the entered number is zero; if so return NumDecimals unchanged.
'
   If CDbl(varNumber) = 0 Then
      GetLeastDec = NumDecimals
      Exit Function
   End If
    
   NumDec = modRound.GetDec(varNumber, NumSigFigs)
   If NumDec = -1 Then GoTo errGetLeastDec
   If NumDec < NumDecimals Then
     GetLeastDec = NumDec
   Else
     GetLeastDec = NumDecimals
   End If
    Exit Function
errGetLeastDec:
    GetLeastDec = "-1"
End Function

Public Function GetDec(ByVal varNumber As String, _
                       ByVal NumSigFigs As String) As Integer
Attribute GetDec.VB_Description = "Determines number of decimals in a number."
'##SUMMARY Determines number of decimals in a number.
'##DATE 1998
'##AUTHOR Todd Augenstein
'##PARAM varNumber (I) String containing Number.
'##PARAM NumSigFigs (I) Number of significant figures.
'##RETURNS Integer, number of decimal places in varNumber.
'##ON_ERROR Returns -1 if there was an error.

On Error GoTo errGetDec

Dim NumDec As Long             ' Number of decimals
Dim tempString As String       ' Temp
Dim TestDouble As Double       ' Tests for double
'
' See if the entered number is a number
'
   TestDouble = CDbl(varNumber)
'
' Find decimal
'
    tempString = Trim(modRound.RoundSg(TestDouble, NumSigFigs))
    While True
       Select Case Left(tempString, 1)
          Case "."
              tempString = Mid(tempString, 2)
              GoTo wendEnd1
       End Select
       tempString = Mid(tempString, 2)
       If tempString = "" Then GoTo wendEnd1
    Wend
wendEnd1:
'
' then pass back the length of TempString.
'
    GetDec = Len(tempString)
    Exit Function
errGetDec:
    GetDec = "-1"
End Function

Public Function GetNum(ByVal varNumber As String) As Long
Attribute GetNum.VB_Description = "Determines number of digits excluding leading zeros of the number part of a real number as string."
'##SUMMARY Determines number of digits excluding leading zeros _
           of the number part of a real number as string.
'##DATE 1998
'##AUTHOR Todd Augenstein
'##PARAM varNumber (I) Long containing number.
'##REMARKS Leading sign + or - are not counted and the decimal is _
not counted. Only the numbers to the left of the decimal _
are counted.</p><p> _
Leading zeros are ignored. i.e. 0.4 returns 0
'##RETURNS String, number of digits in number.
'##ON_ERROR Returns -1 if a runtime error occures or if the input string number is not a valid number, or the _
 input precision can not be determined.

On Error GoTo errGetNum

Dim NumDigits As Long          ' Number of digits
Dim LeadingZero As Boolean     ' True if leading zero
Dim tempString As String       ' Temp
Dim TestDouble As Double       ' Tests for double
Const CodeList = "123456789:;<=>?@"
'
' See if the entered number is a number
'
   TestDouble = CDbl(varNumber)
'
' Zeros to the left of the first nonzero digit are not significant.
' Find first nonzero number
'
    NumDigits = 0
    LeadingZero = True
    tempString = Trim(varNumber)
    While True
       Select Case Left(tempString, 1)
          Case "-", "+": ' do nothing
              
          Case "." ' exit
              GoTo wendEnd1
              
          Case "0": ' count zeros
              If Not LeadingZero Then
                NumDigits = NumDigits + 1
              End If
          Case Else: '
              LeadingZero = False
              NumDigits = NumDigits + 1
       End Select
       tempString = Mid(tempString, 2)
       If tempString = "" Then GoTo wendEnd1
    Wend
wendEnd1:
'
' Thats it !
'
    GetNum = NumDigits
Exit Function
errGetNum:
    GetNum = -1
End Function

