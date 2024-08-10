Attribute VB_Name = "Module1"
Global ORANGECONN_SERVER As New ADODB.Connection
Global ORANGECONN_LOCAL As New ADODB.Connection
Global ORANGECONN_SOURCE As New ADODB.Connection
Global ORANGECONN_TARGET As New ADODB.Connection
Global ORANGECONN_VERSION As New ADODB.Connection
Global GLB_UID As String
Global GLB_PWD As String
Global GLB_PWD_SELECT As String
Global GLB_DATABASENAME As String
Global GLB_CONNSTRING As String
Global GLB_CONNSTRING_VERINFO As String
Global GLB_CONNSTRING_CLIENTS As String
'Global REPORTS_CONNECTION As New Connection
Global GLB_BRANCH_CODE As Integer
Global GLB_COMPANY_NAME As String
Global GLB_COMPANY_ID As String
Global GLB_BRANCH_ADDRESS As String
Global GLB_GODOWN_NAME As String
Global GLB_SHORTNAME As String
Global GLB_GODOWN_TIN As String
Global GLB_GODOWN_MANAGER As String
Global GLB_GODOWN_MANAGER_PHONE As String
Global GLB_USER_PSWD As String
Global GLB_FINANCIAL_YEAR As String
Global GLB_YEAR_BEGINNING_DATE As String
Global GLB_YEAR_END_DATE As String
Global GLB_DATABASE_NAME As String
Global GLB_USER_NAME As String
Global GLB_SETOFF_ACNO As Long
Global GLB_CASH_ACNO As Long
Global GLB_BANK_ACNO As Long
Global GLB_CARD_ACNO As Long
Global GLB_CASHBOOKNO As Long
Global GLB_PAPERSIZE As Integer
Global GLB_NEW_CUST_ORDER As Boolean
Global GLB_CLEARING_DAYS As Integer
Global GLB_TAXAMT As Currency
Global GLB_DISCAMT As Currency
Global GLB_SPDISC_SUM As Currency
Global IP_ADDRESS As String
Global GLB_NEW_SALES As Boolean
Global GLB_NEW_PURCHASE As Boolean
Global GLB_MANFUNC_VISIBLE As String
Global GLB_VOUCHER_VISIBLE As String
Global GLB_INVENTORY_VISIBLE As String
Global GLB_ACCOUNTS_VISIBLE As String
Global GLB_PROFIT_VISIBLE As String
Global GLB_NEW_BUTTON_ENABLE As String
Global GLB_MODIDEL_BUTTON_ENABLE As String
Global GLB_EXCEL_BUTTON_ENABLE As String
Global GLB_GRIDVIEW_ENABLE As String
Global Const INVOICE_TABLE_CONST = "GoodsTransactions_Invoice"
Global GLB_VERSION_NUMBER As Integer
Global GLB_CGST_SGST_SALE As Integer
Global GLB_TAXCGST_SALE As Integer
Global GLB_TAXSGST_SALE As Integer
Global GLB_IGST_SALE As Integer
Global GLB_TAXIGST_SALE As Integer
Global GLB_CGST_SGST_PUR As Integer
Global GLB_TAXCGST_PUR As Integer
Global GLB_TAXSGST_PUR As Integer
Global GLB_IGST_PUR As Integer
Global GLB_TAXIGST_PUR As Integer
Global GLB_CGST_PRINT As Currency
Global GLB_SGST_PRINT As Currency
Global GLB_IGST_PRINT As Currency

Public Function NUMERICDOT(KeyAscii As Integer, CurserPos As Integer, TBox As TextBox) As Integer
Dim POS As Integer
Dim ln As Integer
Dim TMP As String
If Not ((KeyAscii > 47 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 46) Then
    MsgBox "Key Pressed Is Not Numeric. Enter Numeral Only.....", vbCritical, "Message"
    NUMERICDOT = 0
Else
    If InStr(TBox.Text, ".") = 0 Then
        If KeyAscii <> 46 Then
            NUMERICDOT = KeyAscii
        Else
            TMP = Mid(TBox.Text, CurserPos + 1)
            ln = Len(TMP)
            If ln > 2 Then
                NUMERICDOT = 0
            Else
                NUMERICDOT = KeyAscii
            End If
        End If
    Else
        If KeyAscii = 46 Then
            NUMERICDOT = 0
        ElseIf KeyAscii = 8 Then
            NUMERICDOT = KeyAscii
        Else
            POS = InStr(TBox.Text, ".")
            If CurserPos >= POS Then
                TMP = Mid(TBox.Text, POS + 1)
                ln = Len(TMP)
                If ln >= 2 Then
                    NUMERICDOT = 0
                Else
                    NUMERICDOT = KeyAscii
                End If
            Else
                NUMERICDOT = KeyAscii
            End If
        End If
    End If
End If
End Function

Public Function NUMERICDOTNEW(KeyAscii As Integer, CurserPos As Integer, TBox As TextBox) As Integer
Dim POS As Integer
Dim ln As Integer
Dim TMP As String
If Not ((KeyAscii > 47 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 46) Then
    MsgBox "Key Pressed Is Not Numeric. Enter Numeral Only.....", vbCritical, "Message"
    NUMERICDOTNEW = 0
Else
    If InStr(TBox.Text, ".") = 0 Then
        If KeyAscii <> 46 Then
            NUMERICDOTNEW = KeyAscii
        Else
            TMP = Mid(TBox.Text, CurserPos + 1)
            ln = Len(TMP)
            If ln > 4 Then
                NUMERICDOTNEW = 0
            Else
                NUMERICDOTNEW = KeyAscii
            End If
        End If
    Else
        If KeyAscii = 46 Then
            NUMERICDOTNEW = 0
        ElseIf KeyAscii = 8 Then
            NUMERICDOTNEW = KeyAscii
        Else
            POS = InStr(TBox.Text, ".")
            If CurserPos >= POS Then
                TMP = Mid(TBox.Text, POS + 1)
                ln = Len(TMP)
                If ln >= 4 Then
                    NUMERICDOTNEW = 0
                Else
                    NUMERICDOTNEW = KeyAscii
                End If
            Else
                NUMERICDOTNEW = KeyAscii
            End If
        End If
    End If
End If
End Function
Public Function restrictTextBoxToLength(ByRef txtBox As TextBox, kAscii As Integer, Optional restrictLength As Integer = 10) As Integer
    
    If Len(txtBox.Text) >= restrictLength And kAscii <> 8 And kAscii <> 13 And kAscii <> 9 Then
        kAscii = 0
    End If
    restrictTextBoxToLength = kAscii

End Function
Public Function CAPITAL(key As Integer) As Integer
If key = 34 Or key = 39 Or key = 124 Or key = 92 Then
    key = 0
End If

If key > 96 And key <= 122 Then
    key = key - 32
End If
CAPITAL = key
End Function
Public Function REMOVESLASH(dt As String) As String
Dim dte_text, DD, MM, YY, GDATE3 As String
Dim ln, i As Integer
    dte_text = ""
    ln = Len(dt)
    For i = 1 To ln
        If Not (Mid(dt, i, 1) = "/" Or Mid(dt, i, 1) = "-") Then
            dte_text = dte_text + Mid(dt, i, 1)
        End If
    Next
    dt = dte_text
    REMOVESLASH = dt
Exit Function
End Function

Public Function NTOD(INDATE As Long) As String
    YR = Mid(CStr(INDATE), 1, 4)
    MTH = Mid(CStr(INDATE), 5, 2)
    DY = Mid(CStr(INDATE), 7, 2)
    NTOD = DY & "/" & MTH & "/" & YR
Exit Function
End Function
Public Function DTON(OUTDATE1) As Long
    YR = Mid(CStr(OUTDATE1), 7, 4)
'    MsgBox yr & "year"
    MTH = Mid(CStr(OUTDATE1), 4, 2)
'    MsgBox mth & "month"
    DY = Mid(CStr(OUTDATE1), 1, 2)
'    MsgBox dy & "day"
    YR1 = YR & MTH & DY
'    MsgBox Val(yr1) & "all"
    DTON = Val(YR1)
Exit Function
End Function

Public Function DISPLAYDATE(dt As String, Optional OperationMode As Integer = 1, Optional yearCheck As Integer = 0) As String
Dim DATE1, Date2, date3 As Date
Dim DD, MM, YY, GDATE3 As String
    dt = REMOVESLASH(dt)
    DD = Mid(dt, 1, 2)
    MM = Mid(dt, 3, 2)
    YY = Mid(dt, 5)
    GDATE3 = dt
    dt = DD + "/" + MM + "/" + YY
    If (Val(Len(GDATE3)) <> 8 Or (IsDate(dt) = False Or Val(MM) > 12)) Then
        If OperationMode = 1 Then
            MsgBox "Invalid Date Or Date Format Not Correct. Enter An Eight Digit DATE. Example.: 05092010 ddmmyyyy", vbInformation, "Message"
        End If
        DISPLAYDATE = ""
    Else
        DISPLAYDATE = dt
    End If
Exit Function
End Function
Public Function TWODECIMAL(DECIMAL_NUMBER) As Double
'MsgBox Len(DECIMAL_NUMBER)
'MsgBox Mid(DECIMAL_NUMBER, 1, InStr(1, DECIMAL_NUMBER, ".") - 1)
'MsgBox Mid(DECIMAL_NUMBER, InStr(1, DECIMAL_NUMBER, "."), 3)
If InStr(1, DECIMAL_NUMBER, ".") > 0 Then
    TWODECIMAL = Mid(DECIMAL_NUMBER, 1, InStr(1, DECIMAL_NUMBER, ".") - 1) & Mid(DECIMAL_NUMBER, InStr(1, DECIMAL_NUMBER, "."), 3)
ElseIf InStr(1, DECIMAL_NUMBER, ".") = 0 Then
    TWODECIMAL = DECIMAL_NUMBER
End If
End Function
Public Function NUMERIC(KeyAscii As Integer) As Integer
If Not ((KeyAscii > 47 And KeyAscii <= 57) Or KeyAscii = 13 Or KeyAscii = 8 Or KeyAscii = 9) Then
    MsgBox "Key Pressed Is Not Numeric. Enter Numeral Only.....", vbCritical, "Message"
    NUMERIC = 0
Else
    NUMERIC = KeyAscii
End If
End Function
Public Function NUMBFORM(DBamount As Variant) As String 'ravi
On Error GoTo myerror
   Dim ValueNeg As Boolean
   Dim DBamt As String
   DBamt = DBamount
   ValueNeg = False
   If Left(Trim(DBamt), 1) = "-" Then
    ValueNeg = True
    DBamt = Mid(Trim(DBamt), 2)
   End If
   
    CNTR = 4
      myst = DBamt
      dotpos = InStr(1, myst, ".")
      Deci = Mid(myst, dotpos)
      myst = Mid(myst, 1, dotpos - 1)
      If Len(Deci) = 1 Then
          Deci = Deci & "00"
      ElseIf Len(Deci) = 2 Then
          Deci = Deci & "0"
      End If
      strrev = StrReverse(myst)
      If Len(strrev) > 3 Then
          fnst = Mid(strrev, 1, 3) & ","
          dummyst = Replace(fnst, ",", "")
              Do While Len(dummyst) <= Len(strrev)
                 If Len(strrev) - Len(dummyst) > 2 Then
                    fnst = fnst & Mid(strrev, CNTR, 2) & ","
                    CNTR = CNTR + 2
                    dummyst = Replace(fnst, ",", "")
                 End If

                  If Len(strrev) - Len(dummyst) <= 2 Then
                    fnst = fnst & Mid(strrev, CNTR, 2)
                    Exit Do
                  End If

             Loop
      Else
         NUMBFORM = myst & Deci
         If ValueNeg = True Then
            NUMBFORM = "-" & NUMBFORM
         End If
         Exit Function
      End If

      NUMBFORM = StrReverse(fnst) & Deci
      If ValueNeg = True Then
          NUMBFORM = "-" & NUMBFORM
      End If
      Exit Function

myerror:
      strrev = StrReverse(DBamt) '(myst)
      If Len(strrev) > 3 Then
         fnst = Mid(strrev, 1, 3) & ","
         dummyst = Replace(fnst, ",", "")
             Do While Len(dummyst) <= Len(strrev)
                 If Len(strrev) - Len(dummyst) > 2 Then
                    fnst = fnst & Mid(strrev, CNTR, 2) & ","
                    CNTR = CNTR + 2
                    dummyst = Replace(fnst, ",", "")
                 End If
                 If Len(strrev) - Len(dummyst) <= 2 Then
                    fnst = fnst & Mid(strrev, CNTR, 2)
                    Exit Do
                  End If
             Loop
      Else
         NUMBFORM = myst & ".00"
         If ValueNeg = True Then
            NUMBFORM = "-" & NUMBFORM
         End If
         Exit Function
      End If
      NUMBFORM = StrReverse(fnst) & ".00"
      If ValueNeg = True Then
         NUMBFORM = "-" & NUMBFORM
      End If
 End Function
 
Public Function REMOVECOMA(onestring As String) As String
On Error GoTo ErrorHandler
    Dim tempstring As String
    Dim i As Integer
    tempstring = ""
    For i = 1 To Len(onestring)
        If Mid(onestring, i, 1) <> "," Then
            tempstring = tempstring + Mid(onestring, i, 1)
        End If
    Next
    REMOVECOMA = tempstring
Exit Function
ErrorHandler:
    Call logError("Module1", "Public Function removeComa", err, 0)
    
End Function
Public Function logError(FormName As String, FunctionName As String, ByRef ErrorObject As ErrObject, LogMode As Integer, Optional invoiceFileName As String)
'    Dim dt As Date
'    Dim filenum As Long
'    Dim TF As Boolean
'    Dim f As New FileSystemObject
'    Dim TextObj As TextStream
'    Dim ErrorFileName As String
'    Dim ErrorDirName As String
'    Dim fileName As String
'
'    If LogMode = 1 Then
'        ErrorReport.Label7 = ErrorObject.Number
'        ErrorReport.Label8 = ErrorObject.Description
'        ErrorReport.FEEDBACK = ""
'        ErrorReport.Show vbModal
'    End If
'
'    ErrorDirName = "\ErrorLog"
'    If Trim(fileName) = "" Then
'        ErrorFileName = "\Errors.log"
'    Else
'        ErrorFileName = "\" & invoiceFileName
'    End If
'    fileName = App.Path & ErrorDirName & ErrorFileName
'
'    If Not f.FolderExists(App.Path & ErrorDirName) Then
'        MkDir (App.Path & ErrorDirName)
'    End If
'    If f.FileExists(fileName) Then
'        Set TextObj = f.OpenTextFile(fileName, ForAppending, True, TristateMixed)
'    Else
'        Set TextObj = f.CreateTextFile(fileName, True, False)
'    End If
'
'    'TextObj.AtEndOfStream
'    TextObj.WriteBlankLines (2)
'    TextObj.Write ("--------------------------------------------")
'    TextObj.WriteBlankLines (1)
'    TextObj.Write ("date and time : " & Date$ & " , " & Time$)
'    TextObj.WriteBlankLines (1)
'    TextObj.Write ("Error Number : " & Str(ErrorObject.Number))
'    TextObj.WriteBlankLines (1)
'    TextObj.Write ("Error Description : " & ErrorObject.Description)
'    TextObj.WriteBlankLines (1)
'    TextObj.Write ("HelpContext : " & ErrorObject.HelpContext)
'    TextObj.WriteBlankLines (1)
'    TextObj.Write ("HelpFile : " & ErrorObject.HelpFile)
'    TextObj.WriteBlankLines (1)
'    TextObj.Write ("LastDllError : " & ErrorObject.LastDllError)
'    TextObj.WriteBlankLines (1)
'    TextObj.Write ("Source : " & ErrorObject.Source)
'    TextObj.WriteBlankLines (1)
'    TextObj.Write ("Form : " & FormName)
'    TextObj.WriteBlankLines (1)
'    TextObj.Write ("Function / Module: " & FunctionName)
'    TextObj.WriteBlankLines (1)
'    TextObj.Write ("Use Case : " & getUserFeedBack)
'    TextObj.WriteBlankLines (1)
'    TextObj.Write ("--------------------------------------------")
'    TextObj.Close
'
'    setUserFeedBack ("")
End Function
'Public Function CAPITAL(KeyAscii As Integer) As Integer
'If Not ((KeyAscii > 64 And KeyAscii <= 90) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 13) Then
'        MsgBox ("YOU HAVE ENTERED AN INVALID LETTER OR A LETTER IN LOWER CASE. CHANGE TO UPPERCASE BY PRESSING CapsLock KEY ONCE..")
'        CAPITAL = 0
'Else
'    CAPITAL = KeyAscii
'End If
'End Function

Public Function getRecordSet(tableName As String) As ADODB.Recordset
On Error GoTo err
    Dim generalRecord As New ADODB.Recordset
    Dim con As ADODB.Connection
    
'    Set con = REPORTSCONN()
    
    generalRecord.Open "delete from GOODSTRANSACTIONS_INVOICE", REPORTSCONN, adOpenDynamic, adLockOptimistic
    generalRecord.Open "select * from GOODSTRANSACTIONS_INVOICE", REPORTSCONN, adOpenDynamic, adLockOptimistic
    
    Set getRecordSet = generalRecord
Exit Function
err:
    Call logError("Module1", "Public Function getRecordSet", err, 0)
    Set getRecordSet = Nothing
End Function
Public Function NUMTOWORDS(ByVal MyNumber)

Dim temp

         Dim Rupees, Paise

         Dim DecimalPlace, count



         ReDim Place(9) As String

         Place(2) = " Thousand "

         Place(3) = " Lakh "

         Place(4) = " Crore "

         Place(5) = " Billion "



         ' Convert MyNumber to a string, trimming extra spaces.

         MyNumber = Trim(Str(MyNumber))



         ' Find decimal place.

         DecimalPlace = InStr(MyNumber, ".")



         ' If we find decimal place...

         If DecimalPlace > 0 Then

            ' Convert Paise

            temp = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)

            Paise = ConvertTens(temp)



            ' Strip off Paise from remainder to convert.

            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))

         End If



         count = 1

         Do While MyNumber <> ""

            ' Convert last 3 digits of MyNumber to English Rupees.

            If count = 1 Then
            
                temp = ConvertHundreds(Right(MyNumber, 3))
                
            Else
            
            If Len(MyNumber) >= 2 Then
                temp = ConvertTens(Right(MyNumber, 2))
            Else
                temp = ConvertDigit(Right(MyNumber, 1))
            End If
                
            End If

            If temp <> "" Then Rupees = temp & Place(count) & Rupees

            If count = 1 Then
            
                If Len(MyNumber) > 3 Then
    
                   ' Remove last 3 converted digits from MyNumber.
    
                   MyNumber = Left(MyNumber, Len(MyNumber) - 3)
    
                Else
    
                   MyNumber = ""
    
                End If
                
            Else
            
                If Len(MyNumber) > 2 Then
    
                   ' Remove last 2 converted digits from MyNumber.
    
                   MyNumber = Left(MyNumber, Len(MyNumber) - 2)
    
                Else
    
                   MyNumber = ""
    
                End If
                
            End If

            count = count + 1

         Loop



         ' Clean up Rupees.

         Select Case Rupees

            Case ""

               Rupees = "Zero"

            Case "One"

               Rupees = "One"

            Case Else

               Rupees = Rupees

         End Select



         ' Clean up Paise.

         Select Case Paise

            Case ""
' uncomment the below line if paise zero only is required. Rajendra made this below comment.
'               Paise = " And Paise Zero Only"
               Paise = " Only."
            Case "One"

               Paise = " And Paise One Only"

            Case Else

               Paise = " And Paise " & Paise & " Only"

         End Select
         NUMTOWORDS = Rupees & Paise
End Function







Private Function ConvertHundreds(ByVal MyNumber)

Dim result As String



         ' Exit if there is nothing to convert.

         If Val(MyNumber) = 0 Then Exit Function



         ' Append leading zeros to number.

         MyNumber = Right("000" & MyNumber, 3)



         ' Do we have a hundreds place digit to convert?

         If Left(MyNumber, 1) <> "0" Then

            result = ConvertDigit(Left(MyNumber, 1)) & " Hundred "

         End If



         ' Do we have a tens place digit to convert?

         If Mid(MyNumber, 2, 1) <> "0" Then

            result = result & ConvertTens(Mid(MyNumber, 2))

         Else

            ' If not, then convert the ones place digit.

            result = result & ConvertDigit(Mid(MyNumber, 3))

         End If



         ConvertHundreds = Trim(result)

End Function







Private Function ConvertTens(ByVal MyTens)

Dim result As String



         ' Is value between 10 and 19?

         If Val(Left(MyTens, 1)) = 1 Then

            Select Case Val(MyTens)

               Case 10: result = "Ten"

               Case 11: result = "Eleven"

               Case 12: result = "Twelve"

               Case 13: result = "Thirteen"

               Case 14: result = "Fourteen"

               Case 15: result = "Fifteen"

               Case 16: result = "Sixteen"

               Case 17: result = "Seventeen"

               Case 18: result = "Eighteen"

               Case 19: result = "Nineteen"

               Case Else

            End Select

         Else

            ' .. otherwise it's between 20 and 99.

            Select Case Val(Left(MyTens, 1))

               Case 2: result = "Twenty "

               Case 3: result = "Thirty "

               Case 4: result = "Forty "

               Case 5: result = "Fifty "

               Case 6: result = "Sixty "

               Case 7: result = "Seventy "

               Case 8: result = "Eighty "

               Case 9: result = "Ninety "

               Case Else

            End Select



            ' Convert ones place digit.

            result = result & ConvertDigit(Right(MyTens, 1))

         End If



         ConvertTens = result

End Function
Private Function ConvertDigit(ByVal MyDigit)

Select Case Val(MyDigit)

            Case 1: ConvertDigit = "One"

            Case 2: ConvertDigit = "Two"

            Case 3: ConvertDigit = "Three"

            Case 4: ConvertDigit = "Four"

            Case 5: ConvertDigit = "Five"

            Case 6: ConvertDigit = "Six"

            Case 7: ConvertDigit = "Seven"

            Case 8: ConvertDigit = "Eight"

            Case 9: ConvertDigit = "Nine"

            Case Else: ConvertDigit = ""

         End Select

End Function
'-----------------------


