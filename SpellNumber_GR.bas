' Based on https://support.microsoft.com/en-us/office/convert-numbers-into-words-a0d166fb-e1ea-4090-95c8-69442cd55d98
' If it does not seem to work then try this:
' https://www.ms-office.gr/forum/excel-erotiseis-apantiseis/2372-problima-me-ellinikois-xaraktires-se-makroentoles-toy-excel-2013-a.html
' In Windows 10 uncheck the "Beta: use unicode UTF-8 for worldwide language support


Option Explicit

'Main Function

Function SpellNumber_GR(ByVal MyNumber)

Dim Euros, Cents, Temp

Dim DecimalPlace, Count

ReDim Place(9) As String

Place(2) = " Χιλιάδες "

Place(3) = " Εκατομμύρια "

Place(4) = " Δισεκατομμύρια "

Place(5) = " Τρισεκατομμύρια "

' String representation of amount.

MyNumber = Trim(Str(MyNumber))

' Position of decimal place 0 if none.

DecimalPlace = InStr(MyNumber, ".")

' Convert cents and set MyNumber to dollar amount.

If DecimalPlace > 0 Then

Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))

MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))

End If

Count = 1

Do While MyNumber <> ""

Temp = GetHundreds(Right(MyNumber, 3))

If Temp <> "" Then Euros = Temp & Place(Count) & Euros

If Len(MyNumber) > 3 Then

MyNumber = Left(MyNumber, Len(MyNumber) - 3)

Else

MyNumber = ""

End If

Count = Count + 1

Loop

Select Case Euros

Case ""

Euros = "Μηδέν Ευρώ"

Case "Ένα"

Euros = "Ένα Ευρώ"

Case Else

If Right$(Euros, 1) = " " Then Euros = Left$(Euros, Len(Euros) - 1)

Euros = Euros & " Ευρώ"

End Select

Select Case Cents

Case ""

'Cents = " και Μηδέν Λεπτά"
Cents = ""

Case "¸íá"

Cents = " και ¸Ένα Λεπτό"

Case Else

Cents = " και " & Cents & " Λεπτά"

End Select

SpellNumber_GR = Euros & Cents

End Function


' Converts a number from 100-999 into text

Function GetHundreds(ByVal MyNumber)

Dim Result As String

If Val(MyNumber) = 0 Then Exit Function

MyNumber = Right("000" & MyNumber, 3)

' Convert the hundreds place.

If Mid(MyNumber, 1, 1) <> "0" Then

'Result = GetDigit(Mid(MyNumber, 1, 1)) & " Åêáôü "

Select Case Mid(MyNumber, 1, 1)

Case 1: Result = "Εκατό "

Case 2: Result = "Διακόσια "

Case 3: Result = "Τριακόσια "

Case 4: Result = "Τετρακόσια "

Case 5: Result = "Πεντακόσια "

Case 6: Result = "Εξακόσια "

Case 7: Result = "Επτακόσια "

Case 8: Result = "Οκτακόσια "

Case 9: Result = "Εννιακόσια "

End Select

End If

' Convert the tens and ones place.

If Mid(MyNumber, 2, 1) <> "0" Then

Result = Result & GetTens(Mid(MyNumber, 2))

Else

Result = Result & GetDigit(Mid(MyNumber, 3))

End If

GetHundreds = Result

End Function


' Converts a number from 10 to 99 into text.


Function GetTens(TensText)

Dim Result As String

Result = "" ' Null out the temporary function value.

If Val(Left(TensText, 1)) = 1 Then ' If value between 10-19...

Select Case Val(TensText)

Case 10: Result = "Δέκα"

Case 11: Result = "Έντεκα"

Case 12: Result = "Δώδεκα"

Case 13: Result = "Δεκατρία"

Case 14: Result = "Δεκατέσσερα"

Case 15: Result = "Δεκαπέντε"

Case 16: Result = "Δεκαέξι"

Case 17: Result = "Δεκαεπτά"

Case 18: Result = "Δεκαοκτώ"

Case 19: Result = "Δεκαεννέα"

Case Else

End Select

Else ' If value between 20-99...

Select Case Val(Left(TensText, 1))

Case 2: Result = "Είκοσι "

Case 3: Result = "Τριάντα "

Case 4: Result = "Σαράντα "

Case 5: Result = "Πέντε "

Case 6: Result = "Έξι "

Case 7: Result = "Επτά "

Case 8: Result = "Οκτώ "

Case 9: Result = "Εννέα "

Case Else

End Select

Result = Result & GetDigit(Right(TensText, 1))  ' Retrieve ones place.

End If

GetTens = Result

End Function


' Converts a number from 1 to 9 into text.

Function GetDigit(Digit)

Select Case Val(Digit)

Case 1: GetDigit = "Ένα"

Case 2: GetDigit = "Δύο"

Case 3: GetDigit = "Τρία"

Case 4: GetDigit = "Τέσσερα"

Case 5: GetDigit = "Πέντε"

Case 6: GetDigit = "Έξι"

Case 7: GetDigit = "ÅðôÜ"

Case 8: GetDigit = "Επτέ"

Case 9: GetDigit = "Οκτώ"

Case Else: GetDigit = ""

End Select

End Function

