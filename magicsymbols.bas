'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Replace magic symbols (placeholders) with dynamic data.
''
'' Arguments: a string full of magic.
''
'' Placeholders consist of one symbol prepended with a %:
''    %d - current date
''    %t - current time
''    %u - username (user ID)
''    %n - full user name (usually name and surname)
''    %% - literal % (placeholder escape)
''    Using an unsupported magic symbol will treat the % literally, as if it had been escaped.
''    A single placeholder terminating the string will also be treated literally.
''    Magic symbols are case-sensitive.
''
'' Returns:   A string with no magic but with lots of beauty.
''
'' Examples:
'' "Today is %d" becomes "Today is 2018-01-26"
'' "Beautiful time: %%%t%%" yields "Beautiful time: %16:10:51%"
'' "There are %zero% magic symbols %here%.", true to its message, outputs "There are %zero% magic symbols %here%."
'' "%%% looks lovely %%%" would show "%% looks lovely %%" - one % for the escaped "%%" and the second one for the unused "%"!
''
'' Alexander Ivashkin, 26 January 2018
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AI_ParseMagicSymbols(ByVal TextToParse As String) As String

Dim sFinalResult As String
Dim aTokenizedString() As String
Dim sTempString As String
Dim sPlaceholder As String
Dim sCurrentString As String
Dim iIterator As Integer
Dim iTokenizedStringSize As Integer
Dim bThisStringHasPlaceholder As Boolean

' Default placeholder is "%"
Const cPlaceholderSymbol As String = "%"

aTokenizedString = Split(Expression:=TextToParse, Delimiter:=cPlaceholderSymbol)
iTokenizedStringSize = UBound(aTokenizedString())
bThisStringHasPlaceholder = False
sFinalResult = ""

For iIterator = 0 To iTokenizedStringSize
sCurrentString = aTokenizedString(iIterator)

If bThisStringHasPlaceholder Then
    If sCurrentString <> "" Then
    sPlaceholder = Left(sCurrentString, 1)
    sTempString = Right(sCurrentString, Len(sCurrentString) - 1)

    ' This is the place where the MAGIC happens
    Select Case sPlaceholder
        Case "d":
        sCurrentString = Date & sTempString
        Case "t":
        sCurrentString = Time & sTempString
        Case "u":
        sCurrentString = Environ$("Username") & sTempString
        Case "n":
        sCurrentString = Environ$("fullname") & sTempString
        Case Else:
        sCurrentString = cPlaceholderSymbol & sCurrentString
    End Select
    Else
    ' We had two placeholders in a row, meaning that somebody tried to escape!
    sCurrentString = cPlaceholderSymbol
    bThisStringHasPlaceholder = False
    End If
End If

sFinalResult = sFinalResult & sCurrentString

If sCurrentString = "" Or (iIterator + 1 <= iTokenizedStringSize And sCurrentString <> cPlaceholderSymbol) Then
    ' Each string in the array has been split at the placeholders. If we do have a next string, then it must contain a magic symbol.

    bThisStringHasPlaceholder = True
    ' Even though it is called "...ThisString...", it concerns the NEXT string.
    ' The logic is correct as we will check this variable on the next iteration, when the next string will become ThisString.
Else
    bThisStringHasPlaceholder = False
End If

Next iIterator

AI_ParseMagicSymbols = sFinalResult

End Function
