Option Strict On
Option Explicit On
'Aftanom Anfilofieff
'RCET0265
'Spring 2021
'Code Challenge
'https://github.com/AftaAnfi/AftanomCodeChallenge.git
Module Module1

    Sub Main()

        GetUserInput()
        TestValidateAndConvert()

    End Sub

    'Code to test the ValidateAndConvert function
    Private Sub TestValidateAndConvert()
        'number of successful tests
        Dim count As Integer = 0

        'integer to passback result of ValidateAndConvert function
        Dim result As Integer = 0
        Dim pad As Integer = 15
        Dim report As String = ""
        Dim temp As String = ""

        'data to passthrough ValidateAndConvert function and test
        Dim testData = New String(4, 4) {
            {"5", "2", "17", "8", "42"},
            {"6.7", "3.14", "5.4", "5.5", "0.125"},
            {"-21", "-32.1", "-4", "-4.5", "-4.4"},
            {"", "", "", "", ""},
            {"True", "False", "lOOlO", "9O2lO", "dog"}}

        'loop all test values from testData
        For row = 0 To 4
            For column = 0 To 4
                result = 0
                temp = ValidateAndConvert(testData(row, column), result)
                report &= ("Trying: " & testData(row, column)).PadRight(pad)
                If row < 3 Then
                    If CStr(CInt(testData(row, column))) <> CStr(result) Or temp <> "" Then
                        report &= " TEST FAIL" & vbNewLine
                        report &= ("Result is: " & CStr(result)).PadRight(pad) & " : " & temp & vbNewLine
                        report &= ("Should be: " & CStr(CInt(testData(row, column)))).PadRight(pad) & " : " _
                        & "<Empty>" & vbNewLine
                    Else
                        report &= " TEST PASS" & vbNewLine
                        count += 1
                    End If
                ElseIf temp <> "is empty" And row = 3 Then
                    report &= " TEST FAIL" & vbNewLine
                    report &= ("Result is: " & CStr(result)).PadRight(pad) & " : " & temp & vbNewLine
                    report &= ("Should be: " & CStr(0)).PadRight(pad) & " : " & "is empty" & vbNewLine
                ElseIf temp <> "Must contain a number" And row > 3 Then
                    report &= " TEST FAIL" & vbNewLine
                    report &= ("Result is: " & CStr(result)).PadRight(pad) & " : " & temp & vbNewLine
                    report &= ("Should be: " & CStr(0)).PadRight(pad) & " : " & "Must contain a number" _
                    & vbNewLine
                Else
                    report &= " TEST PASS" & vbNewLine
                    count += 1
                End If
            Next
        Next
        Console.WriteLine(report & "Passed " & CStr(count) & " of 25 tests. Score: " _
            & CStr((count / 25) * 100) & "%")
        MsgBox("Passed " & CStr(count) & " of 25 tests. Score: " _
            & CStr((count / 25) * 100) & "%")
    End Sub

    'Get userinput from an inputbox
    Private Sub GetUserInput()
        Dim tempNum As Integer = 0
        Dim userMessage As String = "Please Enter A Number Between 0 and 15" _
            & vbNewLine & "Type Q to Quit"
        Do
            userMessage = InputBox(userMessage, "Hello", "")
            If userMessage <> "Q" And userMessage <> "" Then
                userMessage = ValidateAndConvert(userMessage, tempNum)
                If userMessage = "" Then userMessage = ShortAndSweet(tempNum)        'Replace this Line
                'If userMessage = "" Then userMessage = ShortAndSweet(tempNum)      'With this one
            End If
        Loop Until userMessage = "Q" Or userMessage = ""
    End Sub

    'Function to validate and convert a string to an integer and provide 
    'a string feedback of how the conversion went
    Private Function ValidateAndConvert(ByVal convertThisString As String, ByRef toThisInteger As Integer) As String
        Dim message As String
        Try
            toThisInteger = CInt(convertThisString)
            message = ""
        Catch ex As Exception
            If convertThisString = "" Then
                message = "is empty"
            Else
                message = "Must contain a number"
            End If
        End Try
        Return message$
    End Function

    'Function to provide word form of values 0 - 15
    Private Function ShortAndSweet(ByVal numberFromZeroToFifteen As Integer) As String
        Dim stringArray As String() = {"Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen"}
        Select Case numberFromZeroToFifteen
            Case 0 To 15
                Return ($"Your Number is: {stringArray(numberFromZeroToFifteen)}")
            Case < 0
                Return ($"Your Number is: Too Low")
            Case > 15
                Return ($"Your Number is: Too High")
        End Select
    End Function

End Module
