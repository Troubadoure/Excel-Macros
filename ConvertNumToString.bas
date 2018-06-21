Attribute VB_Name = "ConvertNumToString"
Option Explicit
Option Private Module
Public Function ConvertNumberToString(ByVal v As Long, Optional LanguageCode As String) As String
    ' Turn on DevelopMode Set to True
    Dim Develop As Boolean: Develop = False
    ' Maximum number of places allowed by Excel is 9
    Const MaxPlaces = 9
    Dim tmp(0 To MaxPlaces) As String, tmpNumber() As String
    Dim StartConvert As Boolean: StartConvert = False
    Dim Negative As Boolean: Negative = False
    Dim RemoveTrailingZero As Boolean: RemoveTrailingZero = False
    Dim Minus As String
    Dim PlaceValue As Long
    Dim Remainder As Long
    
    If LanguageCode = "" Then LanguageCode = "en-gb"
    
    If v = 0 Then
        tmp(1) = LanguageDictionary(v, LanguageCode)
    Else
        If v < 0 Then
            Negative = True
            v = Abs(v)
        End If
        
        For PlaceValue = MaxPlaces To 0 Step -1
            ReDim tmpNumber(1 To 2)
            On Error Resume Next
            If v <> v Mod (1 * 10 ^ PlaceValue) Then StartConvert = True
            On Error GoTo 0
            If StartConvert Then
                If PlaceValue > 0 Then RemoveTrailingZero = True
                
                If Develop Then Debug.Print PlaceValue, v Mod (1 * 10 ^ PlaceValue), WorksheetFunction.RoundDown(v / (1 * 10 ^ PlaceValue), 0),
                
                If PlaceValue > 1 Or v >= 20 Then
                    Remainder = v Mod (1 * 10 ^ PlaceValue)
                    v = WorksheetFunction.RoundDown(v / (1 * 10 ^ PlaceValue), 0)
                End If
            
                tmpNumber(1) = LanguageDictionary(v, LanguageCode)
                If PlaceValue > 1 Then
                    tmpNumber(2) = LanguageDictionary(1 * 10 ^ PlaceValue, LanguageCode)
                ElseIf PlaceValue = 1 Then
                    If v >= 2 Then
                        tmpNumber(1) = LanguageDictionary(v * (1 * 10 ^ PlaceValue), LanguageCode)
                    Else
                        tmpNumber(1) = LanguageDictionary(v, LanguageCode)
                    End If
                End If
                
                If v = 0 And PlaceValue > 0 Then tmpNumber(1) = "": tmpNumber(2) = ""
                
                v = Remainder
                
                
                tmp(MaxPlaces - PlaceValue) = WorksheetFunction.Trim(Join(tmpNumber, " "))
                
                If Develop Then Debug.Print tmp(MaxPlaces - PlaceValue)
            End If
        Next PlaceValue
    End If
    
    If Negative Then Minus = LanguageDictionary("Minus", LanguageCode)
    If RemoveTrailingZero And Trim(tmp(MaxPlaces)) = LanguageDictionary(0, LanguageCode) Then tmp(MaxPlaces) = ""
    
    ConvertNumberToString = WorksheetFunction.Trim(Minus & " " & Join(tmp))
End Function
Private Function LanguageDictionary(v As Variant, LanguageCode As String) As String
    If LanguageCode = "en-gb" Then
        Select Case v
            Case "and": LanguageDictionary = "and"
            Case "Minus": LanguageDictionary = "Minus"
            Case 0: LanguageDictionary = "Zero"
            Case 1: LanguageDictionary = "One"
            Case 2: LanguageDictionary = "Two"
            Case 3: LanguageDictionary = "Three"
            Case 4: LanguageDictionary = "Four"
            Case 5: LanguageDictionary = "Five"
            Case 6: LanguageDictionary = "Six"
            Case 7: LanguageDictionary = "Seven"
            Case 8: LanguageDictionary = "Eight"
            Case 9: LanguageDictionary = "Nine"
            Case 10: LanguageDictionary = "Ten"
            Case 11: LanguageDictionary = "Eleven"
            Case 12: LanguageDictionary = "Twelve"
            Case 13: LanguageDictionary = "Thirteen"
            Case 14: LanguageDictionary = "Fourteen"
            Case 15: LanguageDictionary = "Fifteen"
            Case 16: LanguageDictionary = "Sixteen"
            Case 17: LanguageDictionary = "Seventeen"
            Case 18: LanguageDictionary = "Eighteen"
            Case 19: LanguageDictionary = "Nineteen"
            Case 20: LanguageDictionary = "Twenty"
            Case 30: LanguageDictionary = "Thirty"
            Case 40: LanguageDictionary = "Fourty"
            Case 50: LanguageDictionary = "Fifty"
            Case 60: LanguageDictionary = "Sixty"
            Case 70: LanguageDictionary = "Seventy"
            Case 80: LanguageDictionary = "Eighty"
            Case 90: LanguageDictionary = "Ninety"
            Case 1 * 10 ^ 2: LanguageDictionary = "Hundred"
            Case 1 * 10 ^ 3: LanguageDictionary = "Thousand"
            Case 1 * 10 ^ 4: LanguageDictionary = "Ten Thousand"
            Case 1 * 10 ^ 5: LanguageDictionary = "Hundred Thousand"
            Case 1 * 10 ^ 6: LanguageDictionary = "Million"
            Case 1 * 10 ^ 7: LanguageDictionary = "Ten Million"
            Case 1 * 10 ^ 8: LanguageDictionary = "Billion"
            Case 1 * 10 ^ 9: LanguageDictionary = "Billion"
        End Select
    ElseIf LanguageCode = "fr-fr" Then
        Select Case v
            Case "Minus": LanguageDictionary = "moins"
        End Select
    End If
End Function

Sub Setup()
    Dim i As Long
    For i = 1 To 9
        Debug.Print "Case " & 1 * 10 ^ i & ": LanguageDictionary = """""
    Next i
End Sub


