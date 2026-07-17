Attribute VB_Name = "RhinoCase"
'                _...:::::::::::..._         __________________________________________________________________________         _...:::::::::::..._
'           _.::¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦:::;-..[  VBA TERMINAL TEXT CASE FORMATTING TOOL                                  ]..-;:::¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦::._
'         :¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦|==========================================================================|¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦:
'       /:¦:::'''':::¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦\¦¦¦¦¦¦¦|  Copyright (c) 2026 ChronicRhyno                                         |¦¦¦¦¦¦/¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦:::'''':::¦:\
'    .~:¦'            ':¦¦¦¦¦¦¦¦\¦¦¦¦\\¦¦¦¦¦|                                                                          |¦¦¦¦//¦¦¦¦/¦¦¦¦¦¦¦¦¦:'            '¦:~.
' *~` /                 \¦¦¦¦¦¦¦||¦¦¦¦||¦¦¦¦|  Permission is hereby granted, free of charge, to any person obtaining   |¦¦¦||¦¦¦¦||¦¦¦¦¦¦¦¦/                 \ `~*
'     |                  \¦¦¦¦¦¦||¦¦¦¦||¦¦¦¦|  a copy of this software and associated documentation files, to deal     |¦¦¦||¦¦¦¦||¦¦¦¦¦¦¦/                  |
'     |                   |¦¦¦¦¦||¦¦¦¦||¦¦¦¦|  in the Software without restriction, including without limitation the   |¦¦¦||¦¦¦¦||¦¦¦¦¦¦|                   |
'     |                   |¦¦¦¦¦||¦¦¦¦||¦¦¦¦|  rights to use, copy, modify, merge, publish, distribute, sublicense,    |¦¦¦||¦¦¦¦||¦¦¦¦¦¦|                   |
'      \                  |¦¦¦¦¦/¦¦¦¦¦||¦¦¦¦|  and/or sell copies of the Software.                                     |¦¦¦||¦¦¦¦¦\¦¦¦¦¦¦|                  /
'       |                /¦¦¦¦¦¦¦¦¦¦¦//¦¦¦¦¦|                                                                          |¦¦¦¦\\¦¦¦¦¦¦¦¦¦¦¦¦\                |
'        \            `_/;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦|  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,         |¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦;\_`            /
'         |_         ;_/ ;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦|  EXPRESS OR IMPLIED.                                                     |¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦;  \_;         _|
'          \       #;_|  ``'''''''''''''''''|==========================================================================|''''''''''''''``     |_;#       /
'          |       ._/                                                                                                                        \_.       |
'          |_       |                                              May you have fun in any case.                                               |       _|
'           \_____)_))                                                                                                                        ((_(_____/
Sub VBAterminalTEXTcaseFORMATTINGtool()
    On Error GoTo ErrorHandler
    
    Dim selectedRange As range
    Set selectedRange = Selection.range
    
    If Selection.Type = wdSelectionIP Or Len(selectedRange.Text) = 0 Then
        MsgBox "Please highlight some text before running this tool.", vbExclamation, "No Selection Found"
        Exit Sub
    End If
    
    Dim promptMessage As String
    promptMessage = "  TEXT CASE FORMATTING TOOL" & vbCrLf & vbCrLf & _
                    "   [1]  Camel Case  (lowerFirstWord)" & vbCrLf & _
                    "   [2]  Pascal Case (SmartTitleCase)" & vbCrLf & _
                    "   [3]  Zeebra Case (all|lowercase|piped)" & vbCrLf & _
                    "   [4]  kebab-case  (all-lowercase-dashed)" & vbCrLf & _
                    "   [5]  snake_case  (all_lowercase_joined)" & vbCrLf & _
                    "   [6]  Inchworm Case (alternate per word)" & vbCrLf & _
                    "   [7]  Sarcasm Case  (aLtErNaTiNg caps)" & vbCrLf & _
                    "   [8]  Studly Vowels (vOwEls Up)" & vbCrLf & _
                    "   [9]  Studly Consonants (CoNSoNaNTs uP)" & vbCrLf & _
                    "  [10]  Rhino Case   (camelCase+small caps)" & vbCrLf & _
                    "  [11]  ChronicRhyno (PascalCase+small caps)" & vbCrLf & _
                    "  [12]  Studly Rhino (vOwEls Up+smAll cAps)" & vbCrLf & vbCrLf & _
                    "Enter selection (1-12):"
                    
    Dim userChoice As String
    userChoice = InputBox(promptMessage, "VBA TERMINAL v1.0", "1")
    
    If userChoice = "" Then Exit Sub
    
    Select Case Trim(userChoice)
        Case "1"
            ResetRangeFormatting selectedRange
            selectedRange.Text = CamelCase(selectedRange.Text)
        Case "2"
            ResetRangeFormatting selectedRange
            selectedRange.Text = PascalCase(selectedRange.Text)
        Case "3"
            ResetRangeFormatting selectedRange
            selectedRange.Text = ZeebraCase(selectedRange.Text)
        Case "4"
            ResetRangeFormatting selectedRange
            selectedRange.Text = KebabCase(selectedRange.Text)
        Case "5"
            ResetRangeFormatting selectedRange
            selectedRange.Text = SnakeCase(selectedRange.Text)
        Case "6"
            InchWormCase selectedRange
        Case "7"
            ResetRangeFormatting selectedRange
            selectedRange.Text = SarcasmCase(selectedRange.Text)
        Case "8"
            ResetRangeFormatting selectedRange
            selectedRange.Text = StudlyVowels(selectedRange.Text)
        Case "9"
            ResetRangeFormatting selectedRange
            selectedRange.Text = StudlyConsonants(selectedRange.Text)
        Case "10"
            RhinoCase selectedRange
        Case "11"
            ChronicRhynoCase selectedRange
        Case "12"
            StudlyRhinoCase selectedRange
        Case Else
            MsgBox "Invalid selection. Please enter 1-12.", vbCritical, "Terminal Error"
            Exit Sub
    End Select
    Exit Sub
    
ErrorHandler:
    MsgBox "An unexpected error occurred. Operation aborted.", vbCritical, "System Error"
End Sub

Sub ResetRangeFormatting(ByRef rng As range)
    rng.Font.Subscript = False
    rng.Font.Superscript = False
    rng.Font.SmallCaps = False
End Sub

' 1. camelCase (First word completely lowercase, subsequent words capitalized, spaces removed)
' Coined by Newton Love
Function CamelCase(ByVal txt As String) As String
    On Error GoTo FuncError
    txt = Replace(txt, Chr(160), " ")
    
    Dim words() As String
    words = Split(txt, " ")
    
    Dim firstWordFound As Boolean
    firstWordFound = False
    
    Dim i As Long, tempWord As String
    For i = LBound(words) To UBound(words)
        tempWord = words(i)
        If Len(tempWord) > 0 Then
            If Not firstWordFound Then
                words(i) = LCase(tempWord)
                firstWordFound = True
            Else
                Dim firstLetter As String, restOfWord As String
                firstLetter = UCase(Left(tempWord, 1))
                restOfWord = Mid(tempWord, 2)
                If restOfWord = UCase(restOfWord) Then
                    restOfWord = LCase(restOfWord)
                End If
                words(i) = firstLetter & restOfWord
            End If
        End If
    Next i
    
    CamelCase = Join(words, "")
    Exit Function
FuncError:
    CamelCase = txt
End Function

' 2. PascalCase (Handles existing Pascal/Camel preservation + Normalizing spaces)
' Coined by Anders Hejlsberg
Function PascalCase(ByVal txt As String) As String
    On Error GoTo FuncError
    txt = Replace(txt, Chr(160), " ")
    
    Dim words() As String
    words = Split(txt, " ")
    
    Dim i As Long, tempWord As String, firstLetter As String, restOfWord As String
    For i = LBound(words) To UBound(words)
        tempWord = words(i)
        If Len(tempWord) > 0 Then
            firstLetter = UCase(Left(tempWord, 1))
            restOfWord = Mid(tempWord, 2)
            
            If restOfWord = UCase(restOfWord) Then
                restOfWord = LCase(restOfWord)
            End If
            
            words(i) = firstLetter & restOfWord
        End If
    Next i
    
    PascalCase = Join(words, "")
    Exit Function
FuncError:
    PascalCase = txt
End Function

' 3. Zeebra Case (all lowercase, spaces replaced with single pipes)
' Coined by ChronicRhyno
Function ZeebraCase(ByVal txt As String) As String
    On Error GoTo FuncError
    
    txt = Replace(txt, Chr(160), " ")
    txt = LCase(Trim(txt))
    
    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop
    
    ZeebraCase = Replace(txt, " ", "|")
    
    Exit Function
    
FuncError:
    ZeebraCase = txt
End Function

' 4. Kebab-Case (all lowercase, spaces replaced with dashes)
' Coined by Ben Lee
Function KebabCase(ByVal txt As String) As String
    On Error GoTo FuncError
    txt = Replace(txt, Chr(160), " ")
    txt = LCase(Trim(txt))
    
    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop
    
    KebabCase = Replace(txt, " ", "-")
    Exit Function
    
FuncError:
    KebabCase = txt
End Function

' 5. snake_case (All lowercase, spaces replaced with underscores)
' Coined by Gavin Kistner
Function SnakeCase(ByVal txt As String) As String
    On Error GoTo FuncError
    txt = Replace(txt, Chr(160), " ")
    txt = LCase(Trim(txt))

    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop
    
    SnakeCase = Replace(txt, " ", "_")
    Exit Function
    
FuncError:
    SnakeCase = txt
End Function

' 6. Inchworm Case (alternating subscript/superscript per word, removing spaces)
' Coined by ChronicRhyno
Function InchWormCase(ByRef rng As range) As Boolean
    On Error GoTo FuncError
    
    ResetRangeFormatting rng
    
    Dim originalText As String
    originalText = Replace(rng.Text, Chr(160), " ")
    originalText = Trim(originalText)
    
    Do While InStr(originalText, "  ") > 0
        originalText = Replace(originalText, "  ", " ")
    Loop
    
    If Len(originalText) = 0 Then
        InchWormCase = False
        Exit Function
    End If
    
    Dim words() As String
    words = Split(originalText, " ")
    
    Dim cleanText As String
    cleanText = Join(words, "")
    rng.Text = cleanText
    
    Dim charIndex As Long
    charIndex = 1
    
    Dim i As Long
    Dim wordLen As Long
    Dim wordRange As range
    
    For i = 0 To UBound(words)
        wordLen = Len(words(i))
        
        If wordLen > 0 Then
            Set wordRange = rng.Duplicate
            
            wordRange.start = rng.start + charIndex - 1
            wordRange.End = wordRange.start + wordLen
            
            If i Mod 2 = 0 Then
                wordRange.Font.Subscript = True
                wordRange.Font.Superscript = False
            Else
                wordRange.Font.Superscript = True
                wordRange.Font.Subscript = False
            End If
            
            charIndex = charIndex + wordLen
        End If
    Next i
    
    InchWormCase = True '
    Exit Function
    
FuncError:
    InchWormCase = False
End Function

' 7. Sarcasm Case (Alternating caps character-by-character, skipping spaces)
Function SarcasmCase(ByVal txt As String) As String
    On Error GoTo FuncError
    txt = Replace(txt, Chr(160), " ")
    Dim result As String
    result = ""
    Dim i As Long, charCount As Long
    charCount = 0
    For i = 1 To Len(txt)
        Dim char As String
        char = Mid(txt, i, 1)
        
        If char Like "[a-zA-Z]" Then
            If charCount Mod 2 = 0 Then
                result = result & LCase(char)
            Else
                result = result & UCase(char)
            End If
            charCount = charCount + 1
        Else
            result = result & char
        End If
    Next i
    SarcasmCase = result
    Exit Function
FuncError:
    SarcasmCase = txt
End Function

' 8. Studly Vowels (Vowels UPPER, Consonants lower)
' Coined by ChronicRhyno citing David iIii for "Studly Caps/Capitals"
Function StudlyVowels(ByVal txt As String) As String
    On Error GoTo FuncError
    Dim i As Long, char As String, result As String
    result = ""
    For i = 1 To Len(txt)
        char = Mid(txt, i, 1)
        If char Like "[aeiouyAEIOUY]" Then
            result = result & UCase(char)
        Else
            result = result & LCase(char)
        End If
    Next i
    StudlyVowels = result
    Exit Function
FuncError:
    StudlyVowels = txt
End Function

' 9. Studly Consonants (Consonants UPPER, Vowels lower)
' Coined by ChronicRhyno citing David iIii for "Studly Caps/Capitals"
Function StudlyConsonants(ByVal txt As String) As String
    On Error GoTo FuncError
    Dim i As Long, char As String, result As String
    result = ""
    For i = 1 To Len(txt)
        char = Mid(txt, i, 1)
        If char Like "[aeiouyAEIOUY]" Then
            result = result & LCase(char)
        Else
            result = result & UCase(char)
        End If
    Next i
    StudlyConsonants = result
    Exit Function
FuncError:
    StudlyConsonants = txt
End Function

' 10. rhinoCase (camelCase + Small Caps)
' Coined by ChronicRhyno
Sub RhinoCase(ByRef rng As range)
    On Error GoTo FuncError
    ResetRangeFormatting rng
    Dim cleanTxt As String
    cleanTxt = CamelCase(rng.Text)
    rng.Text = cleanTxt
    rng.Font.SmallCaps = True
    Exit Sub
FuncError:
End Sub

' 11. ChronicRhyno Case (PascalCase + Small Caps)
' Coined by ChronicRhyno July 15, 2026, used since 2005
Sub ChronicRhynoCase(ByRef rng As range)
    On Error GoTo FuncError
    ResetRangeFormatting rng
    Dim txt As String
    txt = Replace(rng.Text, Chr(160), " ")
    Dim words() As String
    words = Split(txt, " ")
    Dim i As Long, tempWord As String
    For i = LBound(words) To UBound(words)
        tempWord = words(i)
        If Len(tempWord) > 0 Then
            words(i) = UCase(Left(tempWord, 1)) & LCase(Mid(tempWord, 2))
        End If
    Next i
    rng.Text = Join(words, "")
    rng.Font.SmallCaps = True
    Exit Sub
FuncError:
End Sub

' 12. Studly Rhino Case (studly vowels + small caps + leave spaces)
' Coined by ChronicRhyno citing David iIii for "Studly Caps/Capitals"
Sub StudlyRhinoCase(ByRef rng As range)
    On Error GoTo FuncError
    ResetRangeFormatting rng
    Dim txt As String
    txt = rng.Text
    Dim i As Long, char As String, result As String
    result = ""
    For i = 1 To Len(txt)
        char = Mid(txt, i, 1)
        If char Like "[aeiouyAEIOUY]" Then
            result = result & UCase(char)
        Else
            result = result & LCase(char)
        End If
    Next i
    rng.Text = result
    rng.Font.SmallCaps = True
    Exit Sub
FuncError:
End Sub
