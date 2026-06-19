' =========================================================================
' Unified Question Paper Formatter (MCQ & CQ)
' Instructions: Copy the entire code block below and paste it into a single 
'               VBA module in your Word Normal project.
' =========================================================================

Sub Format_Question_Paper()
    ' =========================================================================
    ' Macro Name: Format_Question_Paper
    ' Purpose: Asks the user to choose MCQ or CQ and executes formatting
    ' =========================================================================
    Dim userInput As String
    
    ' Displaying Input Box to prevent Unicode character errors in VBA prompt
    userInput = InputBox("Please enter a number to select formatting type:" & vbNewLine & vbNewLine & _
                         "Type 1 : MCQ (Multiple Choice Questions)" & vbNewLine & _
                         "Type 2 : CQ (Creative Questions)" & vbNewLine & vbNewLine & _
                         "Leave blank or click Cancel to exit.", "Select Question Type")
                         
    ' Input verification and routing
    If Trim(userInput) = "1" Then
        Application.ScreenUpdating = False
        Call Execute_MCQ_Formatting
        Application.ScreenUpdating = True
        MsgBox "MCQ Formatting completed successfully!", vbInformation, "Completed"
        
    ElseIf Trim(userInput) = "2" Then
        Application.ScreenUpdating = False
        Call Execute_CQ_Formatting
        Application.ScreenUpdating = True
        MsgBox "CQ Formatting completed successfully!", vbInformation, "Completed"
        
    ElseIf Trim(userInput) = "" Then
        MsgBox "Formatting process canceled.", vbInformation, "Canceled"
    Else
        MsgBox "Invalid choice! Please enter either 1 or 2.", vbExclamation, "Error"
    End If
End Sub

' =========================================================================
' Private Sub: Execute_MCQ_Formatting
' Description: Step-by-step MCQ Formatting
' =========================================================================
Private Sub Execute_MCQ_Formatting()
    ' --- Step 5: Preserve Tab after Digit and Ddari (General Preparation) ---
    Call PerformWildcardReplace("([" & BenDigits() & "]{1,2})" & BenDdari() & "^9", "\1" & BenDdari() & "^t")
    
    ' --- Step 12: Remove Extra Spaces (General Preparation) ---
    Call PerformWildcardReplace("[ ]{2,}", " ")
    
    ' --- Step 1: Question Number Auto-Align (Hanging Indent) ---
    Call PerformWildcardReplace("([" & BenDigits() & "]{1,2})" & BenDdari() & "[ ]{1,}", "\1" & BenDdari() & "^t", , 0.5)
    
    ' --- Step 8: Convert 4-line Options to 2-line ---
    Call PerformWildcardReplace("(\(" & ChrW(&H995) & "\)*)^13(\(" & ChrW(&H996) & "\)*)^13(\(" & ChrW(&H997) & "\)*)^13(\(" & ChrW(&H998) & "\)*)", "\1^t\2^p\3^t\4")
    Call PerformWildcardReplace("(" & ChrW(&H995) & "*)^13(" & ChrW(&H996) & "*)^13(" & ChrW(&H997) & "*)^13(" & ChrW(&H998) & "*)", "\1^t\2^p\3^t\4")
    
    ' --- Step 9: Convert Options to 4 separate lines ---
    Call PerformWildcardReplace("(\(" & ChrW(&H995) & "\)*)(\(" & ChrW(&H996) & "\)*)(\(" & ChrW(&H997) & "\)*)(\(" & ChrW(&H998) & "\)*)", "\1^p\2^p\3^p\4")
    
    ' --- Step 14: Roman Numeral Formatting ---
    Call PerformWildcardReplace("(<[ivx]{1,3}\.)[ ]", "\1^t", , 0.5)
    
    ' --- Step 3: Remove Option Brackets and Set Font ---
    Call PerformWildcardReplace("\(([" & BenLetters() & "])\)", "\1", "NesarulOMR")
    Call PerformWildcardReplace("([" & BenLetters() & "])\)", "\1", "NesarulOMR")
    
    ' --- Step 2: Convert Options on Same Line to 2-line ---
    Call PerformWildcardReplace("(\(" & ChrW(&H995) & "\)*)(\(" & ChrW(&H996) & "\)*)(\(" & ChrW(&H997) & "\)*)(\(" & ChrW(&H998) & "\)*)", "\1^t\2^p\3^t\4")
    
    ' --- Step 4: Option Auto-Indentation (Converting Spaces to Tabs) ---
    Call PerformWildcardReplace("([" & BenLetters() & "])[ ]{1,}", "\1^t")
    Call PerformWildcardReplace("([" & BenLetters() & "]).[ ]{1,}", "\1.^t")
    Call PerformWildcardReplace("\(([" & BenLetters() & "])\)[ ]{1,}", "(\1)^t")
    
    ' --- Step 10: Remove Spaces Inside Option Brackets ---
    Call PerformWildcardReplace("\([ ]{0,}([" & BenLettersRange() & "])[ ]{0,}\)", "(\1)")
    
    ' --- Step 11: Convert Bracketless Options to Bracketed ---
    Call PerformWildcardReplace("([" & BenLetters() & "])[\. ][ ]{1,}", "(\1) ")
    
    ' --- Step 13: Add Dot or Brackets to Option Letters on Separate Lines ---
    Call PerformWildcardReplace("(^13)([" & BenLetters() & "])[ ]", "\1(\2) ")
End Sub

' =========================================================================
' Private Sub: Execute_CQ_Formatting
' Description: Step-by-step CQ Formatting
' =========================================================================
Private Sub Execute_CQ_Formatting()
    ' --- Step 5: Preserve Tab after Digit and Ddari (General Preparation) ---
    Call PerformWildcardReplace("([" & BenDigits() & "]{1,2})" & BenDdari() & "^9", "\1" & BenDdari() & "^t")
    
    ' --- Step 12: Remove Extra Spaces (General Preparation) ---
    Call PerformWildcardReplace("[ ]{2,}", " ")
    
    ' --- Step 1: Question Number Auto-Align (Hanging Indent) ---
    Call PerformWildcardReplace("([" & BenDigits() & "]{1,2})" & BenDdari() & "[ ]{1,}", "\1" & BenDdari() & "^t", , 0.5)
    
    ' --- Step 3: Remove Option Brackets and Set Font ---
    Call PerformWildcardReplace("\(([" & BenLetters() & "])\)", "\1", "NesarulOMR")
    
    ' --- Step 6: Move Marks to the Right ---
    Call PerformWildcardReplace("([\?" & BenDdari() & "])[ ]{1,}([" & BenDigitsAll() & "]{1,2})", "\1^t\2")
    
    ' --- Step 7: Move Marks in Parentheses to the Right ---
    Call PerformWildcardReplace("([\?" & BenDdari() & "])[ ]{1,}(\([" & BenDigits() & "]{1,2}\))", "\1^t\2")
End Sub

' =========================================================================
' Private Sub: PerformWildcardReplace
' Description: Helper method that executes search & replace with formatting.
'              To generate hanging indent, LeftIndent and FirstLineIndent
'              are set programmatically to bypass VBA property limitations.
' =========================================================================
Private Sub PerformWildcardReplace(ByVal findTxt As String, ByVal replaceTxt As String, _
    Optional ByVal fontName As String = "", _
    Optional ByVal hangingIndentVal As Double = -1)
    
    Dim rng As Range
    Set rng = ActiveDocument.Content
    
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        
        .Text = findTxt
        .Replacement.Text = replaceTxt
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        
        If fontName <> "" Or hangingIndentVal >= 0 Then
            .Format = True
            If fontName <> "" Then
                .Replacement.Font.Name = fontName
            End If
            If hangingIndentVal >= 0 Then
                .Replacement.ParagraphFormat.LeftIndent = Application.InchesToPoints(hangingIndentVal)
                .Replacement.ParagraphFormat.FirstLineIndent = Application.InchesToPoints(-hangingIndentVal)
            End If
        Else
            .Format = False
        End If
        
        .Execute Replace:=wdReplaceAll
    End With
End Sub

' =========================================================================
' Unicode Character Helper Functions (Bypasses VBE ANSI restrictions)
' =========================================================================
Private Function BenDdari() As String
    BenDdari = ChrW(&H964) ' । (Bengali Ddari)
End Function

Private Function BenLetters() As String
    BenLetters = ChrW(&H995) & ChrW(&H996) & ChrW(&H997) & ChrW(&H998) ' ক, খ, গ, ঘ
End Function

Private Function BenLettersRange() As String
    BenLettersRange = ChrW(&H995) & "-" & ChrW(&H998) ' ক-ঘ (Range)
End Function

Private Function BenDigits() As String
    BenDigits = ChrW(&H9E6) & "-" & ChrW(&H9EF) ' ০-৯ (Digit Range)
End Function

Private Function BenDigitsAll() As String
    ' ১২৩৪৫৬৭৮৯০
    BenDigitsAll = ChrW(&H9E7) & ChrW(&H9E8) & ChrW(&H9E9) & ChrW(&H9EA) & _
                   ChrW(&H9EB) & ChrW(&H9EC) & ChrW(&H9ED) & ChrW(&H9EE) & _
                   ChrW(&H9EF) & ChrW(&H9E6)
End Function