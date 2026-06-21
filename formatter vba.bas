' =========================================================================
' Unified Question Paper Formatter (MCQ & CQ) with Auto-Shortcut
' Shortcut Key: Ctrl + Alt + Q
' =========================================================================

Sub AutoExec()
    On Error Resume Next
    Application.CustomizationContext = NormalTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyQ), _
        KeyCategory:=wdKeyCategoryMacro, Command:="Format_Question_Paper"
End Sub

Sub SetShortcutNow()
    Call AutoExec
    MsgBox "Shortcut Ctrl+Alt+Q has been assigned to the Macro.", vbInformation
End Sub

' =========================================================================
' Unified Question Paper Formatter (MCQ & CQ) - Customized Version
' Instructions: Copy the entire code block below and paste it into a single 
'               VBA module in your Word Normal project.
' =========================================================================

Sub Format_Question_Paper()
    ' =========================================================================
    ' Macro Name: Format_Question_Paper
    ' Purpose: Asks the user to choose MCQ or CQ and executes formatting
    ' =========================================================================
    On Error GoTo ErrHandler
    
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
    
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbNewLine & vbNewLine & _
           "The macro was interrupted. Screen has been restored.", vbCritical, "Formatting Error"
End Sub

' =========================================================================
' Private Sub: Execute_MCQ_Formatting
' Description: Step-by-step MCQ Formatting
' =========================================================================
Private Sub Execute_MCQ_Formatting()
    On Error GoTo MCQErr
    ' --- General: Replace manual line breaks with paragraph marks ---
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^l"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' --- General: Remove empty paragraphs (single-pass ^13^13@, no loop, avoids hang) ---
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^13^13@"
        .Replacement.Text = "^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    ' --- General: Remove trailing whitespace before paragraph breaks ---
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^w^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' --- General: Remove empty paragraphs (2nd pass, catches new empties from ^w^p) ---
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^13^13@"
        .Replacement.Text = "^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    ' --- Step 1: Question Number Auto-Align (Hanging Indent: 0.3") ---
    Call PerformWildcardReplace(findTxt:="([" & BenDigits() & "]{1,2})" & BenDdari() & "[ ]{1,}", replaceTxt:="\1" & BenDdari() & "^t", hangingIndentVal:=0.3)
    
    ' --- Step 8: Convert 4-line Options to 2-line (VBA loop, avoids ^13 wildcard issues) ---
    Call ConvertFourLinesToTwo
    
    ' --- Step 9: Convert Options to 4 separate lines ---
    Call PerformWildcardReplace("(\(" & ChrW(&H995) & "\))(\(" & ChrW(&H996) & "\))(\(" & ChrW(&H997) & "\))(\(" & ChrW(&H998) & "\))", "\1^p\2^p\3^p\4")
    
    ' --- Step 14: Roman Numeral Formatting ---
    Call PerformWildcardReplace(findTxt:="(<[ivx]{1,3}\.)[ ]", replaceTxt:="\1^t", leftIndentVal:=0.3, hangingIndentVal:=0.3)
    
    ' --- Step 3: Remove Option Brackets/Dots and Set Font (if-else) ---
    ' Use PatternExists to determine which format is present, then apply only that one
    If PatternExists("\(([" & BenLetters() & "])\)") Then
        Call PerformWildcardReplace(findTxt:="\(([" & BenLetters() & "])\)", replaceTxt:="\1", fontName:="NesarulOMR")
    ElseIf PatternExists("([" & BenLetters() & "])[.][ ]") Then
        Call PerformWildcardReplace("([" & BenLetters() & "])[.][ ]", "\1", fontName:="NesarulOMR")
    ElseIf PatternExists("([" & BenLetters() & "])[ ]{1,}") Then
        Call PerformWildcardReplace("([" & BenLetters() & "])[ ]{1,}", "\1", fontName:="NesarulOMR")
    Else
        Call PerformWildcardReplace("([" & BenLetters() & "])\)", "\1", fontName:="NesarulOMR")
    End If
    
' --- Step 3b: Apply left indent + tab stop to merged option lines ---
    Dim p As Paragraph, ptxt As String, firstChar As String
    For Each p In ActiveDocument.Paragraphs
        ptxt = Trim(p.Range.Text)
        If Len(ptxt) > 0 And InStr(ptxt, vbTab) > 0 Then
            firstChar = Left(ptxt, 1)
            If firstChar = "(" Then
                firstChar = Mid(ptxt, 2, 1)
            End If
            
            If InStr(BenLetters(), firstChar) > 0 Then
                With p.Range.ParagraphFormat
                    .LeftIndent = Application.InchesToPoints(0.6)
                    .FirstLineIndent = Application.InchesToPoints(-0.3)
                    .TabStops.ClearAll
                    .TabStops.Add Position:=Application.InchesToPoints(2), Alignment:=wdAlignTabLeft
                End With
            End If
        End If
    Next
    
    Exit Sub
MCQErr:
    MsgBox "Error " & Err.Number & " at MCQ Step " & Err.Description, vbExclamation, "MCQ Error"
End Sub

' =========================================================================
' Private Sub: Execute_CQ_Formatting
' Description: Step-by-step CQ Formatting (Customized)
' =========================================================================
Private Sub Execute_CQ_Formatting()
    On Error GoTo CQErr
    ' --- General: Replace manual line breaks with paragraph marks ---
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^l"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' --- General: Remove empty paragraphs (single-pass ^13^13@, no loop, avoids hang) ---
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^13^13@"
        .Replacement.Text = "^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    ' --- General: Remove trailing whitespace before paragraph breaks ---
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^w^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' --- General: Remove empty paragraphs (2nd pass, catches new empties from ^w^p) ---
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^13^13@"
        .Replacement.Text = "^13"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    
    ' --- Step 5: Preserve Tab after Digit and Ddari (General Preparation) ---
    Call PerformWildcardReplace("([" & BenDigits() & "]{1,2})" & BenDdari() & "^9", "\1" & BenDdari() & "^t")
    
    ' --- Step 12: Remove Extra Spaces (General Preparation) ---
    Call PerformWildcardReplace("[ ]{2,}", " ")
    
    ' --- Step 1: Question Number Auto-Align (Hanging Indent) ---
    ' Applied: Alignment = Justify, Hanging Indent = 0.3"
    Call PerformWildcardReplace(findTxt:="([" & BenDigits() & "]{1,2})" & BenDdari() & "[ ]{1,}", replaceTxt:="\1" & BenDdari() & "^t", hangingIndentVal:=0.3, alignVal:=wdAlignParagraphJustify)
    
    ' --- Step 3: Remove Option Brackets and Set Font ---
    ' Applied: Left Indent = 0.3", Hanging Indent = 0.3", Brackets Kept, Font unmodified
    Call PerformWildcardReplace(findTxt:="\(([" & BenLetters() & "])\)", replaceTxt:="(\1)", leftIndentVal:=0.3, hangingIndentVal:=0.3)
    
    ' --- Step 6: Move Marks to the Right ---
    ' Applied: Tab Stop Position = 5" (Right Aligned)
    Call PerformWildcardReplace(findTxt:="([\?" & BenDdari() & "])[ ]{1,}([" & BenDigitsAll() & "]{1,2})", replaceTxt:="\1^t\2", tabStopPos:=5, tabStopAlign:=wdAlignTabRight)
    
    ' --- Step 7: Move Marks in Parentheses to the Right ---
    ' Applied: Tab Stop Position = 5" (Right Aligned, consistent with Step 6)
    Call PerformWildcardReplace(findTxt:="([\?" & BenDdari() & "])[ ]{1,}(\([" & BenDigits() & "]{1,2}\))", replaceTxt:="\1^t\2", tabStopPos:=5, tabStopAlign:=wdAlignTabRight)
    Exit Sub
CQErr:
    MsgBox "Error " & Err.Number & " at CQ Step " & Err.Description, vbExclamation, "CQ Error"
End Sub

' =========================================================================
' Private Sub: ConvertFourLinesToTwo
' Description: Converts 4 consecutive option lines (ক-ঘ) into 2 lines.
'              Avoids ^13 in wildcard patterns (which causes error 5560).
' =========================================================================
Private Sub ConvertFourLinesToTwo()
    ' Bengali characters (bypass VBE ANSI limitation)
    Dim cK As String, cKh As String, cG As String, cGh As String
    cK = ChrW(&H995): cKh = ChrW(&H996): cG = ChrW(&H997): cGh = ChrW(&H998)
    
    Dim para As Paragraph
    Dim groups As Collection
    Set groups = New Collection
    
    ' Helper: get clean text (remove trailing vbCr/vbLf)
    Dim cleanTxt As String
    
    ' First pass: collect starting paragraphs of each (ক)-(খ)-(গ)-(ঘ) group
    For Each para In ActiveDocument.Paragraphs
        cleanTxt = Trim(para.Range.Text)
        If Len(cleanTxt) > 0 Then
            Do While Right(cleanTxt, 1) = vbCr Or Right(cleanTxt, 1) = vbLf
                cleanTxt = Left(cleanTxt, Len(cleanTxt) - 1)
            Loop
        End If
        If Len(cleanTxt) = 0 Then GoTo NextPara
        
        ' Check if this paragraph starts with (ক) or ক
        Dim startLabel As String
        startLabel = Left(cleanTxt, 1)
        If startLabel = "(" Then startLabel = Left(cleanTxt, 2)
        
        If startLabel = "(" & cK Or startLabel = cK Then
            Dim p2 As Paragraph, p3 As Paragraph, p4 As Paragraph
            
            ' Guard each .Next against Nothing
            If para.Next Is Nothing Then GoTo NextPara
            Set p2 = para.Next
            
            If p2.Next Is Nothing Then GoTo NextPara
            Set p3 = p2.Next
            
            If p3.Next Is Nothing Then GoTo NextPara
            Set p4 = p3.Next
            
            If (Not p2 Is Nothing) And (Not p3 Is Nothing) And (Not p4 Is Nothing) Then
                Dim t2 As String, t3 As String, t4 As String
                t2 = Trim(p2.Range.Text): If Len(t2) > 0 Then Do While Right(t2, 1) = vbCr Or Right(t2, 1) = vbLf: t2 = Left(t2, Len(t2) - 1): Loop
                t3 = Trim(p3.Range.Text): If Len(t3) > 0 Then Do While Right(t3, 1) = vbCr Or Right(t3, 1) = vbLf: t3 = Left(t3, Len(t3) - 1): Loop
                t4 = Trim(p4.Range.Text): If Len(t4) > 0 Then Do While Right(t4, 1) = vbCr Or Right(t4, 1) = vbLf: t4 = Left(t4, Len(t4) - 1): Loop
                
                If Len(t2) > 0 And Len(t3) > 0 And Len(t4) > 0 Then
                    Dim l2 As String, l3 As String, l4 As String
                    l2 = Left(t2, 1): If l2 = "(" Then l2 = Left(t2, 2)
                    l3 = Left(t3, 1): If l3 = "(" Then l3 = Left(t3, 2)
                    l4 = Left(t4, 1): If l4 = "(" Then l4 = Left(t4, 2)
                    
                    If (l2 = "(" & cKh Or l2 = cKh) And (l3 = "(" & cG Or l3 = cG) And (l4 = "(" & cGh Or l4 = cGh) Then
                        groups.Add para
                    End If
                End If
            End If
        End If
NextPara:
    Next para
    
    ' Second pass: merge in reverse order to preserve paragraph indices
    Dim i As Long
    For i = groups.Count To 1 Step -1
        Set para = groups(i)
        
        ' Guard .Next against Nothing
        If para.Next Is Nothing Then GoTo NextMerge
        Set p2 = para.Next
        If p2.Next Is Nothing Then GoTo NextMerge
        Set p3 = p2.Next
        If p3.Next Is Nothing Then GoTo NextMerge
        Set p4 = p3.Next
        
        cleanTxt = para.Range.Text
        If Len(cleanTxt) > 0 Then Do While Right(cleanTxt, 1) = vbCr Or Right(cleanTxt, 1) = vbLf: cleanTxt = Left(cleanTxt, Len(cleanTxt) - 1): Loop
        t2 = p2.Range.Text
        If Len(t2) > 0 Then Do While Right(t2, 1) = vbCr Or Right(t2, 1) = vbLf: t2 = Left(t2, Len(t2) - 1): Loop
        t3 = p3.Range.Text
        If Len(t3) > 0 Then Do While Right(t3, 1) = vbCr Or Right(t3, 1) = vbLf: t3 = Left(t3, Len(t3) - 1): Loop
        t4 = p4.Range.Text
        If Len(t4) > 0 Then Do While Right(t4, 1) = vbCr Or Right(t4, 1) = vbLf: t4 = Left(t4, Len(t4) - 1): Loop
        
        ' Delete paragraphs 4 and 3 first (reverse order, safe; doesn't affect para or p2)
        On Error Resume Next
        p4.Range.Delete
        p3.Range.Delete
        
        ' Overwrite para (original (ক)) with first merged line
        para.Range.Text = cleanTxt & vbTab & t2 & vbCr
        
        ' Overwrite p2 (original (খ)) with second merged line
        p2.Range.Text = t3 & vbTab & t4 & vbCr
        Err.Clear
        On Error GoTo 0
NextMerge:
    Next i
End Sub

' =========================================================================
' Private Sub: PerformWildcardReplace
' Description: Expanded helper method supporting custom alignments and 
'              right tab stops in addition to fonts and indents.
' =========================================================================
Private Sub PerformWildcardReplace(ByVal findTxt As String, ByVal replaceTxt As String, _
    Optional ByVal fontName As String = "", _
    Optional ByVal leftIndentVal As Double = -1, _
    Optional ByVal hangingIndentVal As Double = -1, _
    Optional ByVal alignVal As Long = -1, _
    Optional ByVal tabStopPos As Double = -1, _
    Optional ByVal tabStopAlign As WdTabAlignment = wdAlignTabRight)
    
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
        
        If fontName <> "" Or leftIndentVal >= 0 Or hangingIndentVal >= 0 Or alignVal >= 0 Or tabStopPos >= 0 Then
            .Format = True
            If fontName <> "" Then
                .Replacement.Font.Name = fontName
            End If
            If hangingIndentVal >= 0 Then
                Dim finalLeftIndent As Double
                If leftIndentVal >= 0 Then
                    finalLeftIndent = leftIndentVal + hangingIndentVal
                Else
                    finalLeftIndent = hangingIndentVal
                End If
                .Replacement.ParagraphFormat.LeftIndent = Application.InchesToPoints(finalLeftIndent)
                .Replacement.ParagraphFormat.FirstLineIndent = Application.InchesToPoints(-hangingIndentVal)
            End If
            If alignVal >= 0 Then
                .Replacement.ParagraphFormat.Alignment = alignVal
            End If
            If tabStopPos >= 0 Then
                ' Clears previous tab stops in the paragraph before setting the new customized Right Tab
                .Replacement.ParagraphFormat.TabStops.ClearAll
                .Replacement.ParagraphFormat.TabStops.Add _
                    Position:=Application.InchesToPoints(tabStopPos), _
                    Alignment:=tabStopAlign, _
                    Leader:=wdTabLeaderSpaces
            End If
        Else
            .Format = False
        End If
        
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Force-apply paragraph formatting when replacement text equals found text
    ' (Word skips .Replacement.ParagraphFormat in that case)
    If hangingIndentVal >= 0 And fontName = "" And alignVal < 0 And tabStopPos < 0 Then
        Dim fmtRange As Range
        Set fmtRange = ActiveDocument.Content
        With fmtRange.Find
            .ClearFormatting
            .Text = findTxt
            .MatchWildcards = True
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            Dim safetyCounter As Long
            safetyCounter = 0
            Do While .Execute And safetyCounter < 10000
                .Parent.ParagraphFormat.LeftIndent = Application.InchesToPoints(finalLeftIndent)
                .Parent.ParagraphFormat.FirstLineIndent = Application.InchesToPoints(-hangingIndentVal)
                safetyCounter = safetyCounter + 1
            Loop
        End With
    End If
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

' =========================================================================
' Private Function: PatternExists
' Description: Checks if a specific wildcard pattern exists anywhere in the document
' =========================================================================
Private Function PatternExists(ByVal pattern As String) As Boolean
    Dim rng As Range
    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Text = pattern
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        PatternExists = .Execute
    End With
End Function