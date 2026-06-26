' V2.6.0 — Last updated: 2026-06-26
Private Const APP_VERSION As String = "V2.6.0"

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
    MsgBox "Shortcut Ctrl+Alt+Q has been assigned to the Macro.", vbInformation, "Shortcut - " & APP_VERSION
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
                         "Leave blank or click Cancel to exit.", "Select Question Type - " & APP_VERSION)
                         
    ' Input verification and routing
    If Trim(userInput) = "1" Then
        Application.ScreenUpdating = False
        Call Execute_MCQ_Formatting
        Application.ScreenUpdating = True
        MsgBox "MCQ Formatting completed successfully!", vbInformation, "Completed - " & APP_VERSION
        
    ElseIf Trim(userInput) = "2" Then
        Application.ScreenUpdating = False
        Call Execute_CQ_Formatting
        Application.ScreenUpdating = True
        MsgBox "CQ Formatting completed successfully!", vbInformation, "Completed - " & APP_VERSION
        
    ElseIf Trim(userInput) = "" Then
        MsgBox "Formatting process canceled.", vbInformation, "Canceled - " & APP_VERSION
    Else
        MsgBox "Invalid choice! Please enter either 1 or 2.", vbExclamation, "Error - " & APP_VERSION
    End If
    
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbNewLine & vbNewLine & _
           "The macro was interrupted. Screen has been restored.", vbCritical, "Formatting Error - " & APP_VERSION
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
    
    ' --- General: Clean redundant empty paragraphs ---
    Call GeneralCleanup
    
    ' --- Step 0: Convert 1-line options to 2-line with tabs (style-blind) ---
    ' Split গ to new line: handles (গ), গ., গ), and bare গ (safe: requires space after)
    Call PerformWildcardReplace("[ ]{1,}\(" & ChrW(&H997) & "\)", "^p" & ChrW(&H997))
    Call PerformWildcardReplace("[ ]{1,}" & ChrW(&H997) & "\.", "^p" & ChrW(&H997))
    Call PerformWildcardReplace("[ ]{1,}" & ChrW(&H997) & "\)", "^p" & ChrW(&H997))
    Call PerformWildcardReplace("[ ]{1,}" & ChrW(&H997) & "[ ]", "^p" & ChrW(&H997) & " ")
    
    ' Move খ to column 2: handles (খ), খ., খ), and bare খ (safe: requires space after)
    Call PerformWildcardReplace("[ ]{1,}\(" & ChrW(&H996) & "\)", "^t" & ChrW(&H996))
    Call PerformWildcardReplace("[ ]{1,}" & ChrW(&H996) & "\.", "^t" & ChrW(&H996))
    Call PerformWildcardReplace("[ ]{1,}" & ChrW(&H996) & "\)", "^t" & ChrW(&H996))
    Call PerformWildcardReplace("[ ]{1,}" & ChrW(&H996) & "[ ]", "^t" & ChrW(&H996) & " ")
    
    ' Move ঘ to column 2: handles (ঘ), ঘ., ঘ), and bare ঘ (safe: requires space after)
    Call PerformWildcardReplace("[ ]{1,}\(" & ChrW(&H998) & "\)", "^t" & ChrW(&H998))
    Call PerformWildcardReplace("[ ]{1,}" & ChrW(&H998) & "\.", "^t" & ChrW(&H998))
    Call PerformWildcardReplace("[ ]{1,}" & ChrW(&H998) & "\)", "^t" & ChrW(&H998))
    Call PerformWildcardReplace("[ ]{1,}" & ChrW(&H998) & "[ ]", "^t" & ChrW(&H998) & " ")
    
    ' Clean stray ) after splits (^13 in find, ^p in replace)
    Call PerformWildcardReplace("^13\)", "^p")
    
    ' --- Step 1: Question Number Auto-Align (Hanging Indent: 0.3") ---
    Call PerformWildcardReplace(findTxt:="([" & BenDigits() & "]{1,2})" & BenDdari() & "[ ]{1,}", replaceTxt:="\1" & BenDdari() & "^t", hangingIndentVal:=0.3)
    
    ' --- Step 8: Convert 4-line Options to 2-line (VBA loop, avoids ^13 wildcard issues) ---
    Call ConvertFourLinesToTwo
    
    ' --- Step 9: Convert Options to 4 separate lines ---
    Call PerformWildcardReplace("(\(" & ChrW(&H995) & "\))(\(" & ChrW(&H996) & "\))(\(" & ChrW(&H997) & "\))(\(" & ChrW(&H998) & "\))", "\1^p\2^p\3^p\4")
    
    ' --- Step 14: Roman Numeral Formatting ---
    Call PerformWildcardReplace(findTxt:="(<[ivx]{1,3}\.)[ ]", replaceTxt:="\1^t", leftIndentVal:=0.3, hangingIndentVal:=0.3)
    
    ' --- Step 3: Position-based label cleanup and font formatting ---
    Call FormatMCQLabels
    
    ' --- Step 4: Stimulus paragraph indent ---
    Call FormatStimulusIndent
    
    Exit Sub
MCQErr:
    MsgBox "Error " & Err.Number & " at MCQ Step " & Err.Description, vbExclamation, "MCQ Error - " & APP_VERSION
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
    
    ' --- General: Clean redundant empty paragraphs ---
    Call GeneralCleanup
    
    ' --- Merge: Merge number-only lines with CQ sub-question paragraphs ---
    Call MergeCQNumberAndLabel
    
    ' --- Step 5: Preserve Tab after Digit and Ddari (General Preparation) ---
    Call PerformWildcardReplace("([" & BenDigits() & "]{1,2})" & BenDdari() & "^9", "\1" & BenDdari() & "^t")
    
    ' --- Step 12: Remove Extra Spaces (General Preparation) ---
    Call PerformWildcardReplace("[ ]{2,}", " ")
    
    ' --- Step 1: Question Number Auto-Align (Hanging Indent) ---
    ' Applied: Alignment = Justify, Hanging Indent = 0.3"
    Call PerformWildcardReplace(findTxt:="([" & BenDigits() & "]{1,2})" & BenDdari() & "[ ]{1,}", replaceTxt:="\1" & BenDdari() & "^t", hangingIndentVal:=0.3, alignVal:=wdAlignParagraphJustify)
    
    ' --- Step 3: Format CQ sub-question labels (LeftIndent = 0.3") ---
    Call FormatCQLabels
    
    ' --- Step 6: Move Marks to the Right ---
    ' Applied: Tab Stop Position = 5" (Right Aligned)
    Call PerformWildcardReplace(findTxt:="([\?" & BenDdari() & "])[ ]{1,}([" & BenDigitsAll() & "]{1,2})", replaceTxt:="\1^t\2", tabStopPos:=5, tabStopAlign:=wdAlignTabRight)
    
    ' --- Step 7: Move Marks in Parentheses to the Right ---
    ' Applied: Tab Stop Position = 5" (Right Aligned, consistent with Step 6)
    Call PerformWildcardReplace(findTxt:="([\?" & BenDdari() & "])[ ]{1,}(\([" & BenDigits() & "]{1,2}\))", replaceTxt:="\1^t\2", tabStopPos:=5, tabStopAlign:=wdAlignTabRight)
    
    ' --- Step 9: Right-align trailing marks in CQ paragraphs (no preceding Ddari/?) ---
    Call AlignCQTrailingMarks
    
    ' --- Step 8: Stimulus paragraph indent ---
    Call FormatStimulusIndent
    
    Exit Sub
CQErr:
    MsgBox "Error " & Err.Number & " at CQ Step " & Err.Description, vbExclamation, "CQ Error - " & APP_VERSION
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
' Private Sub: FormatMCQLabels
' Description: Paragraph-loop approach to strip decoration and apply
'              NesarulOMR font to MCQ labels (ক, খ, গ, ঘ). Labels are
'              identified ONLY at position 0 or after a tab, preventing
'              false positives inside words (e.g. পরিবর্তক, দ্বিপাক্ষিক).
'              Also applies paragraph formatting (indent, tab stop).
' =========================================================================
Private Sub FormatMCQLabels()
    Dim para As Paragraph
    Dim rngPara As Range
    Dim ptxt As String
    Dim letters As String
    Dim foundLabel As Boolean
    Dim i As Long, j As Long
    Dim paraStart As Long
    Dim chkPos As Long
    Dim ch1 As String, ch2 As String, ch3 As String
    Dim labelData() As Variant
    Dim lblCount As Long
    
    letters = BenLetters()
    
    For Each para In ActiveDocument.Paragraphs
        Set rngPara = para.Range
        ptxt = rngPara.Text
        Do While Len(ptxt) > 0 And (Right(ptxt, 1) = vbCr Or Right(ptxt, 1) = vbLf)
            ptxt = Left(ptxt, Len(ptxt) - 1)
        Loop
        If Len(ptxt) = 0 Then GoTo NextPara
        
        paraStart = rngPara.Start
        foundLabel = False
        lblCount = 0
        ReDim labelData(1 To 10, 1 To 4)
        
        ' --- Check position 1 (start of paragraph) ---
        chkPos = 1
        If chkPos <= Len(ptxt) Then
            ch1 = Mid(ptxt, chkPos, 1)
            ch2 = Mid(ptxt, chkPos + 1, 1)
            ch3 = Mid(ptxt, chkPos + 2, 1)
            
            ' (ক) at start (with lookahead: 다음 문자는 유효한 구분 기호여야 함)
            If ch1 = "(" And ch2 <> "" And InStr(letters, ch2) > 0 And ch3 = ")" _
               And IsValidLabelNextChar(Mid(ptxt, chkPos + 3, 1)) Then
                lblCount = lblCount + 1
                If lblCount > UBound(labelData, 1) Then ReDim Preserve labelData(1 To lblCount + 10, 1 To 4)
                labelData(lblCount, 1) = chkPos + 1
                labelData(lblCount, 2) = True
                labelData(lblCount, 3) = True
                labelData(lblCount, 4) = ")"
                foundLabel = True
                GoTo CheckTabs
            End If
            
            ' ক, ক., ক), bare ক at start (lookahead: prevents গতি/কোন matches)
            If InStr(letters, ch1) > 0 And IsValidLabelNextChar(ch2) Then
                lblCount = lblCount + 1
                If lblCount > UBound(labelData, 1) Then ReDim Preserve labelData(1 To lblCount + 10, 1 To 4)
                labelData(lblCount, 1) = chkPos
                labelData(lblCount, 2) = False
                If ch2 = "." Or ch2 = ")" Then
                    labelData(lblCount, 3) = True
                    labelData(lblCount, 4) = ch2
                Else
                    labelData(lblCount, 3) = False
                    labelData(lblCount, 4) = ""
                End If
                foundLabel = True
            End If
        End If
        
CheckTabs:
        ' --- Check positions after each tab ---
        For i = 1 To Len(ptxt)
            If Mid(ptxt, i, 1) = vbTab Then
                chkPos = i + 1
                If chkPos <= Len(ptxt) Then
                    ch1 = Mid(ptxt, chkPos, 1)
                    ch2 = Mid(ptxt, chkPos + 1, 1)
                    ch3 = Mid(ptxt, chkPos + 2, 1)
                    
                    ' (ক) after tab (with lookahead)
                    If ch1 = "(" And ch2 <> "" And InStr(letters, ch2) > 0 And ch3 = ")" _
                       And IsValidLabelNextChar(Mid(ptxt, chkPos + 3, 1)) Then
                        lblCount = lblCount + 1
                        If lblCount > UBound(labelData, 1) Then ReDim Preserve labelData(1 To lblCount + 10, 1 To 4)
                        labelData(lblCount, 1) = chkPos + 1
                        labelData(lblCount, 2) = True
                        labelData(lblCount, 3) = True
                        labelData(lblCount, 4) = ")"
                        foundLabel = True
                        GoTo NextTab
                    End If
                    
                    ' ক, ক., ক), bare ক after tab (with lookahead)
                    If InStr(letters, ch1) > 0 And IsValidLabelNextChar(ch2) Then
                        lblCount = lblCount + 1
                        If lblCount > UBound(labelData, 1) Then ReDim Preserve labelData(1 To lblCount + 10, 1 To 4)
                        labelData(lblCount, 1) = chkPos
                        labelData(lblCount, 2) = False
                        If ch2 = "." Or ch2 = ")" Then
                            labelData(lblCount, 3) = True
                            labelData(lblCount, 4) = ch2
                        Else
                            labelData(lblCount, 3) = False
                            labelData(lblCount, 4) = ""
                        End If
                        foundLabel = True
                    End If
                End If
NextTab:
            End If
        Next i
        
        ' Process labels in reverse order (high-to-low) to keep positions stable
        For j = lblCount To 1 Step -1
            Dim docLabelPos As Long
            docLabelPos = paraStart + labelData(j, 1) - 1
            
            ' Delete trailing ) or .
            If labelData(j, 3) Then
                rngPara.Document.Range(docLabelPos + 1, docLabelPos + 2).Delete
            End If
            
            ' Delete leading (
            If labelData(j, 2) Then
                rngPara.Document.Range(docLabelPos - 1, docLabelPos).Delete
                docLabelPos = docLabelPos - 1
            End If
            
            ' Apply NesarulOMR font to the single label character only
            rngPara.Document.Range(docLabelPos, docLabelPos + 1).Font.Name = "NesarulOMR"
        Next j
        
        ' Apply paragraph formatting
        If foundLabel Then
            With rngPara.ParagraphFormat
                .LeftIndent = Application.InchesToPoints(0.6)
                .FirstLineIndent = Application.InchesToPoints(-0.3)
                .TabStops.ClearAll
                .TabStops.Add Position:=Application.InchesToPoints(2), Alignment:=wdAlignTabLeft
            End With
        End If
        
NextPara:
    Next para
End Sub

' =========================================================================
' Private Sub: FormatCQLabels
' Description: Paragraph-loop approach to detect CQ sub-question labels
'              (ক, খ, গ, ঘ) at position 0 in all decoration styles:
'              (ক), ক., খ), গ (bare with space). Inserts a tab after the
'              label delimiter so hanging indent works correctly.
' =========================================================================
Private Sub FormatCQLabels()
    Dim para As Paragraph
    Dim ptxt As String
    Dim rng As Range
    Dim letters As String
    Dim foundLabel As Boolean
    Dim firstCh As String, secondCh As String, thirdCh As String, fourthCh As String
    
    letters = BenLetters()
    
    For Each para In ActiveDocument.Paragraphs
        ptxt = para.Range.Text
        Do While Len(ptxt) > 0 And (Right(ptxt, 1) = vbCr Or Right(ptxt, 1) = vbLf)
            ptxt = Left(ptxt, Len(ptxt) - 1)
        Loop
        If Len(ptxt) = 0 Then GoTo NextPara
        
        foundLabel = False
        firstCh = Left(ptxt, 1)
        secondCh = Mid(ptxt, 2, 1)
        thirdCh = Mid(ptxt, 3, 1)
        fourthCh = Mid(ptxt, 4, 1)
        
        Set rng = para.Range
        
        ' Pattern 1: (ক) — opening paren + letter + closing paren at position 0
        If firstCh = "(" And InStr(letters, secondCh) > 0 And thirdCh = ")" _
           And IsValidLabelNextChar(fourthCh) Then
            rng.Start = rng.Start + 3
            rng.End = rng.Start
            If fourthCh = " " Then
                rng.End = rng.Start + 1
            End If
            rng.Text = vbTab
            foundLabel = True
        End If
        
        ' Pattern 2: ক. or খ) — letter + dot or closing paren at position 0
        If Not foundLabel And InStr(letters, firstCh) > 0 _
           And (secondCh = "." Or secondCh = ")") _
           And IsValidLabelNextChar(thirdCh) Then
            rng.Start = rng.Start + 2
            rng.End = rng.Start
            If thirdCh = " " Then
                rng.End = rng.Start + 1
            End If
            rng.Text = vbTab
            foundLabel = True
        End If
        
        ' Pattern 3: গ (bare letter followed by space at position 0)
        If Not foundLabel And InStr(letters, firstCh) > 0 _
           And secondCh = " " Then
            rng.Start = rng.Start + 1
            rng.End = rng.Start + 1
            rng.Text = vbTab
            foundLabel = True
        End If
        
        ' Skip if paragraph starts with Bengali digit(s) + Ddari (already merged or number-only)
        If firstCh >= ChrW(&H9E6) And firstCh <= ChrW(&H9EF) Then GoTo NextPara
        
        If foundLabel Then
            para.Range.ParagraphFormat.LeftIndent = Application.InchesToPoints(0.3)
        End If
        
NextPara:
    Next para
End Sub

' =========================================================================
' Private Sub: MergeCQNumberAndLabel
' Description: Merges a number-only paragraph (e.g. ১০।) with its next
'              paragraph if the next starts with a CQ sub-question label
'              (ক, খ, গ, ঘ). Replaces paragraph mark with a tab and applies
'              0.6" hanging indent.
' =========================================================================
Private Sub MergeCQNumberAndLabel()
    Dim i As Long
    Dim para As Paragraph
    Dim ptxt As String
    Dim nextPara As Paragraph
    Dim nextTxt As String
    Dim mergeRange As Range
    Dim letters As String
    
    letters = BenLetters()
    i = 1
    
    Do While i < ActiveDocument.Paragraphs.Count
        Set para = ActiveDocument.Paragraphs(i)
        ptxt = Trim(para.Range.Text)
        Do While Len(ptxt) > 0 And (Right(ptxt, 1) = vbCr Or Right(ptxt, 1) = vbLf)
            ptxt = Left(ptxt, Len(ptxt) - 1)
        Loop
        
        ' Check if current paragraph is number-only (digits + Ddari)
        If IsNumberOnlyLine(ptxt) Then
            Set nextPara = ActiveDocument.Paragraphs(i + 1)
            nextTxt = nextPara.Range.Text
            Do While Len(nextTxt) > 0 And (Right(nextTxt, 1) = vbCr Or Right(nextTxt, 1) = vbLf)
                nextTxt = Left(nextTxt, Len(nextTxt) - 1)
            Loop
            
            ' Check if next paragraph starts with a CQ label
            If StartsWithCQLabel(nextTxt) Then
                ' Build range covering both paragraphs (excluding nextPara's final ^p)
                Set mergeRange = ActiveDocument.Range(para.Range.Start, nextPara.Range.End - 1)
                
                ' Replace the ^p separating them with a tab
                With mergeRange.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = "^p"
                    .Replacement.Text = "^t"
                    .Execute Replace:=wdReplaceOne
                End With
                
                ' Apply formatting to merged paragraph
                With ActiveDocument.Paragraphs(i).Range.ParagraphFormat
                    .LeftIndent = Application.InchesToPoints(0.6)
                    .FirstLineIndent = Application.InchesToPoints(-0.6)
                    .Alignment = wdAlignParagraphJustify
                End With
                
                ' Don't increment i — paragraphs shifted up after merge
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
End Sub

' =========================================================================
' Private Sub: AlignCQTrailingMarks
' Description: For CQ sub-question paragraphs ending with Bengali digits
'              (mark value), inserts a tab before the trailing digit and
'              sets a right tab stop at 5". Catches marks that don't have
'              a preceding Ddari or question mark.
' =========================================================================
Private Sub AlignCQTrailingMarks()
    Dim para As Paragraph
    Dim ptxt As String
    Dim rng As Range
    Dim lastCh As String, prevCh As String
    Dim digits As String
    
    digits = BenDigitsAll()
    
    For Each para In ActiveDocument.Paragraphs
        ptxt = para.Range.Text
        Do While Len(ptxt) > 0 And (Right(ptxt, 1) = vbCr Or Right(ptxt, 1) = vbLf)
            ptxt = Left(ptxt, Len(ptxt) - 1)
        Loop
        If Len(ptxt) = 0 Then GoTo NextTPara
        
        ' Only process CQ paragraphs (start with CQ label or digits+Ddari)
        If Not StartsWithCQLabel(ptxt) And Not StartsWithNumberDdari(ptxt) Then
            GoTo NextTPara
        End If
        
        ' Check if last character is a Bengali digit (1-2 digits)
        lastCh = Right(ptxt, 1)
        If InStr(digits, lastCh) = 0 Then GoTo NextTPara
        
        ' Check if there's already a tab before the digit
        If Len(ptxt) >= 2 Then
            prevCh = Mid(ptxt, Len(ptxt) - 1, 1)
            If prevCh = vbTab Then GoTo NextTPara
        End If
        
        ' Also check for 2-digit marks (e.g. ১০, ১২)
        Dim markLen As Long
        markLen = 1
        If Len(ptxt) >= 2 Then
            Dim secondLast As String
            secondLast = Mid(ptxt, Len(ptxt) - 1, 1)
            If InStr(digits, secondLast) > 0 Then
                markLen = 2
            End If
        End If
        
        ' Insert tab before the trailing mark
        Set rng = para.Range
        rng.Start = rng.End - 1
        rng.Collapse Direction:=wdCollapseStart
        
        ' If 2-digit mark, move back one more
        If markLen = 2 Then
            rng.Start = rng.Start - 1
            rng.End = rng.Start
        End If
        
        rng.Text = vbTab
        
        ' Set right tab stop at 5"
        With para.Range.ParagraphFormat
            .TabStops.ClearAll
            .TabStops.Add Position:=Application.InchesToPoints(5), Alignment:=wdAlignTabRight
        End With
        
NextTPara:
    Next para
End Sub

' =========================================================================
' Private Sub: FormatStimulusIndent
' Description: Finds trigger paragraphs ending with ":" and applies 0.3"
'              left indent + justified alignment to the next paragraph.
'              Skips if next paragraph starts with a Bengali digit (০-৯).
' =========================================================================
Private Sub FormatStimulusIndent()
    Dim i As Long
    Dim paraText As String
    Dim firstCh As String
    
    For i = 1 To ActiveDocument.Paragraphs.Count - 1
        paraText = Trim(ActiveDocument.Paragraphs(i).Range.Text)
        Do While Len(paraText) > 0 And (Right(paraText, 1) = vbCr Or Right(paraText, 1) = vbLf)
            paraText = Left(paraText, Len(paraText) - 1)
        Loop
        
        ' Trigger: paragraph ends with colon
        If Right(paraText, 1) = ":" Then
            Dim target As Paragraph
            Set target = ActiveDocument.Paragraphs(i + 1)
            
            ' Boundary check: skip if next para starts with Bengali digit ০-৯
            Dim nextText As String
            nextText = Trim(target.Range.Text)
            If Len(nextText) > 0 Then
                firstCh = Left(nextText, 1)
                If firstCh >= ChrW(&H9E6) And firstCh <= ChrW(&H9EF) Then
                    GoTo NextPara
                End If
            End If
            
            ' Apply: 0.3" left indent, no first-line indent, justified
            With target.Range.ParagraphFormat
                .LeftIndent = Application.InchesToPoints(0.3)
                .FirstLineIndent = 0
                .Alignment = wdAlignParagraphJustify
            End With
        End If
NextPara:
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
' Private Function: IsValidLabelNextChar
' Description: Returns True if the character following a label candidate is
'              a valid delimiter (space, ), ., tab, vbCr, vbLf, or empty).
'              Returns False for Bengali letters/vowel signs (prevents word
'              matches like গতি or কোন).
' =========================================================================
Private Function IsValidLabelNextChar(ByVal ch As String) As Boolean
    IsValidLabelNextChar = (ch = " " Or ch = ")" Or ch = "." Or _
                            ch = vbTab Or ch = vbCr Or ch = vbLf Or ch = "")
End Function

' =========================================================================
' Private Function: IsNumberOnlyLine
' Description: Returns True if the text consists solely of Bengali digits
'              followed by a Ddari (।), e.g. "১০।", "৩।"
' =========================================================================
Private Function IsNumberOnlyLine(ByVal txt As String) As Boolean
    Dim t As String
    Dim ddari As String
    ddari = ChrW(&H964)
    
    t = Trim(txt)
    If Len(t) < 2 Then Exit Function
    If Right(t, 1) <> ddari Then Exit Function
    
    t = Left(t, Len(t) - 1)
    If Len(t) = 0 Then Exit Function
    
    Dim i As Long
    For i = 1 To Len(t)
        Dim c As String
        c = Mid(t, i, 1)
        If c < ChrW(&H9E6) Or c > ChrW(&H9EF) Then Exit Function
    Next i
    
    IsNumberOnlyLine = True
End Function

' =========================================================================
' Private Function: StartsWithNumberDdari
' Description: Returns True if the text starts with 1-2 Bengali digits
'              followed by a Ddari (।), e.g. "১০।", "৩। some text"
' =========================================================================
Private Function StartsWithNumberDdari(ByVal txt As String) As Boolean
    Dim t As String
    Dim ddari As String
    ddari = ChrW(&H964)
    
    t = Trim(txt)
    If Len(t) < 2 Then Exit Function
    
    ' First char must be a Bengali digit
    Dim firstCh As String
    firstCh = Left(t, 1)
    If firstCh < ChrW(&H9E6) Or firstCh > ChrW(&H9EF) Then Exit Function
    
    ' Check for Ddari at position 2 or 3 (1-digit or 2-digit number)
    If Mid(t, 2, 1) = ddari Then
        StartsWithNumberDdari = True
    ElseIf Len(t) >= 3 And Mid(t, 2, 1) >= ChrW(&H9E6) And Mid(t, 2, 1) <= ChrW(&H9EF) And Mid(t, 3, 1) = ddari Then
        StartsWithNumberDdari = True
    End If
End Function

' =========================================================================
' Private Function: StartsWithCQLabel
' Description: Returns True if the text starts with one of the CQ
'              sub-question label patterns: (ক), ক., খ), গ (bare + space).
'              Uses IsValidLabelNextChar to prevent matches inside words.
' =========================================================================
Private Function StartsWithCQLabel(ByVal txt As String) As Boolean
    Dim t As String
    Dim letters As String
    Dim firstCh As String, secondCh As String, thirdCh As String
    
    letters = BenLetters()
    t = Trim(txt)
    If Len(t) = 0 Then Exit Function
    
    firstCh = Left(t, 1)
    secondCh = Mid(t, 2, 1)
    thirdCh = Mid(t, 3, 1)
    
    ' Pattern 1: (ক)
    If firstCh = "(" And InStr(letters, secondCh) > 0 And thirdCh = ")" _
       And IsValidLabelNextChar(Mid(t, 4, 1)) Then
        StartsWithCQLabel = True
        Exit Function
    End If
    
    ' Pattern 2: ক. or খ)
    If InStr(letters, firstCh) > 0 _
       And (secondCh = "." Or secondCh = ")") _
       And IsValidLabelNextChar(thirdCh) Then
        StartsWithCQLabel = True
        Exit Function
    End If
    
    ' Pattern 3: গ (bare letter + space)
    If InStr(letters, firstCh) > 0 And secondCh = " " Then
        StartsWithCQLabel = True
        Exit Function
    End If
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

' =========================================================================
' Private Sub: GeneralCleanup
' Description: Removes redundant empty paragraphs including invisible
'              whitespace between them and leading/trailing empty paras.
'              3-phase find/replace (no wildcards) + start/end check.
' =========================================================================
Private Sub GeneralCleanup()
    ' Phase 1: Remove trailing whitespace before paragraph breaks
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
    
    ' Phase 2: Remove whitespace between paragraph marks (^13[space]^13 → ^13^13)
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^p^w"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    
    ' Phase 3: Collapse 2+ consecutive paragraph marks into 1 (loop, no wildcards)
    Dim safety As Long
    safety = 0
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        Do While .Execute(Replace:=wdReplaceAll) And safety < 100
            safety = safety + 1
        Loop
    End With
    
    ' Phase 4: Clean leading empty paragraph
    If ActiveDocument.Paragraphs.Count >= 2 Then
        Dim firstText As String
        firstText = Trim(ActiveDocument.Paragraphs(1).Range.Text)
        If Len(firstText) = 0 Then
            ActiveDocument.Paragraphs(1).Range.Delete
        End If
    End If
    
    ' Phase 5: Clean trailing empty paragraphs
    If ActiveDocument.Paragraphs.Count >= 2 Then
        Dim lastText As String
        lastText = Trim(ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count).Range.Text)
        If Len(lastText) = 0 Then
            ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count).Range.Delete
        End If
    End If
End Sub