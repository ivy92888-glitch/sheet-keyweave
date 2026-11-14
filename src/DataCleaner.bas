Attribute VB_Name = "DataCleaner"
Option Explicit

' ============================================================================
'  Tabular Data Cleaning and Rule-Based Screening Toolkit
' ----------------------------------------------------------------------------
'  A standalone, dependency-free VBA module for cleaning flat worksheet tables.
'  It provides five composable operations:
'
'    1. NormalizeDetailKeys  - group rows by a key column and fill missing
'                              detail keys, or mark them for manual review
'    2. BuildDerivedColumns  - append computed columns (line amount, group
'                              row count, group amount) as worksheet formulas
'    3. ApplyBandHighlight   - shade numeric values that fall inside two bands
'    4. FlagGroupsByRule     - flag every row of a group when the group matches
'                              a keyword set AND a numeric condition
'    5. VerifyGroupFlags     - re-read flagged rows and report false positives
'
'  Design notes
'    * Column positions are resolved from header text, never hard-coded.
'    * Thresholds, keywords, labels and comparison direction are read from an
'      optional "Settings" sheet (column A = key, column B = value).
'    * Built-in defaults below are neutral placeholders for demonstration.
'      They carry no domain meaning and should be replaced by the user.
'    * No external references are required; Scripting.Dictionary is late-bound.
'    * Application state is restored even when a step raises an error.
'
'  Tested with Microsoft Excel and WPS Spreadsheets (Windows, VBA enabled).
'  Run "RunCleaningPipeline" on the active sheet, or call any single step.
'  This is a general-purpose implementation written from scratch for reuse in
'  any tabular cleaning task; it contains no organization-specific data.
' ============================================================================

Private Const SETTINGS_SHEET    As String = "Settings"

' Neutral placeholder defaults. Override on the Settings sheet.
Private Const DEF_GROUP_HEADER    As String = "RecordID"
Private Const DEF_DETAIL_HEADER   As String = "LineKey"
Private Const DEF_COUNT_HEADER    As String = "ItemCount"
Private Const DEF_DESC_HEADER     As String = "Description"
Private Const DEF_QTY_HEADER      As String = "Quantity"
Private Const DEF_PRICE_HEADER    As String = "UnitPrice"
Private Const DEF_AMOUNT_HEADER   As String = "LineAmount"
Private Const DEF_GCOUNT_HEADER   As String = "GroupRowCount"
Private Const DEF_GSUM_HEADER     As String = "GroupAmount"
Private Const DEF_MEASURE_HEADER  As String = "GroupAmount"
Private Const DEF_FLAG_HEADER     As String = "RecordID"
Private Const DEF_KEYWORDS        As String = "SAMPLE_A,SAMPLE_B"
Private Const DEF_NOTE            As String = "REVIEW_REQUIRED"
Private Const DEF_SEPARATOR       As String = "-"
Private Const DEF_COMPARISON      As String = ">="
Private Const DEF_LOWER_BOUND     As Double = 1000
Private Const DEF_UPPER_BOUND     As Double = 2000
Private Const DEF_RULE_BOUND      As Double = 1000

Private mCalcMode As Long
Private mFastOn As Boolean

' ---------------------------------------------------------------------------
'  Configuration helpers
' ---------------------------------------------------------------------------

Private Function Setting(ByVal key As String, ByVal fallback As String) As String
    Dim ws As Worksheet
    Dim r As Range
    Dim v As Variant

    On Error Resume Next
    Set ws = Worksheets(SETTINGS_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Setting = fallback
        Exit Function
    End If

    Set r = ws.Columns(1).Find(What:=key, LookAt:=xlWhole, LookIn:=xlValues)
    If r Is Nothing Then
        Setting = fallback
        Exit Function
    End If

    v = r.Offset(0, 1).Value
    If IsError(v) Or IsEmpty(v) Then
        Setting = fallback
    Else
        Setting = CStr(v)
    End If
End Function

Private Function SettingNumber(ByVal key As String, ByVal fallback As Double) As Double
    Dim s As String
    s = Trim(Setting(key, ""))
    If Len(s) > 0 And IsNumeric(s) Then
        SettingNumber = CDbl(s)
    Else
        SettingNumber = fallback
    End If
End Function

Private Function ColumnOf(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim r As Range
    Set r = ws.Rows(1).Find(What:=headerText, LookAt:=xlWhole, LookIn:=xlValues)
    If r Is Nothing Then
        ColumnOf = 0
    Else
        ColumnOf = r.Column
    End If
End Function

Private Function ColumnLetter(ByVal ws As Worksheet, ByVal col As Long) As String
    ColumnLetter = Split(ws.Cells(1, col).Address(True, False), "$")(0)
End Function

Private Function LastRowOf(ByVal ws As Worksheet, ByVal col As Long) As Long
    LastRowOf = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Private Function RequireColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim col As Long
    col = ColumnOf(ws, headerText)
    If col = 0 Then
        Err.Raise vbObjectError + 1000, , "Required column not found: " & headerText
    End If
    RequireColumn = col
End Function

Private Sub FastMode(ByVal enable As Boolean)
    If enable Then
        If mFastOn Then Exit Sub
        mCalcMode = Application.Calculation
        mFastOn = True
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    Else
        If Not mFastOn Then Exit Sub
        mFastOn = False
        Application.ScreenUpdating = True
        Application.Calculation = mCalcMode
        Application.EnableEvents = True
    End If
End Sub

Private Function SettingColor(ByVal key As String, _
                               ByVal defR As Long, ByVal defG As Long, _
                               ByVal defB As Long) As Long
    Dim s As String, parts As Variant
    s = Trim(Setting(key, ""))
    If Len(s) > 0 Then
        parts = Split(s, ",")
        If UBound(parts) = 2 Then
            If IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)) Then
                SettingColor = RGB(CLng(parts(0)), CLng(parts(1)), CLng(parts(2)))
                Exit Function
            End If
        End If
    End If
    SettingColor = RGB(defR, defG, defB)
End Function

Private Function KeywordList() As Variant
    KeywordList = Split(Setting("Keywords", DEF_KEYWORDS), ",")
End Function

Private Function HasKeyword(ByVal text As String, ByRef words As Variant) As Boolean
    Dim kw As Variant
    For Each kw In words
        If Len(Trim(CStr(kw))) > 0 Then
            If InStr(1, text, Trim(CStr(kw)), vbTextCompare) > 0 Then
                HasKeyword = True
                Exit Function
            End If
        End If
    Next kw
    HasKeyword = False
End Function

Private Function MeetsCondition(ByVal v As Variant, ByVal bound As Double, _
                                ByVal comparison As String) As Boolean
    If Not IsNumeric(v) Or IsEmpty(v) Then
        MeetsCondition = False
        Exit Function
    End If
    Select Case UCase(Trim(comparison))
        Case ">=", "GE", "AT_LEAST": MeetsCondition = (CDbl(v) >= bound)
        Case ">", "GT":              MeetsCondition = (CDbl(v) > bound)
        Case "<", "LT", "BELOW":     MeetsCondition = (CDbl(v) < bound)
        Case "<=", "LE":             MeetsCondition = (CDbl(v) <= bound)
        Case "=", "EQ":              MeetsCondition = (CDbl(v) = bound)
        Case Else:                   MeetsCondition = (CDbl(v) >= bound)
    End Select
End Function

' ---------------------------------------------------------------------------
'  1. Group rows and fill missing detail keys
' ---------------------------------------------------------------------------
'  Rows sharing the same group key are expected to be contiguous. The first
'  row of each group carries the expected item count. When the number of blank
'  detail keys equals that count, they are numbered sequentially; otherwise
'  they are filled with a review note and shaded.
' ---------------------------------------------------------------------------
Public Sub NormalizeDetailKeys()
    Dim ws As Worksheet
    Dim groupCol As Long, detailCol As Long, countCol As Long
    Dim lastRow As Long, i As Long, j As Long, groupEnd As Long
    Dim groupKey As String, expectedCount As Long, blankCount As Long
    Dim firstRow As Long, seq As Long
    Dim separator As String, note As String

    On Error GoTo EH
    Set ws = ActiveSheet
    groupCol = RequireColumn(ws, Setting("GroupHeader", DEF_GROUP_HEADER))
    detailCol = RequireColumn(ws, Setting("DetailHeader", DEF_DETAIL_HEADER))
    countCol = RequireColumn(ws, Setting("ItemCountHeader", DEF_COUNT_HEADER))
    separator = Setting("KeySeparator", DEF_SEPARATOR)
    note = Setting("ReviewNote", DEF_NOTE)

    lastRow = LastRowOf(ws, groupCol)
    If lastRow < 2 Then Exit Sub

    FastMode True
    i = 2
    Do While i <= lastRow
        groupKey = CStr(ws.Cells(i, groupCol).Value)
        firstRow = i
        blankCount = 0
        expectedCount = 0

        j = i
        Do While j <= lastRow And CStr(ws.Cells(j, groupCol).Value) = groupKey
            If Len(Trim(CStr(ws.Cells(j, detailCol).Value))) = 0 Then
                blankCount = blankCount + 1
            End If
            If j = firstRow And IsNumeric(ws.Cells(j, countCol).Value) Then
                expectedCount = CLng(ws.Cells(j, countCol).Value)
            End If
            j = j + 1
        Loop
        groupEnd = j - 1

        If blankCount > 0 Then
            If expectedCount = blankCount Then
                seq = 1
                For j = firstRow To groupEnd
                    If Len(Trim(CStr(ws.Cells(j, detailCol).Value))) = 0 Then
                        ws.Cells(j, detailCol).Value = groupKey & separator & CStr(seq)
                        seq = seq + 1
                    End If
                Next j
            Else
                For j = firstRow To groupEnd
                    If Len(Trim(CStr(ws.Cells(j, detailCol).Value))) = 0 Then
                        With ws.Cells(j, detailCol)
                            .Value = note
                            .Interior.Color = SettingColor("ColorReview", 255, 235, 156)
                        End With
                    End If
                Next j
            End If
        End If

        i = groupEnd + 1
    Loop

    FastMode False
    MsgBox "Detail keys normalized.", vbInformation
    Exit Sub
EH:
    FastMode False
    MsgBox "Step failed: " & Err.Description, vbExclamation
End Sub

' ---------------------------------------------------------------------------
'  2. Append derived columns
' ---------------------------------------------------------------------------
'  Three computed columns are appended at the right edge of the table:
'    line amount   = quantity * unit price
'    group count   = COUNTIF(basis column, current basis key)
'    group amount  = SUMIF(basis column, current basis key, line amount)
'  The basis column is configurable. Point GroupByHeader at the group key
'  column to obtain true per-group aggregation; leaving it on the detail key
'  column reproduces per-row values after detail keys have been numbered.
' ---------------------------------------------------------------------------
Public Sub BuildDerivedColumns()
    Dim ws As Worksheet
    Dim qtyCol As Long, priceCol As Long, detailCol As Long, basisCol As Long
    Dim amountCol As Long, gCountCol As Long, gSumCol As Long
    Dim lastRow As Long, firstNewCol As Long
    Dim qL As String, pL As String, aL As String, bL As String

    On Error GoTo EH
    Set ws = ActiveSheet
    qtyCol = RequireColumn(ws, Setting("QuantityHeader", DEF_QTY_HEADER))
    priceCol = RequireColumn(ws, Setting("UnitPriceHeader", DEF_PRICE_HEADER))
    detailCol = RequireColumn(ws, Setting("DetailHeader", DEF_DETAIL_HEADER))
    basisCol = RequireColumn(ws, Setting("GroupByHeader", DEF_DETAIL_HEADER))

    firstNewCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    amountCol = firstNewCol
    gCountCol = firstNewCol + 1
    gSumCol = firstNewCol + 2

    ws.Cells(1, amountCol).Value = Setting("LineAmountHeader", DEF_AMOUNT_HEADER)
    ws.Cells(1, gCountCol).Value = Setting("GroupCountHeader", DEF_GCOUNT_HEADER)
    ws.Cells(1, gSumCol).Value = Setting("GroupSumHeader", DEF_GSUM_HEADER)

    lastRow = LastRowOf(ws, detailCol)
    If lastRow < 2 Then Exit Sub

    qL = ColumnLetter(ws, qtyCol)
    pL = ColumnLetter(ws, priceCol)
    aL = ColumnLetter(ws, amountCol)
    bL = ColumnLetter(ws, basisCol)

    FastMode True
    With ws.Range(ws.Cells(2, amountCol), ws.Cells(lastRow, amountCol))
        .Formula = "=" & qL & "2*" & pL & "2"
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlRight
    End With

    With ws.Range(ws.Cells(2, gCountCol), ws.Cells(lastRow, gCountCol))
        .Formula = "=COUNTIF(" & bL & ":" & bL & "," & bL & "2)"
        .NumberFormat = "0"
        .HorizontalAlignment = xlRight
    End With

    With ws.Range(ws.Cells(2, gSumCol), ws.Cells(lastRow, gSumCol))
        .Formula = "=SUMIF(" & bL & ":" & bL & "," & bL & "2," & aL & ":" & aL & ")"
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlRight
    End With

    With ws.Range(ws.Cells(1, gCountCol), ws.Cells(lastRow, gSumCol))
        .Interior.Color = SettingColor("ColorDerived", 226, 235, 248)
    End With
    ws.Range(ws.Cells(1, amountCol), ws.Cells(1, gSumCol)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(1, amountCol), ws.Cells(1, gSumCol)).Font.Bold = True

    FastMode False
    MsgBox "Derived columns added.", vbInformation
    Exit Sub
EH:
    FastMode False
    MsgBox "Step failed: " & Err.Description, vbExclamation
End Sub

' ---------------------------------------------------------------------------
'  3. Shade values that fall inside two numeric bands
' ---------------------------------------------------------------------------
Public Sub ApplyBandHighlight()
    Dim ws As Worksheet
    Dim measureCol As Long, lastRow As Long
    Dim lowerBound As Double, upperBound As Double
    Dim cell As Range, v As Variant

    On Error GoTo EH
    Set ws = ActiveSheet
    measureCol = RequireColumn(ws, Setting("MeasureHeader", DEF_MEASURE_HEADER))
    lowerBound = SettingNumber("LowerBound", DEF_LOWER_BOUND)
    upperBound = SettingNumber("UpperBound", DEF_UPPER_BOUND)
    lastRow = LastRowOf(ws, measureCol)
    If lastRow < 2 Then Exit Sub

    FastMode True
    ws.Range(ws.Cells(2, measureCol), ws.Cells(lastRow, measureCol)).Interior.Pattern = xlNone

    For Each cell In ws.Range(ws.Cells(2, measureCol), ws.Cells(lastRow, measureCol))
        v = cell.Value
        If IsNumeric(v) And Not IsEmpty(v) Then
            If CDbl(v) >= lowerBound And CDbl(v) <= upperBound Then
                With cell.Interior
                    .Color = SettingColor("ColorBandLow", 198, 224, 198)
                    .Pattern = xlSolid
                End With
            ElseIf CDbl(v) > upperBound Then
                With cell.Interior
                    .Color = SettingColor("ColorBandHigh", 255, 244, 204)
                    .Pattern = xlSolid
                End With
            End If
        End If
    Next cell

    FastMode False
    MsgBox "Numeric bands highlighted.", vbInformation
    Exit Sub
EH:
    FastMode False
    MsgBox "Step failed: " & Err.Description, vbExclamation
End Sub

' ---------------------------------------------------------------------------
'  4. Flag whole groups matching a keyword set AND a numeric condition
' ---------------------------------------------------------------------------
'  A group qualifies when at least one of its rows contains a configured
'  keyword in the description column and the measured value satisfies the
'  configured comparison. All rows of a qualifying group are then shaded.
'  The comparison direction is configured, so "at or above" and "below"
'  screens share one routine instead of being duplicated.
' ---------------------------------------------------------------------------
Public Sub FlagGroupsByRule()
    Dim ws As Worksheet
    Dim groupCol As Long, descCol As Long, measureCol As Long, flagCol As Long
    Dim lastRow As Long, i As Long
    Dim groupKey As String, text As String, v As Variant
    Dim bound As Double, comparison As String
    Dim words As Variant
    Dim qualified As Object
    Dim flagged As Long

    On Error GoTo EH
    Set ws = ActiveSheet
    groupCol = RequireColumn(ws, Setting("GroupHeader", DEF_GROUP_HEADER))
    descCol = RequireColumn(ws, Setting("DescriptionHeader", DEF_DESC_HEADER))
    measureCol = RequireColumn(ws, Setting("MeasureHeader", DEF_MEASURE_HEADER))
    flagCol = RequireColumn(ws, Setting("FlagTargetHeader", DEF_FLAG_HEADER))
    bound = SettingNumber("RuleBound", DEF_RULE_BOUND)
    comparison = UCase(Trim(Setting("Comparison", DEF_COMPARISON)))
    words = KeywordList()
    Set qualified = CreateObject("Scripting.Dictionary")

    lastRow = LastRowOf(ws, groupCol)
    If lastRow < 2 Then Exit Sub

    FastMode True
    ws.Range(ws.Cells(2, flagCol), ws.Cells(lastRow, flagCol)).Interior.Pattern = xlNone

    For i = 2 To lastRow
        groupKey = CStr(ws.Cells(i, groupCol).Value)
        If Not qualified.Exists(groupKey) Then
            text = CStr(ws.Cells(i, descCol).Value)
            v = ws.Cells(i, measureCol).Value
            If HasKeyword(text, words) And _
               MeetsCondition(v, bound, comparison) Then
                qualified.Add groupKey, 1
            End If
        End If
    Next i

    flagged = 0
    For i = 2 To lastRow
        If qualified.Exists(CStr(ws.Cells(i, groupCol).Value)) Then
            ws.Cells(i, flagCol).Interior.Color = SettingColor("ColorFlag", 144, 200, 160)
            flagged = flagged + 1
        End If
    Next i

    FastMode False
    VerifyGroupFlags
    MsgBox qualified.Count & " group(s) matched; " & flagged & " row(s) flagged.", vbInformation
    Exit Sub
EH:
    FastMode False
    MsgBox "Step failed: " & Err.Description, vbExclamation
End Sub

' ---------------------------------------------------------------------------
'  5. Re-verify flags and report false positives
' ---------------------------------------------------------------------------
'  Every flagged row is checked against the configured rule. A flagged group
'  is valid when at least one of its rows still satisfies both conditions.
' ---------------------------------------------------------------------------
Public Sub VerifyGroupFlags()
    Dim ws As Worksheet
    Dim groupCol As Long, descCol As Long, measureCol As Long, flagCol As Long
    Dim lastRow As Long, i As Long, j As Long
    Dim groupKey As String, text As String, v As Variant
    Dim bound As Double, comparison As String
    Dim words As Variant
    Dim isMatch As Boolean, falsePositives As Long, checked As Long
    Dim marker As Long

    On Error GoTo EH
    Set ws = ActiveSheet
    groupCol = ColumnOf(ws, Setting("GroupHeader", DEF_GROUP_HEADER))
    descCol = ColumnOf(ws, Setting("DescriptionHeader", DEF_DESC_HEADER))
    measureCol = ColumnOf(ws, Setting("MeasureHeader", DEF_MEASURE_HEADER))
    flagCol = ColumnOf(ws, Setting("FlagTargetHeader", DEF_FLAG_HEADER))
    If groupCol = 0 Or descCol = 0 Or measureCol = 0 Or flagCol = 0 Then Exit Sub

    bound = SettingNumber("RuleBound", DEF_RULE_BOUND)
    comparison = UCase(Trim(Setting("Comparison", DEF_COMPARISON)))
    words = KeywordList()
    marker = SettingColor("ColorFlag", 144, 200, 160)
    lastRow = LastRowOf(ws, groupCol)
    falsePositives = 0
    checked = 0

    For i = 2 To lastRow
        If ws.Cells(i, flagCol).Interior.Color = marker Then
            checked = checked + 1
            groupKey = CStr(ws.Cells(i, groupCol).Value)
            isMatch = False

            For j = 2 To lastRow
                If CStr(ws.Cells(j, groupCol).Value) = groupKey Then
                    text = CStr(ws.Cells(j, descCol).Value)
                    v = ws.Cells(j, measureCol).Value
                    If HasKeyword(text, words) And _
                       MeetsCondition(v, bound, comparison) Then
                        isMatch = True
                        Exit For
                    End If
                End If
            Next j

            If Not isMatch Then falsePositives = falsePositives + 1
        End If
    Next i

    If checked > 0 Then
        Debug.Print "Verified " & checked & " flagged row(s); " & _
                    falsePositives & " false positive(s)."
        If falsePositives > 0 Then
            MsgBox falsePositives & " flagged row(s) did not satisfy the rule.", vbExclamation
        End If
    End If
    Exit Sub
EH:
    MsgBox "Verification failed: " & Err.Description, vbExclamation
End Sub

' ---------------------------------------------------------------------------
'  Convenience entry point
' ---------------------------------------------------------------------------
Public Sub RunCleaningPipeline()
    NormalizeDetailKeys
    BuildDerivedColumns
    ApplyBandHighlight
    FlagGroupsByRule
    MsgBox "Cleaning pipeline finished.", vbInformation
End Sub
