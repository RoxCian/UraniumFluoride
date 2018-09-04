Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
#Region "Imports Macros"
Imports ExcelRange = System.Object
Imports ExcelNumber = System.Object
Imports ExcelLogical = System.Object
Imports ExcelDate = System.Object
Imports ExcelString = System.Object
Imports ExcelVariant = System.Object
#End Region

Public Module UtilityFunctions
#Region "ExcelError"
    Public ReadOnly ExcelErrorDiv0 As ExcelError = ExcelError.ExcelErrorDiv0
    Public ReadOnly ExcelErrorNa As ExcelError = ExcelError.ExcelErrorNA
    Public ReadOnly ExcelErrorName As ExcelError = ExcelError.ExcelErrorName
    Public ReadOnly ExcelErrorNull As ExcelError = ExcelError.ExcelErrorNull
    Public ReadOnly ExcelErrorNum As ExcelError = ExcelError.ExcelErrorNum
    Public ReadOnly ExcelErrorRef As ExcelError = ExcelError.ExcelErrorRef
    Public ReadOnly ExcelErrorValue As ExcelError = ExcelError.ExcelErrorValue
#End Region

    Public ReadOnly Property Application As Excel.Application
        Get
            Return ExcelDnaUtil.Application
        End Get
    End Property

    <ExcelFunction(Description:="Banker round")>
    Public Function BankerRound(<MarshalAs(UnmanagedType.Currency)> num As Decimal, pre As Integer, Optional isSignificant As Boolean = False) As ExcelNumber
        If num = 0 Then Return 0
        If isSignificant Then
            Dim power As Integer = Math.Floor(Math.Log10(Math.Abs(num)))
            Return Math.Round(Math.Round(num / 10 ^ power, 14), pre - 1, MidpointRounding.ToEven) * 10 ^ power
        Else
            Return Math.Round(Math.Round(num, 14), pre, MidpointRounding.ToEven)
        End If
    End Function
    <ExcelFunction(IsMacroType:=True)>
    Public Function AverageByMean(<ExcelArgument(AllowReference:=True)> nums As ExcelRange, Optional ratio As Double = 0.1) As ExcelNumber
        Dim nums_Range As Excel.Range = ConvertToRange(nums)
        Try
            If Count(nums_Range) < 3 Then Return Average(nums_Range)
            Dim f As Boolean = False
            Dim ave = Average(nums_Range)
            If ave = 0 Then Return 0
            Dim i, j As Integer
            For i = 1 To nums_Range.Rows.Count
                For j = 1 To nums_Range.Columns.Count
                    If nums_Range(i, j).Value <> vbEmpty And Math.Abs((nums_Range(i, j).Value - ave) / ave) >= ratio Then
                        f = True
                        Exit For
                    End If
                Next
                If f Then Exit For
            Next
            If f Then
                Dim ave2, sum2 As Decimal, len2 As Integer
                For k = 1 To nums_Range.Rows.Count
                    For l = 1 To nums_Range.Columns.Count
                        If nums_Range(k, l).Value <> vbEmpty And Not (i = k And j = l) Then
                            sum2 += nums_Range(k, l).Value
                            len2 += 1
                        End If
                    Next
                Next
                If sum2 = 0 Then Return 0
                ave2 = sum2 / len2
                For k = 1 To nums_Range.Rows.Count
                    For l = 1 To nums_Range.Columns.Count
                        If Not (i = k And j = l) And nums_Range(k, l).value <> vbEmpty And Math.Abs((nums_Range(k, l).Value - ave) / ave) >= ratio Then Return ExcelErrorValue
                    Next
                Next
                Return ave2
            End If
            Return ave
        Finally
        End Try
    End Function
    <ExcelFunction(IsMacroType:=True)>
    Public Function VerifyByMean(<ExcelArgument(AllowReference:=True)> nums As ExcelRange, Optional ratio As Double = 0.1) As ExcelNumber
        Dim nums_Range As Excel.Range = ConvertToRange(nums)
        Try
            If Count(nums_Range) < 3 Then Return 1
            Dim f As Boolean = False
            Dim ave = Application.WorksheetFunction.Average(nums_Range)
            If ave = 0 Then Return -65536
            Dim i, j As Integer
            For i = 1 To nums_Range.Rows.Count
                For j = 1 To nums_Range.Columns.Count
                    If nums_Range(i, j).Value <> vbEmpty And Math.Abs((nums_Range(i, j).Value - ave) / ave) >= ratio Then
                        f = True
                        Exit For
                    End If
                Next
                If f Then Exit For
            Next
            If f Then
                Dim ave2, sum2 As Decimal, len2 As Integer
                For k = 1 To nums_Range.Rows.Count
                    For l = 1 To nums_Range.Columns.Count
                        If nums_Range(k, l).Value <> vbEmpty And Not (i = k And j = l) Then
                            sum2 += nums_Range(k, l).Value
                            len2 += 1
                        End If
                    Next
                Next
                If sum2 = 0 Then Return -65536
                ave2 = sum2 / len2
                For k = 1 To nums_Range.Rows.Count
                    For l = 1 To nums_Range.Columns.Count
                        If Not (i = k And j = l) And nums_Range(k, l).Value <> vbEmpty Then If Math.Abs(nums_Range(k, l).Value - ave2) / ave2 >= ratio Then Return -1
                    Next
                Next
                Return 0
            End If
            Return 1
        Finally
        End Try
    End Function
    <ExcelFunction(IsMacroType:=True)>
    Public Function AverageByMedian(<ExcelArgument(AllowReference:=True)> nums As ExcelRange, Optional ratio As Double = 0.15) As ExcelNumber
        Dim nums_Range As Excel.Range = ConvertToRange(nums)
        Try
            If Count(nums_Range) <= 2 Then Return Average(nums_Range)
            Dim min As Decimal = UtilityFunctions.Min(nums_Range)
            Dim max As Decimal = UtilityFunctions.Max(nums_Range)
            Dim med As Decimal = UtilityFunctions.Med(nums_Range)
            If med = 0 Then Return ExcelErrorDiv0
            If Math.Abs((max - med) / med) >= CDec(ratio) And Math.Abs((med - min) / med) >= CDec(ratio) Then Return ExcelErrorValue
            If Math.Abs((max - med) / med) >= CDec(ratio) Or Math.Abs((med - min) / med) >= CDec(ratio) Then Return med
            Return Average(nums_Range)
        Finally
        End Try
    End Function
    <ExcelFunction(IsMacroType:=True)>
    Public Function VerifyByMedian(<ExcelArgument(AllowReference:=True)> nums As ExcelRange, Optional ratio As Double = 0.15) As ExcelNumber
        Dim nums_Range As Excel.Range = ConvertToRange(nums)
        Try
            If Count(nums_Range) <= 2 Then Return 1
            Dim min As Decimal = UtilityFunctions.Min(nums_Range)
            Dim max As Decimal = UtilityFunctions.Max(nums_Range)
            Dim med As Decimal = UtilityFunctions.Med(nums_Range)
            If med = 0 Then Return -1
            If Math.Abs((max - med) / med) >= CDec(ratio) And Math.Abs((med - min) / med) >= CDec(ratio) Then Return -1
            If Math.Abs((max - med) / med) >= CDec(ratio) Or Math.Abs((med - min) / med) >= CDec(ratio) Then Return 0
            Return 1
        Finally
        End Try
    End Function
    <ExcelFunction(IsVolatile:=True)>
    Public Function RandNoRepeat(bottom As Integer, top As Integer, Optional memorySet As Integer = 0, Optional memories As Integer = 30, Optional unrepeatPossibility As Integer = 0.95) As ExcelNumber
        Static memory As New Helper.ValueCircularListCollection(1024, 1024, Nothing)
        Static randomer As Random
        randomer = New Random(Now.Millisecond)
        If top - bottom < 2 Then Return bottom
        If memories < 2 Then memories = 2
        If memories > memory.ListCapacity Then memories = memory.ListCapacity
        If unrepeatPossibility > 1 Then unrepeatPossibility = 1
        If top - bottom < memories Then memories = top - bottom - 1
        If memorySet > memory.ListCapacity - 1 Or memorySet < 0 Then memorySet = 0

        Dim result As Integer
        Dim f As Boolean = False
        Do Until f
            f = True
            result = randomer.Next(bottom, top)
            For i = -memories + 1 To 0
                If memory(memorySet)(i) IsNot Nothing AndAlso memory(memorySet)(i) = result Then If randomer.Next(0, 1000) / 1000 < unrepeatPossibility Then f = False
            Next
        Loop
        memory(memorySet).MoveNext(result)
        Return result
    End Function
    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function PageLocalize(<ExcelArgument(AllowReference:=True)> r As ExcelRange, pageRowsCount As Integer, pageColumnsCount As Integer, locationRow As Integer, locationColumn As Integer, pageIndex As Integer) As ExcelVariant
        Dim _Range As Excel.Range = ConvertToRange(r)
        If _Range Is Nothing Then Return ExcelErrorNa
        If (locationRow > pageRowsCount Or locationRow < 1 Or locationColumn > pageColumnsCount Or locationColumn < 1) Or
        (pageRowsCount > _Range.Rows.Count Or pageRowsCount < 1 Or pageColumnsCount > _Range.Columns.Count Or pageColumnsCount < 1) Then _
            Return Nothing
        Dim pageCount, pageCountInRow, pageCountInColumn As Integer
        pageCountInRow = _Range.Rows.Count \ pageRowsCount
        pageCountInColumn = _Range.Columns.Count \ pageColumnsCount
        pageCount = pageCountInRow * pageCountInColumn
        If pageIndex > pageCount Or pageIndex < 1 Then Return Nothing
        Dim pageIndexInRow, pageIndexInColumn As Integer
        pageIndex -= 1
        pageIndexInRow = pageIndex \ pageCountInColumn
        pageIndexInColumn = pageIndex Mod pageCountInColumn
        Dim row, column As Integer
        row = pageRowsCount * pageIndexInRow + locationRow
        column = pageColumnsCount * pageIndexInRow + locationColumn
        Return _Range(row, column).Value
    End Function
    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function PageLocalize2(<ExcelArgument(AllowReference:=True)> r As ExcelRange, <ExcelArgument(AllowReference:=True)> rPage As ExcelRange, <ExcelArgument(AllowReference:=True)> rCell As ExcelRange, pageIndex As Integer) As ExcelVariant
        Dim page_Range As Excel.Range = ConvertToRange(rPage), cell_Range As Excel.Range = ConvertToRange(rCell)
        Dim cellRow, cellColumn, pageRow, pageColumn As Integer
        pageRow = page_Range.Rows.Count
        pageColumn = page_Range.Columns.Count
        cellRow = cell_Range.Row
        cellColumn = cell_Range.Column
        Do Until cellRow <= pageRow
            cellRow -= pageRow
        Loop
        Do Until cellColumn <= pageColumn
            cellColumn -= pageColumn
        Loop
        Return PageLocalize(r, pageRow, pageColumn, cellRow, cellColumn, pageIndex)
    End Function
    '<ExcelFunction>
    'Public  Function PageLocalize3(<ExcelArgument(AllowReference:=True)> rCell As ExcelRange, pageIndex As Integer) As ExcelVariant
    '    Dim r As ExcelRange = rCell.Worksheet.UsedRange
    '    Dim cellRow, cellColumn, pageRow, pageColumn As Integer

    'End Function
    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function PageLocalizeByKeyword(<ExcelArgument(AllowReference:=True)> r As ExcelRange, <ExcelArgument(AllowReference:=True)> rPage As ExcelRange, <ExcelArgument(AllowReference:=True)> rCell As ExcelRange, keyword As Object, Optional isValueMatching As Boolean = False) As ExcelVariant
        Dim _Range As Excel.Range = ConvertToRange(r), page_Range As Excel.Range = ConvertToRange(rCell), cell_Range As Excel.Range = ConvertToRange(rCell)
        Dim locationRow = FindRow(_Range, keyword, isValueMatching)
        Dim locationColumn = FindColumn(_Range, keyword, isValueMatching)
        If locationRow = 0 Then Return ExcelErrorNa
        Dim pageIndex As Integer = Math.Ceiling(_Range.Columns.Count / page_Range.Columns.Count) * (locationRow \ page_Range.Rows.Count) + Math.Ceiling(locationColumn \ page_Range.Columns.Count)
        Return PageLocalize2(r, rPage, rCell, pageIndex)
    End Function
    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function FindRow(<ExcelArgument(AllowReference:=True)> r As ExcelRange, keyword As Object, Optional isValueMatching As Boolean = False) As Integer
        Dim _Range As Excel.Range = ConvertToRange(r)
        For i = 1 To _Range.Rows.Count
            For j = 1 To _Range.Columns.Count
                If keyword IsNot Nothing AndAlso keyword <> vbEmpty AndAlso If(isValueMatching, _Range(i, j).Value, _Range(i, j).Text) = keyword Then Return i
            Next
        Next
        Return 0
    End Function
    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function FindColumn(<ExcelArgument(AllowReference:=True)> r As ExcelRange, keyword As Object, Optional isValueMatching As Boolean = False) As Integer
        Dim _Range As Excel.Range = ConvertToRange(r)
        For i = 1 To _Range.Rows.Count
            For j = 1 To _Range.Columns.Count
                If keyword IsNot Nothing AndAlso keyword <> vbEmpty AndAlso If(isValueMatching, _Range(i, j).Value, _Range(i, j).Text) = keyword Then Return j
            Next
        Next
        Return 0
    End Function
    <ExcelFunction(IsMacroType:=True)>
    Public Function RangeToString(<ExcelArgument(AllowReference:=True)> r As ExcelRange) As String
        Dim _Range As Excel.Range = ConvertToRange(r)
        _Range(1, 1).Calculate
        Return _Range(1, 1).Text
    End Function
    <ExcelFunction(IsMacroType:=True)>
    Public Function MergedCellRows(<ExcelArgument(AllowReference:=True)> cell As ExcelRange)
        Return ConvertToRange(cell)(1, 1).MergeArea.Rows.count
    End Function
    <ExcelFunction(IsMacroType:=True)>
    Public Function MergedCellColumns(<ExcelArgument(AllowReference:=True)> cell As ExcelRange)
        Return ConvertToRange(cell)(1, 1).MergeArea.Columns.count
    End Function
    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function DataFitter(formula As String, <MarshalAs(UnmanagedType.Currency)> formulaResult As Decimal, variantIndexToReturn As Integer, <ExcelArgument(AllowReference:=True)> rMinValues As ExcelRange, <ExcelArgument(AllowReference:=True)> rMaxValues As ExcelRange, <ExcelArgument(AllowReference:=True)> rSteps As ExcelRange, Optional isForceRecalculating As Boolean = False) As ExcelNumber
        Static ResultCollection As New Dictionary(Of Integer, ExcelNumber())

        Dim minValues_Range As Excel.Range = ConvertToRange(rMinValues), maxValues_Range As Excel.Range = ConvertToRange(rMaxValues), steps_Range As Excel.Range = ConvertToRange(rSteps)
        Dim minValues() As Decimal = GetNumeric(minValues_Range)
        Dim maxValues() As Decimal = GetNumeric(maxValues_Range)
        Dim steps() As Decimal = GetNumeric(steps_Range)
        Dim valuesCount As Integer = Min(minValues.Count, maxValues.Count, steps.Count)
        If variantIndexToReturn > valuesCount Then Return ExcelErrorNa
        Dim fittedValues(valuesCount - 1) As ExcelNumber
        For i = 0 To valuesCount - 1
            fittedValues(i) = minValues(i)
        Next

        For i = 0 To valuesCount - 1
            If steps(i) < 0 Or maxValues(i) - minValues(i) < steps(i) Then Return ExcelErrorValue
        Next

        Dim info As New Text.StringBuilder
        info.Append("The ")
        info.Append(variantIndexToReturn)
        info.Append(Switch(variantIndexToReturn - (variantIndexToReturn \ 10) * 10, 1, "st", 2, "nd", 3, "rd", "th"))
        info.Append(" variant in the formula {")
        info.Append(formula)
        info.Append(";")
        For i = 0 To valuesCount - 1
            info.Append("$$")
            info.Append(i + 1)
            info.Append("=")
            info.Append(minValues(i))
            info.Append("-")
            info.Append(maxValues(i))
            info.Append("|")
            info.Append(steps(i))
            If i <> valuesCount - 1 Then info.Append(",")
        Next
        info.Append("}.")


        Dim argumentsHash As Integer = formula.GetHashCode Xor formulaResult.GetHashCode Xor GetNumericArrayHash(minValues) Xor GetNumericArrayHash(maxValues) Xor GetNumericArrayHash(steps)
        If ResultCollection.ContainsKey(argumentsHash) And Not isForceRecalculating Then Return ResultCollection(argumentsHash)(variantIndexToReturn - 1)

        Dim w As New WattingWindow(info.ToString)
        w.Show()
        Dim f As Boolean
        f = True

        w.Dispatcher.BeginInvoke(Sub()
                                     Dim calculationCount As ULong = 0
                                     Do While f
                                         Dim formulaForExecute As New Text.StringBuilder(formula, formula.Length * 2)
                                         For j = 0 To valuesCount - 1
                                             formulaForExecute = formulaForExecute.Replace("$$" & j + 1, fittedValues(j))
                                         Next
                                         If formulaResult = Application.Evaluate(formulaForExecute.ToString) Then Exit Do
                                         calculationCount += 1
                                         If calculationCount - (calculationCount \ 50000) * 50000 = 0 Then
                                             If MsgBox("We have calculated " & calculationCount & " times. Do you want to continue?", MsgBoxStyle.OkCancel Xor MsgBoxStyle.Question, "It tooks a long time...") = MsgBoxResult.Cancel Then
                                                 f = False
                                                 Exit Do
                                             End If
                                         End If
                                         For i = 0 To valuesCount - 1
                                             fittedValues(i) = fittedValues(i) + steps(i)
                                             If fittedValues.Last > maxValues(valuesCount - 1) Then
                                                 f = False
                                                 Exit For
                                             End If
                                             If fittedValues(i) <= maxValues(i) Then Exit For Else fittedValues(i) = minValues(i)
                                         Next
                                     Loop
                                     If Not f Then
                                         For i = 0 To fittedValues.Count - 1
                                             fittedValues(i) = ExcelErrorNa
                                         Next
                                     End If
                                 End Sub).Wait()
        ResultCollection.Add(argumentsHash, fittedValues)
        w.Close()
        Return fittedValues(variantIndexToReturn - 1)
    End Function
    <ExcelFunction>
    Public Function RegExFind(text As String, pattern As String, Optional index As Integer = 1, Optional isCaseIgnore As Boolean = True) As ExcelNumber
        Dim e As New Text.RegularExpressions.Regex(pattern, If(isCaseIgnore, System.Text.RegularExpressions.RegexOptions.IgnoreCase, System.Text.RegularExpressions.RegexOptions.None))
        Dim m As System.Text.RegularExpressions.MatchCollection = e.Matches(text)
        If m.Count < index Then Return -1 Else Return m(index - 1).Index + 1
    End Function
    <ExcelFunction>
    Public Function RegExMatch(text As String, pattern As String, Optional index As Integer = 1, Optional isCaseIgnore As Boolean = True) As ExcelString
        Dim e As New Text.RegularExpressions.Regex(pattern, If(isCaseIgnore, System.Text.RegularExpressions.RegexOptions.IgnoreCase, System.Text.RegularExpressions.RegexOptions.None))
        Dim m As System.Text.RegularExpressions.MatchCollection = e.Matches(text)
        If m.Count < index Then Return ExcelErrorNull Else Return m(index).Value
    End Function
    <ExcelFunction(IsMacroType:=True)>
    Public Function RelativeReference(worksheetName As String, rangeText As String, Optional path As String = "") As ExcelRange
        Static memory As New Helper.CircularList(Of (Path As String, Workbook As Workbook))(1024)
        Dim wb As Workbook = Nothing
        If path = "" Then
            wb = Application.ThisWorkbook
        Else
            Dim f As Boolean = False
            For Each i In memory
                If i.Path = path Then
                    wb = i.Workbook
                    f = True
                    Exit For
                End If
            Next
            If Not f Then
                If IO.File.Exists(path) Then wb = Application.Workbooks.Open(path) Else Return ExcelErrorNa
                memory.MoveNext((path, wb))
            End If
        End If
        For Each i As Worksheet In wb.Worksheets
            If i.Name = worksheetName Then Return ConvertToExcelReference(i.Range(rangeText))
        Next
        Return ExcelErrorNa
    End Function
    <ExcelFunction>
    Public Function RegExMatchesCount(input As String, pattern As String, Optional startat As Integer = 0) As ExcelNumber
        Return New Text.RegularExpressions.Regex(pattern).Matches(input, startat).Count
    End Function
    <ExcelFunction(IsMacroType:=True)>
    Public Function VLookUpByRank(<ExcelArgument(AllowReference:=True)> r As ExcelRange, rank As Integer, rankColumn As Integer, lookupColumn As Integer) As ExcelVariant
        Dim _Range As Excel.Range = ConvertToRange(r)
        If _Range.Columns.Count < rankColumn Or _Range.Columns.Count < lookupColumn Or rankColumn < 1 Or lookupColumn < 1 Then Return ExcelErrorRef
        Dim ranktable As New Dictionary(Of Integer, ExcelVariant)
        For i = 1 To _Range.Rows.Count
            ranktable.Add(i, _Range(i, rankColumn).value)
        Next
        'Bad sorting implementation, will be rewrite.
        For i = 1 To _Range.Rows.Count
            For j = 1 To _Range.Rows.Count - 1
                If IsError(ranktable(j)) Or Application.WorksheetFunction.isblank(ranktable(j)) Then
                    Call Swap(ranktable.Values(j), ranktable.Values(j + 1))
                    Call Swap(ranktable.Keys(j), ranktable.Keys(j + 1))
                ElseIf IsError(ranktable(j + 1)) Then
                ElseIf ranktable(j) > ranktable(j + 1) Then
                    Call Swap(ranktable.Values(j), ranktable.Values(j + 1))
                    Call Swap(ranktable.Keys(j), ranktable.Keys(j + 1))
                End If
            Next
        Next
        Return _Range(ranktable.Keys(rank), lookupColumn).Value
    End Function
    <ExcelFunction(IsMacroType:=True)>
    Public Function HLookUpByRank(<ExcelArgument(AllowReference:=True)> r As ExcelRange, rank As Integer, rankRow As Integer, lookupRow As Integer) As ExcelVariant
        Dim _Range As Excel.Range = ConvertToRange(r)
        If _Range.Rows.Count < rankRow Or _Range.Rows.Count < lookupRow Or rankRow < 1 Or lookupRow < 1 Then Return ExcelErrorRef
        Dim ranktable As New Dictionary(Of Integer, ExcelVariant)
        For i = 1 To _Range.Columns.Count
            ranktable.Add(i, _Range(i, rankRow).value)
        Next
        'Bad sorting implementation, will be rewrite.
        For i = 1 To _Range.Rows.Count
            For j = 1 To _Range.Rows.Count - 1
                If IsError(ranktable(j)) Or Application.WorksheetFunction.isblank(ranktable(j)) Then
                    Call Swap(ranktable.Values(j), ranktable.Values(j + 1))
                    Call Swap(ranktable.Keys(j), ranktable.Keys(j + 1))
                ElseIf IsError(ranktable(j + 1)) Then
                ElseIf ranktable(j) > ranktable(j + 1) Then
                    Call Swap(ranktable.Values(j), ranktable.Values(j + 1))
                    Call Swap(ranktable.Keys(j), ranktable.Keys(j + 1))
                End If
            Next
        Next
        Return _Range(lookupRow, ranktable.Keys(rank)).Value
    End Function
    <ExcelFunction>
    Public Function Contents(array As ExcelVariant, searching As ExcelVariant) As ExcelLogical
        If IsArray(array) Then
            For Each i In array
                If searching Is array Then
                    Dim j As Integer = LBound(searching)
                    If i = searching(j) Then Return True
                Else
                    If i = searching Then Return True
                End If
            Next
        End If
        Return False
    End Function
    <ExcelFunction>
    Public Function Switch(expression As ExcelVariant, ParamArray args() As ExcelVariant) As ExcelVariant
        Dim i As Integer
        If Not IsArray(args) Then Return ExcelErrorNull
        For i = 0 To args.Count - 1 Step 2
            If expression = args(i) Then Return args(i + 1)
            On Error Resume Next
        Next
        If Application.IsOdd(args.Count) Then Return args(args.Count - 1)
        Return ExcelErrorNa
    End Function
    <ExcelFunction>
    Public Function MinIndex(ParamArray args() As ExcelVariant) As ExcelNumber
        If Not IsArray(args) Then Return 0
        Dim min As Double = Double.MinValue
        Dim result As Integer
        For i = 0 To args.Count - 1
            If IsNumeric(args(i)) Then
                If args(i) < min Then
                    min = args(i)
                    result = i
                End If
            End If
        Next
        Return result + 1
    End Function
    <ExcelFunction>
    Public Function MaxIndex(ParamArray args() As ExcelVariant) As ExcelNumber
        If Not IsArray(args) Then Return 0
        Dim max As Double = Double.MaxValue
        Dim result As Integer
        For i = 0 To args.Count - 1
            If IsNumeric(args(i)) Then
                If args(i) > max Then
                    max = args(i)
                    result = i
                End If
            End If
        Next
        Return result + 1
    End Function
    <ExcelFunction>
    Public Function Guid(Optional isCompress As Boolean = False) As ExcelString
        Dim result As String = System.Guid.NewGuid.ToString("N").ToUpper
        If isCompress Then
            Return result
        Else 'A lossy compress
            Dim result2 As New Text.StringBuilder
            For i = 0 To result.Count - 1 Step 2
                Dim left As Integer = If(Asc(result(i)) <= Asc("9"), Asc(result(i)) - Asc("0"), Asc(result(i)) - Asc("A") + 10)
                Dim right As Integer = If(Asc(result(i + 1)) <= Asc("9"), Asc(result(i + 1)) - Asc("0"), Asc(result(i + 1)) - Asc("A") + 10)
                Dim sum As Integer = left + right
                If sum <= 9 Then result2.Append(Chr(sum + Asc("0"))) Else result2.Append(Chr(sum - 10 + Asc("A")))
            Next
            Return result2.ToString
        End If
    End Function
    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function CopyTextbox(<ExcelArgument(AllowReference:=True)> r As ExcelRange, textBoxName As String, Optional removeAllRecordedTextBoxes As Boolean = False, Optional left As Double = 0, Optional top As Double = 0, Optional width As Double = 0, Optional height As Double = 0) As ExcelVariant
        Static attachedObjects As New Dictionary(Of String, String)
        Dim _Range As Excel.Range = ConvertToRange(r)
        Dim ws As Worksheet
        ws = _Range.Worksheet
        Dim r1st As Excel.Range
        r1st = _Range(1, 1)
        If removeAllRecordedTextBoxes Then
            For o = 0 To attachedObjects.Count - 1
                RemoveObject(attachedObjects.Item(o), ws.Name, True)
            Next
            attachedObjects.Clear()
        End If

        Do While attachedObjects.ContainsKey(r1st.Address)
            RemoveObject(attachedObjects(r1st.Address), ws.Name, True)
            attachedObjects.Remove(r1st.Address)
        Loop

        Dim t As Shape = Nothing
        Dim f As Boolean
        f = False
        For Each t In ws.Shapes
            If t.Name = textBoxName Then
                If t.Type <> Microsoft.Office.Core.MsoShapeType.msoTextBox Then Resume Next
                f = True
                Exit For
            End If
        Next
        If Not f Then Return "NOTHING TO COPY"
        Dim oName = t.Name
        Dim nName = t.Name & "_" & Guid(True)
        Dim nt As Excel.Shape
        Dim textBoxScale As Double
        textBoxScale = 1
        If Not (t.Width <= r1st.MergeArea.Width And t.Height <= r1st.MergeArea.Height) Then
            textBoxScale = Min(r1st.MergeArea.Width / t.Width, r1st.MergeArea.Height / t.Height)
        End If

        nt = ws.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 0, 0)
        attachedObjects.Add(r1st.Address, nName)
        nt.Name = nName
        ReserveClipboard()
        Dim cb
        t.TextFrame2.TextRange.Copy()
        nt.TextFrame2.TextRange.Paste()
        RestoreClipboard()
        nt.TextFrame.MarginTop = 0
        nt.TextFrame.MarginBottom = 0
        nt.TextFrame.MarginLeft = 0
        nt.TextFrame.MarginRight = 0
        nt.TextFrame.AutoSize = False
        nt.TextFrame.VerticalOverflow = XlOartVerticalOverflow.xlOartVerticalOverflowOverflow
        nt.TextFrame.HorizontalOverflow = XlOartHorizontalOverflow.xlOartHorizontalOverflowOverflow
        nt.Placement = XlPlacement.xlFreeFloating
        nt.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
        nt.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
        If width > 0 Then nt.Width = width Else nt.Width = t.Width * textBoxScale
        If height > 0 Then nt.Height = height Else nt.Height = t.Height * textBoxScale
        If top > 0 Then nt.Top = top Else nt.Top = r1st.MergeArea.Top + (r1st.MergeArea.Height - nt.Height) / 2
        If left > 0 Then nt.Left = left Else nt.Left = r1st.MergeArea.Left + (r1st.MergeArea.Width - nt.Width) / 2
        Return 0
    End Function
    <ExcelFunction>
    Public Function RemoveObject(objectName As String, Optional worksheetName As String = "", Optional continued As Boolean = False) As ExcelVariant
        Dim ws As Worksheet = Nothing
        If worksheetName = "" Then
            Application.ThisWorkbook.Activate()
            ws = Application.ActiveSheet
        Else
            Dim a As Worksheet
            For Each a In Application.ActiveWorkbook.Worksheets
                If a.Name = worksheetName Then ws = a
            Next
        End If
        If ws Is Nothing Then Return "NOTHING TO REMOVE"
        Dim s As Shape = Nothing
        Dim f As Boolean
        f = False
        For Each s In ws.Shapes
            If s.Name = objectName Then
                s.Delete()
                f = True
                If Not continued Then Exit For
            End If
        Next
        If Not f Then Return "NOTHING TO REMOVE" Else Return 0
    End Function

    'Questionable
    '<ExcelFunction>
    'Public Function FormulaRegister(formulaName As String, formula As String) As ExcelVariant
    '    Static FormulaDictionary As New Dictionary(Of String, String)
    '    If formulaName = "__ExtractDictionary" Then 'It will be rewrite.
    '        Return FormulaDictionary
    '        Exit Function
    '    End If
    '    If Not FormulaDictionary.ContainsKey(formulaName) Then FormulaDictionary.Add(formulaName, formula) Else If FormulaDictionary(formulaName) <> formula Then FormulaDictionary(formulaName) = formula Else Return -1
    '    Return 0
    'End Function
    '<ExcelFunction(Description:="Call formula registered.", IsVolatile:=True, IsMacroType:=False)>
    'Public Function FormulaCall(formulaName As String, ParamArray macros As String()) As ExcelVariant
    '    Dim d As Dictionary(Of String, String)
    '    d = FormulaRegister("__ExtractDictionary", "")
    '    If d.ContainsKey(formulaName) Then
    '        Dim f As String
    '        f = d(formulaName)
    '        Dim m As Integer
    '        m = 1
    '        For i = LBound(macros) To UBound(macros)
    '            f = Replace(f, "$$" & m, macros(i))
    '        Next
    '        Return Application.Run(f)
    '    Else
    '        Return ExcelErrorName
    '    End If
    'End Function
End Module
