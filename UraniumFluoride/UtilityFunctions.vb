Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports UraniumFluoride.Helper
#Region "Imports Macros"
Imports ExcelRange = System.Object
Imports ExcelNumber = System.Object
Imports ExcelLogical = System.Object
Imports ExcelDate = System.Object
Imports ExcelString = System.Object
Imports ExcelVariant = System.Object
Imports ExcelBoolean = System.Object
Imports ExcelIndex = System.Int32
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

    Public ReadOnly Property CallerWorksheet As Excel.Worksheet
        Get
            Return ConvertToRange(XlCall.Excel(XlCall.xlfCaller)).Worksheet
        End Get
    End Property

    Public ReadOnly Property CallerWorkbook As Excel.Workbook
        Get
            Return CallerWorksheet.Parent
        End Get
    End Property

    Private Function CheckErrorCode(value As ExcelVariant) As ExcelVariant
        Static ErrorCodeDictionary As New Dictionary(Of Integer, ExcelError) From {{-2146826281, ExcelErrorDiv0},
        {-2146826246, ExcelErrorNa}, {-2146826259, ExcelErrorName}, {-2146826288, ExcelErrorNull},
        {-2146826252, ExcelErrorNum}, {-2146826265, ExcelErrorRef}, {-2146826273, ExcelErrorValue}}
        If IsNumeric(value) AndAlso ErrorCodeDictionary.ContainsKey(value) Then Return ErrorCodeDictionary(value)
        Return value
    End Function

    <ExcelFunction(Description:="Banker round")>
    Public Function BankerRound(<MarshalAs(UnmanagedType.Currency)> num As Decimal, pre As Integer, Optional isSignificant As Boolean = False) As ExcelNumber
        If num = 0 Then Return 0
        If isSignificant Then
            Dim power As Integer = Math.Floor(Math.Log10(Math.Abs(num)))

            Return CheckErrorCode(Math.Round(Math.Round(num / 10 ^ power, 14), pre - 1, MidpointRounding.ToEven) * 10 ^ power)
        Else
            Return CheckErrorCode(Math.Round(Math.Round(num, 14), pre, MidpointRounding.ToEven))
        End If
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function AverageByMean(num As ExcelVariant(,), Optional ratio As Double = 0.1) As ExcelNumber
        Try
            Dim value As ExcelVariant() = TrimNumericArray(MatrixToArray(num))
            If Count(value) = 0 Then Return ExcelErrorNull
            If Count(value) < 3 Then Return Average(value)
            Dim f As Boolean = False
            Dim ave = Average(value)
            If ave = 0 Then Return 0
            Dim i As Integer
            For i = 0 To value.Count - 1
                If Math.Abs(value(i) - ave) / ave >= ratio Then f = True
                If f Then Exit For
            Next
            If f Then
                value(i) = ExcelEmpty.Value
                Dim ave2 As Decimal = Average(value)
                If ave2 = 0 Then Return 0
                If Count(value) <= 3 Then Return ave2
                For j = 0 To value.Count - 1
                    If TypeOf value(j) IsNot ExcelEmpty AndAlso Math.Abs(value(j) - ave) / ave >= ratio Then Return ExcelErrorValue
                Next
                Return ave2
            End If
            Return ave
        Finally
        End Try
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function VerifyByMean(num As ExcelVariant(,), Optional ratio As Double = 0.1) As ExcelNumber
        Try
            Dim value As ExcelVariant() = TrimNumericArray(MatrixToArray(num))
            If Count(value) < 3 Then Return 65535
            Dim f As Boolean = False
            Dim ave = Average(value)
            If ave = 0 Then Return -65536
            Dim i As Integer
            For i = 0 To value.Count - 1
                If Math.Abs(value(i) - ave) / ave >= ratio Then f = True
                If f Then Exit For
            Next
            If f Then
                value(i) = ExcelEmpty.Value
                Dim ave2 As Decimal = Average(value)
                If ave2 = 0 Then Return -65536
                For j = 0 To value.Count - 1
                    If TypeOf value(j) IsNot ExcelEmpty And Math.Abs(value(j) - ave) / ave >= ratio Then Return -1
                Next
                Return 0
            End If
            Return 1
        Finally
        End Try
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function AverageByMedian(num As ExcelVariant(,), Optional ratio As Double = 0.15) As ExcelNumber
        Try
            Dim value As ExcelVariant() = TrimNumericArray(MatrixToArray(num))
            If Count(value) <= 2 Then Return Average(value)
            Dim min As Decimal = UtilityFunctions.Min(value)
            Dim max As Decimal = UtilityFunctions.Max(value)
            Dim med As Decimal = UtilityFunctions.Med(value)
            If med = 0 Then Return ExcelErrorDiv0
            If Math.Abs((max - med) / med) >= CDec(ratio) And Math.Abs((med - min) / med) >= CDec(ratio) Then Return ExcelErrorValue
            If Math.Abs((max - med) / med) >= CDec(ratio) Or Math.Abs((med - min) / med) >= CDec(ratio) Then Return med
            Return Average(value)
        Finally
        End Try
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function VerifyByMedian(num As ExcelVariant(,), Optional ratio As Double = 0.15) As ExcelNumber
        Try
            Dim value As ExcelVariant() = TrimNumericArray(MatrixToArray(num))
            If Count(value) <= 2 Then Return 65535
            Dim min As Decimal = UtilityFunctions.Min(value)
            Dim max As Decimal = UtilityFunctions.Max(value)
            Dim med As Decimal = UtilityFunctions.Med(value)
            If med = 0 Then Return -65536
            If Math.Abs((max - med) / med) >= CDec(ratio) And Math.Abs((med - min) / med) >= CDec(ratio) Then Return -1
            If Math.Abs((max - med) / med) >= CDec(ratio) Or Math.Abs((med - min) / med) >= CDec(ratio) Then Return 0
            Return 1
        Finally
        End Try
    End Function

    <ExcelFunction(IsVolatile:=True)>
    Public Function RandNoRepeat(bottom As Integer, top As Integer, Optional memorySet As Integer = 0, Optional memories As Integer = 30, Optional unrepeatPossibility As Integer = 0.99, Optional seed As ExcelNumber = "") As ExcelNumber
        Static memory As New ValueCircularListCollection(1024, 1024, Nothing)
        If top - bottom < 2 Then Return bottom
        If memories < 2 Then memories = 2
        If memories > memory.ListCapacity Then memories = memory.ListCapacity
        If unrepeatPossibility > 1 Then unrepeatPossibility = 1
        If top - bottom < memories Then memories = top - bottom - 1
        If memorySet > memory.ListCapacity - 1 Or memorySet < 0 Then memorySet = 0
        Dim result As Integer

        If IsNumeric(seed) Then
            Dim randomer As New Random(CDec(seed).GetHashCode), randomer2 As New Random(Not CDec(seed).GetHashCode)

            Dim memorySeeded(memories) As Integer
            For i = 0 To memories
                Dim f As Boolean = False
                Do Until f
                    f = True
                    memorySeeded(i) = randomer.Next(bottom, top)
                    For j = 0 To i - 1
                        If memorySeeded(i) = result Then If randomer2.Next(0, 1000) / 1000 < unrepeatPossibility Then f = False
                    Next
                Loop
            Next
            result = memorySeeded.Last
        Else
            Static randomer As New Random(Now.Millisecond), randomer2 As New Random(Not Now.Millisecond)

            Dim f As Boolean = False
            Do Until f
                f = True
                result = randomer.Next(bottom, top)
                For i = -memories + 1 To 0
                    If memory(memorySet)(i) IsNot Nothing AndAlso memory(memorySet)(i) = result Then If randomer2.Next(0, 1000) / 1000 < unrepeatPossibility Then f = False
                Next
            Loop
            memory(memorySet).MoveNext(result)
        End If
        Return result
    End Function

    <ExcelFunction(IsVolatile:=True)>
    Public Function RandDiscreted(bottom As Integer, top As Integer, Optional memorySet As Integer = 0, Optional memories As Integer = 30, Optional unrepeatPossibility As Integer = 0.99, Optional seed As ExcelNumber = "") As ExcelNumber
        Static memory As New ValueCircularListCollection(1024, 1024, Nothing)
        If top - bottom < 2 Then Return bottom
        If memories < 2 Then memories = 2
        If memories > memory.ListCapacity Then memories = memory.ListCapacity
        If unrepeatPossibility > 1 Then unrepeatPossibility = 1
        If top - bottom < memories Then memories = top - bottom - 1
        If memorySet > memory.ListCapacity - 1 Or memorySet < 0 Then memorySet = 0
        Dim result As Integer
        If IsNumeric(seed) Then
            Dim randomer As New Random(CDec(seed).GetHashCode), randomer2 As New Random(Not CDec(seed).GetHashCode)

            Dim memorySeeded(memories) As Integer
            For i = 0 To memories
                Dim f As Boolean = False
                Do Until f
                    f = True
                    memorySeeded(i) = randomer.Next(bottom, top)
                    For j = 0 To i - 1
                        If randomer2.Next(0, 1000) / 1000 < unrepeatPossibility * (1 - Math.Abs(result - memorySeeded(i) / (top - bottom))) Then
                            f = False
                            Continue Do
                        End If
                    Next
                Loop
            Next
            result = memorySeeded.Last
        Else
            Static randomer As New Random(Now.Millisecond), randomer2 As New Random(Now.Millisecond And Now.Millisecond)

            Dim f As Boolean = False
            Do Until f
                f = True
                result = randomer.Next(bottom, top)
                For i = -memories + 1 To 0
                    If memory(memorySet)(i) IsNot Nothing AndAlso randomer2.Next(0, 1000) / 1000 < unrepeatPossibility * (1 - Math.Abs(result - CDbl(memory(memorySet)(i)) / (top - bottom))) Then
                        f = False
                        Continue Do
                    End If
                Next
            Loop
            memory(memorySet).MoveNext(result)
        End If
        Return result
    End Function

    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function PageLocalize(<ExcelArgument(AllowReference:=True)> r As ExcelRange, pageRowsCount As Integer, pageColumnsCount As Integer, locationRow As Integer, locationColumn As Integer, pageIndex As ExcelIndex) As ExcelVariant
        Dim _Range As Excel.Range = ConvertToRange(r)
        If _Range Is Nothing Then Return ExcelErrorNa
        If Not (New CloseInterval2(Of Integer)(1, 1, pageRowsCount, pageColumnsCount).Contains(locationRow, locationColumn) And
                New CloseInterval2(Of Integer)(1, 1, _Range.Rows.Count, _Range.Columns.Count).Contains(pageRowsCount, pageColumnsCount)) Then _
           Return Nothing
        Dim pageCount, pageCountInRow, pageCountInColumn As Integer
        Try
            pageCountInRow = _Range.Rows.Count \ pageRowsCount
            pageCountInColumn = _Range.Columns.Count \ pageColumnsCount
            pageCount = pageCountInRow * pageCountInColumn
            If Not New CloseInterval(Of Integer)(1, pageCount).Contains(pageIndex) Then Return Nothing
            Dim pageIndexInRow, pageIndexInColumn As Integer
            pageIndex -= 1
            pageIndexInRow = pageIndex \ pageCountInColumn
            pageIndexInColumn = pageIndex Mod pageCountInColumn
            Dim row, column As Integer
            row = pageRowsCount * pageIndexInRow + locationRow
            column = pageColumnsCount * pageIndexInColumn + locationColumn
            Return _Range(row, column).Value
        Catch ex As NullReferenceException
            Return ExcelErrorNa
        End Try
    End Function

    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function PageLocalizeAbbr(<ExcelArgument(AllowReference:=True)> rPage As ExcelRange, <ExcelArgument(AllowReference:=True)> rCell As ExcelRange, pageIndex As Integer) As ExcelVariant
        Dim page_Range As Excel.Range = ConvertToRange(rPage), cell_Range As Excel.Range = ConvertToRange(rCell)
        Dim cellRow, cellColumn, pageRow, pageColumn As Integer
        Try
            pageRow = page_Range.Rows.Count
            pageColumn = page_Range.Columns.Count
            cellRow = cell_Range.Row
            cellColumn = cell_Range.Column
        Catch ex As NullReferenceException
            Return ExcelErrorNa
        End Try
        Do Until cellRow <= pageRow
            cellRow -= pageRow
        Loop
        Do Until cellColumn <= pageColumn
            cellColumn -= pageColumn
        Loop
        Dim pageLocation = GetPageLocationLRTB(rPage, pageIndex)
        Dim locationRow = pageRow * (pageLocation.X - 1) + cellRow
        Dim locationColumn = pageColumn * (pageLocation.Y - 1) + cellColumn
        Dim usedRow = page_Range.Worksheet.UsedRange.Rows.Count
        Dim usedColumn = page_Range.Worksheet.UsedRange.Columns.Count
        If locationRow <= usedRow And locationColumn <= usedColumn And locationRow > 0 And locationColumn > 0 Then Return page_Range.Worksheet.UsedRange(locationRow, locationColumn).Value Else Return ""
    End Function
    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function GetPageCount(<ExcelArgument(AllowReference:=True)> rPage As ExcelRange) As ExcelVariant
        Dim page_Range As Excel.Range = ConvertToRange(rPage)
        Try
            Return GetPageCountInRow(rPage) * GetPageCountInColumn(rPage)
        Catch
            Return ExcelErrorNa
        End Try
    End Function
    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function GetPageCountInRow(<ExcelArgument(AllowReference:=True)> rPage As ExcelRange) As ExcelVariant
        Dim page_Range As Excel.Range = ConvertToRange(rPage)
        Try
            Return Math.Ceiling(page_Range.Worksheet.UsedRange.Columns.Count / page_Range.Columns.Count)
        Catch
            Return ExcelErrorNa
        End Try
    End Function
    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function GetPageCountInColumn(<ExcelArgument(AllowReference:=True)> rPage As ExcelRange) As ExcelVariant
        Dim page_Range As Excel.Range = ConvertToRange(rPage)
        Try
            Return Math.Ceiling(page_Range.Worksheet.UsedRange.Rows.Count / page_Range.Rows.Count)
        Catch
            Return ExcelErrorNa
        End Try
    End Function

    Public Function GetPageLocationLRTB(<ExcelArgument(AllowReference:=True)> rPage As ExcelRange, index As ExcelIndex) As (X As ExcelIndex, Y As ExcelIndex)
        Dim countInRow = GetPageCountInRow(rPage)
        Return (Math.Ceiling(index / countInRow), index Mod countInRow)
    End Function

    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function PageLocalizeSorted(<ExcelArgument(AllowReference:=True)> rPage As ExcelRange, <ExcelArgument(AllowReference:=True)> rSortingCell As ExcelRange, <ExcelArgument(AllowReference:=True)> rSearchingCell As ExcelRange, Optional rank As ExcelIndex = 0, Optional isAscending As Boolean = True) As ExcelVariant
        Dim page_Range As Excel.Range = ConvertToRange(rPage), sortingCell_Range As Excel.Range = ConvertToRange(rSortingCell), serchingCell_Range As Excel.Range = ConvertToRange(rSearchingCell)
        Dim cellRow, cellColumn, pageRow, pageColumn As Integer
        If rank <= 0 Then rank = 1
        Try
            pageRow = page_Range.Rows.Count
            pageColumn = page_Range.Columns.Count
            cellRow = serchingCell_Range.Row
            cellColumn = serchingCell_Range.Column
        Catch ex As NullReferenceException
            Return ExcelErrorNa
        End Try
        Do Until cellRow <= pageRow
            cellRow -= pageRow
        Loop
        Do Until cellColumn <= pageColumn
            cellColumn -= pageColumn
        Loop
        Try
            Dim pageCount = GetPageCount(rPage)
            Dim sortingDictionary As New Dictionary(Of ExcelVariant, Integer)
            For i = 1 To pageCount
                sortingDictionary.Add(PageLocalizeAbbr(rPage, rSortingCell, i), i)
            Next
            If isAscending Then sortingDictionary = sortingDictionary.OrderBy(Function(o) o.Key).ToDictionary(Function(o) o.Key, Function(o) o.Value) Else sortingDictionary = sortingDictionary.OrderByDescending(Function(o) o.Key).ToDictionary(Function(o) o.Key, Function(o) o.Value)
            Return PageLocalizeAbbr(rPage, rSearchingCell, sortingDictionary.Values(rank - 1))
        Catch
            Return ExcelErrorValue
        End Try
    End Function

    '<ExcelFunction(IsMacroType:=True)>
    'Public Function PageLocalize2(<ExcelArgument(AllowReference:=True)> rCell As ExcelRange, pageIndex As Integer) As ExcelVariant
    '    Dim cell_Range As Excel.Range = ConvertToRange(rCell)
    '    Dim r As ExcelRange = cell_Range.Worksheet.UsedRange
    '
    'End Function

    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function PageLocalizeByKeyword(<ExcelArgument(AllowReference:=True)> rPage As ExcelRange, <ExcelArgument(AllowReference:=True)> rCell As ExcelRange, keyword As ExcelVariant, Optional isValueMatching As Boolean = False) As ExcelVariant
        Dim page_Range As Excel.Range = ConvertToRange(rCell), cell_Range As Excel.Range = ConvertToRange(rCell)
        Dim locationRow = FindRow(page_Range.Worksheet.UsedRange, keyword, isValueMatching)
        Dim locationColumn = FindColumn(page_Range.Worksheet.UsedRange, keyword, isValueMatching)
        If locationRow = 0 Then Return ExcelErrorNa
        Dim pageIndex As Integer = Math.Ceiling(page_Range.Worksheet.UsedRange.Columns.Count / page_Range.Columns.Count) * (locationRow \ page_Range.Rows.Count) + Math.Ceiling(locationColumn \ page_Range.Columns.Count)
        Return PageLocalizeAbbr(rPage, rCell, pageIndex)
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
    Public Function RangeText(<ExcelArgument(AllowReference:=True)> r As ExcelRange) As String
        Dim _Range As Excel.Range = ConvertToRange(r)
        _Range(1, 1).Calculate
        Return _Range(1, 1).Text
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function MergedCellRows(<ExcelArgument(AllowReference:=True)> cell As ExcelRange) As ExcelNumber
        Return ConvertToRange(cell)(1, 1).MergeArea.Rows.Count
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function MergedCellColumns(<ExcelArgument(AllowReference:=True)> cell As ExcelRange) As ExcelNumber
        Return ConvertToRange(cell)(1, 1).MergeArea.Columns.Count
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function GetRangeMerged(<ExcelArgument(AllowReference:=True)> cell As ExcelRange) As ExcelRange
        Return ConvertToExcelReference(ConvertToRange(cell)(1, 1).MergeArea(1, 1))
    End Function

    <ExcelFunction(IsMacroType:=True, IsVolatile:=True)>
    Public Function ReferencedMergedCellRows(Optional rangeText As String = "A1", Optional path As String = "", Optional worksheetName As String = "") As ExcelNumber
        Dim r = RelativeReferenceInternal(rangeText, path, worksheetName)
        If TypeOf r Is ClosedXML.Excel.IXLRange Then
            Return CType(r, ClosedXML.Excel.IXLRange).CellsUsed.First.MergedRange.RowCount
        Else
            Dim _range As Excel.Range = ConvertToRange(r)
            Return MergedCellRows(_range)
        End If
        Return ExcelErrorNull
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function ReferencedMergedCellColumns(Optional rangeText As String = "A1", Optional path As String = "", Optional worksheetName As String = "") As ExcelNumber
        Dim r = RelativeReferenceInternal(rangeText, path, worksheetName)
        If TypeOf r Is ClosedXML.Excel.IXLRange Then
            Return CType(r, ClosedXML.Excel.IXLRange).CellsUsed.First.MergedRange.ColumnCount
        Else
            Dim _range As Excel.Range = ConvertToRange(r)
            Return MergedCellColumns(_range)
        End If
        Return ExcelErrorNull
    End Function

    <ExcelFunction(IsVolatile:=True, IsMacroType:=True)>
    Public Function DataFitter(formula As String, <MarshalAs(UnmanagedType.Currency)> formulaResult As Decimal, variantIndexToReturn As Integer, minValues As ExcelVariant(,), maxValues As ExcelVariant(,), stepValues As ExcelVariant(,), Optional isForceRecalculating As Boolean = False) As ExcelNumber
        Static ResultCollection As New Dictionary(Of Integer, ExcelNumber())

        Dim minArray = TrimNumericArray(MatrixToArray(minValues))
        Dim maxArray = TrimNumericArray(MatrixToArray(maxValues))
        Dim stepArray = TrimNumericArray(MatrixToArray(stepValues))
        Dim valuesCount As Integer = Min({minArray.Count, maxArray.Count, stepArray.Count})
        If variantIndexToReturn > valuesCount Then Return ExcelErrorNa
        Dim fittedValues(valuesCount - 1) As ExcelNumber
        For i = 0 To valuesCount - 1
            fittedValues(i) = minArray(i)
        Next

        For i = 0 To valuesCount - 1
            If stepArray(i) < 0 Or maxArray(i) - minArray(i) < stepArray(i) Then Return ExcelErrorValue
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
            info.Append(minArray(i))
            info.Append("-")
            info.Append(maxArray(i))
            info.Append("|")
            info.Append(stepArray(i))
            If i <> valuesCount - 1 Then info.Append(",")
        Next
        info.Append("}.")

        Dim argumentsHash As Integer = formula.GetHashCode Xor formulaResult.GetHashCode Xor GetNumericArrayHash(minArray) Xor GetNumericArrayHash(maxArray) Xor GetNumericArrayHash(stepArray)
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
                                             If MsgBox("We have calculated " & calculationCount & " times. Do you want to continue?", MsgBoxStyle.OkCancel Xor MsgBoxStyle.Question, "It took so long...") = MsgBoxResult.Cancel Then
                                                 f = False
                                                 Exit Do
                                             End If
                                         End If
                                         For i = 0 To valuesCount - 1
                                             fittedValues(i) = fittedValues(i) + stepArray(i)
                                             If fittedValues.Last > maxArray(valuesCount - 1) Then
                                                 f = False
                                                 Exit For
                                             End If
                                             If fittedValues(i) <= maxArray(i) Then Exit For Else fittedValues(i) = minArray(i)
                                         Next
                                     Loop
                                     If Not f Then
                                         For i = 0 To fittedValues.Count - 1
                                             fittedValues(i) = ExcelErrorNa
                                         Next
                                     End If
                                 End Sub).Wait()
        w.Close()
        If ResultCollection.ContainsKey(argumentsHash) Then ResultCollection(argumentsHash) = fittedValues Else ResultCollection.Add(argumentsHash, fittedValues)
        Return fittedValues(variantIndexToReturn - 1)
    End Function

    <ExcelFunction>
    Public Function RegExFind(input As String, pattern As String, Optional index As Integer = 1, Optional isCaseIgnore As Boolean = True) As ExcelNumber
        If Not IsNumeric(index) Or index < 1 Then index = 1
        Dim e As New Text.RegularExpressions.Regex(pattern, If(isCaseIgnore, System.Text.RegularExpressions.RegexOptions.IgnoreCase, System.Text.RegularExpressions.RegexOptions.None))
        Dim m As System.Text.RegularExpressions.MatchCollection = e.Matches(input)
        If m.Count <= index - 1 Then Return -1 Else Return m(index - 1).Index + 1
    End Function

    <ExcelFunction>
    Public Function RegExMatch(input As String, pattern As String, Optional index As Integer = 1, Optional isCaseIgnore As Boolean = True) As ExcelString
        If Not IsNumeric(index) Or index < 1 Then index = 1
        Dim e As New Text.RegularExpressions.Regex(pattern, If(isCaseIgnore, System.Text.RegularExpressions.RegexOptions.IgnoreCase, System.Text.RegularExpressions.RegexOptions.None))
        Dim m As System.Text.RegularExpressions.MatchCollection = e.Matches(input)
        If m.Count <= index - 1 Then Return ExcelErrorNull Else Return m(index - 1).Value
    End Function

    <ExcelFunction>
    Public Function RegExReplace(input As String, pattern As String, replacement As String) As ExcelString
        Return System.Text.RegularExpressions.Regex.Replace(input, pattern, replacement)
    End Function

    <ExcelFunction(IsMacroType:=True, IsVolatile:=True)>
    Public Function RelativeReference(Optional rangeText As String = "A1", Optional path As String = "", Optional worksheetName As String = "") As ExcelRange
        Dim result = RelativeReferenceInternal(rangeText, path, worksheetName)
        If TypeOf result Is ClosedXML.Excel.IXLRange Then
            Dim r = CType(result, ClosedXML.Excel.IXLRange)
            If r.FirstCell.NeedsRecalculation Then r.Worksheet.RecalculateAllFormulas()
            Return r.FirstCell.CachedValue
        Else
            Return ConvertToExcelReference(result)
        End If
    End Function

    Private Function RelativeReferenceInternal(Optional rangeText As String = "A1", Optional path As String = "", Optional worksheetName As String = "") As ExcelRange
        Dim wb As Workbook = Nothing
        If path = "" Then
            Try
                wb = CallerWorkbook
            Catch ex As COMException
                wb = Application.ActiveWorkbook
            End Try
        Else
            If IO.File.Exists(path) Then
                Try
                    Dim wbc = From currentWb As Workbook In Application.Workbooks Where currentWb.Name = path.Split("\").Last Select currentWb
                    If wbc.Count > 0 Then
                        wb = wbc.First
                        If wb.FullName <> path Then Return ExcelErrorNa
                    Else
                        Dim closedWb = Helper.ClosedXMLWorkbookLibrary.Create(path)
                        If worksheetName = "" Then Return closedWb.Worksheets(0).Range(rangeText)
                        For Each i In closedWb.Worksheets
                            If i.Name = worksheetName Then Return i.Range(rangeText)
                        Next
                    End If
                Catch
                    Dim tempFolder As String = Environment.GetEnvironmentVariable("TEMP")
                    Dim tempPath As String = tempFolder & "\" & New IO.FileInfo(path).Name
                    Try
                        FileIO.FileSystem.CopyFile(path, tempPath, True)
                    Catch
                    End Try
                    Dim closedWb = Helper.ClosedXMLWorkbookLibrary.Create(tempPath, path)
                        If worksheetName = "" Then Return closedWb.Worksheets(0).Range(rangeText)
                        For Each i In closedWb.Worksheets
                            If i.Name = worksheetName Then Return i.Range(rangeText)
                        Next
                    End Try
                    Else Return ExcelErrorNa
            End If
        End If
        If worksheetName = "" Then Return CallerWorksheet.Range(rangeText)
        For Each i As Worksheet In wb.Worksheets
            If i.Name = worksheetName Then Return i.Range(rangeText)
        Next
        Return ExcelErrorNa
    End Function

    <ExcelFunction(IsVolatile:=True)>
    Public Function ReferencedVLookUp(keyword As ExcelVariant, Optional rangeText As String = "A1", Optional path As String = "", Optional worksheetName As String = "", Optional targetColumn As Integer = 1, Optional isApproximateMatching As Boolean = False) As ExcelVariant
        Dim r = RelativeReferenceInternal(rangeText, path, worksheetName)
        If TypeOf r Is ClosedXML.Excel.IXLRange Then
            For Each i In CType(r, ClosedXML.Excel.IXLRange).FirstColumn.CellsUsed
                Dim t = If(TypeOf i.Value Is Date And IsNumeric(keyword), Date.FromOADate(keyword), keyword)
                If i.Value.GetType = t.GetType AndAlso i.Value = t Then Return CType(r, ClosedXML.Excel.IXLRange).Cell(i.Address.RowNumber, targetColumn).CachedValue
            Next
        Else
            Return XlCall.Excel(XlCall.xlfVlookup, If(IsDate(keyword), CDate(keyword).ToOADate, keyword), r.Value, targetColumn, isApproximateMatching)
        End If
        Return ExcelErrorNull
    End Function

    <ExcelFunction(IsVolatile:=True)>
    Public Function ReferencedHLookUp(keyword As ExcelVariant, Optional rangeText As String = "A1", Optional path As String = "", Optional worksheetName As String = "", Optional targetRow As Integer = 1, Optional isApproximateMatching As Boolean = False) As ExcelVariant
        Dim r = RelativeReference(rangeText, path, worksheetName)
        If TypeOf r Is ClosedXML.Excel.IXLRange Then
            For Each i In CType(r, ClosedXML.Excel.IXLRange).FirstRow.CellsUsed
                If i.Value = keyword Then Return CType(r, ClosedXML.Excel.IXLRange).Cell(targetRow, i.Address.ColumnNumber).CachedValue
            Next
        Else
            Return XlCall.Excel(XlCall.xlfHlookup, If(IsDate(keyword), CDate(keyword).ToOADate, keyword), r.Value, targetRow, isApproximateMatching)
        End If
        Return ExcelErrorNull
    End Function

    <ExcelFunction>
    Public Function RegExMatchesCount(input As String, pattern As String, Optional startat As Integer = 0) As ExcelNumber
        If Not IsNumeric(startat) Or startat < 1 Then startat = 1
        Return New Text.RegularExpressions.Regex(pattern).Matches(input, startat).Count
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function VLookUpByRank(value As ExcelVariant(,), rank As Integer, rankColumn As Integer, lookupColumn As Integer) As ExcelVariant
        If value.GetLength(1) < rankColumn Or value.GetLength(1) < lookupColumn Or rankColumn < 1 Or lookupColumn < 1 Then Return ExcelErrorRef
        Dim rankTable As New Dictionary(Of Integer, ExcelVariant)
        For i = 0 To value.GetLength(0) - 1
            rankTable.Add(i, value(i, rankColumn).value)
        Next
        'Bad sorting implementation, will be rewrite.
        For i = 0 To value.GetLength(0) - 1
            For j = 0 To value.GetLength(0) - 2
                If IsError(rankTable(j)) Or Application.WorksheetFunction.isblank(rankTable(j)) Then
                    Swap(rankTable.Values(j), rankTable.Values(j + 1))
                    Swap(rankTable.Keys(j), rankTable.Keys(j + 1))
                ElseIf IsError(rankTable(j + 1)) Then
                ElseIf rankTable(j) > rankTable(j + 1) Then
                    Swap(rankTable.Values(j), rankTable.Values(j + 1))
                    Swap(rankTable.Keys(j), rankTable.Keys(j + 1))
                End If
            Next
        Next
        Return value(rankTable.Keys(rank), lookupColumn).Value
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function HLookUpByRank(value As ExcelVariant(,), rank As Integer, rankRow As Integer, lookupRow As Integer) As ExcelVariant
        If value.GetLength(0) < rankRow Or value.GetLength(0) < lookupRow Or rankRow < 1 Or lookupRow < 1 Then Return ExcelErrorRef
        Dim rankTable As New Dictionary(Of Integer, ExcelVariant)
        For i = 1 To value.GetLength(1)
            rankTable.Add(i, value(rankRow, i).value)
        Next
        'Bad sorting implementation, will be rewrite.
        For i = 0 To value.GetLength(1) - 1
            For j = 0 To value.GetLength(1) - 2
                If IsError(rankTable(j)) Or Application.WorksheetFunction.isblank(rankTable(j)) Then
                    Swap(rankTable.Values(j), rankTable.Values(j + 1))
                    Swap(rankTable.Keys(j), rankTable.Keys(j + 1))
                ElseIf IsError(rankTable(j + 1)) Then
                ElseIf rankTable(j) > rankTable(j + 1) Then
                    Swap(rankTable.Values(j), rankTable.Values(j + 1))
                    Swap(rankTable.Keys(j), rankTable.Keys(j + 1))
                End If
            Next
        Next
        Return value(lookupRow, rankTable.Keys(rank)).Value
    End Function

    <ExcelFunction>
    Public Function Contains(arg As ExcelVariant, searching As ExcelVariant) As ExcelLogical
        If IsArray(arg) AndAlso (LBound(searching, 1) > 1 OrElse LBound(searching, 2) > 1) Then
            For Each i In arg
                If searching Is arg Then
                    Dim j As Integer = LBound(searching)
                    If i = searching(j) Then Return True
                Else
                    If i = searching Then Return True
                End If
            Next
        Else
            Try
                If IsArray(arg) Then Return arg(0).Contains(searching) Else Return arg.Contains(searching)
            Catch ex As Exception
                Return ExcelErrorValue
            End Try
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
    Public Function MinIndex(ParamArray args As ExcelVariant()) As ExcelNumber
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
    Public Function MaxIndex(ParamArray args As ExcelVariant()) As ExcelNumber
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
        If Not isCompress Then
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

    <ExcelFunction(IsVolatile:=False, IsMacroType:=True)>
    Public Function CopyTextbox(<ExcelArgument(AllowReference:=True)> r As ExcelRange, textBoxName As String, Optional removeAllRecordedTextBoxes As Boolean = False, Optional left As Double = 0, Optional top As Double = 0, Optional width As Double = 0, Optional height As Double = 0) As ExcelVariant
        Static attachedObjects As New Dictionary(Of String, Shape)
        Dim _Range As Excel.Range = ConvertToRange(r)
        Dim ws As Worksheet = _Range.Worksheet
        Dim r1st As Excel.Range = _Range(1, 1)
        If Not TryCast(_Range.Worksheet.Parent, Workbook)?.FullName = Application.ActiveSheet.Parent.Fullname OrElse Not _Range.Worksheet.CodeName = Application.ActiveSheet.Codename Then
            If attachedObjects.ContainsKey(r1st.Address) Then Return 0 Else Return "HANGED"
        End If

        If removeAllRecordedTextBoxes Then
            For o = 0 To attachedObjects.Count - 1
                attachedObjects.Values(o).Delete()
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(attachedObjects.Values(o))
            Next
            attachedObjects.Clear()
        End If

        Dim t As Shape = Nothing
        Dim f As Boolean = False
        For Each t In ws.Shapes
            If t.Name = textBoxName Then
                If t.Type <> Microsoft.Office.Core.MsoShapeType.msoTextBox Then Continue For
                f = True
                Exit For
            End If
        Next
        If Not f Then Return "NOTHING TO COPY"

        Dim nName = t.Name & "_" & Guid(True)
        Dim nt As Excel.Shape
        Dim textBoxScale As Double = 1
        Try
            If Not (t.Width <= r1st.MergeArea.Width And t.Height <= r1st.MergeArea.Height) Then textBoxScale = Min({r1st.MergeArea.Width / t.Width, r1st.MergeArea.Height / t.Height})
        Catch ex As AccessViolationException
            textBoxScale = 1
        End Try
        If width <= 0 Then width = t.Width * textBoxScale
        If height <= 0 Then height = t.Height * textBoxScale
        If top <= 0 Then top = r1st.MergeArea.Top + (r1st.MergeArea.Height - height) / 2
        If left <= 0 Then left = r1st.MergeArea.Left + (r1st.MergeArea.Width - width) / 2

        Do While attachedObjects.ContainsKey(r1st.Address)
            Try
                If width = attachedObjects(r1st.Address).Width And height = attachedObjects(r1st.Address).Height And top = attachedObjects(r1st.Address).Top And left = attachedObjects(r1st.Address).Left Then Return " "
                attachedObjects(r1st.Address).Delete()
            Catch ex As COMException
            Finally
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(attachedObjects(r1st.Address))
                attachedObjects.Remove(r1st.Address)
            End Try
        Loop

        Try
            t.Name = t.Name
        Catch ex As COMException
            Return "-"
        End Try
        Try
            nt = t.Duplicate
            nt.Name = nName
        Catch ex As Exception
            Throw ex
        End Try
        attachedObjects.Add(r1st.Address, nt)
        Dim timer As New Threading.Timer(Sub() Throw New TimeoutException, Nothing, 2000, Threading.Timeout.Infinite)
        Try
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
        Catch ex As TimeoutException
            Return "#TIMEOUT"
        Catch ex As Exception
            Return ExcelErrorValue
        Finally
            timer.Dispose()
        End Try
        Return 0
    End Function

    Public Class TimeoutException
        Inherits Exception
        Public Overrides ReadOnly Property Message As String
            Get
                Return "Function calling timeout."
            End Get
        End Property
    End Class

    <ExcelFunction>
    Public Function RemoveObject(objectName As String, Optional worksheetName As String = "", Optional continued As Boolean = False) As ExcelVariant
        Dim ws As Worksheet = Nothing
        If worksheetName = "" Then
            ws = CallerWorksheet
        Else
            For Each a In CallerWorkbook.Worksheets
                If a.Name = worksheetName Then ws = a
            Next
        End If
        If ws Is Nothing Then Return "NOTHING TO REMOVE"
        Dim f As Boolean
        f = False
        For Each s In ws.Shapes
            If s.Name = objectName Then
                s.Delete()
                s = Nothing
                f = True
                If Not continued Then Exit For
            End If
        Next
        If Not f Then Return "NOTHING TO REMOVE" Else Return 0
    End Function

    <ExcelFunction>
    Public Function ReserveTextbox(objectNames As String, Optional worksheetName As String = "", Optional isForceExecute As Boolean = False) As ExcelNumber
        Static firstRun As Boolean = True
        If firstRun Or isForceExecute Then
            Dim ws As Worksheet = Nothing
            If worksheetName = "" Then
                ws = CallerWorksheet
            Else
                For Each a In CallerWorkbook.Worksheets
                    If a.Name = worksheetName Then ws = a
                Next
            End If
            If ws Is Nothing Then Return "NOTHING TO REMOVE"
            Dim f As Boolean
            f = False
            Dim objectNameList = objectNames.Split(",").ToList
            For Each s In ws.Shapes
                If Not objectNameList.Contains(s.Name) And s.Type = Microsoft.Office.Core.MsoShapeType.msoTextBox Then
                    s.Delete()
                    s = Nothing
                    f = True
                End If
            Next
            firstRun = False
            If Not f Then Return "NOTHING TO REMOVE"
        End If
        Return 0
    End Function

    <ExcelFunction>
    Public Function CrLf() As ExcelString
        Return Chr(13) & Chr(10)
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function Dir(<ExcelArgument(AllowReference:=True)> Optional r As ExcelRange = Nothing) As ExcelString
        If TypeOf r Is ExcelDna.Integration.ExcelMissing Then Return Application.ActiveWorkbook.Path & "\"
        Return ConvertToRange(r).Worksheet.Parent.Path & "\"
    End Function

    <ExcelFunction>
    Public Function ColumnLetter(columnNumber As Integer) As String
        columnNumber = Math.Abs(columnNumber)
        Dim bits(5) As Char
        Dim i As Integer = 5
        Do While (columnNumber - 1) \ 26 > 0
            Dim r As Integer
            columnNumber = Math.DivRem(columnNumber - 1, 26, r)
            bits(i) = Chr(r + Asc("A"))
            i -= 1
        Loop
        bits(i) = Chr(columnNumber + Asc("A") - 1)
        Return New String(bits, i, 6 - i)
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function StringSplit(target As String, separator As String, index As Integer) As ExcelString
        If target.Split(separator).Count > index - 1 And index >= 1 Then Return target.Split(separator)(index - 1) Else Return ExcelErrorNull
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function StringSplitArray(target As String, separator As String) As ExcelString()
        Return target.Split(separator)
    End Function
    <ExcelFunction(IsMacroType:=True)>
    Public Function GetHashCode(arg As ExcelVariant) As ExcelNumber
        If TypeOf arg Is ExcelError Then
            Return CType(arg, ExcelError).GetHashCode
        ElseIf IsNumeric(arg) Then
            Return CDec(arg).GetHashCode
        ElseIf IsDate(arg) Then
            Return CDate(arg).GetHashCode
        ElseIf TypeOf arg Is ExcelEmpty Then
            Return ExcelEmpty.Value.GetHashCode
        Else
            Return arg.ToString.GetHashCode
        End If
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function UsedRange() As ExcelRange
        Return ConvertToExcelReference(CallerWorksheet.UsedRange)
    End Function

    <ExcelFunction>
    Public Function ArrayedAnd(ParamArray values As ExcelVariant()) As ExcelVariant
        Dim result(,) As ExcelVariant
        Dim count = ParamArrayCount(values)

        ReDim result(Max(count.Select(Function(x) x.RowsCount).ToArray) - 1, Max(count.Select(Function(x) x.ColumnsCount).ToArray) - 1)
        For i = 0 To result.GetLength(0) - 1
            For j = 0 To result.GetLength(1) - 1
                Dim resultElement As Boolean = True
                For Each k In values
                    If IsArray(k) AndAlso (i < k.GetLength(0) And j < k.GetLength(1)) Then resultElement = resultElement And k(i, j) Else resultElement = resultElement And k
                    If Not resultElement Then Exit For
                Next
                result(i, j) = resultElement
            Next
        Next
        Return result
    End Function

    <ExcelFunction>
    Public Function ArrayedOr(ParamArray values As ExcelVariant()) As ExcelVariant
        Dim result(,) As ExcelVariant
        Dim count = ParamArrayCount(values)

        ReDim result(Max(count.Select(Function(x) x.RowsCount).ToArray) - 1, Max(count.Select(Function(x) x.ColumnsCount).ToArray) - 1)
        For i = 0 To result.GetLength(0) - 1
            For j = 0 To result.GetLength(1) - 1
                Dim resultElement As Boolean = False
                For Each k In values
                    If IsArray(k) AndAlso (i < k.GetLength(0) And j < k.GetLength(1)) Then resultElement = resultElement Or k(i, j) Else resultElement = resultElement Or k
                    If resultElement Then Exit For
                Next
                result(i, j) = resultElement
            Next
        Next
        Return result
    End Function

    <ExcelFunction>
    Public Function ArrayedXor(ParamArray values As ExcelVariant()) As ExcelVariant
        Dim result(,) As ExcelVariant
        Dim count = ParamArrayCount(values)

        ReDim result(Max(count.Select(Function(x) x.RowsCount).ToArray) - 1, Max(count.Select(Function(x) x.ColumnsCount).ToArray) - 1)
        For i = 0 To result.GetLength(0) - 1
            For j = 0 To result.GetLength(1) - 1
                Dim resultElement As Boolean = values.First
                For k = 1 To values.Count - 1
                    If IsArray(values(k)) AndAlso (i < values(k).GetLength(0) And j < values(k).GetLength(1)) Then resultElement = resultElement Xor values(k)(i, j) Else resultElement = resultElement Xor values(k)
                Next
                result(i, j) = resultElement
            Next
        Next
        Return result
    End Function

    <ExcelFunction>
    Public Function ArrayedIf(expression As ExcelVariant(,), expressionIfTrue As ExcelVariant(,), expressionIfFalse As ExcelVariant(,)) As ExcelVariant
        Dim result(Max({expression.GetLength(0), expressionIfTrue.GetLength(0), expressionIfFalse.GetLength(0)}) - 1, Max({expression.GetLength(1), expressionIfTrue.GetLength(1), expressionIfFalse.GetLength(1)}) - 1) As ExcelVariant
        For i = 0 To result.GetLength(0) - 1
            For j = 0 To result.GetLength(1) - 1
                result(i, j) = If(If(expression.GetLength(0) > i + 1 And expression.GetLength(1) > j + 1, expression(i + 1, j + 1), ""),
                                  If(expressionIfTrue.GetLength(0) > i + 1 And expressionIfTrue.GetLength(1) > j + 1, expressionIfTrue(i + 1, j + 1), ""),
                                  If(expressionIfFalse.GetLength(0) > i + 1 And expressionIfFalse.GetLength(1) > j + 1, expressionIfFalse(i + 1, j + 1), "")
                                 )
            Next
        Next
        Return result
    End Function

    <ExcelFunction>
    Public Function ArrayedChoose(expressionConditions As ExcelVariant(,), expressionOptions As ExcelVariant(,)) As ExcelVariant
        Dim result As New List(Of ExcelVariant)
        For i = 0 To Min({expressionConditions.GetLength(0) - 1, expressionOptions.GetLength(0) - 1})
            For j = 0 To Min({expressionConditions.GetLength(1) - 1, expressionOptions.GetLength(1) - 1})
                If expressionConditions(i, j) Then result.Add(expressionOptions(i, j))
            Next
        Next
        Return result.ToArray
    End Function

    <ExcelFunction>
    Public Function CCRSplineAnalyze(xValues As ExcelNumber(,), yValues As ExcelNumber(,), command As ExcelString, ParamArray parameters As ExcelNumber()) As ExcelNumber
        Static memory As New CircularList(Of CatmullRomSpline)(256)
        Static [step] = 0.01
        Static alpha = 0.7
        Dim pl As New List(Of Numerics.Vector2)
        Dim cSpline As CatmullRomSpline
        Dim mf As Boolean = False
        For i = 0 To Min(xValues.GetLength(0), yValues.GetLength(0)) - 1
            For j = 0 To Min(xValues.GetLength(1), yValues.GetLength(1)) - 1
                If (IsNumeric(xValues(i, j)) Or IsDate(xValues(i, j))) And (IsNumeric(yValues(i, j)) Or IsDate(yValues(i, j))) Then _
                    pl.Add(New Numerics.Vector2(If(IsDate(xValues(i, j)), CDate(xValues(i, j)).ToOADate, xValues(i, j)), If(IsDate(yValues(i, j)), CDate(yValues(i, j)).ToOADate, yValues(i, j))))
            Next
        Next
        If pl.Count < 3 Then Return ExcelErrorNull
        Dim cIndex As Integer
        For cIndex = -memory.Count + 1 To 0
            If memory(cIndex).P.Count = pl.Count Then
                Dim j As Integer = 0
                Do
                    If memory(cIndex).P(0) <> pl(0) Then Continue For
                    j += 1
                Loop Until j > memory(cIndex).P.Count - 1
                mf = True
                Exit For
            End If
        Next
        If Not mf Then
            cSpline = New CatmullRomSpline(pl.ToArray, alpha)
            memory.MoveNext(cSpline)
        Else
            cSpline = memory(cIndex)
        End If
        Select Case command
            Case "GetX"
                Return cSpline.GetXValue(parameters(0), [step])
            Case "GetY"
                Return cSpline.GetYValue(parameters(0), [step])
            Case "GetYMax"
                Return cSpline.GetYMaxPlot([step]).Y
            Case "GetXInYMax"
                Return cSpline.GetYMaxPlot([step]).X
            Case "GetXMax"
                Return cSpline.GetXMaxPlot([step]).X
            Case "GetYInXMax"
                Return cSpline.GetXMaxPlot([step]).Y
            Case "GetXSeries"
                Return (From i In cSpline.GetAllPlots([step]) Select CObj(i.X)).ToArray
            Case "GetYSeries"
                Return (From i In cSpline.GetAllPlots([step]) Select CObj(i.Y)).ToArray
            Case "PlotCount"
                Return cSpline.GetAllPlots([step]).Count
            Case Else
                Return ExcelErrorValue
        End Select
    End Function

    <ExcelFunction>
    Public Function IntervalToString(intervalCode As String) As ExcelString
        Try
            Dim interval As New Interval(intervalCode)
            Return interval.ToString
        Catch
            Return ExcelErrorValue
        End Try
    End Function

    <ExcelFunction>
    Public Function IsInInterval(intervalCode As String, number As Double) As ExcelBoolean
        Try
            Dim interval As New Interval(intervalCode)
            Return interval.IsInInterval(number)
        Catch
            Return ExcelErrorValue
        End Try
    End Function

    <ExcelFunction>
    Public Function [For](start As Double, finish As Double, [step] As Double) As ExcelVariant()
        Dim result As New List(Of ExcelVariant)
        For i = start To finish Step [step]
            result.Add(i)
        Next
        Return result.ToArray
    End Function

    <ExcelFunction>
    Public Function Deduplicate(values As ExcelVariant(,)) As ExcelVariant()
        Dim result As New List(Of ExcelVariant)
        For Each i In values
            If Not result.Contains(i) And TypeOf i IsNot ExcelMissing Then result.Add(i)
        Next
        Return result.ToArray
    End Function

    <ExcelFunction(IsVolatile:=True)>
    Public Function Concat2(values As ExcelVariant(,), Optional charBetweenColumn As ExcelString = " ", Optional charBetweenRow As ExcelString = Chr(13), Optional isContainsEmpty As Boolean = False) As ExcelString
        If TypeOf charBetweenColumn Is ExcelMissing Then charBetweenColumn = " "
        If TypeOf charBetweenRow Is ExcelMissing Then charBetweenRow = Chr(13)
        Dim concatElements As New List(Of ConcatModelElement)
        Dim result As New Text.StringBuilder

        For i = values.GetLowerBound(0) To values.GetUpperBound(0)
            For j = values.GetLowerBound(1) To values.GetUpperBound(1)
                If TypeOf values(i, j) Is ExcelMissing Or (Not isContainsEmpty AndAlso (TypeOf values(i, j) Is ExcelEmpty OrElse CStr(values(i, j)) = "" OrElse (IsNumeric(values(i, j)) AndAlso Val(values(i, j)) = 0))) Then Continue For
                If concatElements.Count > 0 AndAlso Not concatElements.Last.IsColumnSplitter AndAlso Not concatElements.Last.IsRowSplitter Then concatElements.Add(New ConcatModelElement(charBetweenColumn, False, True))
                concatElements.Add(values(i, j).ToString)
            Next
            If concatElements.Count > 0 AndAlso Not concatElements.Last.IsColumnSplitter AndAlso Not concatElements.Last.IsRowSplitter Then concatElements.Add(New ConcatModelElement(charBetweenRow, True, False))
        Next
        Do While concatElements.Count > 0 AndAlso (concatElements.Last.IsRowSplitter Or concatElements.Last.IsColumnSplitter)
            concatElements.RemoveAt(concatElements.Count - 1)
        Loop
        For Each i In concatElements
            result.Append(i)
        Next
        Return result.ToString
    End Function

    <ExcelFunction>
    Public Function LinqWhere(<ExcelArgument(AllowReference:=True)> r As ExcelRange, expression As ExcelString) As ExcelRange
        If TypeOf r Is ExcelReference Then
            Dim _range = TrimRange(r)
            If TypeOf _range IsNot Excel.Range Then Return ExcelErrorValue
            Dim _result As Range
            For i = 1 To _range.Rows.Count
                For j = 1 To _range.Columns.Count
                    If CheckErrorCode(_range.Worksheet.Evaluate(CStr(expression).Replace("$$var", _range(i, j).Address(,, , True)))) Then
                        If _result Is Nothing Then _result = _range(i, j) Else _result = Application.Union(_result, _range(i, j))
                    End If
                Next
            Next
            Return ConvertToExcelReference(_result)
        ElseIf IsArray(r) Then
            Dim _r = DirectCast(r, Array)
            Dim _result As New List(Of ExcelVariant)
            For Each i In _r
                Dim _resultElement = CallerWorksheet.Evaluate(CStr(expression).Replace("$$var", """" & i & """"))
                If CheckErrorCode(If(TypeOf _resultElement Is Range, _resultElement.Value, _resultElement)) Then _result.Add(i)
            Next
            Return _result.ToArray
        Else
            If CheckErrorCode(CallerWorksheet.Evaluate(CStr(expression).Replace("$$var", """" & r & """"))) Then Return r Else Return ""
        End If
    End Function

    <ExcelFunction>
    Public Function LinqSelect(<ExcelArgument(AllowReference:=True)> r As ExcelRange, expression As ExcelString) As ExcelVariant
        If TypeOf r Is ExcelReference Then
            Dim _range = TrimRange(r)
            If TypeOf _range IsNot Excel.Range Then Return ExcelErrorValue
            Dim _result(_range.Rows.Count - 1, _range.Columns.Count - 1) As ExcelVariant
            For i = 1 To _range.Rows.Count
                For j = 1 To _range.Columns.Count
                    Dim _resultElement = _range.Worksheet.Evaluate(CStr(expression).Replace("$$var", _range(i, j).Address(,,, True)))
                    _result(i - 1, j - 1) = CheckErrorCode(If(TypeOf _resultElement Is Range, _resultElement.Value, If(TypeOf _resultElement Is Array, _resultElement(1), _resultElement)))
                Next
            Next
            Return _result
        ElseIf IsArray(r) Then
            Dim _r = DirectCast(r, Array)
            Dim _result As New List(Of ExcelVariant)
            For Each i In _r
                Dim _resultElement = CallerWorksheet.Evaluate(CStr(expression).Replace("$$var", """" & i & """"))
                _result.Add(CheckErrorCode(If(TypeOf _resultElement Is Range, _resultElement.Value, _resultElement)))
            Next
            Return _result.ToArray
        Else
            If CheckErrorCode(CallerWorksheet.Evaluate(CStr(expression).Replace("$$var", """" & r & """"))) Then Return r Else Return ""
        End If
    End Function

    Private objcache As New Dictionary(Of String, Object)
    <ExcelFunction>
    Public Function SetObjectCache(objectName As ExcelString, o As ExcelVariant) As ExcelNumber
        If objcache.ContainsKey(objectName) Then objcache(objectName) = o Else objcache.Add(objectName, o)
        Return 0
    End Function
    <ExcelFunction>
    Public Function GetObjectCache(objectName As ExcelString, objectSetter As ExcelVariant) As ExcelVariant
        If objcache.ContainsKey(objectName) Then Return objcache(objectName) Else Return ExcelErrorNull
    End Function

    ''Questionable
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

    '<ExcelFunction(Description:="Call a formula registered.", IsVolatile:=True, IsMacroType:=False)>
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
    '        Return Application.Evaluate(f)             '?????????????????????
    '    Else
    '        Return ExcelErrorName
    '    End If
    'End Function
End Module