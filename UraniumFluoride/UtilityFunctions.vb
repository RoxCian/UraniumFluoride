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

    Private Function IrregularValueHandler(value As ExcelVariant) As ExcelVariant
        Static IrregularValueArray As ExcelVariant() = {-6798283198, -2039484959.4}
        If IsNumeric(value) AndAlso IrregularValueArray.Contains(value) Then Return ExcelErrorValue
        Return value
    End Function

    <ExcelFunction(Description:="Banker round")>
    Public Function BankerRound(<MarshalAs(UnmanagedType.Currency)> num As Decimal, pre As Integer, Optional isSignificant As Boolean = False) As ExcelNumber
        If num = 0 Then Return 0
        If isSignificant Then
            Dim power As Integer = Math.Floor(Math.Log10(Math.Abs(num)))

            Return IrregularValueHandler(Math.Round(Math.Round(num / 10 ^ power, 14), pre - 1, MidpointRounding.ToEven) * 10 ^ power)
        Else
            Return IrregularValueHandler(Math.Round(Math.Round(num, 14), pre, MidpointRounding.ToEven))
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
    Public Function RandNoRepeat(bottom As Integer, top As Integer, Optional memorySet As Integer = 0, Optional memories As Integer = 30, Optional unrepeatPossibility As Integer = 0.99) As ExcelNumber
        Static memory As New ValueCircularListCollection(1024, 1024, Nothing)
        Dim x = memory
        Static randomer As New Random(Now.Millisecond), randomer2 As New Random(Now.Millisecond And Now.Millisecond)
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
                If memory(memorySet)(i) IsNot Nothing AndAlso memory(memorySet)(i) = result Then If randomer2.Next(0, 1000) / 1000 > unrepeatPossibility Then f = False
            Next
        Loop
        memory(memorySet).MoveNext(result)
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
        Return PageLocalize(page_Range.Worksheet.UsedRange, pageRow, pageColumn, cellRow, cellColumn, pageIndex)
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
        Dim valuesCount As Integer = Min(minArray.Count, maxArray.Count, stepArray.Count)
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
    Public Function RegExFind(text As String, pattern As String, Optional index As Integer = 1, Optional isCaseIgnore As Boolean = True) As ExcelNumber
        Dim e As New Text.RegularExpressions.Regex(pattern, If(isCaseIgnore, System.Text.RegularExpressions.RegexOptions.IgnoreCase, System.Text.RegularExpressions.RegexOptions.None))
        Dim m As System.Text.RegularExpressions.MatchCollection = e.Matches(text)
        If m.Count <= index - 1 Then Return -1 Else Return m(index - 1).Index + 1
    End Function

    <ExcelFunction>
    Public Function RegExMatch(text As String, pattern As String, Optional index As Integer = 1, Optional isCaseIgnore As Boolean = True) As ExcelString
        If index < 1 Then Return ExcelErrorNull
        Dim e As New Text.RegularExpressions.Regex(pattern, If(isCaseIgnore, System.Text.RegularExpressions.RegexOptions.IgnoreCase, System.Text.RegularExpressions.RegexOptions.None))
        Dim m As System.Text.RegularExpressions.MatchCollection = e.Matches(text)
        If m.Count <= index - 1 Then Return ExcelErrorNull Else Return m(index - 1).Value
    End Function

    <ExcelFunction(IsMacroType:=True)>
    Public Function RelativeReference(Optional rangeText As String = "A1", Optional path As String = "", Optional worksheetName As String = "") As ExcelRange
        Dim result = RelativeReferenceInternal(rangeText, path, worksheetName)
        If TypeOf result Is ClosedXML.Excel.IXLRange Then
            Dim r = CType(result, ClosedXML.Excel.IXLRange)
            'If r.FirstCell.NeedsRecalculation Then r.Worksheet.RecalculateAllFormulas()
            Return r.FirstCell.CachedValue
        Else
            Return result
        End If
    End Function

    Private Function RelativeReferenceInternal(Optional rangeText As String = "A1", Optional path As String = "", Optional worksheetName As String = "") As ExcelRange
        Dim wb As Workbook = Nothing
        If path = "" Then
            Try
                wb = Application.ThisWorkbook
            Catch ex As COMException
                wb = Application.ActiveWorkbook
            End Try
        Else
            If IO.File.Exists(path) Then
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
            Else Return ExcelErrorNa
            End If
        End If
        If worksheetName = "" Then Return ConvertToExcelReference(wb.Worksheets(1).Range(rangeText))
        For Each i As Worksheet In wb.Worksheets
            If i.Name = worksheetName Then Return ConvertToExcelReference(i.Range(rangeText))
        Next
        Return ExcelErrorNa
    End Function

    <ExcelFunction>
    Public Function ReferencedVLookUp(keyword As String, Optional rangeText As String = "A1", Optional path As String = "", Optional worksheetName As String = "", Optional targetColumn As Integer = 1, Optional isApproximateMatching As Boolean = False) As ExcelVariant
        Dim r = RelativeReferenceInternal(rangeText, path, worksheetName)
        If TypeOf r Is ClosedXML.Excel.IXLRange Then
            For Each i In CType(r, ClosedXML.Excel.IXLRange).FirstColumn.CellsUsed
                Dim t = If(TypeOf i.Value Is Date And IsNumeric(keyword), Date.FromOADate(keyword), keyword)
                If i.Value = t Then Return CType(r, ClosedXML.Excel.IXLRange).Cell(i.Address.RowNumber, targetColumn).Value
            Next
        Else
            Dim _range As Excel.Range = ConvertToRange(r)
            Return Application.VLookup(keyword, _range, targetColumn, isApproximateMatching)
        End If
        Return ExcelErrorNull
    End Function

    <ExcelFunction>
    Public Function ReferencedHLookUp(keyword As String, Optional rangeText As String = "A1", Optional path As String = "", Optional worksheetName As String = "", Optional targetRow As Integer = 1, Optional isApproximateMatching As Boolean = False) As ExcelVariant
        Dim r = RelativeReference(rangeText, path, worksheetName)
        If TypeOf r Is ClosedXML.Excel.IXLRange Then
            For Each i In CType(r, ClosedXML.Excel.IXLRange).FirstRow.CellsUsed
                If i.Value = keyword Then Return CType(r, ClosedXML.Excel.IXLRange).Cell(targetRow, i.Address.ColumnNumber).Value
            Next
        Else
            Dim _range As Excel.Range = ConvertToRange(r)
            Return Application.WorksheetFunction.HLookup(keyword, _range, targetRow, isApproximateMatching)
        End If
        Return ExcelErrorNull
    End Function

    <ExcelFunction>
    Public Function RegExMatchesCount(input As String, pattern As String, Optional startat As Integer = 0) As ExcelNumber
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
        If IsArray(arg) AndAlso arg.Count > 1 Then
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
    Dim log As New Text.StringBuilder
    <ExcelFunction(IsVolatile:=False, IsMacroType:=True)>
    Public Function CopyTextbox(<ExcelArgument(AllowReference:=True)> r As ExcelRange, textBoxName As String, Optional removeAllRecordedTextBoxes As Boolean = False, Optional left As Double = 0, Optional top As Double = 0, Optional width As Double = 0, Optional height As Double = 0) As ExcelVariant
        Static attachedObjects As New Dictionary(Of String, Shape)
        Dim x = attachedObjects
        Dim _Range As Excel.Range = ConvertToRange(r)
        Dim ws As Worksheet = _Range.Worksheet
        Dim r1st As Excel.Range = _Range(1, 1)
        If r1st.Address = "$BM$31" Then log.Append("executed;")
        If Not CType(_Range.Worksheet.Parent, Workbook).FullName = Application.ActiveSheet.Parent.Fullname OrElse Not _Range.Worksheet.CodeName = Application.ActiveSheet.Codename Then If attachedObjects.ContainsKey(r1st.Address) Then Return 0 Else Return "HANGED"

        If removeAllRecordedTextBoxes Then
            For o = 0 To attachedObjects.Count - 1
                attachedObjects.Values(o).Delete()
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(attachedObjects.Values(o))
            Next
            attachedObjects.Clear()
        End If

        Do While attachedObjects.ContainsKey(r1st.Address)
            Try
                attachedObjects(r1st.Address).Delete()
            Catch ex As COMException
            Finally
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(attachedObjects(r1st.Address))
                attachedObjects.Remove(r1st.Address)
            End Try
        Loop

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
            If Not (t.Width <= r1st.MergeArea.Width And t.Height <= r1st.MergeArea.Height) Then textBoxScale = Min(r1st.MergeArea.Width / t.Width, r1st.MergeArea.Height / t.Height)
        Catch ex As AccessViolationException
            textBoxScale = 1
        End Try
        Try
            t.Name = t.Name
        Catch ex As COMException
            Return "-"
        End Try
        Try
            nt = t.Duplicate
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
                s = Nothing
                f = True
                If Not continued Then Exit For
            End If
        Next
        If Not f Then Return "NOTHING TO REMOVE" Else Return 0
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
        If target.Split(separator).Count > index - 1 Then Return target.Split(separator)(index - 1) Else Return ExcelErrorNull
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
    '
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
