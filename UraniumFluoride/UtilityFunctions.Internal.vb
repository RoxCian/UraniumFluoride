Imports System.Runtime.InteropServices
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
#Region "Imports Macros"
Imports ExcelRange = System.Object
Imports ExcelNumber = System.Object
Imports ExcelLogical = System.Object
Imports ExcelDate = System.Object
Imports ExcelString = System.Object
Imports ExcelVariant = System.Object
#End Region

Partial Public Module UtilityFunctions

#Region "Range converting functions"
    Public Function ConvertToRange(ref As ExcelRange) As Excel.Range
        If TypeOf ref Is ExcelReference Then
            Dim refSheetFullName As String = XlCall.Excel(XlCall.xlSheetNm, ref)
            Dim refWorkbookName As String = Text.RegularExpressions.Regex.Match(refSheetFullName, "(?<=\[).*(?=\])").Value
            Dim refSheetName As String = Text.RegularExpressions.Regex.Match(refSheetFullName, "(?<=\[.*\]).*").Value
            Dim refAddress = GetExcelReferenceAddress_A1(ref)
            Return Application.Workbooks(refWorkbookName).Worksheets(refSheetName).Range(refAddress.TopLeft, refAddress.BottomRight)
        Else
            If TypeOf ref Is Excel.Range Then Return ref Else Return Nothing
        End If
    End Function
    Public Function ConvertToExcelReference(r As ExcelRange) As ExcelReference
        If TypeOf r Is Excel.Range Then Return New ExcelReference(r.Row - 1, r.Row - 1 + r.Rows.Count - 1, r.Column - 1, r.Column - 1 + r.Columns.Count - 1, CType(XlCall.Excel(XlCall.xlSheetId, "[" + r.Parent.Parent.Name + "]" + r.Worksheet.Name), ExcelReference).SheetId) Else If TypeOf r Is ExcelReference Then Return r Else Return Nothing
    End Function
    Public Function GetExcelReferenceAddress_A1(ref As ExcelReference) As (TopLeft As String, BottomRight As String)
        Dim r As (TopLeft As String, BottomRight As String)
        r.TopLeft = ColumnLetter(ref.ColumnFirst + 1) & (ref.RowFirst + 1)
        r.BottomRight = ColumnLetter(ref.ColumnLast + 1) & (ref.RowLast + 1)
        Return r
    End Function
#End Region

#Region "Internal implementation of some excel functions"
    Public Function TrimRange(r As ExcelRange) As Excel.Range
        Dim _Range As Excel.Range = ConvertToRange(r)
        Try
            Return Application.Intersect(_Range, _Range.Worksheet.UsedRange)
        Catch ex As COMException
            Return _Range
        Catch ex As NullReferenceException
            Return _Range
        End Try
    End Function
    Public Function TrimArray(value As ExcelVariant()) As ExcelVariant()
        Static result As New List(Of ExcelVariant)
        For Each i In value
            If TypeOf i IsNot ExcelEmpty Then result.Add(i)
        Next
        TrimArray = result.ToArray
        result.Clear()
    End Function
    Public Function TrimNumericArray(value As ExcelVariant()) As ExcelVariant()
        Static result As New List(Of ExcelVariant)
        For Each i In value
            If IsNumeric(i) Then result.Add(i) Else If IsDate(i) Then result.Add(CDate(i).ToOADate)
        Next
        TrimNumericArray = result.ToArray
        result.Clear()
    End Function

    Public Function RangeToArray(r As ExcelRange) As ExcelVariant()
        Dim _Range As Excel.Range = TrimRange(r)
        Dim result(_Range.Count - 1) As ExcelVariant

        Dim p As Integer = 0
        For i = 1 To _Range.Rows.Count
            For j = 1 To _Range.Columns.Count
                result(p) = _Range(i, j).Value
                p += 1
            Next
        Next
        Return result
    End Function
    Public Function RangeToMatrix(r As ExcelRange) As ExcelVariant(,)
        Dim _Range As Excel.Range = TrimRange(r)
        Dim result(_Range.Rows.Count - 1, _Range.Columns.Count - 1) As ExcelVariant
        For i = 1 To _Range.Rows.Count
            For j = 1 To _Range.Columns.Count
                result(i - 1, j - 1) = _Range(i, j).Value
            Next
        Next
        Return result
    End Function
    Public Function MatrixToArray(value As ExcelVariant(,)) As ExcelVariant()
        Static result As New List(Of ExcelVariant)
        For i = 0 To value.GetLength(0) - 1
            For j = 0 To value.GetLength(1) - 1
                result.Add(value(i, j))
            Next
        Next
        MatrixToArray = result.ToArray
        result.Clear()
    End Function

    Public Function Min(r As Excel.Range) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Return Min(RangeToArray(r))
    End Function
    Public Function Min(Of T)(ParamArray value() As T) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Dim result As Decimal = Decimal.MaxValue
        For Each i In value
            If IsNumeric(i) AndAlso CTypeDynamic(Of Decimal)(i) < result Then result = CTypeDynamic(Of Decimal)(i) Else If IsDate(i) AndAlso CTypeDynamic(Of Date)(i).ToOADate < result Then result = CTypeDynamic(Of Date)(i).ToOADate
        Next
        Return result
    End Function
    Public Function Max(r As Excel.Range) As <MarshalAs(UnmanagedType.Currency)> Decimal
        If IsArray(r) Then Return Max(r) Else Return Max(RangeToArray(r))
    End Function
    Public Function Max(Of T)(ParamArray value() As T) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Dim result As Decimal = Decimal.MinValue
        For Each i In value
            If IsNumeric(i) AndAlso CTypeDynamic(Of Decimal)(i) > result Then result = CTypeDynamic(Of Decimal)(i) Else If IsDate(i) AndAlso CTypeDynamic(Of Date)(i).ToOADate > result Then result = CTypeDynamic(Of Date)(i).ToOADate
        Next
        Return result
    End Function
    Public Function Med(r As Excel.Range) As <MarshalAs(UnmanagedType.Currency)> Decimal
        If IsArray(r) Then Return Med(r) Else Return Med(RangeToArray(r))
    End Function
    Public Function Med(Of T)(ParamArray value() As T) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Dim result As Decimal
        Dim substraction As Decimal = Decimal.MaxValue
        Dim average As Decimal = UtilityFunctions.Average(value)
        For Each i In value
            If IsNumeric(i) AndAlso Math.Abs(CTypeDynamic(Of Decimal)(i) - average) < substraction Then
                result = CTypeDynamic(Of Decimal)(i)
                substraction = Math.Abs(CTypeDynamic(Of Decimal)(i) - average)
            ElseIf IsDate(i) AndAlso Math.Abs(CTypeDynamic(Of Date)(i).ToOADate - average) < result Then
                result = CTypeDynamic(Of Date)(i).ToOADate
                substraction = Math.Abs(CTypeDynamic(Of Date)(i).ToOADate - average)
            End If
        Next
        Return result
    End Function
    Public Function Count(r As Excel.Range) As Integer
        Return Count(RangeToArray(r))
    End Function
    Public Function Count(Of T)(ParamArray value() As T) As Integer
        Dim result As Integer
        For Each i In value
            If Not IsBlank(i) And Not IsError(i) And (IsDate(i) Or IsNumeric(i)) Then result += 1
        Next
        Return result
    End Function
    Public Function ParamArrayCount(ParamArray values()) As (RowsCount As Integer, ColumnsCount As Integer)()
        Dim count(values.Count - 1) As (RowsCount As Integer, ColumnsCount As Integer)
        For i = 0 To values.Count - 1
            If IsArray(values(i)) Then count(i) = (CType(values(i), Array).GetLength(0), CType(values(i), Array).GetLength(1))
        Next
        Return count
    End Function
    Public Function Sum(r As ExcelRange) As <MarshalAs(UnmanagedType.Currency)> Decimal
        If IsArray(r) Then Return Sum(r) Else Return Sum(RangeToArray(r))
    End Function
    Public Function Sum(Of T)(ParamArray value() As T) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Dim result As Decimal
        For Each i In value
            If IsNumeric(i) Then result += CTypeDynamic(Of Decimal)(i) Else If IsDate(i) Then result += CTypeDynamic(Of Date)(i).ToOADate
        Next
        Return result
    End Function
    Public Function Average(r As ExcelRange) As Decimal
        Dim c = Count(r)
        If c = 0 Then Return 0 Else Return Sum(r) / c
    End Function
    Public Function Average(ParamArray value()) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Dim c = Count(value)
        If c = 0 Then Return 0 Else Return Sum(value) / c
    End Function
    Public Function GetNumeric(r As ExcelRange) As Decimal()
        Return GetNumeric(RangeToArray(r))
    End Function
    Public Function GetNumeric(ParamArray value()) As Decimal()
        Dim result As New List(Of Decimal)
        For Each i In value
            If IsNumeric(i) Then result.Add(i) Else If IsDate(i) Then result.Add(CDate(i).ToOADate)
        Next
        Return result.ToArray
    End Function
    Public Function GetNumericArrayHash(r As ExcelRange) As Integer
        Return GetNumericArrayHash(GetNumeric(r))
    End Function
    Public Function GetNumericArrayHash(ParamArray value() As ExcelVariant) As Integer
        Dim result As Integer = 0
        For Each i In value
            If IsNumeric(i) Then result = result Xor i.GetHashCode
        Next
        Return result
    End Function
    Public Function IsBlank(value) As Boolean
        If IsNumeric(value) Then Return value = 0
        If IsDate(value) Then Return value = New Date("1899/12/31")
        If TypeOf value Is Range Then Return Application.WorksheetFunction.IsBlank(value)
        If TypeOf value Is ExcelDna.Integration.ExcelEmpty Then Return True
        Return False
    End Function
    Public Function IsError(value) As Boolean
        If TypeOf value Is Range Then Return Application.WorksheetFunction.IsError(value)
        If TypeOf value Is ExcelDna.Integration.ExcelError Then Return True
        Return False
    End Function
#End Region

#Region "Miscellaneous"
    Public Sub Swap(Of T)(ByRef obj1 As T, ByRef obj2 As T)
        Dim obj3 As T
        obj3 = obj1
        obj1 = obj2
        obj2 = obj3
    End Sub
#End Region
End Module
