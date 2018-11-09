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
        If TypeOf ref Is ExcelReference Then Return Application.Range(XlCall.Excel(XlCall.xlfReftext, ref, True)) Else If TypeOf ref Is Excel.Range Then Return ref Else Return Nothing
    End Function
    Public Function ConvertToExcelReference(r As ExcelRange) As ExcelReference
        If TypeOf r Is Excel.Range Then Return New ExcelReference(r.Row - 1, r.Row - 1 + r.Rows.Count - 1, r.Column - 1, r.Column - 1 + r.Columns.Count - 1, CType(XlCall.Excel(XlCall.xlSheetId, "[" + r.Parent.Parent.Name + "]" + r.Worksheet.Name), ExcelReference).SheetId) Else If TypeOf r Is ExcelReference Then Return r Else Return New ExcelReference(0, 0)
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

    Public Function RangeToArray(<ExcelArgument(AllowReference:=True)> r As ExcelRange) As ExcelVariant()
        Dim _Range As Excel.Range = TrimRange(r)
        Dim result(_Range.Count - 1) As ExcelVariant

        Dim p As Integer = 0
        For i = 1 To _Range.Rows.Count
            For j = 1 To _Range.Columns.Count
                result(p) = _Range(i, j).Value
                p = p + 1
            Next
        Next
        Return result
    End Function
    Public Function RangeToMatrix(<ExcelArgument(AllowReference:=True)> r As ExcelRange) As ExcelVariant(,)
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
    Public Function Min(r As ExcelRange) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Return Min(RangeToArray(ConvertToRange(r)))
    End Function
    Public Function Min(ParamArray value()) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Dim result As Decimal = Decimal.MaxValue
        For Each i In value
            If IsNumeric(i) AndAlso i < result Then result = i Else If IsDate(i) AndAlso CDate(i).ToOADate < result Then result = CDate(i).ToOADate
        Next
        Return result
    End Function
    Public Function Max(r As ExcelRange) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Return Max(RangeToArray(r))
    End Function
    Public Function Max(ParamArray value()) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Dim result As Decimal = Decimal.MinValue
        For Each i In value
            If IsNumeric(i) AndAlso i > result Then result = CDec(i) Else If IsDate(i) AndAlso CDate(i).ToOADate > result Then result = CDate(i).ToOADate
        Next
        Return result
    End Function
    Public Function Med(r As ExcelRange) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Return Med(RangeToArray(r))
    End Function
    Public Function Med(ParamArray value()) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Dim result As Decimal
        Dim substraction As Decimal = Decimal.MaxValue
        Dim average As Decimal = UtilityFunctions.Average(value)
        For Each i In value
            If IsNumeric(i) AndAlso Math.Abs(i - average) < substraction Then
                result = CDec(i)
                substraction = Math.Abs(i - average)
            ElseIf IsDate(i) AndAlso Math.Abs(CDate(i).ToOADate - average) < result Then
                result = CDate(i).ToOADate
                substraction = Math.Abs(CDate(i).ToOADate - average)
            End If
        Next
        Return result
    End Function
    Public Function Count(r As ExcelRange) As Integer
        Return Count(RangeToArray(r))
    End Function
    Public Function Count(ParamArray value()) As Integer
        Dim result As Integer
        For Each i In value
            If Not IsBlank(i) And Not IsError(i) And (IsDate(i) Or IsNumeric(i)) Then result += 1
        Next
        Return result
    End Function
    Public Function Sum(r As ExcelRange) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Return Sum(RangeToArray(r))
    End Function
    Public Function Sum(ParamArray value()) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Dim result As Decimal
        For Each i In value
            If IsNumeric(i) Then result += CDec(i) Else If IsDate(i) Then result += CDate(i).ToOADate
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
    Private Function GetShape(shapeName As String, Optional worksheetName As String = "") As Shape
        Dim ws As Worksheet = Nothing
        If worksheetName = "" Then
            Application.ThisWorkbook.Activate()
            ws = Application.ActiveSheet
        Else
            Dim a As Worksheet
            For Each a In Application.ThisWorkbook.Worksheets
                If a.Name = worksheetName Then ws = a
            Next
        End If
        If ws Is Nothing Then Return Nothing
        Dim s As Excel.Shape = Nothing
        Dim f As Boolean
        f = False
        For Each s In ws.Shapes
            If s.Name = shapeName Then
                f = True
                Exit For
            End If
        Next
        If f Then Return s Else Return Nothing
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
