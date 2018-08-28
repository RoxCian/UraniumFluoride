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

#Region "Internal helper functions"
    Public Function ConvertToRange(ref As ExcelRange) As Excel.Range
        If TypeOf ref Is ExcelReference Then Return Application.Range(XlCall.Excel(XlCall.xlfReftext, ref, True)) Else If TypeOf ref Is Excel.Range Then Return ref Else Return Nothing
    End Function
    Public Function ConvertToExcelReference(r As ExcelRange) As ExcelReference
        If TypeOf r Is Excel.Range Then Return New ExcelReference(r.Row - 1, r.Row - 1 + r.Rows.Count - 1, r.Column - 1, r.Column - 1 + r.Columns.Count - 1, CType(XlCall.Excel(XlCall.xlSheetId, "[" + r.Parent.Parent.Name + "]" + r.Worksheet.Name), IntPtr)) Else Return New ExcelReference(0, 0)
    End Function
    Public Sub Swap(Of T)(ByRef obj1 As T, ByRef obj2 As T)
        Dim obj3 As T
        obj3 = obj1
        obj1 = obj2
        obj2 = obj3
    End Sub
#End Region

#Region "Internal implementation of some excel functions"
    Public Function Min(r As Excel.Range) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Return Min(RangeToValueArray(r))
    End Function
    Public Function Min(r As ExcelRange) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Return Min(RangeToValueArray(ConvertToRange(r)))
    End Function
    Public Function Min(ParamArray value()) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Dim result As Decimal = Decimal.MaxValue
        For Each i In value
            If IsNumeric(i) AndAlso i < result Then result = i Else If IsDate(i) AndAlso CDate(i).ToOADate < result Then result = CDate(i).ToOADate
        Next
        Return result
    End Function
    Public Function Max(r As ExcelRange) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Return Max(RangeToValueArray(r))
    End Function
    Public Function Max(ParamArray value()) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Dim result As Decimal = Decimal.MinValue
        For Each i In value
            If IsNumeric(i) AndAlso i > result Then result = CDec(i) Else If IsDate(i) AndAlso CDate(i).ToOADate > result Then result = CDate(i).ToOADate
        Next
        Return result
    End Function
    Public Function Med(r As ExcelRange) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Return Med(RangeToValueArray(r))
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
        Return Count(RangeToValueArray(r))
    End Function
    Public Function Count(ParamArray value()) As Integer
        Dim result As Integer
        For Each i In value
            If IsNumeric(i) Or IsDate(i) Then result += 1
        Next
        Return result
    End Function
    Public Function Sum(r As ExcelRange) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Return Sum(RangeToValueArray(r))
    End Function
    Public Function Sum(ParamArray value()) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Dim result As Decimal
        For Each i In value
            If IsNumeric(i) Then result += CDec(i) Else If IsDate(i) Then result += CDate(i).ToOADate
        Next
        Return result
    End Function
    Public Function Average(r As ExcelRange) As Decimal
        If Count(r) = 0 Then Return 0 Else Return Sum(r) / Count(r)
    End Function
    Public Function Average(ParamArray value()) As <MarshalAs(UnmanagedType.Currency)> Decimal
        Return Sum(value) / Count(value)
    End Function
    Public Function GetNumeric(r As ExcelRange) As Decimal()
        Return GetNumeric(RangeToValueArray(r))
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
    Public Function GetNumericArrayHash(ParamArray value() As Decimal) As Integer
        Dim result As Integer = 0
        For Each i In value
            result = result Xor i.GetHashCode
        Next
        Return result
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
End Module
