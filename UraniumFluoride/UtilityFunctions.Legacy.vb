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

Partial Public Module UtilityFunctions
    <ExcelFunction(Description:="Banker round, for compatibility of old formulas and codes", IsVolatile:=True)>
    Public Function BKROUND(<MarshalAs(UnmanagedType.Currency)> num As Decimal, pre As Integer) As ExcelNumber
        Return BankerRound(num, pre, False)
    End Function
    <ExcelFunction(Description:="Banker round for significant number usage, for compatibility of old formulas and codes", IsVolatile:=True)>
    Public Function BKROUNDEFFECTIVE(<MarshalAs(UnmanagedType.Currency)> num As Decimal, pre As Integer) As ExcelNumber
        Return BankerRound(num, pre, True)
    End Function
    <ExcelFunction(Description:="Average of a group of numbers without element which is 10% to mean value, for compatibility of old formulas and codes", IsVolatile:=True, IsMacroType:=True)>
    Public Function AVERAGE10(<ExcelArgument(AllowReference:=True)> num As ExcelVariant(,)) As ExcelNumber
        Return AverageByMean(num)
    End Function
    <ExcelFunction(IsMacroType:=True)>
    Public Function CHECKER10(<ExcelArgument(AllowReference:=True)> num As ExcelVariant(,)) As ExcelNumber
        Return VerifyByMean(num)
    End Function
    <ExcelFunction(IsMacroType:=True)>
    Public Function AVERAGE15BYMEDIAN(<ExcelArgument(AllowReference:=True)> num As ExcelVariant(,)) As ExcelNumber
        Return AverageByMedian(num)
    End Function
    <ExcelFunction(IsMacroType:=True, IsVolatile:=True)>
    Public Function PAGELOCALIZER(<ExcelArgument(AllowReference:=True)> r As ExcelRange, pageRowsCount As Integer, pageColumnsCount As Integer, locationRow As Integer, locationColumn As Integer, index As Integer) As ExcelVariant
        Return PageLocalize(r, pageRowsCount, pageColumnsCount, locationRow, locationColumn, index)
    End Function
    <ExcelFunction>
    Public Function REGEXPFIND(text As String, pattern As String, Optional index As Integer = 1, Optional isCaseIgnore As Boolean = True) As ExcelNumber
        Return RegExFind(text, pattern, index, isCaseIgnore)
    End Function
    <ExcelFunction>
    Public Function REGEXPMATCH(text As String, pattern As String, Optional index As Integer = 1, Optional isCaseIgnore As Boolean = True) As ExcelString
        Return RegExMatch(text, pattern, index, isCaseIgnore)
    End Function
    <ExcelFunction(IsMacroType:=True)>
    Public Function FILERELATIVEREFERENCE(workbookPath As String, Optional worksheetName As String = "", Optional rangeText As String = "A1") As ExcelRange
        Dim wb As Workbook = Nothing
        If workbookPath = "" Then
            Try
                wb = Application.ThisWorkbook
            Catch ex As COMException
                wb = Application.ActiveWorkbook
            End Try
        Else
            If IO.File.Exists(workbookPath) Then
                Dim wbc = From currentWb As Workbook In Application.Workbooks Where currentWb.Name = workbookPath.Split("\").Last Select currentWb
                If wbc.Count > 0 Then
                    wb = wbc.First
                    If wb.FullName <> workbookPath Then Return ExcelErrorNa
                Else
                    wb = Application.Workbooks.Open(workbookPath)
                    For Each i As Window In wb.Windows
                        i.Visible = False
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
    Public Function STRINGCOUNT(s As String, p As String) As Integer
        Return RegExMatchesCount(s, p, 0)
    End Function
    <ExcelFunction>
    Public Function FIND2(findText As String, withinText As String, Optional startNum As Integer = 1) As ExcelNumber
        If withinText.Contains(findText) Then Return withinText.IndexOf(findText, startNum) + 1 Else Return -1
    End Function
    <ExcelFunction>
    Public Function FINDB2(findText As String, withinText As String, Optional startNum As Integer = 1) As ExcelNumber
        If withinText.Contains(findText) Then Return System.Text.Encoding.Default.GetByteCount(withinText.Remove(withinText.IndexOf(findText, startNum))) + 1 Else Return -1
    End Function
End Module
