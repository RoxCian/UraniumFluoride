Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
#Region "Imports Macros"
Imports ExcelRange = System.Object
Imports ExcelDecimal = System.Object
Imports ExcelInteger = System.Object
Imports ExcelString = System.Object
Imports ExcelVariant = System.Object
#End Region

Partial Public Module UtilityFunctions
    <ExcelFunction(Description:="Banker round, obsoletted", IsVolatile:=True)>
    Public Function BKROUND(<MarshalAs(UnmanagedType.Currency)> num As Decimal, pre As Integer) As ExcelDecimal
        Return BankerRound(num, pre, False)
    End Function
    <ExcelFunction(Description:="Banker round for significant number usage, obsoletted", IsVolatile:=True)>
    Public Function BKROUNDEFFECTIVE(<MarshalAs(UnmanagedType.Currency)> num As Decimal, pre As Integer) As ExcelDecimal
        Return BankerRound(num, pre, True)
    End Function
    <ExcelFunction(Description:="Average of a group of numbers without element which is 10% to mean value, obsoletted", IsVolatile:=True, IsMacroType:=True)>
    Public Function AVERAGE10(<ExcelArgument(AllowReference:=True)> nums As ExcelRange) As ExcelDecimal
        Return AverageByMean(nums)
    End Function
    <ExcelFunction>
    Public Function CHECKER10(<ExcelArgument(AllowReference:=True)> nums As ExcelRange) As ExcelInteger
        Return VerifyByMean(nums)
    End Function
    <ExcelFunction>
    Public Function AVERAGE15BYMEDIAN(<ExcelArgument(AllowReference:=True)> nums As ExcelRange) As ExcelDecimal
        Return AverageByMedian(nums)
    End Function
    <ExcelFunction>
    Public Function PAGELOCALIZER(<ExcelArgument(AllowReference:=True)> r As Range, pageRowsCount As Integer, pageColumnsCount As Integer, locationRow As Integer, locationColumn As Integer, index As Integer) As ExcelVariant
        Return PageLocalize(r, pageRowsCount, pageColumnsCount, locationRow, locationColumn, index)
    End Function
    <ExcelFunction>
    Public Function REGEXPFIND(text As String, pattern As String, Optional index As Integer = 1, Optional isCaseIgnore As Boolean = True) As ExcelInteger
        Return RegExFind(text, pattern, index, isCaseIgnore)
    End Function
    <ExcelFunction>
    Public Function REGEXPMATCH(text As String, pattern As String, Optional index As Integer = 1, Optional isCaseIgnore As Boolean = True) As ExcelString
        Return RegExMatch(text, pattern, index, isCaseIgnore)
    End Function
    <ExcelFunction>
    Public Function FILERELATIVEREFERENCE(workbookPath As String, Optional worksheetName As String = "", Optional rangeText As String = "A1") As ExcelRange
        Return ConvertToExcelReference(RelativeReference(worksheetName, rangeText, workbookPath))
    End Function
    <ExcelFunction>
    Public Function STRINGCOUNT(s As String, p As String) As Integer
        Return RegExMatchesCount(s, p, 0)
    End Function
End Module
