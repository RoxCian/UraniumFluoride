Imports System.Runtime.InteropServices
Imports ExcelDna.Integration
Imports ExcelDna.Registration

Public Class AddIn
    Implements IExcelAddIn

    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        ExcelIntegration.RegisterUnhandledExceptionHandler(Function(ex) "#EXCEPTION--We're sorry, but here is an unhandled exception: " & ex.ToString & "@ Uranium Fluoride")
        ExcelRegistration.GetExcelFunctions.ProcessParamsRegistrations.RegisterFunctions

    End Sub

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose
    End Sub
End Class