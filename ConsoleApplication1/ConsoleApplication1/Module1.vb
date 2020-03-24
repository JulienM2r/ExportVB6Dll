Imports System.Runtime.InteropServices
Imports RGiesecke.DllExport

Module Module1
    Dim txbValueReturned As String

    <DllExport("fn_calc")>
    Public Function Fn_calc(a As Integer, b As Integer) As Integer
        Return a + b
    End Function

    <DllExport("testmethod")>
    Public Sub Testmethod()
        txbValueReturned = FunctionCalled("Hello")
    End Sub

    <DllExport("getData")>
    Public Function GetData() As String
        Return txbValueReturned
    End Function

    Public Function FunctionCalled _
    (ByVal strValuePassed As String) As String
        FunctionCalled = strValuePassed + "_changed_in_DLL___^^"
    End Function

    Sub Main()
        testmethod()
    End Sub
End Module
