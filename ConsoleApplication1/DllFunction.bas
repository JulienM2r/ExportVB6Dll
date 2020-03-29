Attribute VB_Name = "Module1"
Option Explicit

Public Const DLL_PROCESS_DETACH = 0
Public Const DLL_PROCESS_ATTACH = 1
Public Const DLL_THREAD_ATTACH = 2
Public Const DLL_THREAD_DETACH = 3

Public txbValueReturned As String

Public Function DllMain(hInst As Long, fdwReason As Long, lpvReserved As Long) As Boolean
    Select Case fdwReason
        Case DLL_PROCESS_DETACH
        ' No per-process cleanup needed

        Case DLL_PROCESS_ATTACH
            DllMain = True

        Case DLL_THREAD_ATTACH
            ' No per-thread initialization needed

        Case DLL_THREAD_DETACH
        ' No per-thread cleanup needed

    End Select
End Function

    '<DllExport("fn_calc")>
    Public Function Fn_calc(ByVal a As Integer, ByVal b As Integer) As Integer
        Fn_calc = a + b
        Debug.Print Fn_calc
    End Function

    '<DllExport("testmethod")>
    Private Function Testmethodint() As Integer
'        Dim tryS As String
'        tryS = "Hello"
        txbValueReturned = "Hello" 'Utf8BytesFromString(tryS)
        'Debug.Print txbValueReturned
'        FunctionCalled tryS
'        Debug.Print TypeName(tryS)
        Testmethodint = 10
    End Function

    Public Function Testmethod() As Integer
       
       
       Testmethod = Testmethodint()
        
    End Function

    '<DllExport("getData")>
    Public Function GetData() As String
        GetData = txbValueReturned   'Utf8BytesToString(txbValueReturned)
        Debug.Print GetData
    End Function

    Public Function FunctionCalled(ByRef strValuePassed() As Variant) As String
        Dim b() As Byte
        b = Utf8BytesFromString("Hello")
        Debug.Print TypeName(b)
        FunctionCalled = Utf8BytesToString(b) 'dead
        Debug.Print TypeName(FunctionCalled)
        
        'FunctionCalled = transString(strValuePassed) 'dead
        'FunctionCalled = 25 'ok
        'FunctionCalled = StrConv(strValuePassed, vbFromUnicode) 'dead+
        'FunctionCalled = StrConv(strValuePassed, vbUnicode) 'dead
        'FunctionCalled = Utf8BytesToString(strValuePassed)
    End Function

    
    
    

