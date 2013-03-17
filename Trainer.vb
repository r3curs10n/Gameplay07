Public Class Trainer

    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal Classname As String, ByVal WindowName As String) As IntPtr
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As IntPtr
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As IntPtr) As Integer
    Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As IntPtr, ByVal lpBaseAddress As Integer, ByRef lpBuffer As Integer, ByVal nSize As Short, ByVal lpNumberOfBytesWritten As Integer) As Integer
    Private Declare Function WriteProcessMemory1 Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As IntPtr, ByVal lpBaseAddress As Integer, ByRef lpBuffer As Single, ByVal nSize As Short, ByVal lpNumberOfBytesWritten As Integer) As Integer
    Private Declare Function WriteProcessMemory Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As IntPtr, ByVal lpBaseAddress As Integer, ByRef lpBuffer As Int32, ByVal nSize As Short, ByVal lpNumberOfBytesWritten As Integer) As Integer
    Private Const PROCESS_VM_ALL As Integer = &H1F0FFF
    Private Const PROCESS_VM_READ As Integer = &H10
    Private Const PROCESS_VM_WRITE As Integer = &H20
    Private Const PROCESS_VM_OPERATION As Integer = &H8

    Dim hWnd As IntPtr, pHandle As IntPtr, val As Int32, val1 As Single
    Dim processID As Integer, WindowText As String

    Sub New(ByVal wt As String)
        WindowText = wt
    End Sub

    Function GetSValue(ByVal Address As Integer) As Integer

        hWnd = FindWindow(vbNullString, WindowText)
        If Not (hWnd = 0) Then
            GetWindowThreadProcessId(hWnd, processID)
            pHandle = OpenProcess(PROCESS_VM_READ + PROCESS_VM_OPERATION, False, processID)
            If Not (pHandle = 0) Then
                ReadProcessMemory(pHandle, Address, val, 4, 0&)
            End If
            CloseHandle(pHandle)

        End If
        Return val
    End Function

    Sub SetSValue(ByVal Address As Integer, ByVal Buffer As Int32)

        hWnd = FindWindow(vbNullString, WindowText)
        If Not (hWnd = 0) Then
            GetWindowThreadProcessId(hWnd, processID)
            pHandle = OpenProcess(PROCESS_VM_WRITE + PROCESS_VM_OPERATION, False, processID)
            If Not (pHandle = 0) Then
                WriteProcessMemory(pHandle, Address, Buffer, 4, 0&)
            End If
            CloseHandle(pHandle)
        End If
    End Sub

    Function OpenProcess(ByVal write As Boolean) As Boolean

        hWnd = FindWindow(vbNullString, WindowText)
        GetWindowThreadProcessId(hWnd, processID)
        If write = False Then
            pHandle = OpenProcess(PROCESS_VM_READ + PROCESS_VM_OPERATION, False, processID)
        Else
            pHandle = OpenProcess(PROCESS_VM_WRITE + PROCESS_VM_OPERATION, False, processID)
            If pHandle = IntPtr.Zero Then
                Return False
                Exit Function
            End If
        End If
        Return True
    End Function

    Sub SetValue(ByVal Address As Integer, ByVal Buffer As Single)

        WriteProcessMemory1(pHandle, Address, Buffer, 4, 0&)

    End Sub


    Sub CloseProcess()

        CloseHandle(pHandle)

    End Sub

End Class
