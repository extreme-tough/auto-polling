Attribute VB_Name = "Module1"
Option Explicit
    'Add the following code to a module
Private Const INFINITE As Long = &HFFFFFFFF
Private Const SEE_MASK_FLAG_NO_UI As Long = &H400
Private Const SEE_MASK_NOCLOSEPROCESS As Long = &H40


Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
    End Type

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function WaitForSingleObject Lib "Kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long


Public Declare Function GetLastError Lib "Kernel32.dll" () As Long


Public Declare Function ShellExecuteEx Lib "Shell32.dll" (ByRef lpExecInfo As SHELLEXECUTEINFO) As Long


Public Function ShellExecuteWait(ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

    Dim lReturn As Long, lResult As Long
    Dim tExecuteInfo As SHELLEXECUTEINFO
    'Fill the SHELLEXECUTEINFO structure
    tExecuteInfo.cbSize = Len(tExecuteInfo)
    tExecuteInfo.fMask = SEE_MASK_NOCLOSEPROCESS
    tExecuteInfo.hWnd = hWnd
    tExecuteInfo.lpVerb = lpOperation
    tExecuteInfo.lpFile = lpFile
    tExecuteInfo.lpParameters = lpParameters
    tExecuteInfo.lpDirectory = lpDirectory
    tExecuteInfo.nShow = nShowCmd
    'Call the API with the specified paramet
    '     ers
    lReturn = ShellExecuteEx(tExecuteInfo)
    If lReturn = 0 Then lReturn = GetLastError Else lReturn = tExecuteInfo.hInstApp
    'If there's a new process wait while it
    '     terminates


    If tExecuteInfo.hProcess <> 0 Then
        lResult = WaitForSingleObject(tExecuteInfo.hProcess, INFINITE)
    End If

    'Return the ShellExecuteEx return value
    ShellExecuteWait = lReturn
End Function

'And the following code to a form. Also,
'     you must add a CommandButton named "Comm
'     and1"


Public Sub Example()
    Dim hWnd As Long
    MsgBox "gg"
    Call ShellExecuteWait(hWnd, "open", "C:\Windows\Notepad.exe", "", "", vbNormalFocus)
End Sub
