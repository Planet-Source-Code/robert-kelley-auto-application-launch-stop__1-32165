Attribute VB_Name = "modAPI"
Option Explicit
Dim objForm As Form                 ' Form Object Variable
Dim AppInfo As SHELLEXECUTEINFO     ' Place holder for Application Information
Public Weekdays(6) As Integer

'   API CALLS
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Boolean
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetLastError Lib "kernel32" () As Long

' CONSTANTS
Public Const WM_USER = &H400&
Public Const WM_CLOSE = &H10
Public Const WM_PAINT = &HF&
Public Const WM_STYLECHANGED = &H7D&
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWNORMAL = 1
Global Const SEE_MASK_NOCLOSEPROCESS = &H40
Global Const SEE_MASK_FLAG_DDEWAIT = &H100

' PRIVATE TYPES
Type SHELLEXECUTEINFO
        cbSize As Long
        fMask As Long
        hwnd As Long
        lpVerb As String
        lpFile As String
        lpParameters As String
        lpDirectory As String
        nShow As Long
        hInstApp As Long
        '  Optional fields
        lpIDList As Long
        lpClass As String
        hkeyClass As Long
        dwHotKey As Long
        hIcon As Long
        hProcess As Long
End Type

Public Function StartProcess(strFile As String)
Dim ret As Boolean      ' Return Value

    AppInfo.cbSize = Len(AppInfo)   ' Length
    AppInfo.lpFile = strFile ' File set from a common dialog box
    AppInfo.hwnd = frmMain.hwnd          ' HWnd of Calling Object
    AppInfo.lpVerb = "open"         ' String how to handle app
    AppInfo.nShow = SW_SHOWNORMAL   ' How do we show it?
    AppInfo.hProcess = 0            ' set Initial ProcessId
    AppInfo.fMask = SEE_MASK_NOCLOSEPROCESS
    
' Execute the Application
    ret = ShellExecuteEx(AppInfo)
    
End Function

Public Function KillProcess()
Dim ret As Boolean
    
' Kill the Process we opened earlier
    ret = TerminateProcess(AppInfo.hProcess, 0)
End Function

Sub Main()
' Create Instance of Form and Show it on load
    Set objForm = New frmMain
    objForm.Show
End Sub

Public Function CheckTime(strValue As String)
' Cruddy little function to see if we have a valid time format
    If Not InStr(strValue, ":") > 0 Or Len(Trim$(strValue)) <> 5 Then
        MsgBox "Invalid Time Format! Use ##:##", vbOKOnly, "Bad Time"
    Else
    End If
End Function

Public Function FillHelp(varBox As TextBox)
Dim HelpText As String
Dim strFile As String

' Open the Help Text File and fill the form
    Open App.Path & "\help.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, strFile
            HelpText = HelpText & strFile & vbCrLf
        Loop
    Close #1
        
varBox.Text = HelpText
End Function

Public Function CheckWeekDay() As Boolean
Dim intWeekDay As Integer

' Get the DayNumber of the week (IE: Monday = 1)
    intWeekDay = DatePart("w", Date)
    
' See if We run on this day
    If Weekdays(intWeekDay - 1) = 1 Then
        CheckWeekDay = True
    Else
        CheckWeekDay = False
    End If
    
    'Debug.Print Weekdays(intWeekDay - 1) & ":" & CheckWeekDay
End Function
