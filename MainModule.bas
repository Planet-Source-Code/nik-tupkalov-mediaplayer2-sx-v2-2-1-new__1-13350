Attribute VB_Name = "MainModule"
Option Explicit

Public Type NotifyIconData
    cbSize          As Long
    hwnd            As Long
    uID             As Long
    uFlags          As Long
    uCallBackMsg    As Long
    hIcon           As Long
    szTip           As String * 64
End Type

Public Const STI_ADD = 0
Public Const STI_MODIFY = 1
Public Const STI_DELETE = 2
Public Const STI_MESSAGE = 1
Public Const STI_ICON = 2
Public Const STI_TIP = 4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const HKCR = &H80000000
Public Const REG_SZ = 1                          ' Unicode nul terminated string

Public Enum SysTrayAction
    AddIcon = STI_ADD
    ModifyIcon = STI_MODIFY
    DeleteIcon = STI_DELETE
End Enum

Declare Function SysTrayIcon Lib "SHELL32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NotifyIconData) As Long
Declare Function OSRegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function OSRegCloseKey Lib "advapi32.dll" Alias "RegCloseKey" (ByVal hKey As Long) As Long
Declare Function OSRegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function OSRegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function OSRegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Const WM_SYSCOMMAND = &H112
Public Const SC_MOVE = &HF012

Public RegLastError As Long, atStart As Boolean
Private KeyPath$, hHnd&, Temp&, TempEx&
Private TempExB&, TempExA$, TempExC%
Public guIconData As NotifyIconData

Sub Main()

Static strSubKey As String
    strSubKey = "Files " & App.Title

If RegReadString(HKCR, ".mpc", "", "") <> strSubKey Then
    RegCreateKey HKCR, ".mpc"
        RegWriteString HKCR, ".mpc", "", "", strSubKey
            RegCreateKey HKCR, strSubKey, "shell\open\command"
        RegWriteString HKCR, strSubKey, "", "", "Files catalog MediaPlayer 2.0 SX"
    RegWriteString HKCR, strSubKey, "shell\open", "", "Play for MediaPlayer 2.0 SX"
RegWriteString HKCR, strSubKey, "shell\open\command", "", Chr(34) _
               & App.Path & "\" & App.Title & ".exe" & Chr(34) & " %1"
     RegCreateKey HKCR, strSubKey, "DefaultIcon"
        RegWriteString HKCR, strSubKey, "DefaultIcon", "", _
                       App.Path & "\" & App.Title & ".exe,5"
End If
Load Remote
Remote.Show
If atStart Then Remote.mnuSysTray_Click
End Sub

Public Function RegCreateKey(ByVal hKey As Long, ByVal Key As String, Optional SubKey As Variant) As Boolean
    
    ' Create Key If It Doesn't Exist
    If Not IsMissing(SubKey) Then
        Temp& = OSRegCreateKey(hKey, Key & "\" & SubKey, hHnd&)
    Else
        Temp& = OSRegCreateKey(hKey, Key, hHnd&)
    End If
    
    ' Process Returned Information
    If RegCheckError(Temp&) Then GoTo CreateKeyError
   
    ' Close Handle To Key
    Temp& = OSRegCloseKey(hHnd&)
    
    ' Operation Was Successful
    RegCreateKey = -1

    ' Exit Function With Passed Value
    Exit Function

CreateKeyError:
    
    ' Store Error In Variable
    RegLastError = Temp&
    
    ' Operation Was Not Successful
    RegCreateKey = 0
    
    ' Close Handle To Key
    Temp& = OSRegCloseKey(hHnd&)
    
End Function

Public Function RegWriteString(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String, ByVal Value As String) As Boolean

    ' Combine The Key And SubKey Paths
    If Not SubKey = "" Then KeyPath$ = _
    Key + "\" + SubKey Else KeyPath$ = Key
    
    ' Create Key If It Doesn't Exist
    Temp& = OSRegCreateKey(hKey, KeyPath$, hHnd&)
    
    ' Process Returned Information
    If RegCheckError(Temp&) Then GoTo WriteStringError
    
    ' Set New Value For The Opened Key
    Temp& = OSRegSetValueEx(hHnd&, ValueName, 0&, REG_SZ, ByVal Value, Len(Value))
     
     ' Process Returned Information
    If RegCheckError(Temp&) Then GoTo WriteStringError

    ' Close Handle To Key
    Temp& = OSRegCloseKey(hHnd&)

    ' Operation Was Successful
    RegWriteString = -1

    ' Exit Function With Passed Value
    Exit Function

WriteStringError:
    
    ' Store Error In Variable
    RegLastError = Temp&
    
    ' Operation Was Not Successful
    RegWriteString = 0
    
    ' Close Handle To Key
    Temp& = OSRegCloseKey(hHnd&)
    
End Function

Private Function RegCheckError(ByRef ErrorValue As Long) As Boolean

    If ((ErrorValue < 8) And (ErrorValue > 1)) Or _
       (ErrorValue = 87) Or (ErrorValue = 259) Then _
       RegCheckError = -1 Else RegCheckError = 0

End Function

Public Function RegReadString(ByVal hKey As Long, ByVal Key As String, ByVal SubKey As String, ByVal ValueName As String) As String

    ' Combine The Key And SubKey Paths
    If Not SubKey = "" Then KeyPath$ = _
    Key + "\" + SubKey Else KeyPath$ = Key
    
    ' Open The Key For Operations
    Temp& = OSRegOpenKey(hKey, KeyPath$, hHnd&)
    
    ' Process Returned Information
    If RegCheckError(Temp&) Then GoTo ReadStringError
    
    ' Read In Information In Unicode Format
    Temp& = OSRegQueryValueEx(hHnd&, ValueName, 0&, TempEx&, Temp&, TempExB&)
    
    ' Process Returned Information
    If RegCheckError(Temp&) Then GoTo ReadStringError
    
    ' Operation Was Successful
    If TempEx& = REG_SZ Then
         
        ' Create ASCIIZ Based String
        TempExA$ = String(TempExB&, " ")
        
        ' Convert Information To String Format
        Temp& = OSRegQueryValueEx(hHnd&, ValueName, 0&, 0&, ByVal TempExA$, TempExB&)

        ' Process Returned Information
        If RegCheckError(Temp&) Then GoTo ReadStringError
        
        ' Find Unicode String NULL Terminator
        TempExC% = InStr(TempExA$, Chr$(0))
        
        ' Return All Characters Before NULL
        If TempExC% > 0 Then
            RegReadString = Left$(TempExA$, TempExC% - 1)
        Else
            RegReadString = TempExA$
        End If

    End If

    ' Close Handle To Key
    Temp& = OSRegCloseKey(hHnd&)

    ' Exit Function With Passed Value
    Exit Function

ReadStringError:
    
    ' Store Error In Variable
    RegLastError = Temp&
    
    ' Operation Was Not Successful
    RegReadString = vbNullString
    
    ' Close Handle To Key
    Temp& = OSRegCloseKey(hHnd&)
    
End Function

Public Sub SetSysTrayIcon(eAction As SysTrayAction, hMsgWnd As Long, hIcon As Long, sToolTip As String)

'Example:
'   Call SetSysTrayIcon(ModifyIcon, chkHiddenCheckBox.hWnd, Me.Icon, "My Program Name")
'   chkHiddenCheckBox_MouseMove() event will receive all messages from the System Tray Icon.
'   chkHiddenCheckBox could be any control with a MouseMove event.

Dim lRet As Long

    With guIconData
        .cbSize = Len(guIconData)
        .hwnd = hMsgWnd
        .uID = vbNull
        .uFlags = STI_MESSAGE Or STI_ICON Or STI_TIP
        .uCallBackMsg = WM_MOUSEMOVE
        .hIcon = hIcon
        .szTip = sToolTip & Chr$(0)
    End With
    
    lRet = SysTrayIcon(eAction, guIconData)
    
End Sub
