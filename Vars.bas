Attribute VB_Name = "Vars"
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Global FileOK As Boolean
Global NetGame As Boolean
Global URL As String
Global Protected As Boolean
Global Pic(1 To 100) As Boolean
Global Row(1 To 10) As String
Global Col(1 To 10) As String
Global PicEDIT(1 To 100) As Boolean
Global NumTrue As Integer
Global IsTrue As Integer
Global Time As Integer
Global Description As String
Global StageSize As String
Global Tutorial As Integer
Global Stage As String
Global EasyStagePass(1 To 10) As Boolean
Global EasyStageTime(1 To 10) As Integer
Global HardStagePass(1 To 10) As Boolean
Global HardStageTime(1 To 10) As Integer
Global Custom As Boolean
Global file As String
Global Done As Boolean
Global StageNum As Integer

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1
Public Const REG_DWORD = 4

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long

Declare Function RegCreateKey Lib "advapi32.dll" Alias _
        "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, _
        phkResult As Long) As Long

Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
        "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long

Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
        "RegDeleteValueA" (ByVal Hkey As Long, _
        ByVal lpValueName As String) As Long

Declare Function RegOpenKey Lib "advapi32.dll" Alias _
        "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, _
        phkResult As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
        "RegQueryValueExA" (ByVal Hkey As Long, _
        ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, lpcbData As Long) As Long

Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
        "RegSetValueExA" (ByVal Hkey As Long, _
        ByVal lpValueName As String, ByVal Reserved As Long, _
        ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Sub SaveKey(Hkey As Long, strPath As String)
    Dim keyhand&
    r = RegCreateKey(Hkey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub

Public Function GetString(Hkey As Long, strPath As String, strValue As String)
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    r = RegOpenKey(Hkey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

Public Sub SaveString(Hkey As Long, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub

Function GetDWord(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDWord = lBuf
        End If
    End If
    r = RegCloseKey(keyhand)
End Function

Function SaveDWord(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    r = RegCloseKey(keyhand)
End Function

Public Function DeleteKey(ByVal Hkey As Long, ByVal strKey As String)
    Dim r As Long
    r = RegDeleteKey(Hkey, strKey)
End Function

Public Function DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim keyhand As Long
    r = RegOpenKey(Hkey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function

Public Sub EasyRefresh()
Dim Tmp As Integer
Tmp = 1

Open file For Input As #1
Input #1, EasyStagePass(1), EasyStageTime(1), EasyStagePass(2), EasyStageTime(2), EasyStagePass(3), EasyStageTime(3), EasyStagePass(4), EasyStageTime(4), EasyStagePass(5), EasyStageTime(5), EasyStagePass(6), EasyStageTime(6), EasyStagePass(7), EasyStageTime(7), EasyStagePass(8), EasyStageTime(8), EasyStagePass(9), EasyStageTime(9), EasyStagePass(10), EasyStageTime(10)
Close #1

Do While Tmp < 11
    If EasyStagePass(Tmp) = True Then
        frm5x5Select.ckStage(Tmp).Value = Checked
        frm5x5Select.lblTime(Tmp).Caption = EasyStageTime(Tmp)
        frm5x5Select.lblStage(Tmp).Font.Strikethrough = True
    Else
        frm5x5Select.ckStage(Tmp).Value = Unchecked
        frm5x5Select.lblTime(Tmp).Caption = "N/A"
        frm5x5Select.lblStage(Tmp).Font.Strikethrough = False
    End If
    Tmp = Tmp + 1
Loop

End Sub

Public Sub HardRefresh()
Dim Tmp As Integer
Tmp = 1

Open file For Input As #1
Input #1, HardStagePass(1), HardStageTime(1), HardStagePass(2), HardStageTime(2), HardStagePass(3), HardStageTime(3), HardStagePass(4), HardStageTime(4), HardStagePass(5), HardStageTime(5), HardStagePass(6), HardStageTime(6), HardStagePass(7), HardStageTime(7), HardStagePass(8), HardStageTime(8), HardStagePass(9), HardStageTime(9), HardStagePass(10), HardStageTime(10)
Close #1

Do While Tmp < 11
    If HardStagePass(Tmp) = True Then
        frm10x10select.ckStage(Tmp).Value = Checked
        frm10x10select.lblTime(Tmp).Caption = HardStageTime(Tmp)
        frm10x10select.lblStage(Tmp).Font.Strikethrough = True
    Else
        frm10x10select.ckStage(Tmp).Value = Unchecked
        frm10x10select.lblTime(Tmp).Caption = "N/A"
        frm10x10select.lblStage(Tmp).Font.Strikethrough = False
    End If
    Tmp = Tmp + 1
Loop

End Sub

