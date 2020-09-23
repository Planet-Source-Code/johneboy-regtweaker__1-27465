VERSION 5.00
Begin VB.Form Warning 
   Caption         =   "Warning!"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3990
   Icon            =   "OS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Um, &No Thanks"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Yes, Go Ahead"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Are you sure you want to continue?"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   3255
   End
End
Attribute VB_Name = "warning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

      Private Declare Function ShellExecute Lib _
     "shell32.dll" Alias "ShellExecuteA" _
     (ByVal hwnd As Long, ByVal lpOperation _
     As String, ByVal lpFile As String, ByVal _
     lpParameters As String, ByVal lpDirectory _
     As String, ByVal nShowCmd As Long) As Long
     
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long


Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long


Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long


Private Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long


Private Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long


Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long


Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long


Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Private Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long


Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long


Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long


Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long


Private Declare Function RegUnLoadKey Lib "advapi32.dll" Alias "RegUnLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
    Private Const REG_NONE = 0
    Private Const REG_SZ = 1
    Private Const REG_EXPAND_SZ = 2
    Private Const REG_BINARY = 3
    Private Const REG_DWORD = 4
    Private Const REG_DWORD_LITTLE_ENDIAN = 4
    Private Const REG_DWORD_BIG_ENDIAN = 5
    Private Const REG_LINK = 6
    Private Const REG_MULTI_SZ = 7
    Private Const REG_RESOURCE_LIST = 8
    Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9
    Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10
    Private Const REG_CREATED_NEW_KEY = &H1
    Private Const REG_OPENED_EXISTING_KEY = &H2
    Private Const REG_WHOLE_HIVE_VOLATILE = &H1
    Private Const REG_REFRESH_HIVE = &H2
    Private Const REG_NOTIFY_CHANGE_NAME = &H1
    Private Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
    Private Const REG_NOTIFY_CHANGE_LAST_SET = &H4
    Private Const REG_NOTIFY_CHANGE_SECURITY = &H8
    Private Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
    Private Const REG_OPTION_RESERVED = 0
    Private Const REG_OPTION_NON_VOLATILE = 0
    Private Const REG_OPTION_VOLATILE = 1
    Private Const REG_OPTION_CREATE_LINK = 2
    Private Const REG_OPTION_BACKUP_RESTORE = 4
    Private Const STANDARD_RIGHTS_READ = &H20000
    Private Const STANDARD_RIGHTS_WRITE = &H20000
    Private Const STANDARD_RIGHTS_EXECUTE = &H20000
    Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
    Private Const STANDARD_RIGHTS_ALL = &H1F0000
    Private Const DELETE = &H10000
    Private Const READ_CONTROL = &H20000
    Private Const WRITE_DAC = &H40000
    Private Const WRITE_OWNER = &H80000
    Private Const SYNCHRONIZE = &H100000
    Private Const KEY_QUERY_VALUE = &H1
    Private Const KEY_SET_VALUE = &H2
    Private Const KEY_CREATE_SUB_KEY = &H4
    Private Const KEY_ENUMERATE_SUB_KEYS = &H8
    Private Const KEY_NOTIFY = &H10
    Private Const KEY_CREATE_LINK = &H20
    Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))


Private Function GetString(hKey As Long, strpath As String, strvalue As String)
    Dim keyhand&
    Dim datatype&
    r = RegOpenKey(hKey, strpath, keyhand&)
    GetString = RegQueryStringValue(keyhand&, strvalue)
    r = RegCloseKey(keyhand&)
End Function


Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String)
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    On Error GoTo 0
    lResult = RegQueryValueEx(hKey, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)


    If lResult = ERROR_SUCCESS Then


        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, " ")
            lResult = RegQueryValueEx(hKey, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)


            If lResult = ERROR_SUCCESS Then
                RegQueryStringValue = StripTerminator(strBuf)
            End If
        End If
    End If
End Function


Private Sub SaveKey(hKey As Long, strpath As String)
    Dim keyhand&
    r = RegCreateKey(hKey, strpath, keyhand&)
    r = RegCloseKey(keyhand&)
End Sub


Private Sub SaveString(hKey As Long, strpath As String, strvalue As String, strdata As String)
    Dim keyhand&
    r = RegCreateKey(hKey, strpath, keyhand&)
    r = RegSetValueEx(keyhand&, strvalue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand&)
End Sub


Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))


    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
'HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion

Function GetWindowsVersion() As String
    GetWindowsVersion = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "Version")

End Function

Private Sub Command1_Click()
Main.Show
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
RemoveCloseMenu Me
Label1.Caption = "" + GetWindowsVersion + ""
If Label1.Caption = "" Then
Label1.Caption = "Windows NT/2000/XP"
End If
Label2.Caption = "regTweaker has detected you are using Microsoft " + Label1.Caption + ".  regTweaker was made on and designed for the Windows 98 environment.  Although some tweaks may work, others may damage your system"
If Label1.Caption = "Windows 98" Then
Main.Show
Unload Me
Else
End If

End Sub


