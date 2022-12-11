VERSION 5.00
Begin VB.Form frmMain
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Silent Echoes WinSpy"
   ClientHeight    =   3240
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7620
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3
      Caption         =   "Exit"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2760
      Width           =   4095
   End
   Begin VB.TextBox WNTKey
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Top             =   600
      Width           =   3615
   End
   Begin VB.TextBox W9xKey
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox txtSet
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label3
      Caption         =   "Your system variables are:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label2
      Caption         =   "Your Windows NT/2000/XP/2003 Product Key is:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label1
      Caption         =   "Your Windows 95/98/ME Product Key Is:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Menu winspy
      Caption         =   "&WinSpy"
      Begin VB.Menu run
         Caption         =   "&Run..."
      End
      Begin VB.Menu browse
         Caption         =   "&Browse..."
      End
      Begin VB.Menu cpanel
         Caption         =   "Launch &Control Panel"
      End
      Begin VB.Menu telnet
         Caption         =   "Run &Telnet"
      End
      Begin VB.Menu exit
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu sem
      Caption         =   "&Silent Echoes Media"
      Begin VB.Menu soft
         Caption         =   "Get More &Software..."
      End
      Begin VB.Menu site
         Caption         =   "&Web Site"
      End
   End
   Begin VB.Menu help
      Caption         =   "&Help"
      Begin VB.Menu contents
         Caption         =   "&Contents..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This module reads and writes registry keys.  Unlike the
' internal registry access methods of VB, it can read and
' write any registry keys with string values.

'---------------------------------------------------------------
'-Registry API Declarations...
'---------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

'---------------------------------------------------------------
'- Registry Api Constants...
'---------------------------------------------------------------
' Reg Data Types...
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

' Reg Create Type Values...
Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Reg Key ROOT Types...
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004

' Return Value...
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- Registry Security Attributes TYPE...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

' The resource string will be loaded into a control's property as follows:
' Object      Property
' Form        Caption
' Menu        Caption
' TabStrip    Caption, ToolTipText
' Toolbar     ToolTipText
' ListView    ColumnHeader.Text

Sub LoadResStrings(frm As Form)
  On Error Resume Next

  Dim ctl As Control
  Dim obj As Object

  'set the form's caption
  If IsNumeric(frm.Tag) Then
    frm.Caption = LoadResString(CInt(frm.Tag))
  End If

  'set the controls' captions using the caption
  'property for menu items and the Tag property
  'for all other controls
  For Each ctl In frm.Controls
    Err.Clear
    If TypeName(ctl) = "Menu" Then
      If IsNumeric(ctl.Caption) Then
        If Err = 0 Then
          ctl.Caption = LoadResString(CInt(ctl.Caption))
        End If
      End If
    ElseIf TypeName(ctl) = "TabStrip" Then
      For Each obj In ctl.Tabs
        Err.Clear
        If IsNumeric(obj.Tag) Then
          obj.Caption = LoadResString(CInt(obj.Tag))
        End If
        'check for a tooltip
        If IsNumeric(obj.ToolTipText) Then
          If Err = 0 Then
            obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
          End If
        End If
      Next
    ElseIf TypeName(ctl) = "Toolbar" Then
      For Each obj In ctl.Buttons
        Err.Clear
        If IsNumeric(obj.Tag) Then
          obj.ToolTipText = LoadResString(CInt(obj.Tag))
        End If
      Next
    ElseIf TypeName(ctl) = "ListView" Then
      For Each obj In ctl.ColumnHeaders
        Err.Clear
        If IsNumeric(obj.Tag) Then
          obj.Text = LoadResString(CInt(obj.Tag))
        End If
      Next
    Else
      If IsNumeric(ctl.Tag) Then
        If Err = 0 Then
          ctl.Caption = LoadResString(CInt(ctl.Tag))
        End If
      End If
      'check for a tooltip
      If IsNumeric(ctl.ToolTipText) Then
        If Err = 0 Then
          ctl.ToolTipText = LoadResString(CInt(ctl.ToolTipText))
        End If
      End If
    End If
  Next

End Sub

'-------------------------------------------------------------------------------------------------
'sample usage - Debug.Print UpodateKey(HKEY_CLASSES_ROOT, "keyname", "newvalue")
'-------------------------------------------------------------------------------------------------
Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
    Dim rc As Long                                      ' Return Code
    Dim hKey As Long                                    ' Handle To A Registry Key
    Dim hDepth As Long                                  '
    Dim lpAttr As SECURITY_ATTRIBUTES                   ' Registry Security Type

    lpAttr.nLength = 50                                 ' Set Security Attributes To Defaults...
    lpAttr.lpSecurityDescriptor = 0                     ' ...
    lpAttr.bInheritHandle = True                        ' ...

    '------------------------------------------------------------
    '- Create/Open Registry Key...
    '------------------------------------------------------------
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)                   ' Create/Open //KeyRoot//KeyName

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Errors...

    '------------------------------------------------------------
    '- Create/Modify Key Value...
    '------------------------------------------------------------
    If (SubKeyValue = "") Then SubKeyValue = " "        ' A Space Is Needed For RegSetValueEx() To Work...

    ' Create/Modify Key Value
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))

    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError   ' Handle Error
    '------------------------------------------------------------
    '- Close Registry Key...
    '------------------------------------------------------------
    rc = RegCloseKey(hKey)                              ' Close Key

    UpdateKey = True                                    ' Return Success
    Exit Function                                       ' Exit
CreateKeyError:
    UpdateKey = False                                   ' Set Error Return Code
    rc = RegCloseKey(hKey)                              ' Attempt To Close Key
End Function

'-------------------------------------------------------------------------------------------------
'sample usage - Debug.Print GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
'-------------------------------------------------------------------------------------------------
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable

    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...

    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size

    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors

    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' Search Data Types...
    Case REG_SZ, REG_EXPAND_SZ                              ' String Registry Key Data Type
        sKeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' Convert Double Word To String
    End Select

    GetKeyValue = sKeyVal                                   ' Return Value
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit

GetKeyError:    ' Cleanup After An Error Has Occured...
    GetKeyValue = vbNullString                              ' Set Return Val To Empty String
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


Private Sub browse_Click()
frmBrowse.Show
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub contents_Click()
MsgBox "Wehk Eet Aeet!", vbOKOnly, "Silent Echoes WinSpy"
End Sub

Private Sub cpanel_Click()
Shell "control", vbNormalFocus
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
W9xKey.Text = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProductKey")
WNTKey.Text = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "ProductKey")




End Sub

Private Sub run_Click()
frmRun.Show
End Sub

Private Sub site_Click()
Shell "start http://www.silentechoes.com/", vbNormalFocus
End Sub

Private Sub soft_Click()
Shell "start http://www.silentechoes.com/software", vbNormalFocus
End Sub

Private Sub telnet_Click()
Shell "Telnet", vbNormalFocus
End Sub
