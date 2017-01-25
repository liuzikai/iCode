Attribute VB_Name = "modPublic"
Option Explicit

Private Const SettingFileName As String = "iCode\Settings.ini"
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Public IDELoaded As Boolean

Public AddInIns As AddIn

Public Windows_Linker As clsWindowsHandler
Public IDEEnhancer As New clsIDEEnhancer
Public CodeStatistic As clsCodeStatistic
Public TipsBarHandler As clsTipsBarHandler
Public CodeIndent As clsCodeIndent
Public AutoComplete As clsAutoComplete
Public ColorCode As clsColorCode

Public Function Settings_Get(ByVal Section As String, ByVal KeyName As String, ByVal DefaultValue As String) As String
    Dim Buffer As String * 255
    Call GetPrivateProfileString(Section, KeyName, DefaultValue, Buffer, 255, Environ("APPDATA") & "\" & SettingFileName)
    Settings_Get = Left$(Buffer, InStr(Buffer, Chr$(0)) - 1)
End Function

Public Sub Settings_Write(ByVal Section As String, ByVal KeyName As String, ByVal KeyValue As String)
    If Dir(Environ("APPDATA") & "\iCode", vbDirectory) = "" Then
        MkDir Environ("APPDATA") & "\iCode"
    End If
    Call WritePrivateProfileString(Section, KeyName, KeyValue, Environ("APPDATA") & "\" & SettingFileName)
End Sub


