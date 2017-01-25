Attribute VB_Name = "modPublic"
Option Explicit

Public CodeOpe As CodeOperator

Public Enum AC_ItemKind
    kCollection = 10
    kMethod = 1
    kProperty = 2
    kVar = 3
    kEvent = 4
    kConst = 5
    kClass = 6
    kModule = 7
    kType = 8
    kEnum = 9
    kUnUsedMethod = 11
    kUnUsedConst = 12
    kUnUsedType = 13
End Enum

Public Const SignOfChoose As String = "!^*()-+=;:,<.>/ "
Public Const SignOfEnd As String = "`~@#$%&[{]}\|'?"""

Public AddInIns As AddIn
