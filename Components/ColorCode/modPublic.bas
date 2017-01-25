Attribute VB_Name = "modPublic"
Option Explicit

Public Type Color_Set
    Caption As String
    Creator As String
    Color(0 To 29) As Long
End Type

Public Const Color_Max = 36

Public Color_Count As Long
Public Colors(Color_Max) As Color_Set
