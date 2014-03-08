Attribute VB_Name = "mod_require_vars"
Option Explicit


Public Type OS_Info_Struct
    CurrentAccount As String
    ComputerName As String
    OS As String
    Build As String
    Version_Minor As String
    Version_Major As String
    Version As String
    RootDir As String
End Type

Public Type Require_Struct
    Type As String 'registry, folder, file, etc to check for
    Location As String
    Value As String
    Exists As Boolean
End Type


Public Type Requirements_Struct
    Exchange_Version As String
    OS_Good As Boolean
    OS As OS_Info_Struct
    IIS() As Require_Struct
    SysTools() As Require_Struct
    ExchangeTools() As Require_Struct
    ExMerge() As Require_Struct
End Type

Public Req As Requirements_Struct
