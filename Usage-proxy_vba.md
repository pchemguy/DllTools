---
layout: default
title: Proxied call - VBA procedure
nav_order: 3
parent: Usage examples
permalink: /usage/proxy-vba
---

#### DllCallDemoVBACall

```vb
Option Explicit

Private Type TModuleState
    LongVal As Long
    LongRef As Long
    ByteVal As Byte
    ByteRef As Byte
    StrVal As String
    StrRef As String
End Type
Private this As TModuleState


Private Sub Main()
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(vbNullString)
    DllMan.CacheProcPtr "DllCallDemoVBACall", "In3Out3Ret1", AddressOf In3Out3Ret1
    
    With this
        .ByteVal = 10
        .LongVal = 30
        .StrVal = "StrVal"
        .ByteRef = 20
        .LongRef = 40
        .StrRef = "StrRef"
    End With
        
    With this
        Dim Arguments As Variant
        Arguments = Array( _
            .ByteVal, _
            .LongVal, _
            .StrVal, _
            VarPtr(.ByteRef), _
            VarPtr(.LongRef), _
            VarPtr(.StrRef) _
        )
    End With
    
    Debug.Print "==================== In3Out3Ret1 ===================="
    Dim Result As Long
    Result = DllMan.IndirectCall("DllCallDemoVBACall", "In3Out3Ret1", _
                                 CC_STDCALL, vbLong, Arguments)
    
    Debug.Print vbNewLine & "----- VERIFYING RETURNED VALUES -----"
    With this
        Debug.Print "ByteVal = " & CStr(.ByteVal) & vbTab & vbTab & IIf(.ByteVal = 10, "OK/UNCHANGED", "BAD")
        Debug.Print "ByteRef = " & CStr(.ByteRef) & vbTab & vbTab & IIf(.ByteRef = 200, "OK/UPDATED", "BAD")
        Debug.Print "LongVal = " & CStr(.LongVal) & vbTab & vbTab & IIf(.LongVal = 30, "OK/UNCHANGED", "BAD")
        Debug.Print "LongRef = " & CStr(.LongRef) & vbTab & vbTab & IIf(.LongRef = 400, "OK/UPDATED", "BAD")
        Debug.Print "StrVal  = " & CStr(.StrVal) & vbTab & IIf(.StrVal = "StrVal", "OK/UNCHANGED", "BAD")
        Debug.Print "StrRef  = " & CStr(.StrRef) & vbTab & IIf(.StrRef = "StrRefNew", "OK/UPDATED", "BAD")
    End With
    Debug.Print "Result  = " & CStr(Result) & vbTab & vbTab & IIf(Result = 70, "OK", "BAD")
    Debug.Print "-------------------- In3Out3Ret1 --------------------"
End Sub


Private Function In3Out3Ret1(ByVal ByteVal As Byte, ByVal LongVal As Long, ByVal StrVal As String, _
                             ByRef ByteRef As Byte, ByRef LongRef As Long, ByRef StrRef As String) As Long
    Debug.Print "----- VERIFYING RECEIVED ARGUEMNTS -----"
    Debug.Print "ByteVal = " & CStr(ByteVal) & vbTab & vbTab & IIf(ByteVal = 10, "OK", "BAD")
    Debug.Print "ByteRef = " & CStr(ByteRef) & vbTab & vbTab & IIf(ByteRef = 20, "OK", "BAD")
    Debug.Print "LongVal = " & CStr(LongVal) & vbTab & vbTab & IIf(LongVal = 30, "OK", "BAD")
    Debug.Print "LongRef = " & CStr(LongRef) & vbTab & vbTab & IIf(LongRef = 40, "OK", "BAD")
    Debug.Print "StrVal  = " & CStr(StrVal) & vbTab & IIf(StrVal = "StrVal", "OK", "BAD")
    Debug.Print "StrRef  = " & CStr(StrRef) & vbTab & IIf(StrRef = "StrRef", "OK", "BAD")
    In3Out3Ret1 = LongVal + LongRef
    
    LongVal = 300
    LongRef = 400
    ByteVal = 100
    ByteRef = 200
    StrVal = "StrValNew"
    StrRef = "StrRefNew"
End Function
```
