---
layout: default
title: Proxied call - CDECL function
nav_order: 5
parent: Usage examples
permalink: /usage/proxy-cdecl
---

Finally, the target of this example is a CDECL WINAPI function [wsprintfW][], which is an analog of the C swprintf function. Here,  
{"Param1 = %d , Param2 = %s", 1048576, "ABC"}  
is transformed into  
"Param1 = 1048576 , Param2 = ABC".

#### DllCallDemoCDECL

```vb
Option Explicit

Private Sub Main()
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(vbNullString, "user32", False)
    
    Dim Buffer As String
    Buffer = String(1024, vbNullChar)
    Dim Template As String
    Template = "Param1 = %d , Param2 = %s"
    Dim NumArg As Long
    NumArg = 1048576
    Dim StrArg As String
    StrArg = "ABC"
    
    Dim Arguments As Variant
    Arguments = Array( _
        StrPtr(Buffer), _
        StrPtr(Template), _
        NumArg, _
        StrPtr(StrArg) _
    )

    Dim Result As Long
    Debug.Print "==================== DLL-CDECL ===================="
    Result = DllMan.IndirectCall("user32", "wsprintfW", CC_CDECL, vbLong, Arguments)
    Buffer = Left$(Buffer, Result)
    Debug.Print "ResultL = " & CStr(Result) & String(9, vbTab) & IIf(Result = 31, "MATCHED", "MISMATCHED")
    Debug.Print "Result  = """ & Buffer & """" & vbTab & vbTab & _
                IIf(Buffer = "Param1 = 1048576 , Param2 = ABC", "MATCHED", "MISMATCHED")
    Debug.Print "-------------------- DLL-CDECL --------------------"
End Sub
```

This example is based on this blog [post][source].


<!-- References -->

[wsprintfW]: https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-wsprintfw
[source]: https://akihitoyamashiro.blogspot.com/2020/07/how-to-use-function-pointer-in-vba-3.html
