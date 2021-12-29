---
layout: default
title: Proxied call - SQLite
nav_order: 2
parent: Usage examples
permalink: /usage/proxy-sqlite
---

The following examples (see the DllTools.Demo.DllCall folder in the RD Code Explorer) demonstrate the use of the *IndirectCall* method, which wraps the *DispCallFunc* API and facilitates indirect DLL calls. The primary use case for this API is calling an x32 *CDECL* DLL routine, which is not available as an STDCALL/WINAPI version. A word of warning: the VBA compiler cannot perform any correctness checks on DLL calls placed via the DispCallFunc API. While IndirectCall partially automates the setup process where possible, any caller's mistake can easily crash the host application.

This example calls the same STDCALL SQLite DLLs accessible via the Declare statement, as illustrated previously. Its primary purpose is to compare the setup processes and results. (The only change necessary for a CDECL version of SQLite is changing the CC_STDCALL parameter below to CC_CDECL.) This code calls two SQLite functions, sqlite3_libversion_number, as before, and sqlite3_libversion. These functions take no arguments, so the setup is similar to the previous example, except that the Declare statement is gone, and the calls go through IndirectCall rather than directly.

#### DllCallDemoSQLite

```vb
'@Folder "DllTools.Demo.DllCall"
Option Explicit

Private Const LIB_NAME As String = "DllTools"
Private Const PATH_SEP As String = "\"
Private Const LIB_RPREFIX As String = _
    "Library" & PATH_SEP & LIB_NAME & PATH_SEP & "dll" & PATH_SEP

Private Sub Main()
    Dim PtrType As VbVarType
    Dim DllNames As Variant
    #If Win64 Then
        PtrType = vbLongLong
        DllNames = "sqlite3.dll"
    #Else
        PtrType = vbLong
        DllNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", _
                         "icuio68.dll", "icutu68.dll", "sqlite3.dll")
    #End If
    Dim DllPath As String
    DllPath = LIB_RPREFIX & ARCH
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DllPath, DllNames)
    Dim SQLiteVerLng As Long
    SQLiteVerLng = DllMan.IndirectCall("SQLite3", "sqlite3_libversion_number", CC_STDCALL, vbLong, Empty)
    Debug.Print "SQLite version: " & CStr(SQLiteVerLng)
    Dim SQLiteVerStr As String
    SQLiteVerStr = UTFlib.StrFromUTF8Ptr( _
        DllMan.IndirectCall("SQLite3", "sqlite3_libversion", CC_STDCALL, PtrType, Empty))
    Debug.Print "SQLite version: " & SQLiteVerStr
    If Replace(Replace(SQLiteVerStr, ".", "0"), "0", vbNullString) = _
       Replace(CStr(SQLiteVerLng), "0", vbNullString) Then
        Debug.Print "VERSIONS MATCHED"
    Else
        Debug.Print "VERSIONS MISMATCHED"
    End If
End Sub
```
