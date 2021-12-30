---
layout: default
title: Load DLLs from user folder. I
nav_order: 1
parent: Usage examples
permalink: /usage/load-dlls1
---

### Load custom-compiled SQLite DLLs from a user folder using a standard module

DllManager is a predeclared class employing the Factory/Constructor (Create/Init) pattern. The DllManager factory takes the default DLL path as the first required argument. It can be blank for target DLLs located in a preset location within the project folder optionally checked by DllManager. The second optional argument specifies the names of the DLLs to be loaded, and, if not provided, Load/LoadMultiple methods can load the libraries after instantiation.

This demo, DllManagerDemo, is a standard module located in the DllTools.Manager.Demo folder in the RubberDuckVBA Code Explorer. It uses the *sqlite3_libversion_number* function to compare a conventional system call and a similar call to a DLL located in a user directory. Specifically, Windows 10 includes a system copy of the SQLite library named WinSQLite3.dll directly accessible via the Declare statement. Separately, this project provides a custom-compiled SQLite library file SQLite3.dll loaded via DllManager. A module-level attribute, *this.DllMan*, must keep a reference to the DllManager object. Otherwise, after exiting the *SQLiteLoadMultipleArray* Sub, VBA would destroy the DllManager instance, calling its *Class_Terminate* Sub, which, in turn, causes a call to the FreeLibrary API freeing the loaded library.

#### DllManagerDemo

```vb
Option Explicit

#If Win64 Then
Private Const ARCH As String = "x64"
#Else
Private Const ARCH As String = "x32"
#End If

Private Const LIB_NAME As String = "DllTools"
Private Const PATH_SEP As String = "\"
Private Const LIB_RPREFIX As String = "Library" & "\" & LIB_NAME & "\dll\"

#If VBA7 Then
Private Declare PtrSafe Function winsqlite3_libversion_number Lib "WinSQLite3" Alias "sqlite3_libversion_number" () As Long
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
#Else
Private Declare Function winsqlite3_libversion_number Lib "WinSQLite3" Alias "sqlite3_libversion_number" () As Long
Private Declare Function sqlite3_libversion_number Lib "SQLite3" () As Long
#End If

Private Type TModuleState
    DllMan As DllManager
End Type
Private this As TModuleState


Private Sub GetWinSQLite3VersionNumber()
    Debug.Print winsqlite3_libversion_number()
End Sub

Private Sub GetSQLite3VersionNumber()
    SQLiteLoadMultipleArray    
    Debug.Print sqlite3_libversion_number()
    Set this.DllMan = Nothing
End Sub

Private Sub SQLiteLoadMultipleArray()
    Dim DllPath As String
    DllPath = ThisWorkbook.Path & PATH_SEP & LIB_RPREFIX & ARCH
    Dim DllNames As Variant
    If ARCH = "x64" Then
        DllNames = "sqlite3.dll"
    Else
        DllNames = Array( _
            "icudt68.dll", "icuuc68.dll", "icuin68.dll", _
            "icuio68.dll", "icutu68.dll", "sqlite3.dll" _
        )
    End If
    Dim DllMan As DllManager
    Set DllMan = DllManager.Create(DllPath)
    Set this.DllMan = DllMan
    DllMan.LoadMultiple DllNames
End Sub
```

Usually, a DLL located in a user directory can still be accessed directly via the Declare statement by specifying the absolute path to the file. However, the automatic Windows DLL loading process may fail if the library has dependencies. For example, the target DLL may require a particular library version, which differs from the one installed on the system. Or perhaps, a custom dependency resides in a different non-system directory.

In this case, Windows could not load dependencies for the custom-compiled x32 SQLite copy automatically. As demonstrated in this example, the DllManager accepts multiple DLL names and loads them in the order provided, successfully overcoming this issue. Moreover, DllManager can load libraries from several non-standard locations via its Load method.
