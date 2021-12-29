---
layout: default
title: Load DLLs from user folder
nav_order: 1
parent: Usage examples
permalink: /usage/load-dlls
---

### Load custom-compiled SQLite DLLs from a user folder

Below is a stripped-down version of the SQLiteC class from the SQLiteCAdo library responsible for loading an SQLite DLL with dependencies (DllManagerDemoSQLiteC class in the *DllTools.Manager.Demo* RD Code Explorer folder). This class, similar to DllManager, has the predeclared attribute set and employs the Factory/Constructor (Create/Init) pattern. For illustrative purposes, only one SQLite function declaration remains in the class. Private structure *TObjectState* encapsulates private instance fields (variable `this`), including a reference to a DllManager instance. The caller should keep this reference while using the DLL, because DllManager's Class_Terminate calls the FreeLibrary API freeing the loaded library.

The DllManager factory takes the default DLL path as the first required argument. It can be blank for target DLLs located in a preset location within the project folder optionally checked by DllManager. The second optional argument specifies the names of the DLLs to be loaded, and if not provided, the Init constructor uses the default values. Also, note that the SQLite build used by SQLiteCAdo includes several dependencies. In the case of the x32 version, Windows fails to load these dependencies automatically, so an array of DLL names explicitly specifies all DLLs for the x32 environment.

#### DllManagerDemoSQLiteC.cls

```vb
'@PredeclaredId
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function sqlite3_libversion_number Lib "SQLite3" () As Long
#Else
Private Declare Function sqlite3_libversion_number Lib "SQLite3" () As Long
#End If

Private Type TObjectState
    DllMan As DllManager
    Connections As Scripting.Dictionary
End Type
Private this As TObjectState

Public Function Create(ByVal DllPath As String, _
              Optional ByVal DllNames As Variant = Empty) As DllManagerDemoSQLiteC
    Dim Instance As DllManagerDemoSQLiteC
    Set Instance = New DllManagerDemoSQLiteC
    Instance.Init DllPath, DllNames
    Set Create = Instance
End Function

Friend Sub Init(ByVal DllPath As String, _
       Optional ByVal DllNames As Variant = Empty)
    Dim FileNames As Variant
    If Not IsEmpty(DllNames) Then
        FileNames = DllNames
    Else
        #If Win64 Then
            FileNames = "sqlite3.dll"
        #Else
            FileNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", _
                              "icuio68.dll", "icutu68.dll", "sqlite3.dll")
        #End If
    End If
    Set this.DllMan = DllManager.Create(DllPath, FileNames)
    Set this.Connections = New Scripting.Dictionary
    this.Connections.CompareMode = TextCompare
End Sub

Public Function Version() As Long
    Version = sqlite3_libversion_number()
End Function
```

The RubberDuck Addin, if available, can activate the predeclared class attribute. Otherwise, an auto-assigned module- or project-level variable named after the class can act as a predeclared instance:  
`Private/Public DllManagerDemoSQLiteC as New DllManagerDemoSQLiteC`  
In the former case, this command executed from the *immediate pane* prints the SQLite version number:  
`?DllManagerDemoSQLiteC.Create("").Version`  
This class can be instantiated from a standard module, for example: 

```vb
Private Sub InitDBQC()
    Dim DllPath As String
    Dim DllNames As Variant
    #If Win64 Then
        DllPath = ThisWorkbook.Path & "\Library\DllTools\dll\x64"
        DllNames = "sqlite3.dll"
    #Else
        DllPath = ThisWorkbook.Path & "\Library\DllTools\dll\x32"
        DllNames = Array("icudt68.dll", "icuuc68.dll", "icuin68.dll", _
                         "icuio68.dll", "icutu68.dll", "sqlite3.dll")
    #End If
    Dim dbm As DllManagerDemoSQLiteC
    Set dbm = DllManagerDemoSQLiteC.Create(DllPath, DllNames)
    If dbm Is Nothing Then
        Err.Raise ErrNo.ObjectCreateErr, "SQLiteCExamples", _
                  "Failed to create an DllManagerDemoSQLiteC instance."
    Else
        Debug.Print "Database manager instance (DllManagerDemoSQLiteC class) is ready"
        Debug.Print "SQLite version number: " & CStr(dbm.Version())
    End If
```

*DllManagerDemo* from the same Code Explorer folder (*DllTools.Manager.Demo*) demonstrates a similar functionality executed from a standard module directly.
