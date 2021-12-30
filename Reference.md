---
layout: default
title: Reference
nav_order: 3
permalink: /reference
---

DllManager is a predeclared class employing the Factory/Constructor (Create/Init) pattern. It wraps several DLL-related Windows APIs and a Scripting.Dictionary object for the \<DLL name\>&nbsp;&rarr;&nbsp;\<DLL handle\> mapping.

**DllManager.Create** & **DllManager.Init**
The factory takes three optional arguments, passed to the constructor:  

  * *DefaultPath* (String) - the default DLL directory,  
  * *DllNames* (Variant, defaults to empty) - a single DLL name (String) or multiple DLL names (array),  
  * *ResolvePath* (Boolean, defaults to true) - indicates whether the factory should resolve the library path.  

If *DefaultPath* is not provided or set to blank string, it defaults to "Library\" & ThisWorkbook.VBProject\.Name & "\dll\" & ARCH, where ARCH is a conditionally declared constant set to "x32" for VBA-x32 and "x64" for VBA-x64. *DefaultPath* can be set/changed via its setter (Property Let) and reset with a blank string. The setter checks if the parameter holds a valid absolute or relative (w.r.t. ThisWorkbook.Path) path. If this check succeeds, SetDllDirectory API sets the default DLL search path.

*DllNames* accepts either a String containing a single DLL name (passed to *DllManager.Load*) or an array of String names (passed to *DllManager.LoadMultiple*).  

*ResolvePath* is only honored if the second argument is a String (see *DllManager.Load* for further details).  

**DllManager.Load** & **DllManager.Free**
*DllManager.Load* loads individual libraries, taking three arguments:

  * *DllName* (String, required) - the DLL name to be loaded,
  * *Path* (String, defaults to *DefaultPath*) - the target DLL directory,
  * *ResolvePath* (Boolean, defaults to true) - indicates whether *Load* should resolve the library path.

If *Path* is provided, it is used instead of the *DefaultPath*.  

If *ResolvePath* is true, it checks if *Path* is a valid absolute or relative (w.r.t. ThisWorkbook.Path) path and, if this check passes, appends the *DllName*. Otherwise, *Load* appends *DllName* to *Path* without any checks.

Eventually, the LoadLibrary API loads the library.

*DllManager.Free* frees previously loaded library.

**DllManager.LoadMultiple** & **DllManager.FreeMultiple**
take a 0-based array of DLL names to be loaded/freed (a ParamArray list is also accepted). If no argument is provided, *LoadMultiple* does nothing, but *FreeMultiple* frees all loaded libraries.

**DllManager.GetDllPath**
takes  a name of a previously loaded DLL and returns its file path.

**DllManager.ProcAddressGet** & **DllManager.CacheProcPtr**
*DllManager.ProcAddressGet* takes two String arguments, the name of a previously loaded DLL library and a function name, and retrieves the entry pointer (RAM address) via Windows API. It also saves the retrieved address in the internal cache and uses it on subsequent requests. *DllManager.CacheProcPtr* also takes a function pointer as the third argument and saves it into the internal cache directly. This API can be used, for example, for tests involving *DllManager.ProcAddressGet*.

**DllManager.IndirectCall**
wraps the DispCallFunc Windows API to make proxied DLL calls. It takes five arguments:

  * *ModuleName* (String) - same as ProcAddressGet/CacheProcPtr,
  * *ProcName* (String) - same as ProcAddressGet/CacheProcPtr,
  * *CallnigConv* (CALLCONV Enum) - indicates calling convention of the called function,
  * *ReturnType* (VbVarType Enum) - indicates type of the returned value,
  * *Arguments* (array of Variant) - contains arguments to be passed to the called function from left to right.

The arguments should be prepared according to the called function specifications, providing addresses where necessary; *IndirectCall* cannot provide any assistance here.

**DllManager.Class_Terminate**
calls *FreeMultiple*, freeing all loaded libraries.


<!-- References -->

[DLL API]: https://docs.microsoft.com/en-us/windows/win32/dlls/dynamic-link-library-functions
[SQLite VBA]: https://pchemguy.github.io/SQLite-ICU-MinGW/stdcall
