---
layout: default
title: Summary of functionality
nav_order: 1
permalink: /
---

The DllManager class, the main component of DllTools, facilitates VBA calls to

  * DLLs in non-standard locations,  
  * CDECL routines from VBA-x32 hosts, and  
  * DLLs with multiple non-system dependencies.  

#### DLLs with dependencies in non-standard locations

While the Declare statement is sufficient for making VBA-compatible DLL calls, this statement must include an absolute path to the library for non-standard locations, making the declarations ugly and the code fragile. The alternative approach involves explicit loading of DLLs via the LoadLibrary Windows API. Furthermore, Windows may sometimes fail to load dependencies automatically, raising obscure errors. In such a case, DllMnager can take an ordered sequence of DLL names for loading in the provided order.

#### Proxying calls to CDECL routines from VBA-x32 hosts

VBA-x32 only supports calls to DLL routines that follow the WINAPI/STDCALL calling convention. If a VBA-x32 application needs functionality provided by a DLL, a WINAPI version is always preferable. However, some libraries may only be available as CDECL binaries. Additionally, variadic functions must follow the CDECL calling convention and are not directly accessible from VBA-x32. DllManager wraps the DispCallFunc Windows API, which can act as a calling proxy in such cases.
