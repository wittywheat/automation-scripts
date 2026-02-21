Option Explicit

' Add any type libraries to be used.
Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")

' Get the Application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
