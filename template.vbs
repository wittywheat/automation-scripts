Option Explicit

' Add any type libraries to be used.
Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")

' Get the Application object
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' License the document
ValidateServer(pcbDocObj)

' ### Your Code Will Go Here!


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Local functions

' Server validation function
Function ValidateServer(docObj)

  Dim keyInt
  Dim licenseTokenInt
  Dim licenseServerObj

  ' Ask Expedition’s document for the key
  keyInt = docObj.Validate(0)

  ' Get license server
  Set licenseServerObj = CreateObject("MGCPCBAutomationLicensing.Application")

  ' Ask the license server for the license token
  licenseTokenInt = licenseServerObj.GetToken(keyInt)

  ' Release license server
  Set licenseServerObj = nothing

  ' Turn off error messages (validate may fail if the token is incorrect)
  On Error Resume Next
  Err.Clear

  ' Ask the document to validate the license token
  docObj.Validate(licenseTokenInt)
  If Err Then
    ValidateServer = 0
  Else
    ValidateServer = 1
  End If

End Function
