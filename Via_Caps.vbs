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


' Set the unit to be used
pcbDocObj.CurrentUnit = epcbUnitMils

' Start a transaction so that all changes
' are on a single undo level
pcbDocObj.TransactionStart

' Delete all the existing gfx on Test
Dim usrLayerGfxColl
Set usrLayerGfxColl = pcbDocObj.UserLayerGfxs(epcbSelectAll, "Test")

usrLayerGfxColl.Delete

' Get the collection of Vias
Dim viaColl
Set viaColl = pcbDocObj.Vias

' Collect parameters for placing graphics on user layer

' Get the user layer
Dim userLayerObj
Set userLayerObj = pcbDocObj.FindUserLayer("Test")

' Fill the graphics
Dim filledBool: filledBool = True

' 0 width
Dim widthReal: widthReal = 0

' Don't tie to a component
Dim cmpObj: Set cmpObj = Nothing

' via cap diameter
Dim radiusDbl: radiusDbl = 25

' Add circle for each via on user layer "Test"
Dim viaObj
For Each viaObj In viaColl
' Create a circle points array at the via location
Dim pntsArr
pntsArr = pcbAppObj.Utility.CreateCircleXYR(viaObj.PositionX, _
viaObj.PositionY, radiusDbl)
Dim numPntsInt
numPntsInt = UBound(pntsArr, 2) + 1

' Add the graphics to the user layer
Call pcbDocObj.PutUserLayerGfx(userLayerObj, widthReal, _
numPntsInt, pntsArr, filledBool, _
cmpObj, epcbUnitCurrent)
Next

Call pcbAppObj.Gui.StatusBarText("Added " & viaColl.Count & _
" via caps on user layer ""Test""", epcbStatusField1)

pcbDocObj.TransactionEnd

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
