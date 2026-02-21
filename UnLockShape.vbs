'==========================================================================
' DISCLAIMER OF WARRANTY:  Unless otherwise agreed in writing,
' Mentor Graphics software and associated files are provided "as is"
' and without warranty.  Mentor Graphics has no obligation to support
' or otherwise maintain software.  Mentor Graphics makes no warranties,
' express or implied with respect to software including any warranty
' of merchantability or fitness for a particular purpose.
'
' LIMITATION OF LIABILITY: Mentor Graphics is not liable for any property
' damage, personal injury, loss of profits, interruption of business, or for
' any other special, consequential or incidental damages, however caused,
' whether for breach of warranty, contract, tort (including negligence),
' strict liability or otherwise. In no event shall Mentor Graphics'
' liability exceed the amount paid for the product giving rise to the claim.
'==========================================================================
'
'   Source: Mentor Graphics CSD,  Longmont, Co. USA
'   Author: Al Layson
'   App:    None
'   O/S:	Windows XP
'   Lang:	VBScript
'   Status: Work in Progress
'   Date:   8/10/2010
'   Description: 
' 		Remove the lock from Conductive Shapes - DR 645904
'==========================================================================


Option Explicit     ' This means that all variables must be declared using the 'Dim' statement

'---------------------------------------
'Stup Main program
'---------------------------------------
' Get the application object
Dim pcbApp
Set pcbApp = Application
Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")

' Get the active document
Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument

Sub main()
	pcbApp.lockserver
	pcbDoc.transactionStart 

	'........................insert code...........................
	'use file.writeline() to write to transcript
	Dim ca,cnt
	cnt=0
	Dim del
	del=MsgBox("Force Delete?",vbYesnocancel,"Unlock Shape")
	If del=vbCancel Then Exit Sub
	On Error Resume Next
	For Each ca In pcbDoc.ConductiveAreas
		If ca.selected Then
			ca.Anchor=epcbAnchorNone
			If del=vbYes Then ca.delete()
			If Err Then 'although a runtime error, expedition doesn't report the error which is related to existing Hazards
				MsgBox(Err.Number & " : " & Err.Description)
				Err.Clear
			Else
				cnt=cnt+1
			End If
		End If
	Next
	MsgBox(cnt & " Conductive Shapes were Unlocked/Unfixed and/or deleted")
	pcbDoc.transactionEnd
	pcbApp.unlockServer
End Sub
	
' License the document
If (ValidateServer(pcbDoc) = 1) Then Call main
'---------------------------------------
' Begin Validate Server Function
'---------------------------------------
Private Function ValidateServer(doc)
    
    Dim key, licenseServer, licenseToken

    ' Ask Expedition’s document for the key
    key = doc.Validate(0)

    ' Get license server
    Set licenseServer = CreateObject("MGCPCBAutomationLicensing.Application")

    ' Ask the license server for the license token
    licenseToken = licenseServer.GetToken(key)

    ' Release license server
    Set licenseServer = nothing

    ' Turn off error messages.  Validate may fail if the token is incorrect
    On Error Resume Next
    Err.Clear

    ' Ask the document to validate the license token
    doc.Validate(licenseToken)
    If Err Then
        ValidateServer = 0    
    Else
        ValidateServer = 1
    End If

End Function
'---------------------------------------
' End Validate Server Function
'---------------------------------------
