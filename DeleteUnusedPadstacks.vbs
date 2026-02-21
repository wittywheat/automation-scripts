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
'   Author: Nadia Ahmad
'   App:    None
'   O/S:	Windows XP
'   Lang:	VBScript
'   Status: Work in Progress
'   Date:   05/03/2010
'   Description: This script deletes all unused Padstacks in the design.  
'                This script is provided in lieu of DR 148533 -"Description: Using the Forward Annotation option 'Delete Local Library data; 
'                then rebuild all local library data' does not delete the unused Padstack data".  
'                Note: If a padstack is not placed in the design but used in a local cell (unused cell), then that padstack is not deleted unless that unused cell is
'					   deleted first.  This is why it could be necessary to first run Forward Annotation with "Delete local data and rebuild...." option first to remove all 
'                      unused Parts and Cells before running this script.  
'                      If a padstack is used in Setup_Parameters/NetClasses/CES but is not placed in the design, then that padstack is not deleted, unless it is removed from
'                      NetClasses/CES-->Via Assigment tab OR from Allowed via list in Setup Parameters dialog.  
'==========================================================================


Option Explicit   ' This means that all variables must be declared using the 'Dim' statement

' Get the application object
Dim pcbApp
Set pcbApp = Application
Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")

' Get the active document of Expedition PCB
Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument

'---------------------------------------
'Stup Main program
'---------------------------------------
Sub main ()
	On Error Resume Next

	pcbApp.lockserver
	pcbDoc.transactionStart
	
	Dim pstkEditor, pstkDB, pstk, pstkinDesign, unusedpstk, ps
	
	' Getting the Padstack Editor dialog object.  
	Set pstkEditor = pcbDoc.PadstackEditor
	
	pstkEditor.visible = 0
	
	' Lock the Active Padstack Server for writing.  
	pstkEditor.LockServer  
	
	' Getting the active Padstack Database
	Set pstkDB = pstkEditor.ActiveDatabase

	unusedpstk = 0 ' This variable will help us track the padstack which is not used in the design. 
	
	file.writeline("Deleted Padstack: ")
	file.writeline
	
	For Each pstk In pstkDB.padstacks ' Looping through each padstack in the Local Padstack Editor
		
		unusedpstk = 0 
		
		For Each ps In pcbDoc.padstacks ' Looping through each padstack in the Design.
		
			If pstk.name <> ps.name Then  
			
			else
				
			 	unusedpstk = 1 ' Flag to check that this padstack does exists in design.  
			 	Exit For
			End If 
		
		Next 
		
		If unusedpstk = 0 then 
		    file.writeline(pstk.name)
		    file.writeline
		    'msgbox(pstk.name)
			pstk.delete
			If (Err) Then
				Call pcbApp.Gui.StatusBarText("An error occured while deleting padstacks. Please see \PCB\DeletedPadstackReport.txt" , epcbStatusFieldError)
				file.writeline("An error occured while deleting (" + pstk.name + ") padstack.  It is either used in an unused cell in Local Cell Editor or this padstack is being referenced in Setup Parameters or NetClasses or CES.")
				file.writeline("To remove the unused package cell, run Forward Annotation with 'Delete local data and rebuild all local library data'.")
				file.writeline("Remove references to this via from Setup Parameters/NetClasses/CES if needs to be deleted.")
				file.writeline
			End If

			
		End If
	Next
	
	pstkEditor.SaveActiveDatabase
	pstkEditor.UnlockServer
	pstkEditor.Quit
	pcbDoc.transactionEnd
	pcbApp.unlockServer

End Sub

' License the document
If (ValidateServer(pcbDoc) = 1) Then
	'define a filename and concatenate it with pcb dir path
	Dim fso,sFilename
	sFilename=pcbDoc.Path & "\DeletedPadstackReport.txt"
	
	'Create & open the output text file for writing
	Set fso = CreateObject("Scripting.FileSystemObject")
	Dim file: Set file = fso.CreateTextFile(sFilename, True)

	Call main
	
	'Close and open file
	file.Close
    Dim WshShell
    Set WshShell = CreateObject("WScript.Shell")
	WshShell.Run sFilename

End If
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