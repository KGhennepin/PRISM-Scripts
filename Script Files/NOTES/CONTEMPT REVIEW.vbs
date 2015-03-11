option explicit

Dim Beta_Agency

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO				'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF

BeginDialog Contempt_Dialog, 0, 0, 361, 170, "Contempt Review"
  Text 5, 10, 50, 10, "Contempt Review "
  EditBox 65, 5, 105, 15, Contempt_review 
  Text 5, 30, 50, 10, "DLS status "
  DropListBox 60, 30, 105, 15, "<select one>"+chr(9)+"Suspended"+chr(9)+"Revoked"+chr(9)+"Payment Plan", DLS_Status
  Text 5, 45, 40, 15, "DC file review "
  DropListBox 55, 45, 105, 10, "<select one>"+chr(9)+"Yes"+chr(9)+"No", DC_File_Review
  Text 5, 70, 60, 15, "last payment date"
  EditBox 95, 70, 80, 15, Date
  Text 5, 90, 60, 15, "Previous Contempt"
  DropListBox 75, 85, 100, 15, "Yes"+chr(9)+"No"+chr(9)+"Dismissed", Previous_Contempt
  ButtonGroup ButtonPressed
    OkButton 250, 150, 50, 15
    CancelButton 305, 150, 50, 15
EndDialog

 
EMConnect "" 		'Connects to prism

DIM Contempt_Dialog, Contempt_Review, Case_Number, DLS_Status, DC_File_Review, Previous_Contempt, Date, ButtonPressed 

Do	
	Dialog Contempt_Dialog 'ties the dialog from above 
IF ButtonPressed = 0 THEN StopScript 'this will stop the script on it's own 
	IF DC_File_Review = "<select one>" THEN Msgbox "DC File Review must be Complete" 'message box will pop up if 
LOOP UNTIL Dc_File_Review <> "<select one>" 'this will cause the message box to pop up until the user completes the field 



CALL check_for_prism (true)		'Checks for password 

CALL navigate_to_prism_screen ("CAAD")	'Goes to CAAD

PF5		'Creates a new note

EMwritescreen "a", 3, 29 	'Add Mode

EMwriteScreen "Free", 4, 54 	'Set the Cursor on the CAAD note area

EmsetCursor 16, 4		'Header (places cursor in correct position 

CALL write_variable_in_caad (Contempt_Review)
CALL write_bullet_and_variable_in_caad ("dls status", DLS_Status)
CALL write_bullet_and_variable_in_caad ("dc file review", DC_File_Review) 
CALL write_bullet_and_variable_in_caad  ("Last Payment Date", date)
CALL write_bullet_and_variable_in_caad ("Previous contempt", previous_contempt)   'quotes are use to define" and underscore is example = defines my variable
EMSendKey "<enter>" ' allows system to press enter without the user 
EMWaitReady 0, 0


Request auto fill names, roles (ncp/cp), 




















