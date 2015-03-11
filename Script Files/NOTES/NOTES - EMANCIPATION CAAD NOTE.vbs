option explicit

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO, beta_agency					'Declares variables to be good to option explicit users
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

DIM  ButtonPressed, emancipation_dialog, anticipated_graduation_date, enrollment_verified_checkbox, worker_signature
EMConnect ""	'Connects to PRISM
EMFocus


BeginDialog emancipation_dialog, 0, 0, 226, 105, "Emancipation Dialog"		'Beginning of Dialog
  ButtonGroup ButtonPressed
    OkButton 105, 85, 50, 15
    CancelButton 160, 85, 50, 15
  EditBox 115, 15, 75, 15, anticipated_graduation_date
  Text 10, 15, 95, 115, "Anticipated Graduation Date"
  CheckBox 115, 40, 115, 15, "Enrollment in school verified", enrollment_verified_checkbox
  Text 110, 65, 65, 15, "Worker Signature"
  EditBox 175, 65, 45, 15, worker_signature
EndDialog

DO		'Beginning of Looping process			
	dialog emancipation_dialog	'Dialog always in between the DO and the LOOP UNTIL 
	IF worker_signature = "" THEN MsgBox "You must sign!"
LOOP UNTIL worker_signature <> ""

	
IF ButtonPressed = 0 THEN StopScript	


CALL navigate_to_PRISM_Screen ("CAAD")	'Takes you to CAAD
		
PF5		'Pulls up blank CAAD note

EMWriteScreen "A", 3, 2

EMWriteScreen "FREE", 4, 54

EMSetCursor 16, 4

CALL write_bullet_and_variable_in_CAAD ("Anticipated Graduation Date", anticipated_graduation_date)		'Copies the information provided via dialog to the CAAD note
IF enrollment_verified_checkbox = checked THEN CALL write_bullet_and_variable_in_CAAD ("Enrollment Verified", "Checked.")
CALL write_bullet_and_variable_in_CAAD ("Worker Signature", worker_signature) 

StopScript









