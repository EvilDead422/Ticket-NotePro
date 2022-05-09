#SingleInstance, Force

/*
--------------------------------------------------------------------------
Menu Items
--------------------------------------------------------------------------
*/
;~ Sub Menu Items
Menu, FileMenu, Add, &New, ClearForm
Menu, FileMenu, Add, Save &As, SaveAsTicket
Menu, FileMenu, Add, &Load, LoadTicket
Menu, FileMenu, Add  ; Dividing line.
Menu, FileMenu, Add, E&xit, FileExit

Menu, HelpMenu, Add, &About, HelpAbout

Menu, ContactMenu, Add, &Save, SaveContact
Menu, ContactMenu, Add, &Load, LoadContact

;~ Menu Options
Menu, DeafultMenu, Add, &File, :FileMenu
Menu, DeafultMenu, Add, &Contacts, :ContactMenu
Menu, DeafultMenu, Add, &Help, :HelpMenu

;~ Attaching menu to Main Gui
Gui, Main:Menu, DeafultMenu
/*
--------------------------------------------------------------------------
Main Gui Outline
--------------------------------------------------------------------------
*/
Gui, Main:Color, 1DB5E6, A8F3F3, 

;~ Ticket Number Field
	Gui, Main: Add, Text, x12 y9 w80 h20 , Ticket #:
	Gui, Main: Add, Edit, x67 y9 w90 h20 +Center vTNum, 
	Gui, Main: Add, Button, x160 y9 w70 h20 gSaveAsTicket, Save Ticket
	Gui, Main: Add, Button, x235 y9 w70 h20 gLoadTicket, Load Ticket
	
;~ User Information Fields
Gui, Main: Add, GroupBox, x7 y29 w300 h110 , User Information:

	Gui, Main: Add, Text, x12 y49 w55 h20 , First Name:
	Gui, Main: Add, Edit, x67 y49 w90 h20 vUser_Field_1, 
		
	Gui, Main: Add, Text, x158 y49 w60 h20 , Last Name:
	Gui, Main: Add, Edit, x212 y49 w90 h20 vUser_Field_2, 
		
	Gui, Main: Add, Text, x12 y79 w60 h20 , Institution:
	Gui, Main: Add, Edit, x67 y79 w235 h20 vUser_Field_3, 
		
	Gui, Main: Add, Text, x12 y109 w90 h20 , Callback #:
	Gui, Main: Add, Edit, x67 y109 w121 h20 vUser_Field_4, 
	
	Gui, Main: Add, Button, x245 y109 w55 h20 gLoadContact, Load User
	Gui, Main: Add, Button, x190 y109 w50 h20 gCopyNumber, Copy #
		
;~ Ticket Information Fields
Gui, Main: Add, GroupBox, x7 y139 w300 h110 , Ticket Information:
		
	Gui, Main: Add, Text, x12 y159 w70 h20 , Current Issue:
	Gui, Main: Add, Edit, x82 y159 w210 h20 vIssue, 
	
	Gui, Main: Add, Text, x12 y189 w70 h20 , Started Work:
	Gui, Main: Add, Edit, x82 y189 w145 h20 vStWk, 
	
	Gui, Main: Add, Button, x232 y189 w60 h20 gInputTime, Input Time
	
;~ Work Done Field
	Gui, Main: Add, Text, x312 y9 w60 h20 , Work Done:
	Gui, Main: Add, Edit, x312 y25 w460 h180 vWkDn, 
	
;~ Resolution Field
	Gui, Main: Add, Text, x312 y210, Resolution:
	Gui, Main: Add, Edit, x312 y226 w460 vResolution,
	
;~ Case Status DropDown
	Gui, Main: Add, Text, x12 y219 w100 h20 , Case Status:
	Gui, Main: Add, DropDownList, x82 y219 w210 h100 sort vStatus, Closing as "Issue Resolved"|Ticket is still "In Progress"||Placing in "Assigned"|Placing in "Waiting on Client"|Placing in "Waiting on Vendor"|Going to "Escalate to Management"|Going to "Escalate to Tier 2"|Ticket "Needs On Site Scheduled"
	
;~ Template Chooser Dropdown
	Gui, Main: Add, Text, x312 y252 w60 h20 , Template:
	Gui, Main: Add, DropDownList, x362 y252 w120 h100 sort vTemplate, Work Done||Closing Note|Chat|Email|

;~ Write Fast Box
	Gui, Main: Add, CheckBox, x642 y252 w65 h20 vWriteFast, Write Fast
	
;~ Technician Fileds
	Gui, Main: Add, Text, x12 y252 w70 h20, Technician:
	Gui, Main: Add, Edit, x72 y252 w100 h20 vTech, 
	
;~ Bottom Right Buttons
	Gui, Main: Add, Button, x712 y252 w60 h20 Default gWriteIt, Write It
	Gui, Main: Add, Button, x575 y252 w60 h20 gCopyIt, CopyIt
	Gui, Main: Add, Button, x712 y3 w60 h20 gClearForm, Clear Form

GuiControl, Main: Focus, User_Field_1
Gui, Main: Show, w780 h280 , HD Ticket NotePro

return
/*
--------------------------------------------------------------------------
Functions:
--------------------------------------------------------------------------
*/

CopyNumber:
	Gui, Main: Submit, NoHide
	Clipboard = %User_Field_4%
return

CopyIt:
Gui, Main: Submit, NoHide
If(Template="Work Done")
{
	gosub, CopyWorkDone
}
else if(Template="Closing Note")
{
	gosub, CopyClosingNote
}
else if(Template="Email")
{
	gosub, EmailGUI
}
else
{
	gosub, ChatGUI
}
return 

CopyWorkDone:
Gui, Main: Submit, Nohide

WorkTemp =
(
Ticket: %TNum%
User: %User_Field_1% %User_Field_2%
Institution: %User_Field_3%
Started Work: %StWk%
Current Issue: %Issue%
Callback Number: %User_Field_4%

Work Done:
%WkDn%

Signed: %Tech% - %Status%
)

Clipboard := WorkTemp

return

CopyClosingNote:
Gui, Main: Submit, Nohide
ResTemp =
(
User: %User_Field_1% %User_Field_2%
Summary of Issue: %Issue%
Summary of Resolution: %Resolution%

Thanks for working with us on this issue. If this issue persists, please feel free to reply to this email directly with any further information about the issue or it's continuance.

Thanks,
%Tech% - Closing as "Issue Resolved"
)
Clipboard := ResTemp
return

SaveTicket:
return

HelpAbout:
;~ MsgBox [, Options, Title, Text, Timeout]
	AboutMsg =
(
This program was desinged by EvilDead. 

It's original intention was to be used as a supplimentary notation device for Help Desk troubleshooting. The hope is that the software is maintianed for such practice.

It was designed originally in January of 2022.
)
	MsgBox,, About, %AboutMsg%,
return

SaveAsTicket:
	;~ SplitPath, InputVar [, OutFileName, OutDir, OutExtension, OutNameNoExt, OutDrive]
	Gui, Main: Submit, NoHide
	
	TicketName = %User_Field_3% - %TNum% - %Issue%
	
	FileSelectFile, Save_Ticket, S16, %TicketName%
		if(ErrorLevel)
			return
	SplitPath, Save_Ticket,,, ext1
	if (ext1 = "")
		Save_Ticket .= ".txt"
	FileDelete, % Save_Ticket
	FileAppend, % TNum "¥" User_Field_1 "¥" User_Field_2 "¥" User_Field_3 "¥" User_Field_4 "¥" Issue "¥" StWk "¥" WkDn "¥" Resolution, % Save_Ticket
	SplitPath, Save_Ticket, , , , File_Name,
	
return 

SaveContact:
	
	Gui, Main: Submit, NoHide
	
	ContactName = %User_Field_3% - %User_Field_1% %User_Field_2%
	
	FileSelectFile, Save_File_Name, S16, %ContactName%
	if(ErrorLevel)
			return
	SplitPath, Save_File_Name,,, ext
	if (ext = "")
		Save_File_Name .= ".txt"
	FileDelete, % Save_File_Name
	FileAppend, % User_Field_1 "¥" User_Field_2 "¥" User_Field_3 "¥" User_Field_4, % Save_File_Name

return

LoadTicket:

	Gui, Main: Submit, NoHide
	
	FileSelectFile, Load_Ticket,, 
	if(ErrorLevel)
			return
	FileRead, Temp_File_Data, % Load_Ticket
	Loop, Parse, Temp_File_Data, ¥
		{
			Loop_Var_%A_Index% := A_LoopField
		}
		
	GuiControl,, TNum, % Loop_Var_1
	GuiControl,, User_Field_1, % Loop_Var_2
	GuiControl,, User_Field_2, % Loop_Var_3
	GuiControl,, User_Field_3, % Loop_Var_4
	GuiControl,, User_Field_4, % Loop_Var_5
	GuiControl,, Issue, % Loop_Var_6
	GuiControl,, StWk, % Loop_Var_7
	GuiControl,, WkDn, % Loop_Var_8
	GuiControl,, Resolution, % Loop_Var_9
	GuiControl,, WriteFast, 0
	
	GuiControl, Main: Focus, User_Field_1
	SplitPath, Load_Ticket, , , , File_Name,
	
	return

return 

LoadContact:
	;~ FileSelectFile, OutputVar [, Options, RootDir\Filename, Title, Filter]
	Gui, Main: Submit, NoHide
	
	FileSelectFile, Load_File_Name,, 
	if(ErrorLevel)
			return
	FileRead, Temp_File_Data, % Load_File_Name
	Loop, Parse, Temp_File_Data, ¥
		{
			User_Field_%A_Index% := A_LoopField ;A_Index counts each time it collects a variable and adds 1 < This avoids counting variables
			GuiControl,,User_Field_%A_Index%, % User_Field_%A_Index%
		}
	return
	
return

WriteIt:

	Gui, Main: Submit, NoHide
	If(WriteFast=1)
	{
		Gosub TemplatesFast
	}
	else
	{
		Gosub TemplatesSlow
	}
	
return

TemplatesFast:

	Gui, Main: Submit
	If(Template="Work Done")
		{
		SendInput, ^bUser^b: %User_Field_1% %User_Field_2%{Enter}
		SendInput, ^bInstitution^b: %User_Field_3% {Enter}
		SendInput, ^bStarted Work^b: %StWk% {Enter}
		SendInput, ^bCurrent Issue^b: %Issue% {Enter}
		SendInput, ^bCallback Number^b: %User_Field_4% {Enter}
		SendInput, {Enter}
		SendInput, ^bWork Done^b: {Enter}%WkDn% {Enter}
		SendInput, {Enter}
		SendInput, ^bSigned^b: ^u%Tech% - %Status% ^u
		
		Gui, Main: Show
		}
	Else if(Template="Closing Note")
		{
		SendInput, ^bUser^b: %User_Field_1% %User_Field_2%{Enter}
		SendInput, ^bSummary of Issue^b: %Issue% {Enter}
		SendInput, ^bSummary of Resolution^b: %resolution% {Enter}
		SendInput, {Enter}
		SendInput, Thanks for working with us on this issue. If it persists please feel free to reply to this email directly with any further information about the issue or it's continuance.{Enter}
		SendInput, {Enter}
		SendInput, Thanks,{Enter}
		SendInput, ^u%Tech% - Closing as "Issue Resolved"^u
		
		Gui, Main: Show
		}
	Else if(Template="Chat")
		{
			gosub, ChatGUI
		}
	Else 
		{
			gosub, EmailGUI
		}
	
return

TemplatesSlow:

	Gui, Main: Submit
	If(Template="Work Done")
		{
		Send, ^bUser^b: %User_Field_1% %User_Field_2%{Enter}
		Send, ^bInstitution^b: %User_Field_3% {Enter}
		Send, ^bStarted Work^b: %StWk% {Enter}
		Send, ^bCurrent Issue^b: %Issue% {Enter}
		Send, ^bCallback Number^b: %User_Field_4% {Enter}
		Send, {Enter}
		Send, ^bWork Done^b: {Enter}%WkDn% {Enter}
		Send, {Enter}
		Send, ^bSigned^b: ^u%Tech% - %Status% ^u
		
		Gui, Main: Show
		}
	Else if(Template="Closing Note")
		{
		Send, ^bUser^b: %User_Field_1% %User_Field_2%{Enter}
		Send, ^bSummary of Issue^b: %Issue% {Enter}
		Send, ^bSummary of Resolution^b: %resolution% {Enter}
		Send, {Enter}
		Send, Thanks for working with us on this issue. If it persists please feel free to reply to this email directly with any further information about the issue or it's continuance.{Enter}
		Send, {Enter}
		Send, Thanks,{Enter}
		Send, ^u%Tech% - Closing as "Issue Resolved"^u
		
		Gui, Main: Show
		}
	Else if(Template="Chat")
		{
			gosub, ChatGUI
		}
	Else 
		{
			gosub, EmailGUI
		}
	
return

ChatGUI:
;~ Chat GUI-----------------------------------------------------------------
	Gui, Main: Submit
	Gui, Chat:Color, 1DB5E6, A8F3F3, 
	
	Gui, Chat:Add, Text, x12 y9 w80 h20 +Right, User:
	Gui, Chat:Add, Text, x12 y39 w80 h20 +Right, Ticket Number:
	Gui, Chat:Add, Text, x12 y69 w80 h20 +Right, Institution:
	Gui, Chat:Add, Text, x12 y99 w80 h20 +Right, Current Issue:
	Gui, Chat:Add, Text, x12 y129 w80 h20 +Right, Troubleshooting:
	
	Gui, Chat:Add, Edit, x102 y9 w120 h20 vUserC, %User_Field_1% %User_Field_2%
	Gui, Chat:Add, Edit, x102 y39 w120 h20 vTNumC, %TNum%
	Gui, Chat:Add, Edit, x102 y69 w120 h20 vInstC, %User_Field_3%
	Gui, Chat:Add, Edit, x102 y99 w240 h20 vIssueC, %Issue%
	Gui, Chat:Add, Edit, x12 y149 w330 h150 vTrouble, %WkDn% 
	
	Gui, Chat:Add, Button, x262 y309 w80 h20 gWriteChat, Write It
	Gui, Chat:Add, Button, x142 y309 w80 h20 gCopyChat, Copy It
	Gui, Chat:Add, Button, x12 y309 w80 h20 gCancelChat, Cancel

	Gui, Chat: Show, w356 h344, Chat Template
	Gui, Chat: +AlwaysOnTop
	Gui, Main: Hide	
	GuiControl, Chat:Focus, Trouble
	
	

return
	
CopyChat:
Gui, Chat: Submit

ChatTemp =
(
User: %UserC%
Ticket Number: %TNumC%
Institution: %InstC%
Current Issue: %IssueC%
Troubleshooting:
%Trouble%
)

Clipboard := ChatTemp
Gui, Chat: Destroy
Gui, Main: Show
return
	
WriteChat:

	Gui, Chat: Submit
		
		SendInput, ^bUser^b: %UserC% {Enter}
		SendInput, ^bTicket Number^b: %TNumC% {Enter}
		SendInput, ^bClient Institution^b: %InstC% {Enter}
		SendInput, ^bDescription of Issue^b: %IssueC% {Enter}
		SendInput, ^bTroubleshooting Summary^b: {Enter} %Trouble%

	Gui, Chat: Destroy
	Gui, Main: Show
	
return

EmailGUI:
;~ Email GUI-----------------------------------------------------------------
	Gui, Main: Submit,
	Gui, Email:Color, 1DB5E6, A8F3F3, 
	
	Gui, Email:Add, Text, x12 y19 w20 h20 +Right, Hey
	Gui, Email:Add, Edit, x42 y19 w120 h20 vUserE, %User_Field_1%
	Gui, Email:Add, Text, x12 y42 w60 h20 , Email Body:
	Gui, Email:Add, Edit, x12 y59 w290 h110 vBody, 
	Gui, Email:Add, Text, x12 y179 w40 h20 , Thanks`,
	Gui, Email:Add, Text, x12 y199 w100 h20 , %Tech%
	Gui, Email:Add, Button, x232 y199 w70 h20 gWriteEmail, Write It
	Gui, Email:Add, Button, x232 y175 w70 h20 gCopyEmail, Copy It
	Gui, Email:Add, Button, x152 y199 w70 h20 gCancelEmail, Cancel

	Gui, Email:Show, w319 h233, Email Template
	Gui, Email:+AlwaysOnTop
	Gui, Main: Hide
	GuiControl, Email: Focus, Body
	if(ErrorLevel)
		Gui, Main: Show
			return
	
return

CopyEmail:
Gui, Email: Submit
Gui, Main: Submit

EmailTemp =
(
Hey %UserE%,

%Body%

Thanks,
%Tech%
)

Clipboard := EmailTemp
Gui, Email: Destroy
Gui, Main: Show
return

WriteEmail:

	Gui, Email: Submit
		
		send, Hey %UserE%, {Enter}
		send, {Enter}
		send, %Body% {Enter}
		send, {Enter}
		send, Thanks, {Enter}
		send, %Tech% {Enter}
	
	Gui, Email: Destroy
	Gui, Main: Show
	
return

CancelEmail:

	Gui, Email: Destroy
	Gui, Main: Show

return

CancelChat:

	Gui, Chat: Destroy
	Gui, Main: Show

return

ClearForm: 
	
	GuiControl,, TNum
	GuiControl,, User_Field_1
	GuiControl,, User_Field_2
	GuiControl,, User_Field_3
	GuiControl,, User_Field_4
	GuiControl,, Issue
	GuiControl,, StWk
	GuiControl,, WkDn
	GuiControl, Choose, Status, Ticket is still "In Progress"
	GuiControl, Choose, Template, Work Done
	GuiControl,, Resolution
	
	GuiControl, Main: Focus, User_Field_1
	
return

InputTime:
	
	FormatTime, DateTime,, MMMM d, yyyy @ h:mm tt
	GuiControl, Text, StWk, %DateTime%
	return
	
return

FileExit:
GuiClose:
ExitApp