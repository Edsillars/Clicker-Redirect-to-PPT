;Based on code from https://github.com/4000degrees/ClickerOverride and https://www.autohotkey.com/board/topic/108575-how-to-detect-powerpoint-slide-number-using-ahk/

;Combined and edited by Edward Sillars, Acorn AV. London

;  Updated how it identifies the Powerpoint window, changed it from the ID which is unique  - to any window that contains PowerPoint Presenter View - (there is usually only 1), 
;  This enables it to auto find the window and you can open up different presentations without having to refresh the program
;  There was a toggle button that allowed automatically focus on the powerpoint window on click, reprogrammed that to be a text box array that you can specify which slide the videos are on so it just brings those to focus.
;  Took other code that talks to the powerpoint application to find out what slide it is currently on, and used that to compare against the text box array to identify if the current slide is equal to a slide that a video is on, then put powerpoint in focus mode
;  Expanded the key commands, it just covered page up and down, but some clickers work on left, right, up and down so I have included all of those in the program. Added check boxes so you can enable or disable them.


#SingleInstance Force
#Persistent


  objPPT := ComObjCreate("PowerPoint.Application")
  ComObjError(false)

; Sets the title to include any of the text
SetTitleMatchMode, 2



WinGet AllWindows, List, PowerPoint Presenter View,
Gui, Add, Text,, Select which keys you would like redirected to PowerPoint Presenter View Window (Please note this only works when Presenter Mode is enabled)


Gui, Add, CheckBox, vPuPd checked, Page Up and Page Down Redirected
Gui, Add, CheckBox, vLeftandRight checked, Left and Right Redirected
Gui, Add, CheckBox, vUpandDown, Up and Down Redirected
Gui, Add, Text,, 
Gui, Add, Text,, PowerPoint will not play a video with other windows infront of it. 
Gui, Add, Text,, Below type the slide numbers seperated by a comma for the slides with videos on. Powerpoint will be bought to the front to play the video.
Gui, Add, Text,, e.g. 3,6,12,18
Gui, Add, Edit, vvideoslide


Gui, Add, Text,, 
Gui, Add, Text,, Based on code from https://github.com/4000degrees/ClickerOverride 
Gui, Add, Text,, Edited by Ed Sillars, Acorn AV London.




Gui, Show


return



GuiClose:
	ExitApp
	return



GetSelectedWindowId() {
	Gui, Submit, NoHide
	WinGet AllWindows, ID, PowerPoint Presenter View,
	id := AllWindows
	return %id%
}



SelectedPgUpPgDn() {
	Gui, Submit, NoHide
	global PuPd
	if PuPd
		msgbox, pg is selected
		PressPgUp()
}

PressPgUp(){
Checkifvideotofocus()
}





Checkifvideotofocus(){

	current := objPPT.SlideShowWindows(1).View.Slide.SlideIndex

Loop, parse, videoslide, `,
{
loopnumber = %A_LoopField%

if (current = loopnumber){

WinActivate % "ahk_id" GetSelectedWindowId()
}
else{

}

}
}

; PgUp only works with bluetooth clickers, for interspace and others need left/right and also will program up and down as a backup so any of those 6 will work for forwards and back in the same way.

*PgUp::
	
Gui, Submit, NoHide
	global PuPd
	if (PuPd = 1){
		
		ControlSend,ahk_parent,{PgUp},% "ahk_id" GetSelectedWindowId()
		current := objPPT.SlideShowWindows(1).View.Slide.SlideIndex

		Loop, parse, videoslide, `,
			{
			loopnumber = %A_LoopField%

			if (current = loopnumber){

			WinActivate % "ahk_id" GetSelectedWindowId()
				}
			else{

				}

			}
		}
	else{


	}
return

*Up::
	
Gui, Submit, NoHide
	global UpandDown
	if (UpandDown = 1){
		
		ControlSend,ahk_parent,{Up},% "ahk_id" GetSelectedWindowId()
		current := objPPT.SlideShowWindows(1).View.Slide.SlideIndex

		Loop, parse, videoslide, `,
			{
			loopnumber = %A_LoopField%

			if (current = loopnumber){

			WinActivate % "ahk_id" GetSelectedWindowId()
				}
			else{

				}

			}
		}
	else{


	}
return


*Right::
	
Gui, Submit, NoHide
	global LeftandRight
	if (LeftandRight = 1){
		
		ControlSend,ahk_parent,{Right},% "ahk_id" GetSelectedWindowId()
		current := objPPT.SlideShowWindows(1).View.Slide.SlideIndex

		Loop, parse, videoslide, `,
			{
			loopnumber = %A_LoopField%

			if (current = loopnumber){

			WinActivate % "ahk_id" GetSelectedWindowId()
				}
			else{

				}

			}
		}
	else{


	}
return




*PgDn::
	
Gui, Submit, NoHide
	global PuPd
	if (PuPd = 1){
		
		ControlSend,ahk_parent,{PgDn},% "ahk_id" GetSelectedWindowId()
		current := objPPT.SlideShowWindows(1).View.Slide.SlideIndex

		Loop, parse, videoslide, `,
			{
			loopnumber = %A_LoopField%

			if (current = loopnumber){

			WinActivate % "ahk_id" GetSelectedWindowId()
				}
			else{

				}

			}
		}
	else{


	}
return


*Down::
	
Gui, Submit, NoHide
	global UpandDown
	if (UpandDown = 1){
		
		ControlSend,ahk_parent,{Down},% "ahk_id" GetSelectedWindowId()
		current := objPPT.SlideShowWindows(1).View.Slide.SlideIndex

		Loop, parse, videoslide, `,
			{
			loopnumber = %A_LoopField%

			if (current = loopnumber){

			WinActivate % "ahk_id" GetSelectedWindowId()
				}
			else{

				}

			}
		}
	else{


	}
return


*Left::
	
Gui, Submit, NoHide
	global LeftandRight
	if (LeftandRight = 1){
		
		ControlSend,ahk_parent,{Left},% "ahk_id" GetSelectedWindowId()
		current := objPPT.SlideShowWindows(1).View.Slide.SlideIndex

		Loop, parse, videoslide, `,
			{
			loopnumber = %A_LoopField%

			if (current = loopnumber){

			WinActivate % "ahk_id" GetSelectedWindowId()
				}
			else{

				}

			}
		}
	else{


	}
return

	















