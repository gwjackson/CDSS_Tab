; find the HCC descriptor in a large text file
;
PendHCC := "D69.6,E11.29,E11.319,E11.3299,E11.3599,E11.42,E11.49,E11.51,N18.2"
; get rid of the "."
StringReplace, PendHCC, PendHCC, . , , All
;Loop, Read, C:\Users\gjackson\Documents\AHK-Scrips\CDSSTab\HCC_Descriptors.txt
HCC_String := "Format: `r`nHCC`tICD`tDescriptor`r`n"
Loop, Parse, PendHCC, `, 
{
	ICD := A_LoopField
	loopcount := A_Index
	Loop, Read, HCC_Descriptors.txt
	{
		IfInString, A_LoopReadLine, %ICD%
		{
			HCC_String := HCC_String . "`r`n" . A_LoopReadLine
			break
		}
	}
}
; 
gui, add, text, w580,  %HCC_String%
Gui, Add, Button, gGuiClose , CLOSE
gui, show, Autosize Center, Pending HCC with Descriptors
return
;
GuiClose:
ExitApp