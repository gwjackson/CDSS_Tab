;
; Metrics/FOC GUI 
; Walker Jackson, MD 5/20/17
; !m::
;
sVersion := "01.08.15"
;
;04jul2017 MOR 
;	Additional environment settings
;	Handle missing default file location
;	Support .ini file if present
;		If none of the above, then let user select report file location and then save in .ini file
;	Spell check with various spelling corrections
;	Expand name selection form
;	Correct alignment of run date in GUI form
;	Add version number to GUI form
;04jul2017 MOR
;21jul2017 MOR
;	Changes to support new HCC columns 
;		Addressed HCC - K
;		Pending HCC	  - L
;
;		NB: a new column Z was added. Adjust following columns accordingly
;

;04jul2017 MOR
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn UseUnsetLocal ; Enable warnings to assist with detecting common errors.
#SingleInstance Force

SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; *****************************
; findFilePath is in the Scripts\CDSS\Lib dir 
;

global oPBR

;
; 
; User selects and copies the pt. name to the Clipboard and then fires the 
;	script in FK to call this script
;
; Name in PBR is LastName, FirstName
; this is assumption of FindThis for now; may I will need to add code to validate this
;********************
; Test read off the clipboard
myClipboard := Clipboard		; preserve the clipboard contents (at least the Text)
clipboard := ""					; clear the clipboard 
Send ^{vk43}					; Ctrl C (copy) 
ClipWait						;give chance for clipboard to get the data
FindThis := clipboard
clipboard := myClipboard		; restore the clipboard
; ========================
; RegEx test 8/15/17;
; capture the Last name part 
FoundPos := RegExMatch(FindThis, "^(([A-Za-z\- ]+\s?)?)", ptName)
if (FoundPos > 0) {
	FindThis := ptName1
	;MsgBox %FindThis%
} else {
	MsgBox, 
		(
		Sorry not able to find this patient 
		Try just the Last Name, or Fisrt Name
		)
}
; ==========================
;
; MsgBox %FindThis%
; Test "FindThis" pre-defined BROWN, ERNEST or whatever
; FindThis := "jesus"  ; the search term 
;
; How to access Workbook without opening it?
; https://autohotkey.com/board/topic/56987-com-object-breference-autohotkey-v11/:
; //autohotkey.com/board/topic/56987-com-object-reference-autohotkey-v11/page-4?&#entry381256
; 
 ; *************************************
 ; 30July2017 GWJ
 ; get the "FilePath" to the  "OPNE" PRB Excel worksheet
FilePath := findFilePath()
;
; try w/user selection
oPBR := ComObjGet(FilePath) ; access the selected workbook object 
; to test for it to be an Object  	
loopCount := 0
while !isObject(oPBR) 
{
	SplashTextOn,300,20,Opening and reading Excel
	Sleep, 200					; wait 20 mili-seconds for file access, Excel can be really slow
	loopCount =+ 1
	if (loopCount > 10) {
		SplashTextOff
		MsgBox, % "Can't locate or open this file name: " FilePath
		goto, GuiClose
		return
	}
	
}
; succeded!
SplashTextOff
;
;}
;04jul2017 MOR
;***********************************************
;
LastRow := oPBR.Sheets(1).UsedRange.Rows.Count	; gets the last row of the sheet  
;MsgBox % Last row is LastRow 
; for the PBR the data starts are Row - 20 
NameRange := oPBR.Sheets(1).Range("A20:A" LastRow)	;range to search 
;
LastCell := NameRange.Cells(NameRange.Cells.Count)	;last cell in NameRange 
;
;
FoundCell := NameRange.Find(FindThis) 		; consider; LookIn:=x1Values, LookAt := x1Whole
; test for failed search 
if (!FoundCell) {
	MsgBox, Sorry search for %FindThis%, failed`n`nTry just the last name or the first name`nThe PRP has names in the 'LastName, Firstname' format`r`n`r`nThis data base contains only Medical Advantage patients
	goto, GuiClose
	return 
}
FirstAddr := FoundCell.Address
;MsgBox, % "The first found cell in the range is " FirstAddr " with a value of '" FoundCell.Value "'."
; **********
; if only one pt
ptRowNumber:= SubStr(FoundCell.Address, 4)
; **********
;MsgBox, % "Pt row number is " ptRowNumber
; start to collect the ListBox options 
;
ptArray := Object()
ptArray[1] := FoundCell
ptCount :=1
;repeat the search (FoundCell vs FindThis)
try while FoundCell := NameRange.FindNext(FoundCell)
{
	if (FoundCell.address = FirstAddr)  ; loop has wrapped around 
		break
	ptCount += 1
	ptArray[ptCount] := FoundCell
;
;MsgBox, % "The next cell in the range is`r`n" ptArray[ptCount].Address "`nwith a value of`r`n" ptArray[ptCount].Value "`r`nand at cell address of " ptArray[ptCount].Address
;
}
;
aryLength :=  ptArray.MaxIndex() 
;MsgBox, % "the pt. array length is " aryLength
;set up the list box but skip if only one choice;
;FoundPos := InStr(FoundCell.Value, "-")
;match1 := SubStr(FoundCell.Value, 1, FoundPos - 2)
;
if (aryLength > 1) {
	ptChoiceList := ""
	Loop, %aryLength%
		{
;		MsgBox, % ptArray[A_Index].Value
		FoundPos := InStr(ptArray[A_Index].Value, "-")
;		MsgBox, % "Position for " A_Index " is " FoundPos
		match1 := SubStr(ptArray[A_Index].Value, 1, FoundPos - 2)
;		MsgBox, % "match1 is " match1
		ptChoiceList := ptChoiceList match1 "|" 
;		MsgBox, % "Pt. choice list is " ptChoiceList
		}
;
	;MsgBox, %ptChoiceList%
	; got the "|" delimited list of pt's 
	; function call here
	GetUserChoice()
;
;	wait for the user to select their patient 
	WinWaitNotActive, Select your patient
	;
}
;There is some kind of rare condition in the HCC split logic if don't have to ask for a choice
Sleep, 1000	;Wait one second
;
; cont. of after multichoice - 
; Preceed to the data extraction and Gui creation 
; Read the Row of Metrics
; box colors decimal Yellow = 13826810; Fuchsia = 13353215; Red = 255,
; text in the cell and 
; check colors to adjust the GUI text colors
; 
;MsgBox The Patient's row number for data extractions is `n %ptRowNumber%
;
sPtDemo := oPBR.Sheets(1).Range("A" ptRowNumber).Value
sRunDate := oPBR.Sheets(1).Range("A3").Value
;
; HCC status
;
AddrHCC := oPBR.Sheets(1).Range("J" ptRowNumber).Value
Metric31 := oPBR.Sheets(1).Range("J" ptRowNumber).Interior.Color
fmtAddrHCC := RegExReplace(AddrHCC, ",", ", ")
;
PendHCC := oPBR.Sheets(1).Range("K" ptRowNumber).Value
Metric32 := oPBR.Sheets(1).Range("K" ptRowNumber).Interior.Color
fmtPendHCC := RegExReplace(PendHCC, ",", ", ")
;
; Labs/Vitals
;
LabA1c := oPBR.Sheets(1).Range("R" ptRowNumber).Value
Metric22 := oPBR.Sheets(1).Range("R" ptRowNumber).Interior.Color
;
LabBMI := oPBR.Sheets(1).Range("T" ptRowNumber).Value
Metric23 := oPBR.Sheets(1).Range("T" ptRowNumber).Interior.Color
;
LabBP :=  oPBR.Sheets(1).Range("S" ptRowNumber).Value
Metric24 := oPBR.Sheets(1).Range("S" ptRowNumber).Interior.Color
;
; Med Adherence
MedAdhDM := oPBR.Sheets(1).Range("U" ptRowNumber).Value
Metric3 := oPBR.Sheets(1).Range("U" ptRowNumber).Interior.Color
;
MedAdhHTN := oPBR.Sheets(1).Range("V" ptRowNumber).Value
Metric5 := oPBR.Sheets(1).Range("V" ptRowNumber).Interior.Color
;
MedAdhChol := oPBR.Sheets(1).Range("W" ptRowNumber).Value
Metric4 := oPBR.Sheets(1).Range("W" ptRowNumber).Interior.Color
;
; dislike special cases but 2 columns for this metric
MedDMStatin := oPBR.Sheets(1).Range("X" ptRowNumber).Value
Metric6 := oPBR.Sheets(1).Range("X" ptRowNumber).Interior.Color
MedDMSTatin2 := oPBR.Sheets(1).Range("Y" ptRowNumber).Value
MedDMStatin := MedDMStatine . MedDMStatin2						; combine the 2 values
Metric6Z := oPBR.Sheets(1).Range("Y" ptRowNumber).Interior.Color
if (Metric6 > Metric6Z)											; set to highest alert color
	Metric6 := Metric6Z
;
MedHRM  := oPBR.Sheets(1).Range("Z" ptRowNumber).Value
Metric7 := oPBR.Sheets(1).Range("Z" ptRowNumber).Interior.Color
;
; Processes
ProcMammo  := oPBR.Sheets(1).Range("AA" ptRowNumber).Value
Metric8 := oPBR.Sheets(1).Range("AA" ptRowNumber).Interior.Color
;
ProcColo := oPBR.Sheets(1).Range("AB" ptRowNumber).Value
Metric9 := oPBR.Sheets(1).Range("AB" ptRowNumber).Interior.Color
;
ProcDMEye := oPBR.Sheets(1).Range("AC" ptRowNumber).Value
Metric10 := oPBR.Sheets(1).Range("AC" ptRowNumber).Interior.Color
;
ProcOsteo := oPBR.Sheets(1).Range("AD" ptRowNumber).Value
Metric11 := oPBR.Sheets(1).Range("AD" ptRowNumber).Interior.Color
;
ProcRA := oPBR.Sheets(1).Range("AE" ptRowNumber).Value
Metric12 := oPBR.Sheets(1).Range("AE" ptRowNumber).Interior.Color
;
; FOC - primary 
FOCDM := oPBR.Sheets(1).Range("AF" ptRowNumber).Value
Metric13 := oPBR.Sheets(1).Range("AF" ptRowNumber).Interior.Color
;
FOCPHQ9 := oPBR.Sheets(1).Range("AG" ptRowNumber).Value
Metric14 := oPBR.Sheets(1).Range("AG" ptRowNumber).Interior.Color
;
FOCSmoker  := oPBR.Sheets(1).Range("AH" ptRowNumber).Value
Metric15 := oPBR.Sheets(1).Range("AH" ptRowNumber).Interior.Color
;
FOCSpiro  := oPBR.Sheets(1).Range("AI" ptRowNumber).Value
Metric16 := oPBR.Sheets(1).Range("AI" ptRowNumber).Interior.Color
;
FOCABI := oPBR.Sheets(1).Range("AJ" ptRowNumber).Value
Metric17 := oPBR.Sheets(1).Range("AJ" ptRowNumber).Interior.Color
;
FOCDPN := oPBR.Sheets(1).Range("AK" ptRowNumber).Value
Metric18 := oPBR.Sheets(1).Range("AK" ptRowNumber).Interior.Color
;
FOCCKD := oPBR.Sheets(1).Range("AL" ptRowNumber).Value
Metric27 := oPBR.Sheets(1).Range("AL" ptRowNumber).Interior.Color
;
; FOC - secondary
FOCPTH := oPBR.Sheets(1).Range("AM" ptRowNumber).Value
Metric19 := oPBR.Sheets(1).Range("AM" ptRowNumber).Interior.Color
;
FOCCKD2 := oPBR.Sheets(1).Range("AN" ptRowNumber).Value
Metric20 := oPBR.Sheets(1).Range("AN" ptRowNumber).Value
;
FOCThrombo := oPBR.Sheets(1).Range("AO" ptRowNumber).Value
Metric21 := oPBR.Sheets(1).Range("AO" ptRowNumber).Value
;
;
; test for NA "--" if present ignore -> filtering 
Metric_String := "WellMed Metrics: `r`n`r`nLAB MEASURES:`r`n"
IfNotInString, LabA1c, --
	Metric_String := Metric_String "Lab A1c = " LabA1c "`r`n"
IfNotInString, LabBMI, --
	Metric_String := Metric_String "Vital BMI = " LabBMI "`r`n"
Metric_String := Metric_String "`r`nMEDICATION ADHERENCE:`r`n"
IfNotInString, MedAdhHTN, --
	Metric_String := Metric_String "Adh HTN meds " MedAdhHTN "`r`n"
IfNotInString, MedAdhChol, --
	Metric_String := Metric_String "Adh Chol meds " MedAdhChol "`r`n"
IfNotInString, MedDMStatin, --
	Metric_String := Metric_String "Adh Statins in DM " MedDMStatin "`r`n"
IfNotInString, MedHRM, --
	Metric_String := Metric_String "Pt on HRM(S) " MedHRM "`r`n"
Metric_String := Metric_String "`r`nPROCESSES MEASURES:`r`n"
IfNotInString, ProcMammo, --
	Metric_String := Metric_String "Mammo status " ProcMammo "`r`n"
IfNotInString, ProcColo, --
	Metric_String := Metric_String "Colo screen " ProcColo "`r`n"
IfNotInString, ProcDMEye, --
	Metric_String := Metric_String "DM eye exam " ProcDMEye "`r`n"
IfNotInString, ProcRA, --
	Metric_String := Metric_String "RA pt. on DMARD(S) " ProcRA "`r`n"
Metric_String := Metric_String "`r`nFOC MEASURES:`r`n"
IfNotInString, FOCDM, --
	Metric_String := Metric_String "DM screen " FOCDM "`r`n"
IfNotInString, FOCPHQ9, --
	Metric_String := Metric_String "PHQ-9 screen " FOCPHQ9 "`r`n"
IfNotInString, FOCSmoker, --
	Metric_String := Metric_String "Smoking screen " FOCSmoker "`r`n"
IfNotInString, FOCSpiro, --
	Metric_String := Metric_String "Spiro done for pt. w/Sx's " FOCSpiro "`r`n"
IfNotInString, FOCABI, --
	Metric_String := Metric_String "Vascular screen " FOCABI "`r`n"
IfNotInString, FOCDPN, --
	Metric_String := Metric_String "Neuropathy screen " FOCDPN "`r`n"
Metric_String := Metric_String "`r`nSECONDARY FOC MEASURES:`r`n"
IfNotInString, FOCPTH, --
	Metric_String := Metric_String "F/U PTH in pt. w/CKD " FOCPTH "`r`n"
IfNotInString, FOCCKD2, --
	Metric_String := Metric_String "F/U Ur/Cr ratio " FOCCKD2 "`r`n"
IfNotInString, FOCThrombo, --
	Metric_String := Metric_String "F/U CBC fo Thrombo penia/cytosis " FOCThrombo "`r`n"
;
/*
; *********************************************************
; get Addressed / Pending HCC information
; Have to handle case with HCC string is longer than width. Couldn't get +Wrap feature to work
TextSize:=GetTextSize(AddrHCC, "S10,Verdana", false, 0)
TextSplitLen:=400/TextSize*StrLen(AddrHCC) ;400 is the intended width
AddrHCC:=WrapText_Force(AddrHCC,TextSplitLen,",")
;
TextSize:=GetTextSize(PendHCC, "S10,Verdana", false, 0)
TextSplitLen:=400/TextSize*StrLen(PendHCC) ;400 is the intended width
PendHCC:=WrapText_Force(PendHCC,TextSplitLen,",")
; ***************************************************************
*/
; add HCC's to Metric_String here, ? should move all these routines outside of the GUI?
Metric_String := Metric_String . "`r`nPending HCC: `r`n"  . PendHCC . "`r`nAddressed HCC:`r`n" . AddrHCC . "`r`nPlease print Attestation Form for full description / completion"
; copy string to Clipboard for pasting into the note
ClipBoard := Metric_String
;
; ********************
; build the HCC_Descriptions string for display when called
; and write the string to file for persistance / later access
; get rid of the "."
StringReplace, PendHCC, PendHCC, . , , All
;Loop, Read, C:\Users\gjackson\Documents\AHK-Scrips\CDSSTab\HCC_Descriptors.txt 
; HCC_Descriptors must be in the scrips directory
if !(PendHCC = "--") {
	HCC_String := sPtDemo . "`r`n`r`nFormat: `r`nHCC`tICD`tDescriptor`r`n"
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
}
; 
HCC_String := HCC_String . "`r`n`r`nPlease print Attestation Form for full description / completion`r`n"
; now write the file to the disk 
; file name "HCC"A_YYYYA_MMA_DD" ; and append each string to the file for that day 
; seperate by a pipe "|" and read as a CSV file
; create if not there 
HCC_File_Name := "HCC" . A_YYYY . A_MM . A_DD . ".txt"
FileAppend, %HCC_String% . |, %HCC_File_Name%
;
;
; Build the Dashboard GUI
Gui, Font, S7 CDefault, Verdana
;
Gui, Add, Text, x12 y1 w175 h80 vMetric1, %sPtDemo%
;
Gui, Add, Text, x235 y21 vMetric2, Physician report `n%sRunDate% ;04jul2017 MOR Run Date didn't align with Physician Report, fixed
;
;22jul2017 MOR Add Addressed/Pending HCC information
;
Gui, Add, GroupBox, x22 y60 w400 h105 , HCC Status
;Gui, Add, Text, x32 y72 w380 h50 vMetricAddHcc, Addressed HCC %AddrHCC%
Gui, Add, Text, x32 y72 w380 r6 vMetricAddHcc, Addressed HCC %fmtAddrHCC%
Gui, Add, Text, x32 y112 w380 h50 vMetricPendHCC, Pending HCC %fmtPendHCC%
;
Gui, Add, GroupBox, x22 y170 w400 h82, MEDICATION ADHERENCE
Gui, Add, Text, x32 y182 w65 h65 gHEDIS_MA_DM vMetric3, Adh DM %MedAdhDM%
Gui, Add, Text, x112 y182 w65 h65 gHEDIS_MA_Chol vMetric4 , Adh Chol %MedAdhChol%
Gui, Add, Text, x192 y182 w65 h65 gHEDIS_MA_HTN vMetric5 , Adh HTN %MedAdhHTN%
Gui, Add, Text, x272 y182 w65 h65 gHEDIS_MA_DM_Statins vMetric6 , Statin in DM %MedDMStatin%
Gui, Add, Text, x352 y182 w65 h65 vHHRM gHEDIS_HRM vMetric7, HRM Med %MedHRM%
;
Gui, Add, GroupBox, x22 y252 w400 h82 , PROCESSES
Gui, Add, Text, x32 y264 w65 h65 gProc_Mammo vMetric8, Mammo %ProcMammo%
Gui, Add, Text, x112 y264 w65 h65 gProc_Colo vMetric9, Colo screen %ProcColo%
Gui, Add, Text, x192 y264 w65 h65 gProc_DMEye vMetric10, DM Eye screen %ProcDMEye%
Gui, Add, Text, x272 y264 w65 h65 gProc_Tx_Osteo vMetric11, DEXA or Tx Osteo %ProcOsteo%
Gui, Add, Text, x352 y264 w65 h65 gProc_RA_DMARD vMetric12, RA on DMARD %ProcRA% 
;
Gui, Add, GroupBox, x22 y334 w400 h82 , FOC Metrics
Gui, Add, Text, x32 y346 w65 h65 gFOC_DM_Screen vMetric13 , DM screen %FOCDM%
Gui, Add, Text, x112 y346 w65 h65 gFOC_PHQ-9 vMetric14, PHQ-9 screen %FOCPHQ9%
Gui, Add, Text, x192 y346 w65 h65 gFOC_Smoking_Hx vMetric15, Smoking Hx %FOCSmoker%
Gui, Add, Text, x272 y346 w65 h65 gFOC_Spiro vMetric16, Spiro %FOCSpiro%
Gui, Add, Text, x352 y346 w65 h65 gFOC_Vascular vMetric17, Vascular screen %FOCABI%
;
Gui, Add, GroupBox, x22 y416 w155 h82, FOC Metrics
Gui, Add, Text, x32 y428 w62 h65 gFOC_Neuropathy vMetric18, Neuopathy screen %FOCDPN% 
Gui, Add, Text, x112 y424 w62 h65 gFOC_CKD vMetric27,CKD Screen %FOCCKD%
;
; move over 2 steps
Gui, Add, GroupBox, x182 y416 w242 h82 , SECONDARY FOC
Gui, Add, Text, x192 y428 w65 h65 g2FOC_PTH vMetric19 , F/U PTH in CKD %FOCPTH%
Gui, Add, Text, x272 y428 w65 h65 g2FOC_2CKD vMetric20, F/U Ur/Cr ratio %FOCCKD2%
Gui, Add, Text, x352 y428 w65 h65 g2FOC_Plat vMetric21, F/U Abnl Plat. %FOCThrombo%
;
Gui, Add, GroupBox, x22 y498 w244 h82 , LABS
Gui, Add, Text, x32 y510 w65 h65 gLab_A1c_Control vMetric22, A1c < 9 %LabA1c%
Gui, Add, Text, x112 y510 w65 h65 gLab_BMI vMetric23, BMI %LabBMI%
Gui, Add, Text, x192 y510 w65 h65 gLab_BP vMetric24, HTN %LabBP%


Gui, Add, Button, x272 y510 w65 h30 gMetricFAQ vMetric25, Metric FAQ
Gui, Add, Button, x272 y545 w65 h30 gHCCPending, Pending HCCs
Gui, Add, Button, x352 y510 w65 h65 gGuiClose vMetric26, CLOSE
;
Gui, Show, Center w440 h590, Patient's Quality Metric Dashboard   (%sVersion%) ;04jul2017 MOR
;
; ******************** 
; update text color in the Gui 
; 13826810 = "Yellow", but reset yellow to octal FF8800 more of an orange/yellow for better visibility
; 13353215  = "Fuchsia"
; 255 = "Red"
;
; HCC Pending alert 
IfNotInString, PendHCC, --  
{
	GuiControl, +cRed, MetricPendHCC
	GuiControl, , MetricPendHCC, Pending HCC %fmtPendHCC%
}
;
; Medication Adherence
if (Metric3 = 255) { 
	GuiControl, +cRed, Metric3
	GuiControl, , Metric3 ,Adh DM %MedAdhDM%
}
if (Metric3 = 13353215) { 
	GuiControl, +cFuchsia, Metric3
	GuiControl, , Metric3 ,Adh DM %MedAdhDM%
}
if (Metric3 = 13826810 ) { 
	GuiControl, +cff8800, Metric3
	GuiControl, , Metric3 ,Adh DM %MedAdhDM%
}
; Med Adh Chol
if (Metric4 = 255) { 
	GuiControl, +cRed, Metric4
	GuiControl, , Metric4 ,Adh Chol %MedAdhChol%
}
if (Metric4 = 13353215) { 
	GuiControl, +cFuchsia, Metric4
	GuiControl, , Metric4 ,Adh Chol %MedAdhChol%
}
if (Metric4 = 13826810) { 
	GuiControl, +cff8800, Metric4
	GuiControl, , Metric4 ,Adh Chol %MedAdhChol%
}
; Med Adh HTN
if (Metric5 = 255) { 
	GuiControl, +cRed, Metric5
	GuiControl, , Metric5 , Adh HTN %MedAdhHTN%
}
if (Metric5 = 13353215) { 
	GuiControl, +cFuchsia, Metric5
	GuiControl, , Metric5 , Adh HTN %MedAdhHTN%
}
if (Metric5 = 13826810) { 
	GuiControl, +cff8800, Metric5
	GuiControl, , Metric5 , Adh HTN %MedAdhHTN%
}
; Statin w/DM 
if (Metric6 = 255) { 
	GuiControl, +cRed, Metric6
	GuiControl, , Metric6 , Statin in DM %MedDMStatin%
}
if (Metric6 = 13353215) { 
	GuiControl, +cFuchsia, Metric6
	GuiControl, , Metric6 , Statin in DM %MedDMStatin%
}
if (Metric6 = 13826810) { 
	GuiControl, +cff8800, Metric6
	GuiControl, , Metric6 , Statin in DM %MedDMStatin%
}
; HRM Meds
if (Metric7 = 255) { 
	GuiControl, +cRed, Metric7
	GuiControl, , Metric7 ,HRM Med %MedHRM%
}
if (Metric7 = 13353215) { 
	GuiControl, +cFuchsia, Metric7   
	GuiControl, , Metric7 ,HRM Med %MedHRM%	
}
if (Metric7 = 13826810) { 
	GuiControl, +cff8800, Metric7
	GuiControl, , Metric7 ,HRM Med %MedHRM%
}
; Processes
; Mammo
if (Metric8 = 255) { 
	GuiControl, +cRed, Metric8
	GuiControl, , Metric8 ,Mammo %ProcMammo%
}
if (Metric8 = 13353215) { 
	GuiControl, +cFuchsia, Metric8   
	GuiControl, , Metric8 ,Mammo %ProcMammo%	
}
if (Metric8 = 13826810) { 
	GuiControl, +cff8800, Metric8
	GuiControl, , Metric8 ,Mammo %ProcMammo%
}
; Colo Screen; ProcColo
if (Metric9 = 255) { 
	GuiControl, +cRed, Metric9
	GuiControl, , Metric9 ,Colo Screen %ProcColo%
}
if (Metric9 = 13353215) { 
	GuiControl, +cFuchsia, Metric9   
	GuiControl, , Metric9 ,Colo Screen %ProcColo%	
}
if (Metric9 = 13826810) { 
	GuiControl, +cff8800, Metric9
	GuiControl, , Metric9 ,Colo Screen %ProcColo%
}
; DM Eye exam; ProcDMEye
if (Metric10 = 255) { 
	GuiControl, +cRed, Metric10
	GuiControl, , Metric10 ,DM Eye Screen %ProcDMEye%
}
if (Metric10 = 13353215) { 
	GuiControl, +cFuchsia, Metric10   
	GuiControl, , Metric10 ,DM Eye Screen %ProcDMEye%	
}
if (Metric10 = 13826810) { 
	GuiControl, +cff8800, Metric10
	GuiControl, , Metric10 ,DM Eye Screen %ProcDMEye%
}
; DEXA / Osteoporosis
if (Metric11 = 255) { 
	GuiControl, +cRed, Metric11
	GuiControl, , Metric11 ,DEXA or Tx Osteo %ProcOsteo%
}
if (Metric11 = 13353215) { 
	GuiControl, +cFuchsia, Metric11
	GuiControl, , Metric11 ,DEXA or Tx Osteo %ProcOsteo%
}
if (Metric11 = 13826810) { 
	GuiControl, +cff8800, Metric11
	GuiControl, , Metric11 ,DEXA or Tx Osteo %ProcOsteo%
}
; RA on DMARD; ProcRA
if (Metric12 = 255) { 
	GuiControl, +cRed, Metric12
	GuiControl, , Metric12 ,RA on DMARD %ProcRA%
}
if (Metric12 = 13353215) { 
	GuiControl, +cFuchsia, Metric12
	GuiControl, , Metric12 ,RA on DMARD %ProcRA%
}
if (Metric12 = 13826810) { 
	GuiControl, +cff8800, Metric12
	GuiControl, , Metric12 ,RA on DMARD %ProcRA%
}
; FOC Metrics
; DM screen; FOCDM
if (Metric13 = 255) { 
	GuiControl, +cRed, Metric13
	GuiControl, , Metric13 ,DM Screen %FOCDM% 
}
if (Metric13 = 13353215) { 
	GuiControl, +cFuchsia, Metric13
	GuiControl, , Metric13 ,DM Screen %FOCDM% 
} 
if (Metric13 = 13826810) { 
	GuiControl, +cff8800, Metric13
	GuiControl, , Metric13 ,DM Screen %FOCDM% 
} 
; PHQ-9 screen; FOCPHQ9 
if (Metric14 = 255) { 
	GuiControl, +cRed, Metric14
	GuiControl, , Metric14 ,PHQ-9 Screen %FOCPHQ9% 
} 
if (Metric14 = 13353215) { 
	GuiControl, +cFuchsia, Metric14
	GuiControl, , Metric14 ,PHQ-9 Screen %FOCPHQ9% 
}
if (Metric14 = 13826810) { 
	GuiControl, +cff8800, Metric14
	GuiControl, , Metric14 ,PHQ-9 Screen %FOCPHQ9% 
} 
; Smoking Hx; FOCSmoker 
if (Metric15 = 255) { 
	GuiControl, +cRed, Metric15
	GuiControl, , Metric15 ,Smoking Hx %FOCSmoker% 
} 
if (Metric15 = 13353215) { 
	GuiControl, +cFuchsia, Metric15
	GuiControl, , Metric15 ,Smoking Hx %FOCSmoker% 
}
if (Metric15 = 13826810) { 
	GuiControl, +cff8800, Metric15
	GuiControl, , Metric15 ,Smoking Hx %FOCSmoker% 
} 
;Spiro; FOCSpiro 
if (Metric16 = 255) { 
	GuiControl, +cRed, Metric16
	GuiControl, , Metric16 ,Spiro %FOCSpiro% 
} 
if (Metric16 = 13353215) { 
	GuiControl, +cFuchsia, Metric16
	GuiControl, , Metric16 ,Spiro %FOCSpiro% 
}
if (Metric16 = 13826810) { 
	GuiControl, +cff8800, Metric16
	GuiControl, , Metric16 ,Spiro %FOCSpiro% 
} 
; Vascular Screen; FOCABI 
if (Metric17 = 255) { 
	GuiControl, +cRed, Metric17
	GuiControl, , Metric17 ,Vascular Screen %FOCABI% 
}
if (Metric17 = 13353215) { 
	GuiControl, +cFuchsia, Metric17
	GuiControl, , Metric17 ,Vascular Screen %FOCABI% 
}
if (Metric17 = 13826810) { 
	GuiControl, +cff8800, Metric17
	GuiControl, , Metric17 ,Vascular Screen %FOCABI% 
}
; Neuropathy Screen %FOCDPN% 
if (Metric18 = 255) { 
	GuiControl, +cRed, Metric18
	GuiControl, , Metric18 ,Neuropathy Screen %FOCDPN%
}
if (Metric18 = 13353215) { 
	GuiControl, +cFuchsia, Metric18
	GuiControl, , Metric18 ,Neuropathy Screen %FOCDPN% 
} 
if (Metric18 = 13826810) { 
	GuiControl, +cff8800, Metric18
	GuiControl, , Metric18 ,Neuropathy Screen %FOCDPN% 
} 
; Metric27 CKD %FOCCKD%





; Secondary FOC Metrics
; F/U PTH in CKD %FOCPTH% 
if (Metric19 = 255) { 
	GuiControl, +cRed, Metric19
	GuiControl, , Metric19 ,F/U PTH in CKD %FOCPTH% 
} 
if (Metric19 = 13353215) { 
	GuiControl, +cFuchsia, Metric19
	GuiControl, , Metric19 ,F/U PTH in CKD %FOCPTH%
} 
if (Metric19 = 13826810) { 
	GuiControl, +cff8800, Metric19
	GuiControl, , Metric19 ,F/U PTH in CKD %FOCPTH% 
} 
; F/U Ur/Cr ratio %FOCCKD2% 
if (Metric20 = 255) { 
	GuiControl, +cRed, Metric20
	GuiControl, , Metric20 ,F/U Ur/Cr ratio %FOCCKD2% 
} 
if (Metric20 = 13353215) { 
	GuiControl, +cFuchsia, Metric20
	GuiControl, , Metric20 ,F/U Ur/Cr ratio %FOCCKD2%
} 
if (Metric20 = 13826810) { 
	GuiControl, +cff8800, Metric20
	GuiControl, , Metric20 ,F/U Ur/Cr ratio %FOCCKD2% 
} 
; F/U Abnl Plat. %FOCThrombo% 
if (Metric21 = 255) { 
	GuiControl, +cRed, Metric21
	GuiControl, , Metric21 ,F/U Abnl Plat. %FOCThrombo% 
} 
if (Metric21 = 13353215) { 
	GuiControl, +cFuchsia, Metric21
	GuiControl, , Metric21 ,F/U Abnl Plat. %FOCThrombo% 
} 
if (Metric21 = 13826810) { 
	GuiControl, +cff8800, Metric21
	GuiControl, , Metric21 ,F/U Abnl Plat. %FOCThrombo%
} 
; Lab measures
; Lab A1c <9 %LabA1c% 
if (Metric22 = 255) { 
	GuiControl, +cRed, Metric22
	GuiControl, , Metric22 ,Lab A1c <9 %LabA1c% 
} 
if (Metric22 = 13353215) { 
	GuiControl, +cFuchsia, Metric22
	GuiControl, , Metric22 ,Lab A1c <9 %LabA1c% 
} if (Metric22 = 13826810) { 
	GuiControl, +cff8800, Metric22
	GuiControl, , Metric22 ,Lab A1c <9 %LabA1c% 
} 
; BMI %LabBMI% 
if (Metric23 = 255) { 
	GuiControl, +cRed, Metric23
	GuiControl, , Metric23 ,BMI %LabBMI%
}
; Labs 
if (Metric23 = 13353215) { 
	GuiControl, +cFuchsia, Metric23
	GuiControl, , Metric23 ,BMI %LabBMI%
}
if (Metric23 = 13826810) { 
	GuiControl, +cff8800, Metric23
	GuiControl, , Metric23 ,BMI %LabBMI%
}
; HTN %LabBP%
if (Metric24 = 255) { 
	GuiControl, +cRed, Metric24
	GuiControl, , Metric24 ,HTN %LabBP%
}
if (Metric24 = 13353215) { 
	GuiControl, +cFuchsia, Metric24
	GuiControl, , Metric24 ,HTN %LabBP%
}
if (Metric24 = 13826810) { 
	GuiControl, +cff8800, Metric24
	GuiControl, , Metric24 ,HTN %LabBP%
}
;
;
return
;
; ********************
; HEDIS info for each of the buttons 
; template; remember to escape commas (`,) in the plain text 
;button_Name:
;MsgBox, 64, MsgBox Title:,`r`nWHO:`r`nthe text`r`n`r`nEXCLUSSIONS:``rnthe text `r`n`r`nMEASURE IS MET IF:`r`nThe Text, 5 return
;
;Med Adherence
HEDIS_HRM:
MsgBox, 64, HEDIS HRM, High Risk Medications`nWHO:`n65 years of age and older with two or more prescription fills for the same HRM drug in the current year for a RAAS`n`nEXCLUSSIONS:`nHospice at any time during the current year (New!)`nMust be hospice by CMS identification (claim)`n`nMEASURE IS FAILED IF:`nHRM A: based on two fills of drug with same active ingredient (all others-see list)`nHRM B: based on 90 cumulative day supply (NEW!) in a calendar year (nitrofurantoin and non-benzodiazepine hypnotics)`nHRM C: based on average daily dose (reserpine`, digoxin and doxepin)`nComplete list on ePRG (pending) and the DataRAPS Wiki, 5
return 
;
HEDIS_MA_DM:
MsgBox, 64, HEDIS Diabetic Med Adherence, WHO:`n18 years of age and older with at least two fills of oral diabetic medication(s) across any of the drug classes during the year.`n`nEXCEPTIONS:`nPatients with 1 fill or more of Insulin and patients with ESRD captured anytime during the calendar year.`n`nMEASURE IS SUCCESSFULLY MET WHEN:`nPatient has filled the medication at least 80`% or more in the year.`nCaptured via pharmacy claims`n`nDRUG CLASSES INCLUDED:`nBiguanides`, sulfonylureas`, thiazolidinediones`, DPP-IV inhibitors`, incretin mimetic drugs`, meglitinide drugs`, orSGLT2 inhibitors. , 5
return
;
HEDIS_MA_Chol:
MsgBox, 64, Medication Adherence for Cholesterol (Statins), WHO:`n18 years and older with ONE fill of either the same medication or another statin medication during the year.`n`nEXCLUSIONS:`nNone`n`nMEASURE IS SUCCESSFULLY MET WHEN:`nPatient has filled the medication at least 80`% or more during the year.`nCaptured via pharmacy claim`n`nDrug class included:`nStatins , 5
return
;
HEDIS_MA_HTN:
MsgBox, 64, Medication Adherence for HTN,WHO:`n18 and older with at least two fills of either the same medication or medications in the drug class during the current year.`n`nEXCLUSIONS:`nPatients with ESRD in that calendar year and/or patients with one fill or more of sacubitril/valsartan anytime during the year.`n`nMEASURE IS SUCCESSFULLY MET WHEN:`nPatient has filled the RAS antagonist medication at least 80`% or more in the year.`nCaptured via pharmacy claims`n`nDrugs included:`nAngiotensin Converting Enzyme Inhibitors (ACEIs)`nAngiotensin Receptor Blockers (ARBs)`nDirect Renin Inhibitors (DRIs), 5
return
;
HEDIS_MA_DM_Statins: 
MsgBox, 64, Statin Therapy in Diabetics`,NEW metric for 2017:,`n`nWHO:`nThose patients 40–75 years of age during the current year and year prior with diabetes who do not have clinical Atherosclerotic Cardiovascular Disease (ASCVD).`n`nEXCLUSIONS:`nAnyone discharged from inpatient with:`nCABG or PCI`nPregnancy`, in vitro fertilization`, clomiphene`, gestational diabetes`nCirrhosis`nMyalgia`, myositis`, myopathy`, rhabdomyolysis`nCKD stage V`, ESRD`nsteroid-induced diabetes`nIn hospice at any time during the measurement year.`n`nMEASURE IS SUCCESSFULLY MET IF:`nDiabetic Patient Received Statin Therapy:`nThe percentage of diabetic members who were dispensed a statin of any dosage intensity at least 1 time during the measurement year`n`nReminder: Diabetic Patient Had 80`% Statin Medication Adherence: The percentage of diabetic members who filled a statin prescription at least 1 time during the measurement year and continued to fill the medication at least 80`% of the treatment period Patient filled at least ONE statin prescription of any dose during the year.,5
return 
; Processes measures 
Proc_Mammo:
MsgBox, 64,  Breast Cancer Screening:,`r`nWHO:`r`nWomen ages 50-74`r`n`r`nEXCLUSIONS:`r`nBilateral Mastectomy or unilateral mastectomy x 2 in past or current year.`r`nCaptured via submitted claim`r`nUnilateral mastectomy requires MMG of remaining breast.`r`n`r`nMEASURE IS SUCCESSFULLY MET IF:`r`nPatient has MMG between October 1`, two years prior and Dec 31st of the current year (NEW)., 5
return  
;
Proc_Colo:
MsgBox, 64, Colo Screening:,`r`nWHO:The members 50-75 years of age.`r`n`r`nEXCLUSSIONS:If the member has had colorectal cancer or a total colectomy any time during the member’s medical history`n­The exclusion may be captured by submitting a copy of the surgical report or a notation in the medical history section of the medical record.`nThe medical record must include a note clearly indicating colorectal cancer history or date of total colectomy`nHospice`r`n`r`nMEASURE IS MET IF:`r`nFOBT: annual guiac FOBT or FIT. There MUST be a lab result`, It cannot be from a digital rectal exam in office.(NEW!)`nFlex-Sig: current year or 4 years prior (Q5 yrs)`nColonoscopy: current year or 9 years prior (Q10 yrs)`nVirtual Colonography: current year or 4 years prior (Q5yrs) (NEW!)`nFIT-DNA (Cologaurd): current year or 2 years prior (Q2yrs) (NEW!), 10
return
;
Proc_DMEye:
MsgBox, 64, Diabetes Care - Eye Exam:,WHO:18-75 year old with diabetes (type 1 or 2)`r`n`r`nEXCLUSSIONS:`r`nGestational diabetes or steroid-induced diabetes (drug-induced diabetes ICD 10 code)`r`n`r`nMEASURE IS MET IF:`nA retinal or dilated eye exam by an eye care professional (optometrist or ophthalmologist) in the current year`nA negative retinal or dilated eye exam (negative for retinopathy) by an eye care professional in the year prior, 10
return
;
;04jul2015 MOR Change spelling of MEARUSE to MEASURE
Proc_Tx_Osteo:
MsgBox, 64, Osteoporosis Managment in Women with a fracture:,WHO:`nPatients age 67 - 85 who suffered a fracture during current year or a fracture diagnosis between July 1 in previous year`, through June 30 of the current year (7/1/16-6/30/17)`n`r`n`rEXCLUSIONS:`nFractures of fingers`, toes`, face or skull`nBone density test 24 month prior to fracture`nClaim or encounter for osteoporosis Tx 12 months prior to fx.`nFilled or had active rx for osteoporosis treatment 12 months prior to Fx`r`n`r`nMEASURE IS SUCCESSFULLY MET WHEN:`nA bone mineral density (BMD) test or prescription for a drug to treat or prevent osteoporosis within six months after the fracture`nThe following medications qualify for osteoporosis treatment under this measure:`n--Bisphosphonates: alendronate`, alendronate-cholecaciferol`, ibandronate`, risedronate`, and zoledronic acid`n--Other drugs: calcitonin`, denosumab`, raloxifene`, teriparatide., 30
return
;
Proc_RA_DMARD:
MsgBox, 64, Rheumatoid Arthritis Management:,WHO:`n18 years or older with two diagnoses (claims) of RA in the current year.`r`n`r`nEXCLUSIONS:`nAny current or past diagnosis of HIV and pregnancy (current year claim must be submitted)`r`n`r`nMEASURE IS SUCCESSFULLY MED WHEN:`nAt least one filled ambulatory prescription for DMARD`n   Captured via claims and pharmacy data`n   Includes injectables and outpatient IV therapy.,10
return
;
; Primary FOC Measures
;
FOC_DM_Screen:
MsgBox, 64, FOC Diabetes A1c Screening:,WHO:`nAll patients 45 years or older`r`n`r`nSCREEN:`n Lab A1c`r`n`r`nACCEPTABLE DATES:`nCurrent year`r`n`r`nFREQUENCY:`nAnnually, 10
return 
;
FOC_PHQ-9:
MsgBox, 64,Depression Screening:,WHO:`nAll Patients`r`n`r`nSURVEY:`nPHQ-9`r`n`r`nACCEPABLE DATES:`nCurrent year`r`n`r`nFREQUENCY:`nAnnually, 10
return 
;
FOC_Smoking_Hx:
MsgBox, 64, Smoking History:,WHO:`nAll patients;`n(If histroy of smoking >= 5 years`, consider Spirometry)`r`n`r`nSURVEY:`nCurrent or Prior Smoker`r`n`r`nACCEPTABLE DATES:`nCurrent year`r`n`r`nFREQUENCY:`nCurrent and Prior year, 10
return
;
FOC_Spiro:
MsgBox, 64, Spirometry:,WHO:`nHistory of current or prior smoker or other risk`nAnd has chronic symptoms of; `n--Cough`, Wheezing`, Shortness of breath`n--Symptoms last for weeks or months`r`n`r`nSURVEY:`nSpirometry`, FEV1`/FVC and FEV1; remember if FEV1/FVC is <= 70`% consider COPD`r`n`r`nACCEPTABLE DATES:`nPrevious or Current year`r`n`r`nFREQUENCY:`nCurrent or Prior year, 10
return
;
FOC_Vascular:
MsgBox, 64,Vascular Screen:,WHO:`nAll patients 50 - 75 years old and no current diagnosis of vascular disease`r`n`r`nEXAM:`nFloCheck`, QuantaFlo`, or ABI`nBoth sides`, if amputation remaining side`r`n`r`nACCEPTABLE DATES:`nCurrent or Prior year`r`n`r`nFREQUENCY:`nCurrent or Prior year`nFreqency will change to Q3 years`, if prior 2 consecutive years are negative, 20
return
;
FOC_Neuropathy:
MsgBox, 64,Neuropathy Screen:,WHO:`nAll patients without current neuropathy diagnosis`r`n`r`nEXAM:`nDPN`, Sudo Scan`, NCS`r`n`r`nEXAM:`nDPN left or right side; Amplitude and Velosity or`nNCV or SudoSCan or`nVPT (Vibritory Preception Test) & Monofilament; will no longer meet this metric though my be used to assist in making this diagnosis`r`n`r`nACCEPTABLE DATES:`nCurrent or Prior Year`r`n`r`nFREQUENCY:`nCurrent or Prior year, 20
return
;
FOC_CKD:
MsgBox, 64, MsgBox Primary CKD Screening,`r`nWHO:`r`nAll MA (Medical Advantage) patients`r`n`r`nEXCLUSSIONS:`r`n None `r`n`r`nMEASURE IS MET IF:`r`nSerum Creatinine`, eGFR`, and Urine Cr / Microalbumin ratio reported in current year, 5 
return
;
; Secondary FOC measures 
;
2FOC_PTH:
MsgBox, 64,Secondary FOC PTH:,`r`n`r`nSECONDARY SCREEN:`nIf eGFR < 60`, or `nhospital stay > 14 days within the past 6 months`n`nSCREEN:`niPTH`n`nACCEPTABLE DATES:`nCurrent or Prior year, 15
return
;
2FOC_2CKD:
MsgBox, 64, Secondary FOC Screen for CKD:,`n`nSECONDARY SCREEN:`n   If initial eGFR >= 60 and UMA is normal`, no further testing is required this year`n*  If initial eGFR >= 60 and UMA abnormal`, repeat UMA in 3 months`n**   If initial eGFR <60 and UMA normal repeat eGFR in 3 months`n*** If initial eGFR < 60 and UMA abnormal`, repeat both test in 3 months`n`nSCREEN:`n*   Urine Microalbumin/Creat Ratio`n**  eGFR and Serum Creatinine`n*** eGFR`, Serum Creatinine`, and Urine Microalbumin/Creat Ratio`n`nACCEPTABLE DATES:`nCurrent or Prior year, 20
return
;
2FOC_Plat:
MsgBox, 64,Secondary FOC Thrombocytopenia:,`n`nSECONDARY SCREEN:`nIf Platelet count < 120`n`nSCREEN:`nrepeat CBC`n`nACCEPTABLE DATES:`nCurrent year, 20
return
;
; Labs
Lab_A1c_Control:
MsgBox, 64,  HEDIS Diabets Control,`n`nWHO:`n18-75 years identified as having diabetes by:`nSubmitted claim OR`nPharmacy claim`n`nEXCLUSIONS:`nA diagnosis of gestational diabetes or steroid-induced diabetes (use code for drug-induced diabetes)`, in any setting`, during the current year or the year prior.`n`nMEASURE IS SUCCESSFULLY MET WHEN:`nA1c <= 9.0, 5
return  
;
Lab_BMI:
MsgBox, 64, HEDIS BMI,Adult BMI Assessment`n`nWHO:`n18-74 years`n`nEXCLUSIONS:`nPregnancy`n`nMEASURE IS SUCCESSFULLY MET IF:`nAge 20 and above have documented BMI in current year or year prior. `n(NEW!) Age 19 and lower will need documentation of percentile in current year or year prior.`n`nRAP:`nBMI >= 40 or a BMI >= 35 with comorbidity(ies)., 5
return
;
Lab_BP:
MsgBox, 64, HEDIS HTN Control,WHO:`nPatients 18 to 85 years old with a single diagnosis of hypertension.`n`nEXCLUSIONS;`nESRD`, Pregnancy`, members with non-acute inpatient hospital stays during measurement year.`n`nMEASURE:`n`nMost recent blood pressure (BP) reading during measurement year. Reading must occur after the initial diagnosis of hypertension.`n`nMEASURE IS SUCCESSFULLY MET WHEN:`n18-59 years of age whose BP is < 140/90mmgHg`nMembers 60–85 years of age with a diagnosis of diabetes whose BP was < 140/90 mm Hg (New!)`nMembers 60–85 years of age without a diagnosis of diabetes whose BP was < 150/90 mm Hg, 5
return
;
MetricFAQ:
; Info Doc "2017_One-Page_Quality_Guide"
SetTimer, ChangeButtonNames, 50
MsgBox, 4, Metric FAQ, To get information about any of the Metrics just click on the text box in question.`r`nA brief description will be displayed`, click OK to close or will automtically close in ~5-10 seconds.`nClick the "Guide" button to open the One-Page Quality Guide pdf. , 10
IfMsgBox, No 
	run, %A_ScriptDir% \2017_One-Page_Quality_Guide.pdf  
return 
;
; ********************
; HCC Pending routine 
; ?? hummm - we may need to run this even if not called by the user as the info may
; be needed by the provider during the OV
; so may change this routine to just the GUI stuff and generate the string with the main 
; GUI dashboard creation ??
; would put the code in the data extraction section 
HCCPending:
;
gui, new,,Pending HCC with Descriptors
gui, add, text, w580,  %HCC_String%
Gui, Add, Button, gHCCGuiClose , CLOSE
gui, show, Autosize Center, Pending HCC with Descriptors
;
return
;
HCCGuiClose:
Gui, destroy
return
;
; ********************
ChangeButtonNames: 
IfWinNotExist, Metric FAQ
    return  ; Keep waiting.
SetTimer, ChangeButtonNames, Off 
WinActivate 
ControlSetText, Button1, &Close
ControlSetText, Button2, &Guide
return 
;
; ********************
GetUserChoice() 
{
global 
; show the ListBox of pt's matched
;MsgBox, % "GetUserChoice;`n" ptChoiceList	
	;Gui, Add, ListBox, h80 w280 vptChoice AltSubmit gptChoice, %ptChoiceList%
	Gui, Add, ListBox, h280 w280 vptChoice AltSubmit gptChoice, %ptChoiceList%	;04jul2017 Expand height to allow for more names without a scroll bar
	Gui, Show, Center, Select your patient
	return
;
;
	ptChoice:
	Gui, Submit
;	MsgBox, % "You chose  number was " ptChoice "`r`n and the patient is " ptArray[ptChoice].Value "`r`nwho's address is " ptArray[ptChoice].Address
	; ***** New selection of the patient row number *****
	ptRowNumber := SubStr(ptArray[ptChoice].Address, 4)
;	MsgBox, % "Function call - - Row number is " ptRowNumber
	Gui, destroy 
	return 
;
return
}
;
/*
; *******************
;29jul2017 GWJ;	function to read iniCDSS_Tab.ini info 
findFilePath()
{
	; check the dafault location
	
	;	default save name will be 'Copy 1 plus report.xlsx'; and default location will be the users DeskTop 
	;	FilePath := %A_Desktop% "\Copy of 1 Plus Report-Test.xlsx"
	;	Test file and filepath Mitosis
	;	FilePath := "C:\Users\Walker\Documents\Medicine1\CDSSTab\TestPBRMay2017.xlsx"
	;	Test file and filepath WellMed
	;	FilePath := "C:\Users\gjackson\Documents\CDSSTab\TestPBRMay2017.xlsx"
	;	Test file and filepath Kingston | FlashDrive
	;	 FilePath := A_Desktop "\qCopy of 1 Plus Report.xlsx"
	
	FilePath := "C:\DeskTop\Copy of 1 Plus REport.xlsx"
	;
	IfExist, %FilePath%
	{
		return FilePath
	}
	; default path not exist
	;
	Ifexist, %A_MyDocuments%\iniCDSS_Tab.ini
		{
		; file exists read the path out
		IniRead, FilePath, %A_MyDocuments%\iniCDSS_Tab.ini, PBRPath, FilePath
		;
		; check that file exists
		IfExist, %FilePath%
			{
			return FilePath
			
			}
		}
;
	; file not there so ask user where it is 
	FileSelectFile, newPath, , ,Please select your current PRP (Pink Box Report) , *.xlsx
		;MsgBox, % "You chose " newPath
		FilePath := newPath
		; write it to iniCDSS_Tab.ini file
		IniWrite, %FilePath%, %A_MyDocuments%\iniCDSS_Tab.ini, PBRPath, FilePath
;
	return FilePath
}
*/ 


; ********************
GuiClose:
Gui, Destroy
oPBR.Close(0)
oPBR := ""
ExitApp
return 
;
;