/* *******************

 Returns either a valid FilePath to the open PRP or "ERROR" 
 encript .xlsx file and saves to users computer in the iniCDSS_Tab.ini file 
	see .ini file for it's structure
 decrypt .cprt file for use by the CDSS_Tab script
 Walker Jackson, MD 9/14/2017
 
 see "Include" code for citations and sources MUCH appreciated
 https://autohotkey.com//boards/viewtopic.php?f=5&t=20386&hilit=crypt
 https://autohotkey.com//boards/viewtopic.php?f=6&t=23413&hilit=crypt
*/

#Include Crypt.ahk
#Include CryptConst.ahk
#Include CryptFoos.ahk


findFilePath()
{
	/*
		Check default path - 
		FilePath := %A_Desktop% "\Copy of 1 Plus Report-Test.xlsx"
		then check the CDSS_Tab.ini file for file path 
		then ask user to select PRP
		encrypt file for storage 
		decrypt file for use
	*/
	
	
	; 1st test for encrypted file default location
	; FilePath := A_Desktop . "Copy of 1 Plus Report.crpt"
	FilePath := "C:\Users\gjackson\Desktop\1 Plus Report.crpt"

	
	IfExist, %FilePath%
	{
		; decrypt
		MsgBox, % FilePath . " default"
		FilePath := decryptFile(FilePath)
		; return contains either valid open PRP filepath or "ERROR"
		return FilePath
	}
	
	; not in default - > check the .ini file
	Ifexist, %A_MyDocuments%\iniCDSS_Tab.ini
		{
		; file exists read the path out (should be an encrypted file)
		IniRead, FilePath, %A_MyDocuments%\iniCDSS_Tab.ini, PBRPath, FilePath
		;
		; check that file exists
		MsgBox, % FilePath . " to the ini"
		IfExist, %FilePath%
			{
			FilePath := decryptFile(FilePath)
			; return contains either valid open PRP filepath or "ERROR"
			return FilePath
			}
		}

	; file not there so ask user where it is 
	FileSelectFile, newPath, , ,Please select your current PRP (Pink Box Report) , *.xlsx
	MsgBox, Selected file,  %newPath% . " file select"
	; encrypt and save the encrypted PRP
	FilePath := encryptFile(newPath)
	if (FilePath = "ERROR") {
		newPath := "ERROR"
	}
	; returns the newPath or ERROR
	Return newPath 
}

 
encryptFile(FilePathOpen)
{
	; attempts to enrypt the file and Returns "ERROR"  if failed or the file path passed to it FilePathOpen
	; saves the FilePathCrpt to the .ini file for future reference 
	InputBox, userPassPhrase, Please enter the PassPhrase to encript your file (!REMEMBER IT!), HIDE
	
	; FilePathOpen is the path to the open file to be encrypted	
	StringReplace, FilePathCrpt, FilePathOpen, .xlsx, .crpt
	
	MsgBox,Encrypt File,  %FilePathCrpt% . " is the crpt PRP path"
	
	bytesWriten := Crypt.Encrypt.FileEncrypt(FilePathOpen, FilePathCrpt, userPassPhrase)
				   
	If (byetesWriten < 1) {
		MsgBox, 16, Encryption Error, Sorry Encrypting the File failed\nPlease try again\nIf error continues contact support
		return "ERROR"
	}

	iniWrite, %FilePathCrpt%, %A_MyDocuments%\iniCDSS_Tab.ini, PBRPath, FilePath
	
	return FilePathOpen
}


decryptFile(FilePathCrpt)
{
	; decrypt the file / get the PassPhrase from the user 
	
	InputBox, userPassPhrase, Please enter the PassPhrase to De-crypt your file, HIDE, 20, 200
	; FilePathCrpt is encrypted file path to be decrypted
	StringReplace, FilePathOpen, FilePathCrpt, .crpt, .xlsx
	MsgBox, Decrypt File,  %FilePathOpen% . " open PRP path"
	
	bytesWriten := Crypt.Encrypt.FileDecrypt(FilePathCrpt, FilePathOpen, userPassPhrase) 
	
	
	If (byetesWriten < 1) {
		MsgBox, 16, Decryption Error, Sorry Decrypting the File failed\nPlease try again\nIf error continues contact support
		return "ERROR"
	}
	
	return FilePathOpen
}



