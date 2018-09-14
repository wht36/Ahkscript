#NoEnv
#SingleInstance force
; UnRAR.dll/UnRAR64.dll ahk demo, reads password from UTF-8 encoded "password.txt"
; UnRAR dlls at https://rarlab.com/rar_add.htm

UnRAR("English\*.rar", "English", "\del\English")
UnRAR("Korean\*.rar", "Korean", "\del\Korean")
return

UnRar(FileSpec, DestPath="", Delete="")	; FileSpec can include wildcards; if not specified DestPath will be A_WorkingDir, Delete can be -1 for pure delete, 1 for recycle, other to move to folder
{
	global UnPackSize, UnPackFileName, UnRARLog, Progress, TryPassword
	static	hModule, RarCallBack, Passwords, RAROpenArchiveDataEx, RARHeaderDataEx, Version
		, UnRAR := A_ScriptDir "\UnRAR64.dll"
		, ERAR := {11:"Not enough memory", 12:"Bad data (broken header/CRC error)", 13:"Bad archive", 14:"Unknown encryption", 15:"Cannot open file", 16:"Cannot create file", 17:"Cannot close file", 18:"Cannot read file", 19:"Cannot write file", 20:"Buffer too small", 21:"Unknown error", 22:"Missing password", 23:"Reference error", 24:"Invalid password"}
	If (A_PtrSize="" || A_PtrSize="4")
		A_PtrSize := 4, Ptr := "UInt", UnRAR := A_ScriptDir "\UnRAR.dll"
	If !hModule								; initialise DLL & related
		if hModule := DllCall("LoadLibrary", "Str", UnRAR, Ptr)
			VarSetCapacity(RAROpenArchiveDataEx, (A_PtrSize*5) + 132, 0)
			, RarCallBack := RegisterCallBack("RarCallBack","",4)
			, Numput(RarCallBack, RAROpenArchiveDataEx, A_PtrSize*3+24)
			, VarSetCapacity(RARHeaderDataEx, 10224 + A_PtrSize*3, 0)
			, Version := DllCall(UnRAR "\RARGetDllVersion")
		else {
			msgbox Cannot load %UnRAR%
			ExitApp
		}
	If (DestPath="")
		DestPath := A_WorkingDir

/*
struct RAROpenArchiveDataEx
{				;32bit	64bit
  char         *ArcName;        ;0	0	Point to zero terminated Ansi archive name or NULL if Unicode name specified. 	
  wchar_t      *ArcNameW;       ;4      8	Point to zero terminated Unicode archive name or NULL.
  unsigned int  OpenMode;       ;8	16      RAR_OM_LIST = 0 (Read file headers); RAR_OM_EXTRACT = 1 (test/extract); RAR_OM_LIST_INCSPLIT = 2 (read file headers incl split archives)
  unsigned int  OpenResult;     ;12	20	0 Success, ERAR_NO_MEMORY not enough memory, ERAR_BAD_DATA archive header broken, ERAR_UNKNOWN_FORMAT unknown encryption, EBAR_EOPEN open error, ERAR_BAD_PASSWORD invalid password (only for RAR5 archives)
  char         *CmtBuf;         ;16     24	buffer for comments (max 64kb), if nul comment not read
  unsigned int  CmtBufSize;     ;20     32	max size of comment buffer
  unsigned int  CmtSize;        ;24     36	size of comment stored
  unsigned int  CmtState;       ;28     40	0 No comments, 1 Comments read, ERAR_NO_MEMORY Not enough memory to extract comments, ERAR_BAD_DATA Broken comment, ERAR_SMALL_BUF Buffer is too small, comments are not read completely.
  unsigned int  Flags;          ;32     44	1 archive volume, 2 comment present, 4 locked archive, 8 solid, 16 new naming scheme (volname.partN.rar), 32 authenticity info present (obsolete), 64 recovery record present, 128 headers encrypted, 256 first volume (RAR3.0 or later)
  UNRARCALLBACK Callback;       ;36     48	callback address to process UnRAR events
  LPARAM        UserData;       ;40     56	Userdefined data to pass to callback
  unsigned int  Reserved[28];   ;44     64	Reserved for future use, must be zero
				;152    172
};

struct RARHeaderDataEx
{				  ;32 bit	64 bit
  char         ArcName[1024];     ;0   		0
  wchar_t      ArcNameW[1024];    ;1024		1024
  char         FileName[1024];    ;3072		3072
  wchar_t      FileNameW[1024];   ;4096		4096
  unsigned int Flags;             ;6144		6144		; RHDF_SPLITBEFORE=1 Continued from previous volume, RHDF_SPLITAFTER=2 continued on next volume, RHDF_ENCRYPTED=4 encrypted, 8 reserved, 16 RHDF_SOLID, 32 RHDF_DIRECTORY
  unsigned int PackSize;          ;6148         6148
  unsigned int PackSizeHigh;      ;6152         6152
  unsigned int UnpSize;           ;6156         6156
  unsigned int UnpSizeHigh;       ;6160         6160
  unsigned int HostOS;            ;6164         6164
  unsigned int FileCRC;           ;6168         6168
  unsigned int FileTime;          ;6172         6172
  unsigned int UnpVer;            ;6176         6176
  unsigned int Method;            ;6180         6180
  unsigned int FileAttr;          ;6184         6184
  char         *CmtBuf;           ;6188         6192
  unsigned int CmtBufSize;        ;6192         6200
  unsigned int CmtSize;           ;6196         6204
  unsigned int CmtState;          ;6200         6208
  unsigned int DictSize;          ;6204         6212
  unsigned int HashType;          ;6208         6216
  char         Hash[32];          ;6212         6220
  unsigned int RedirType;	  ;6244		6252
  wchar_t      *RedirName;	  ;6248		6256
  unsigned int RedirNameSize;     ;6252		6264
  unsigned int DirTarget;         ;6256         6268
  unsigned int MtimeLow;          ;6260         6272
  unsigned int MtimeHigh;         ;6264         6276
  unsigned int CtimeLow;          ;6268         6280
  unsigned int CtimeHigh;         ;6272         6284
  unsigned int AtimeLow;          ;6276         6288
  unsigned int AtimeHigh;         ;6280         6292
  unsigned int Reserved[988]      ;6284         6296
};                                ;10236	10248
*/

	Gui, UnRAR:+LastFoundExist							; check if can re-use output window
	IfWinNotExist
	{
		Gui, UnRAR:New
		Gui, Add, Edit, x5 vUnRARLog w400 r10
		Gui, Add, Progress, vProgress w400
		Gui, Add, Button, w75 Default gUnRARGuiClose, &OK 
		Gui, Add, Button, w75 x+25 gUnRARGuiEscape, Cancel
		Gui, Show,, UnRAR v%Version%
	}

	OrigDest := RegExReplace(DestPath, "\\$")				; Save original destpath for automatic child folder creation (otherwise will result in chained child paths)
	Loop, Files, %FileSpec%
	{
		DestPath := OrigDest, RarFile := A_LoopFileFullPath
		Numput(&RarFile, RAROpenArchiveDataEx, A_PtrSize)		; can't use A_LoopFileFullPath in NumPut ... why?
		UnRARLog .= "`n" RarFile

		; first pass to see if need to create dir
		Numput(0, RAROpenArchiveDataEx, A_PtrSize*2, "UInt")		; OpenMode, 0=list, 1=test/extract, 2=read headers incl split archives
		Handle := DllCall(UnRAR "\RAROpenArchiveEx", Ptr, &RAROpenArchiveDataEx, Ptr)
		If OpenResult := NumGet(RAROpenArchiveDataEx, A_PtrSize*2+4, "UInt")
		{
			GuiControl,, UnRARLog, %UnRARLog%
			UnRARLog .= " Err#" OpenResult ": " ERAR[OpenResult]
			Continue
		}

		NoDir := 0
		while !HeaderResult := DllCall(UnRAR "\RARReadHeaderEx", Ptr, Handle, Ptr, &RARHeaderDataEx)		; read file headers
		{
			if !(NumGet(RARHeaderDataEx, 6144, "UInt") & 32)						; if file (i.e. not directory)
			{
				If !InStr(UnPackFileName := StrGet(&RARHeaderDataEx+4096 ,"utf-16"), "\")		; count number of files without directory
					NoDir++
				If NoDir>1
				{											; automatically create folder based on archive name (minus parent path & extension)
					DestPath .= "\" RegExReplace(A_LoopFileName, "i)part\d+\.rar|\.[^\.]+$")	; "\" RegExReplace(StrReplace("`n" RarFile, "`n" RegExReplace(FileSpec, "[^\\]+$")), "i)part\d+\.rar|\.[^\.]+$")
					break
				}
			}
			if DllCall(UnRAR "\RARProcessFileW", Ptr, Handle, "Int", 0, Ptr, &DestPath, Ptr, &DestName)	; process & move to next file in archive (RAR_SKIP=0)
      				Break
		}
		UnRARLog .= " ==> " DestPath

		; second pass to extract
		PasswordIdx := 1, TriedPasswords := "`n", Errors := 0				; initialize password attempts for current archive
		DllCall(UnRAR "\RARCloseArchive", Ptr, Handle)					; re-open RAR file
		Numput(1, RAROpenArchiveDataEx, A_PtrSize*2, "UInt")				; OpenMode, 0=list, 1=test/extract, 2=read headers incl split archives
		Handle := DllCall(UnRAR "\RAROpenArchiveEx", Ptr, &RAROpenArchiveDataEx, Ptr)

		while !HeaderResult := DllCall(UnRAR "\RARReadHeaderEx", Ptr, Handle, Ptr, &RARHeaderDataEx)
		{
			UnPackFileName := StrGet(&RARHeaderDataEx+4096 ,"utf-16")
			, UnPackSize := NumGet(RARHeaderDataEx, 6156, "UInt")
			, Progress := 0, DestName := "", Flags := NumGet(RARHeaderDataEx, 6144, "UInt") 
			, UnRARLog .= "`n" UnPackFileName 

			if FileExist(DestPath "\" UnPackFileName)
			{
				FileGetSize, Size, %DestPath%\%UnPackFileName%
				If (Size==UnPackSize)
				{
					DllCall(UnRAR "\RARProcessFileW", Ptr, Handle, "Int", 0, Ptr, &DestPath, Ptr, &DestName)
					UnRARLog .= " skipped -- same size (" UnPackSize " bytes)."
					continue
				}
			}

			if Flags & 4		; If encrypted
			{
				If !Passwords	
				{
					FileRead, src, *P65001 Password.txt	; 1200 unicode or 65001 UTF-8
					Passwords := StrSplit(RegExReplace(src, "[`r`n]+", "`n"), "`n")	; Initialise array, skip blank lines
				}
				If !TryPassword								; Use Last successful password if available
					TryPassword := Passwords[PasswordIdx]				; TryPassword = empty string if beyond password list (callback will then prompt for user password)
			}

			GuiControl,, UnRARLog, %UnRARLog%

			if ProcessResult := DllCall(UnRAR "\RARProcessFileW", Ptr, Handle, "Int", 2, Ptr, &DestPath, Ptr, &DestName)	; RAR_SKIP=0, RAR_TEST=1, RAR_EXTRACT=2
    			{ 
				if (TryPassword) && (ProcessResult=12 || ProcessResult=24)		; if unsuccessful password (crc / password error)
				{
					TriedPasswords .= TryPassword "`n"				; save previously tried passwords to avoid duplicates
					TryPassword := ""						; make password nul to signal that last password failed
					While (Passwords.Length() >= PasswordIdx)			; loop through passwords in password list, starting at first password
					{
						If !Instr(TriedPasswords, "`n" Passwords[PasswordIdx] "`n", 1)	; skip duplicate passwords
							break
						PasswordIdx++
					}
					DllCall(UnRAR "\RARCloseArchive", Ptr, Handle)			; re-open RAR file ... can't find an easy way to restart
					Handle := DllCall(UnRAR "\RAROpenArchiveEx", Ptr, &RAROpenArchiveDataEx, Ptr)
					continue
				}
				UnRARLog .= " Err#" ProcessResult ": " ERAR[ProcessResult]
				Errors++ 
	      			continue
    			}

			UnRARLog .= "`t(" UnPackSize " bytes)"
			if Flags & 4									; save successful password to file & password list if user provided password
			{
				UnRARLog .= "`nPassword #" PasswordIdx ": " TryPassword
				If Passwords.Length() < PasswordIdx
				{
					FileAppend, `n%TryPassword%, Password.txt, UTF-8
					Passwords[PasswordIdx] := TryPassword
				}
			}  
		}
		DllCall(UnRAR "\RARCloseArchive", Ptr, Handle)
		If (HeaderResult>10)
			UnRARLog .= "`nHeader err#" HeaderResult ": " ERAR[HeaderResult]
		Else if !Errors && Delete
		{
			IfEqual, Delete, -1, FileDelete, %RarFile%
			Else IfEqual, Delete, 1, FileRecycle, %RarFile%
			Else {
				FileCreateDir, %Delete%
				FileMove, %RarFile%, %Delete%
			}
		}
		GuiControl,, UnRARLog, %UnRARLog%
	}
}

UnRARGuiEscape:
UnRARGuiClose:
ExitApp

RarCallback(Msg, User, P1, P2){		; Msg UCM_CHANGEVOLUME = 0, UCM_PROCESSDATA = 1, UCM_NEEDPASSWORD = 2, UCM_CHANGEVOLUMEW = 3, UCM_NEEDPASSWORDW = 4
	global UnPackSize, Progress, TryPassword, UnPackFileName
	If (Msg==0 || Msg==3) && (!P2) 	; P1 = next volume name, Param2 = RAR_VOL_ASK = 0, RAR_VOL_NOTIFY = 1
	{
		Vol := StrGet(P1, Msg ? "utf-16" : "cp0")
		InputBox, Path, Next volume %P1% not found, Please enter path of next volume,,,,,,,,%P1%
		IfNotEqual, ErrorLevel, 0, Return, -1
		StrPut(Path, P1, 1024, Msg ? "utf-16" : "cp0")
	} else if (Msg=1) {		; P1 = pointer to unpacked data (read only, do not modify), P2 = size of unpacked data
		Progress += P2
		GuiControl, UnRAR:, Progress, % Progress*400/UnPackSize
	} else if (Msg=2 || Msg=3){	; P1 = pointer to password buffer, P2 = size of buffer
		IfEqual, TryPassword	; Ask for password if no password available
		{
			InputBox, TryPassword, Password required for %UnPackFileName%, Please enter password for %UnPackFileName%,,,,,,,,
			IfNotEqual, ErrorLevel, 0, Return, -1
		}
		StrPut(TryPassword, P1, P2, Msg=3 ? "utf-16" : "cp0")
	}
	return 1
}
