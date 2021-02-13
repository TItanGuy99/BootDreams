; Script generated with the Venis Install Wizard

; Define your application name
!define APPNAME "BootDreams"
!define APPNAMEANDVERSION "BootDreams 1.0.6c"

; Main Install settings
Name "${APPNAMEANDVERSION}"
InstallDir "C:\BootDreams"
InstallDirRegKey HKLM "Software\${APPNAME}" ""
OutFile "C:\Documents and Settings\Cyle\Desktop\BootDreams_106c.exe"

DirText "Choose the folder in which to install ${APPNAMEANDVERSION}."

Section "BootDreams"

	; Set Section properties
	SetOverwrite on

	; Set Section Files and Shortcuts
	SetOutPath "$INSTDIR\cdda\"
	File "C:\BootDreams\cdda\audio.raw"
	SetOutPath "$INSTDIR\iplogos\"
	File "C:\BootDreams\iplogos\BootDreams.mr"
	File "C:\BootDreams\iplogos\Consolevision.mr"
	File "C:\BootDreams\iplogos\DCEmulation.mr"
	SetOutPath "$INSTDIR\tools\"
	File "C:\BootDreams\tools\audio.raw"
	File "C:\BootDreams\tools\cdi2nero.exe"
	File "C:\BootDreams\tools\cdi4dc.exe"
	File "C:\BootDreams\tools\cdirip.exe"
	File "C:\BootDreams\tools\cdrecord.exe"
	File "C:\BootDreams\tools\cygwin1.dll"
	File "C:\BootDreams\tools\IP.TMPL"
	File "C:\BootDreams\tools\lbacalc.exe"
	File "C:\BootDreams\tools\libmp3lame-0.dll"
	File "C:\BootDreams\tools\libogg-0.dll"
	File "C:\BootDreams\tools\libvorbis-0.dll"
	File "C:\BootDreams\tools\libvorbisenc-2.dll"
	File "C:\BootDreams\tools\libvorbisfile-3.dll"
	File "C:\BootDreams\tools\mds4dc.exe"
	File "C:\BootDreams\tools\mkisofs.exe"
	File "C:\BootDreams\tools\newfile.exe"
	File "C:\BootDreams\tools\scramble.exe"
	File "C:\BootDreams\tools\sh-elf-objcopy.exe"
	File "C:\BootDreams\tools\sox.exe"
	File "C:\BootDreams\tools\wnaspi32.dll"
	SetOutPath "$INSTDIR\"
	File "C:\BootDreams\BootDreams.exe"
	File "C:\BootDreams\BootDreams.chm"
	CreateShortCut "$DESKTOP\BootDreams.lnk" "$INSTDIR\BootDreams.exe"
	CreateDirectory "$SMPROGRAMS\BootDreams"
	CreateShortCut "$SENDTO\BootDreams.lnk" "$INSTDIR\BootDreams.exe"
	CreateShortCut "$SMPROGRAMS\BootDreams\BootDreams.lnk" "$INSTDIR\BootDreams.exe"
	CreateShortCut "$SMPROGRAMS\BootDreams\Uninstall.lnk" "$INSTDIR\Uninstall.exe"

SectionEnd

Section -FinishSection

	WriteRegStr HKLM "Software\${APPNAME}" "" "$INSTDIR"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME}"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" "$INSTDIR\Uninstall.exe"
	WriteUninstaller "$INSTDIR\Uninstall.exe"
	
SectionEnd

;Uninstall section
Section Uninstall

	;Remove from registry...
	DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"
	DeleteRegKey HKLM "SOFTWARE\${APPNAME}"

	; Delete self
	Delete "$INSTDIR\Uninstall.exe"

	; Delete Shortcuts
	Delete "$DESKTOP\BootDreams.lnk"
	Delete "$SENDTO\BootDreams.lnk"
	Delete "$SMPROGRAMS\BootDreams\BootDreams.lnk"
	Delete "$SMPROGRAMS\BootDreams\Uninstall.lnk"

	; Clean up BootDreams	
	Delete "$INSTDIR\cdda\audio.raw"
	Delete "$INSTDIR\iplogos\BootDreams.mr"
	Delete "$INSTDIR\iplogos\Consolevision.mr"
	Delete "$INSTDIR\iplogos\DCEmulation.mr"
	Delete "$INSTDIR\tools\audio.raw"
	Delete "$INSTDIR\tools\cdi2nero.exe"
	Delete "$INSTDIR\tools\cdi4dc.exe"
	Delete "$INSTDIR\tools\cdirip.exe"
	Delete "$INSTDIR\tools\cdrecord.exe"
	Delete "$INSTDIR\tools\cygwin1.dll"
	Delete "$INSTDIR\tools\IP.TMPL"
	Delete "$INSTDIR\tools\lbacalc.exe"
	Delete "$INSTDIR\tools\libmp3lame-0.dll"
	Delete "$INSTDIR\tools\libogg-0.dll"
	Delete "$INSTDIR\tools\libvorbis-0.dll"
	Delete "$INSTDIR\tools\libvorbisenc-2.dll"
	Delete "$INSTDIR\tools\libvorbisfile-3.dll"
	Delete "$INSTDIR\tools\mds4dc.exe"
	Delete "$INSTDIR\tools\mkisofs.exe"
	Delete "$INSTDIR\tools\newfile.exe"
	Delete "$INSTDIR\tools\scramble.exe"
	Delete "$INSTDIR\tools\settings.ini"
	Delete "$INSTDIR\tools\sh-elf-objcopy.exe"
	Delete "$INSTDIR\tools\sox.exe"
	Delete "$INSTDIR\tools\wnaspi32.dll"
	Delete "$INSTDIR\BootDreams.exe"
	Delete "$INSTDIR\BootDreams.chm"

	; Remove remaining directories
	RMDir "$INSTDIR\cdda\"
	RMDir "$INSTDIR\iplogos\"
	RMDir "$INSTDIR\tools\"
	RMDir "$INSTDIR\"
	RMDir "$SMPROGRAMS\BootDreams"
	
SectionEnd

Function un.onInit

	MessageBox MB_YESNO|MB_DEFBUTTON2|MB_ICONQUESTION "Remove ${APPNAMEANDVERSION} and all of its components?" IDYES DoUninstall
		Abort
	DoUninstall:

FunctionEnd

BrandingText "${__TIMESTAMP__}"

; eof