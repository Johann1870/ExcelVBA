@ECHO OFF

	CD %WINDIR%\system32
	
	::ADMIN COLOR
	SET ACOLOR=FC
	
	::USER COLOR
	SET UCOLOR=8A
	
	::TITLE
	SET TITLE1= - %USERDOMAIN%\%USERNAME% - %COMPUTERNAME%  LogonServer: %LOGONSERVER%
	SET TITLE= %TITLE1%  Session:%SESSIONNAME%
	
	FSUTIL.exe > nul 2> nul && (COLOR %ACOLOR% & TITLE ADMIN %TITLE%) || (COLOR %UCOLOR% & TITLE NONADMIN %TITLE%)
	
	::Scripting path
	SET path=%PATH%;J:\sc
	
	::Pandoc
	SET path=%PATH%;J:\ex\Pandoc;%USERPROFILE%\AppData\Local\Pandoc
	
	::MiKTeX
	SET path=%PATH%;J:\ex\MiKTeX\miktex-portable-2.9.5105\miktex\bin
	
	::PSTools
	SET path=%PATH%;J:\PSTools
	
	::Git
	SET path=%PATH%;%LOCALAPPDATA%\GitHub\PortableGit_\cmd
	
@ECHO ON
