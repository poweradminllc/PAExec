// PAExec.cpp: Main program file
//
// Copyright (c) Power Admin LLC, 2012 - 2013
//
// This code is provided "as is", with absolutely no warranty expressed
// or implied. Any use is at your own risk.
//
//////////////////////////////////////////////////////////////////////


#include "stdafx.h"
#include "PAExec.h"
#include "CmdLineParser.h"
#include <conio.h>
#include <vector>
#include <lm.h>
#include <UserEnv.h>
#include <WtsApi32.h>

////////////////////////////////////////////////////////////////////////
//
//v1.29 - April 14, 2021
//    will connect with given credentials before automatically trying file copy and service installation
//    support redirected input from file
//    don't show window if -i flag isn't given
//    better cleanup of remote service
//



bool gbODS = false;
CString gLogPath;

// The one and only application object
//CWinApp theApp;
void RegressionTests();
BOOL WINAPI ConsoleCtrlHandler(DWORD dwCtrlType);
bool CheckTimeout(__time64_t start, LPCWSTR serverName, Settings& settings);

#ifdef _DEBUG
bool bWaited = false;
void WaitSomewhere()
{
	if(bWaited)
		return;
	bWaited = true;
	wprintf(L"\r\nDEBUG: Waiting for key press\r\n");
	_getch();
}
#endif

int wmain(int argc, wchar_t* argv[], wchar_t* envp[])
{
	int exitCode = 0;
	bool bPrintExitCode = true;

	//// initialize MFC and print and error on failure
	//if (!AfxWinInit(::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0))
	//{
	//	Log(L"Fatal Error: MFC initialization failed", true);
	//	exitCode = -1;
	//}
	//else
	//{
		VERIFY(SetConsoleCtrlHandler(ConsoleCtrlHandler, TRUE));

		CCmdLineParser cmdParser;
		cmdParser.Parse(::GetCommandLine());

		if(cmdParser.HasKey(L"service"))
		{
			//Service Control Manager launched us -- we're a service!
			return StartLocalService(cmdParser);
		}

		Settings settings;

		//read early so logging is started as soon as possible
		if(cmdParser.HasKey(L"dbg"))
		{
			settings.bODS = true;
			gbODS = true;
		}

		if(cmdParser.HasKey(L"lo"))
		{
			gLogPath = settings.localLogPath = cmdParser.GetVal(L"lo");
			if(settings.localLogPath.IsEmpty())
				Log(L"-lo missing value", true);
		}

		if (cmdParser.HasKey(L"share"))
		{
			settings.targetShare = cmdParser.GetVal(L"share");
			if (settings.targetShare.IsEmpty())
				Log(L"-share missing value", true);
		}

#ifdef _DEBUG
		gbODS = true;
#endif

		PrintCopyright();

#ifdef _DEBUG
		RegressionTests();
#endif

		if(ParseCommandLine(settings, ::GetCommandLine()))
		{
			if(gbStop)
			{
				exitCode = -11;
				goto Exit;
			}

			if(settings.computerList.empty())
			{
				if(settings.bUseSystemAccount || ((DWORD)-1 != settings.sessionToInteractWith))
				{
					//have to run service locally
					settings.computerList.push_back(L".");
				}
			}

			if(settings.computerList.empty())
			{
				//run locally
				bool b = StartProcess(settings, NULL);
				if(b)
				{
					if(false == settings.bDontWaitForTerminate)
					{
						DWORD waitMS = settings.timeoutSeconds * 1000;
						if(waitMS == 0)
							waitMS = INFINITE;
						DWORD ret = WaitForSingleObject(settings.hProcess, waitMS);
						switch(ret)
						{
						case WAIT_TIMEOUT:
							Log(L"PAExec timed out waiting for app to exit -- terminating app", true);
							TerminateProcess(settings.hProcess, (DWORD)-10);
							break;
						case WAIT_OBJECT_0: break;
						default: Log(L"PAExec error waiting for app to exit", GetLastError()); 
							break;
						}
						GetExitCodeProcess(settings.hProcess, (DWORD*)&exitCode);
					}
					else
						Log(StrFormat(L"%s started with process ID %u", settings.app, settings.processID), false);
				}
				if(false == b)
					exitCode = -3;
				CloseHandle(settings.hProcess);
			}
			else
			{
				for(std::vector<CString>::iterator cItr = settings.computerList.begin(); (settings.computerList.end() != cItr) && (false == gbStop); cItr++)
				{
					HANDLE hPipe = INVALID_HANDLE_VALUE;

					settings.bNeedToDetachFromAdmin = false;
					settings.bNeedToDetachFromIPC = false;
					settings.bNeedToDeleteServiceFile = false;
					settings.bNeedToDeleteService = false;
					settings.hProcess = NULL;
					settings.processID = 0;
					settings.hStdErr = NULL;
					settings.hStdIn = NULL;
					settings.hStdOut = NULL;

					bool bNeedToSendFile = false;

					CString shownTargetName = *cItr;
					if(shownTargetName == L".")
						shownTargetName = L"{local server}";

					Log(StrFormat(L"\r\nConnecting to %s...", shownTargetName), false);

					//ADMIN$ and IPC$ connections will be made as needed in CopyPAExecToRemote and InstallAndStartRemoteService

					__time64_t start = _time64(NULL);
					//copy myself
					if(false == CopyPAExecToRemote(settings, *cItr)) //logs on error
					{
						exitCode = -4;
						goto PerServerCleanup;
					}

					if(CheckTimeout(start, *cItr, settings))
					{
						exitCode = -5;
						goto PerServerCleanup;
					}

					Log(StrFormat(L"Starting PAExec service on %s...", shownTargetName), false);
					//install myself as a remote service and start remote service
					if(false == InstallAndStartRemoteService(*cItr, settings)) //logs on error
					{
						exitCode = -6;
						goto PerServerCleanup;
					}

					settings.bNeedToDeleteService = true; //always try to clean up if we've gotten this far

					//Send settings, and find out if file(s) need to be copied
					if(false == SendSettings(*cItr, settings, hPipe, bNeedToSendFile))
					{
						exitCode = -7;
						goto PerServerCleanup;
					}

					//copy target if needed
					if(bNeedToSendFile)
					{
						if(1 == settings.srcFileInfos.size())
							Log(StrFormat(L"Copying %s remotely...", settings.srcFileInfos[0].filenameOnly), false);
						else
							Log(StrFormat(L"Copying %u files remotely...", settings.srcFileInfos.size()), false);

						if(false == SendFilesToRemote(*cItr, settings, hPipe))
						{
							exitCode = -8;
							goto PerServerCleanup;
						}
					}

					Log(L"", false); //blank line

					//establish communication with remote service
					//when this returns, the remote app has shutdown (or we weren't supposed to wait for it)
					StartRemoteApp(*cItr, settings, hPipe, exitCode);

PerServerCleanup:
					if(settings.bNeedToDeleteService)
						StopAndDeleteRemoteService(*cItr, settings); //always cleanup
					if(settings.bNeedToDeleteServiceFile)
						DeletePAExecFromRemote(*cItr, settings);
					if(settings.bNeedToDetachFromAdmin)
						EstablishConnection(settings, *cItr, settings.targetShare, false);
					if(settings.bNeedToDetachFromIPC)
						EstablishConnection(settings, *cItr, L"IPC$", false);
					if(false == gbStop) //if stopping, just bail -- the OS will close these (and some of them might be invalid now anyway)
					{
						if(!BAD_HANDLE(hPipe))
							CloseHandle(hPipe);
						if(!BAD_HANDLE(settings.hProcess))
						{
							CloseHandle(settings.hProcess);
							settings.hProcess = NULL;
						}
						if(!BAD_HANDLE(settings.hStdErr))
						{
							CloseHandle(settings.hStdErr);
							settings.hStdErr = NULL;
						}
						if(!BAD_HANDLE(settings.hStdIn))
						{
							CloseHandle(settings.hStdIn);
							settings.hStdIn = NULL;
						}
						if(!BAD_HANDLE(settings.hStdOut))
						{
							CloseHandle(settings.hStdOut);
							settings.hStdOut = NULL;
						}
						if(!BAD_HANDLE(settings.hUserImpersonated))
						{
							CloseHandle(settings.hUserImpersonated);
							settings.hUserImpersonated = NULL;
						}
					}
				}
			}
		}
		else
		{
			if(gbStop)
			{
				exitCode = -11;
				goto Exit;
			}

			PrintUsage();
			bPrintExitCode = false;
			exitCode = -2;
		}

		//clean up
		if(!BAD_HANDLE(settings.hUserProfile))
		{
			UnloadUserProfile(settings.hUser, settings.hUserProfile);
			settings.hUserProfile = NULL;
		}
		if(!BAD_HANDLE(settings.hUser))
		{
			CloseHandle(settings.hUser);
			settings.hUser = NULL;
		}
	//}

Exit:
	gbStop = true;

	if(bPrintExitCode)
		Log(StrFormat(L"\r\nPAExec returning exit code %d\r\n", exitCode), false);

#ifdef _DEBUG
	WaitSomewhere();
#endif

	return exitCode;
}

void PrintCopyright()
{
	CString ver;
	TCHAR	filename[MAX_PATH * 4];

	if(0 != GetModuleFileName(NULL, filename, MAX_PATH * 4))
	{
		DWORD	handle;
		DWORD verSize = GetFileVersionInfoSize(filename, &handle); 
		BYTE*	verInfo = new BYTE[verSize + 1];
		if(GetFileVersionInfo(filename, NULL, verSize, verInfo))
		{
			VS_FIXEDFILEINFO	*fileInfo;
			UINT				fileInfoSize;
			if(VerQueryValue(verInfo, TEXT("\\"), (void**)&fileInfo, &fileInfoSize))
				ver = StrFormat(L"v%u.%u", ((WORD*)&(fileInfo->dwFileVersionMS))[1], ((WORD*)&(fileInfo->dwFileVersionMS))[0]);
		}

		delete [] verInfo;		
		verInfo = NULL;
	}

	Log(StrFormat(L"\r\nPAExec %s - Execute Programs Remotely\r\nCopyright (c) 2012-2021 Power Admin LLC\r\nwww.poweradmin.com/PAExec\r\n", ver), false);
}


void PrintUsage()
{
	HRSRC hR = FindResource(NULL, MAKEINTRESOURCE(IDR_TEXT1), L"TEXT"); //OK

	DWORD size = SizeofResource(NULL, hR);
	HGLOBAL hG = LoadResource(NULL, hR);
	_ASSERT(NULL != hG);
	char* pCopy = new char[size + 1];
	char* pB = (char*)LockResource(hG); 
	memcpy(pCopy, pB, size);
	pCopy[size] = '\0';
	printf("\r\n");
	printf("%s",pCopy);
	printf("\r\n");
	delete [] pCopy;
}


bool GetTargetFileInfo(FileInfo& fi) //returns whether all files were found or not
{
	_ASSERT(FALSE == fi.fullFilePath.IsEmpty());
	bool bAllFilesFound = true;

	HANDLE hSrcFile = INVALID_HANDLE_VALUE;
	hSrcFile = CreateFile(fi.fullFilePath, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, 0, 0);
	DWORD gle = GetLastError();

	if(BAD_HANDLE(hSrcFile))
	{
#ifdef _DEBUG
		Log(StrFormat(L"DEBUG: Failed to open file %s - can't get version info.", fi.fullFilePath), gle);
#endif
		if(gbInService)
		{
#ifdef _DEBUG
			Log(StrFormat(L"DEBUG: Marking to copy %s", fi.fullFilePath), false);
#endif
			fi.bCopyFile = true; //file doesn't exist, so we need it to be copied
		}
		return false;
	}

	BY_HANDLE_FILE_INFORMATION fiInfo = {0};
	if(FALSE == GetFileInformationByHandle(hSrcFile, &fiInfo))
	{
		gle = GetLastError();
		Log(StrFormat(L"Error getting file info from %s.", fi.fullFilePath), gle);
		CloseHandle(hSrcFile);
		return false;
	}

	fi.fileLastWrite = fiInfo.ftLastWriteTime;
	CloseHandle(hSrcFile);
	hSrcFile = INVALID_HANDLE_VALUE;

	DWORD	handle;
	DWORD verSize = GetFileVersionInfoSize(fi.fullFilePath, &handle); 
	BYTE* verInfo = new BYTE[verSize + 1];
	if(GetFileVersionInfo(fi.fullFilePath, NULL, verSize, verInfo))
	{
		VS_FIXEDFILEINFO *fileInfo = NULL;
		UINT fileInfoSize = 0;
		if(VerQueryValue(verInfo, TEXT("\\"), (void**)&fileInfo, &fileInfoSize))
		{
			fi.fileVersionLS = fileInfo->dwFileVersionLS;
			fi.fileVersionMS = fileInfo->dwFileDateMS;
		}
	}
	delete [] verInfo;
	return true; //got file info
}

void RegressionTests()
{
	Log(L" --- start regression tests --- ", false);

	{
		Settings settings;
		_ASSERT(false == ParseCommandLine(settings, L"/?"));
	}

	{
		Settings settings;
		//dumb example (-c and cmd.exe) but still need to figure out how to handle it
		_ASSERT(ParseCommandLine(settings, L"C:\\PAExec.exe -u doug -p test -c -w \"C:\\Windows\\system32\" cmd.exe"));

		RemMsg msg1;
		settings.Serialize(msg1, true);
		msg1.m_bResetReadItr = true;
		settings.Serialize(msg1, false);

		_ASSERT(settings.bCopyFiles); _ASSERT(settings.workingDir == L"C:\\Windows\\system32"); _ASSERT(0 == settings.app.CompareNoCase(L"cmd.exe"));
		_ASSERT(settings.user == L"doug"); _ASSERT(settings.password == L"test"); 
	}

	{
		Settings settings;
		_ASSERT(ParseCommandLine(settings, L"C:\\PAExec.exe \\\\192.168.7.2 -i calc")); //make sure -i doesn't eat calc

		_ASSERT(0 == settings.app.CompareNoCase(L"calc"));
	}

	{
		Settings settings1, settings2;
		_ASSERT(ParseCommandLine(settings1, L"C:\\PAExec.exe \\\\192.168.7.2 -i calc")); //make sure -i doesn't eat calc
		_ASSERT(ParseCommandLine(settings2, L"C:\\PAExec.exe \\\\192.168.7.2 calc")); //make sure -i doesn't eat calc

		_ASSERT(settings1.app == settings2.app);
	}

	{
		Settings settings;
		_ASSERT(ParseCommandLine(settings, L"\"C:\\dir with space\\PAExec.exe\" -u doug -p test -c -w \"C:\\Windows space\\system32\" cmd.exe"));

		RemMsg msg1;
		settings.Serialize(msg1, true);
		msg1.m_bResetReadItr = true;
		settings.Serialize(msg1, false);

		_ASSERT(settings.bCopyFiles); _ASSERT(settings.workingDir == L"C:\\Windows space\\system32"); _ASSERT(0 == settings.app.CompareNoCase(L"cmd.exe"));
		_ASSERT(settings.user == L"doug"); _ASSERT(settings.password == L"test");
	}

	{
		Settings settings;
		_ASSERT(false == ParseCommandLine(settings, L"\"C:\\dir with space\\PAExec.exe\" /?"));
	}

	{
		Settings settings;
		_ASSERT(ParseCommandLine(settings, L"\\\\remote_server taskmgr.exe /?"));
		_ASSERT(settings.appArgs == L"/?");
		_ASSERT(settings.app == L"taskmgr.exe");
	}

	{
		Settings settings;
		_ASSERT(ParseCommandLine(settings, L"\\\\remote_server taskmgr.exe /c"));
		_ASSERT(settings.appArgs == L"/c");
		_ASSERT(settings.app == L"taskmgr.exe");
	}

	{
		Settings settings;
		_ASSERT(ParseCommandLine(settings, L"\"C:\\dir with space\\PAExec.exe\" -a 1,2 -d -s cmd.exe"));

		RemMsg msg1;
		settings.Serialize(msg1, true);
		msg1.m_bResetReadItr = true;
		settings.Serialize(msg1, false);

		_ASSERT(settings.bDontWaitForTerminate); _ASSERT(settings.allowedProcessors.size() == 2); _ASSERT(0 == settings.app.CompareNoCase(L"cmd.exe"));
	}

	{
		Settings settings;
		_ASSERT(ParseCommandLine(settings, L"\"C:\\dir with space\\PAExec.exe\" -a 1,2 -d -u Doug -p Test cmd.exe ping me"));

		RemMsg msg1;
		settings.Serialize(msg1, true);
		msg1.m_bResetReadItr = true;
		settings.Serialize(msg1, false);

		_ASSERT(settings.bDontWaitForTerminate); 
		_ASSERT(settings.allowedProcessors.size() == 2); 
		_ASSERT(0 == settings.app.CompareNoCase(L"cmd.exe"));
		_ASSERT(settings.appArgs == L"ping me");
	}

	{
		Settings settings;
		//bad path, but -c not specified so we won't catch it at parse time
		_ASSERT(ParseCommandLine(settings, L"\"C:\\dir with space\\PAExec.exe\" -s \"C:\\path with\\space \\cmd.exe\" ping me"));

		RemMsg msg1;
		settings.Serialize(msg1, true);
		msg1.m_bResetReadItr = true;
		settings.Serialize(msg1, false);

		_ASSERT(settings.app == L"C:\\path with\\space \\cmd.exe"); _ASSERT(settings.appArgs == L"ping me");
	}

	{
		Settings settings;
		//bad path
		_ASSERT(false == ParseCommandLine(settings, L"\"C:\\dir with space\\PAExec.exe\" -c -v -s \"C:\\path with\\space \\cmd.exe\""));

		RemMsg msg1;
		settings.Serialize(msg1, true);
		msg1.m_bResetReadItr = true;
		settings.Serialize(msg1, false);

		_ASSERT(settings.bCopyFiles); 
		_ASSERT(settings.bCopyIfNewerOrHigherVer); 
		_ASSERT(settings.app == L"C:\\path with\\space \\cmd.exe");
		_ASSERT(settings.appArgs.IsEmpty());
	}

	{
		Settings settings;
		_ASSERT(ParseCommandLine(settings, L"\"C:\\dir with space\\PAExec.exe\" -u Doug -P TEST -c -v cmd.exe"));

		_ASSERT(settings.bCopyFiles); 
		_ASSERT(settings.bCopyIfNewerOrHigherVer); 
		_ASSERT(0 == settings.app.CompareNoCase(L"cmd.exe"));
		_ASSERT(settings.appArgs.IsEmpty()); 

		RemMsg msg1;
		settings.Serialize(msg1, true);
		msg1.m_bResetReadItr = true;
		settings.Serialize(msg1, false);

		_ASSERT(settings.bCopyFiles); 
		_ASSERT(settings.bCopyIfNewerOrHigherVer); 
		_ASSERT(0 == settings.app.CompareNoCase(L"cmd.exe"));
		_ASSERT(settings.appArgs.IsEmpty()); 
	}

	{
		Settings settings;
		_ASSERT(ParseCommandLine(settings, L"\"C:\\dir with space\\PAExec.exe\" -csrc C:\\Windows\\notepad.exe -c cmd.exe"));

		_ASSERT(settings.bCopyFiles); _ASSERT(false == settings.bCopyIfNewerOrHigherVer); _ASSERT(0 == settings.app.CompareNoCase(L"cmd.exe"));
		_ASSERT(settings.appArgs.IsEmpty()); 
		_ASSERT(0 == settings.srcFileInfos[0].fullFilePath.CompareNoCase(L"C:\\windows\\notepad.exe"));
		_ASSERT(0 == settings.destFileInfos[0].filenameOnly.CompareNoCase(L"cmd.exe"));
		_ASSERT(settings.destFileInfos[0].fullFilePath.IsEmpty());
	}

	{
		Settings settings;
		//false because all the files in regression1 can't be found
		_ASSERT(ParseCommandLine(settings, L"\"C:\\dir with space\\PAExec.exe\" -clist debug\\regression1.txt -c myapp.exe"));

		_ASSERT(settings.bCopyFiles); _ASSERT(false == settings.bCopyIfNewerOrHigherVer); _ASSERT(0 == settings.app.CompareNoCase(L"myapp.exe"));
		_ASSERT(settings.appArgs.IsEmpty()); 

		_ASSERT(0 == settings.srcFileInfos[0].filenameOnly.CompareNoCase(L"paexec.exe")); //using paexec.exe so it is found in the debug folder
		_ASSERT(0 == settings.srcFileInfos[1].filenameOnly.CompareNoCase(L"paexec.obj"));
		_ASSERT(0 == settings.srcFileInfos[2].filenameOnly.CompareNoCase(L"paexec.pdb"));

		_ASSERT(0 == settings.destFileInfos[0].filenameOnly.CompareNoCase(L"myapp.exe")); //file name changed in dest based on settings.app
		_ASSERT(0 == settings.destFileInfos[1].filenameOnly.CompareNoCase(L"paexec.obj"));
		_ASSERT(0 == settings.destFileInfos[2].filenameOnly.CompareNoCase(L"paexec.pdb"));

		_ASSERT(FALSE == settings.srcFileInfos[0].fullFilePath.IsEmpty());
		_ASSERT(FALSE == settings.srcFileInfos[1].fullFilePath.IsEmpty());
		_ASSERT(FALSE == settings.srcFileInfos[2].fullFilePath.IsEmpty());

		_ASSERT(FALSE == settings.srcDir.IsEmpty());

		_ASSERT(0 == settings.destFileInfos[0].filenameOnly.CompareNoCase(L"myapp.exe"));

		_ASSERT(settings.destFileInfos.size() == settings.srcFileInfos.size());
	}

	{
		Settings settings;
		_ASSERT(ParseCommandLine(settings, L"\"C:\\dir with space\\PAExec.exe\" -clist debug\\regression2.txt -c myapp.exe"));

		_ASSERT(settings.bCopyFiles); _ASSERT(false == settings.bCopyIfNewerOrHigherVer); _ASSERT(0 == settings.app.CompareNoCase(L"myapp.exe"));
		_ASSERT(settings.appArgs.IsEmpty()); 

		_ASSERT(0 == settings.srcFileInfos[0].filenameOnly.CompareNoCase(L"paexec.exe"));
		_ASSERT(0 == settings.destFileInfos[0].filenameOnly.CompareNoCase(L"myapp.exe")); //dest using different filename based on settings.app

		_ASSERT(FALSE == settings.srcFileInfos[0].fullFilePath.IsEmpty());

		_ASSERT(FALSE == settings.srcDir.IsEmpty());

		_ASSERT(settings.destFileInfos.size() == settings.srcFileInfos.size());
	}

	{
		Settings settings;

		_ASSERT(ParseCommandLine(settings, L"\\dell8 C:\\Windows\\System32\\wevtutil.exe qe System /q:*^[System^[^(EventID=6008^)^]^] /f:text /c:1 /rd:true > c:\\temp\\tmpfile.txt 2>NUL"));

		_ASSERT(false == settings.bCopyFiles); //ensure /c in the wevtutil.exe command line is not seen by PAExec's command line parsing 

		_ASSERT(false == settings.bCopyIfNewerOrHigherVer); 
		_ASSERT(0 == settings.app.CompareNoCase(L"C:\\Windows\\System32\\wevtutil.exe"));
		_ASSERT(FALSE == settings.appArgs.IsEmpty()); 

		_ASSERT(settings.destFileInfos.size() == settings.srcFileInfos.size());
	}

	{
		Settings settings;

		_ASSERT(ParseCommandLine(settings, L"\\\\dell8 C:\\Windows\\System32\\wevtutil.exe qe System /q:*^[System^[^(EventID=6008^)^]^] /f:text -c:1 /rd:true > c:\\temp\\tmpfile.txt 2>NUL"));

		_ASSERT(false == settings.bCopyFiles); //ensure -c in the wevtutil.exe command line is not seen by PAExec's command line parsing 

		_ASSERT(false == settings.bCopyIfNewerOrHigherVer); 
		_ASSERT(0 == settings.app.CompareNoCase(L"C:\\Windows\\System32\\wevtutil.exe"));
		_ASSERT(FALSE == settings.appArgs.IsEmpty()); 

		_ASSERT(settings.destFileInfos.size() == settings.srcFileInfos.size());
	}

	{
		Settings settings;

		_ASSERT(ParseCommandLine(settings, L"\\\\dell8 CSCRIPT C:\\Windows\\System32\\eventquery.vbs /fi \"id eq 6008\" /l system 2^>NUL ^| find \"EventLog\""));

		_ASSERT(false == settings.bCopyFiles); 

		_ASSERT(false == settings.bCopyIfNewerOrHigherVer); 
		_ASSERT(0 == settings.app.CompareNoCase(L"CSCRIPT"));
		_ASSERT(FALSE == settings.appArgs.IsEmpty()); 

		_ASSERT(settings.destFileInfos.size() == settings.srcFileInfos.size());
	}

	{
		Settings settings;

		_ASSERT(ParseCommandLine(settings, L"\\\\dell8 -c -cnodel notepad.exe"));
	}

	
	Log(L" --- end regression tests --- \r\n", false);
}


BOOL WINAPI ConsoleCtrlHandler(DWORD dwCtrlType)
{
	switch( dwCtrlType )
	{
	case CTRL_C_EVENT:
	case CTRL_BREAK_EVENT:
	case CTRL_CLOSE_EVENT:
	case CTRL_LOGOFF_EVENT:
	case CTRL_SHUTDOWN_EVENT:
		printf("^C (stopping)\r\n");
		gbStop = true;
		return TRUE;
	}

	return FALSE;
}


bool CheckTimeout(__time64_t start, LPCWSTR serverName, Settings& settings)
{
	if((0 != settings.remoteCompConnectTimeoutSec) && ((_time64(NULL) - start) > settings.remoteCompConnectTimeoutSec))
	{
		Log(StrFormat(L"Timeout connecting to %s...", serverName), true);
		return true;
	}
	return false;
}




bool Settings::ResolveFilePaths()
{
	bool bAllFilesFound = true;

	//dir may or may not be set.  if not set, grab it from first file (which will have to be found on the path)
	_ASSERT(false == srcFileInfos.empty());
	_ASSERT(false == destFileInfos.empty());

	CString* pDir = &srcDir;
	std::vector<FileInfo>* pFileList = &srcFileInfos;
	
	if(gbInService)
	{
#ifdef _DEBUG
		Log(L"DEBUG: ResolveFilePaths in service", false);
#endif
		pDir = &destDir;
		pFileList = &destFileInfos;
	}

	std::vector<CString> results;
	for(std::vector<FileInfo>::iterator itr = pFileList->begin(); pFileList->end() != itr; itr++)
	{
		CString path = *pDir;
		if((FALSE == path.IsEmpty()) && (path.Right(1) != L"\\"))
			path += L"\\";
		path += (*itr).filenameOnly;

		wchar_t expanded[_MAX_PATH * 4] = {0};
		ExpandEnvironmentStrings(path, expanded, sizeof(expanded)/sizeof(wchar_t));
		if(0 != wcslen(expanded))
			path = expanded;

		if(0 != _waccess(path, 0))
		{
			path = ExpandToFullPath(path); //search on path if no path specified

			if(0 != _waccess(path, 0))
				bAllFilesFound = false;
		}

		if(pFileList->begin() == itr) 
		{
			if((false == bAllFilesFound) && (false == gbInService))
				return false; //can't find the very first (executable target) file, so bail
			
			if(pDir->IsEmpty())
			{
				//figure out source/target dir based on this first file
				LPWSTR cPtr = (LPWSTR)wcsrchr(path, L'\\');
				if(NULL != cPtr)
				{
					*cPtr = L'\0'; //truncate path
					*pDir = (LPCWSTR)path;
					*cPtr = L'\\'; //restore
				}
				else
				{
					//file must not exist, so assume windows directory
					GetWindowsDirectory(pDir->GetBuffer(_MAX_PATH), _MAX_PATH);
					pDir->ReleaseBuffer();
					//and now repair path
					path = *pDir;
					if((FALSE == path.IsEmpty()) && (path.Right(1) != L"\\"))
						path += L"\\";
					path += (*itr).filenameOnly;
				}
			}
		}

		(*itr).fullFilePath = path;

#ifdef _DEBUG
		Log(StrFormat(L"DEBUG: Setting full path to %s", path), false);
#endif

		bAllFilesFound &= GetTargetFileInfo(*itr); //non-existent files in the service will get marked as bCopy = true
	}

	_ASSERT(FALSE == pDir->IsEmpty());

	return bAllFilesFound;
}
