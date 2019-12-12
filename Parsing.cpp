// Parsing.cpp: Commandline parsing and interpretting
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

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


bool GetComputerList(Settings& settings, LPCWSTR& cmdLine)
{
	DWORD gle = 0;

	// [\\computer[,computer2[,...] | @file]]
	//@file PsExec will execute the command on each of the computers listed	in the file.
	//a computer name of "\\*" PsExec runs the applications on all computers in the current domain.

	//first part of command line is the executable itself, so skip past that
	if(L'"' == *cmdLine)
	{
		//path is quoted, so skip to end quote
		cmdLine = wcschr(cmdLine + 1, L'"');
		if(NULL != cmdLine)
			cmdLine++;
	}
	else
	{
		//no quotes, skip forward to whitespace
		while(!iswspace(*cmdLine) && *cmdLine)
			cmdLine++;
	}

	//now skip past white space
	while(iswspace(*cmdLine) && *cmdLine)
		cmdLine++;

	//if we see -accepteula or /accepteula, skip over it
	if( (0 == _wcsnicmp(cmdLine, L"/accepteula", 11)) || (0 == _wcsnicmp(cmdLine, L"-accepteula", 11)) )
	{
		cmdLine += 11; //get past Eula
		//now skip past white space
		while(iswspace(*cmdLine) && *cmdLine)
			cmdLine++;
	}

	if(0 == wcsncmp(cmdLine, L"\\\\*", 3))
	{
		cmdLine += 3; 
		//get server list from domain
		SERVER_INFO_100* pInfo = NULL;
		DWORD numServers = 0, total = 0;
		DWORD ignored = 0;
		NET_API_STATUS stat = NetServerEnum(NULL, 100, (LPBYTE*)&pInfo, MAX_PREFERRED_LENGTH, &numServers, &total, SV_TYPE_SERVER | SV_TYPE_WINDOWS, NULL, &ignored);
		if(NERR_Success == stat)
		{
			for(DWORD i = 0; i < numServers; i++)
				settings.computerList.push_back(pInfo[i].sv100_name);
		}
		else
			Log(L"Got error from NetServerEnum: ", (DWORD)stat);
		NetApiBufferFree(pInfo);
		pInfo = NULL;
		if(settings.computerList.empty())
			Log(L"No computers could be found", true);
		return !settings.computerList.empty();
	}
	else
	{
		if(L'@' == *cmdLine)
		{
			//read server list from file.  Assumes UTF8
			LPCWSTR fileStart = cmdLine + 1;
			while(!iswspace(*cmdLine) && *cmdLine)
				cmdLine++;
			CString file = CString(fileStart).Left(cmdLine - fileStart);
			
			file = ExpandToFullPath(file); //search on path if no path specified

			CString content;
			if(false == ReadTextFile(file, content))
				return false;

			wchar_t* pC = wcstok(content.LockBuffer(), L"\r\n");
			while(NULL != pC)
			{
				CString s = pC;
				s.Trim();
				if(FALSE == s.IsEmpty())
					settings.computerList.push_back(s);
				pC = wcstok(NULL, L"\r\n");
			}
			if(settings.computerList.empty())
				Log(L"Computer list file empty", true);
			return !settings.computerList.empty();
		}
		else
		{
			if(0 == wcsncmp(cmdLine, L"\\\\", 2))
			{
				//get, possibly comma-delimited, computer list
				LPCWSTR compListStart = cmdLine + 2;
				//skip forward to whitespace
				while(!iswspace(*cmdLine) && *cmdLine)
					cmdLine++;
				CString compList = CString(compListStart).Left(cmdLine - compListStart);
				wchar_t* pC = wcstok(compList.LockBuffer(), L",");
				while(NULL != pC)
				{
					CString s = pC;
					s.Trim();
					if(FALSE == s.IsEmpty())
						settings.computerList.push_back(s);
					pC = wcstok(NULL, L",");
				}

				if(settings.computerList.empty())
					Log(L"Computer not specified", true);
				return !settings.computerList.empty();
			}
			return true; //no server specified
		}
	}
}

LPCWSTR SkipForward(LPCWSTR cmdLinePastCompList, LPCWSTR arg, bool bCanHaveArg)
{
	LPCWSTR pStart = cmdLinePastCompList;
Top:
	LPCWSTR pC = wcsstr(pStart, arg);
	if(NULL != pC)
	{
		pC += wcslen(arg);
		if(false == iswspace(*pC))
		{
			//we're in a larger delimiter, so skip past this
			pStart = pC + 1;
			goto Top;
		}			

		while(iswspace(*pC) && *pC)
			pC++;
		if(bCanHaveArg)
			if(*pC != L'-')
			{
				bool bInQuote = false;
				while((!iswspace(*pC) || (*pC == L'"') || bInQuote) && *pC)
				{
					if(*pC == L'"')
						bInQuote = !bInQuote;
					pC++;
				}
				//to end of arg now
				while(iswspace(*pC) && *pC)
					pC++;
			}
			if(pC > cmdLinePastCompList)
				cmdLinePastCompList = pC;
	}
	return cmdLinePastCompList;
}

typedef struct 
{
	LPCWSTR cmd;
	bool	bCanHaveArgs;
	bool	bMustHaveArgs;
}CommandList;

CommandList gSupportedCommands[] = 
{
	{L"u", true, true},
	{L"p", true, false},
	{L"p@", true, true},
	{L"p@d", false, false},
	{L"n", true, true},
	{L"l", false, false},
	{L"h", false, false},
	{L"s", false, false},
	{L"e", false, false},
	{L"x", false, false},
	{L"i", true, false},
	{L"c", false, false},
	{L"cnodel", false, false},
	{L"f", false, false},
	{L"v", false, false},
	{L"w", true, true},
	{L"d", false, false},
	{L"low", false, false},
	{L"belownormal", false, false},
	{L"abovenormal", false, false},
	{L"high", false, false},
	{L"realtime", false, false},
	{L"background", false, false},
	{L"a", true, true},
	{L"csrc", true, true},
	{L"clist", true, true},
	{L"dfr", false, false},
	{L"lo", true, true},
	{L"rlo", true, true},
	{L"dbg", false, false},
	{L"to", true, true},
	{L"noname", false, false},
	{L"sname", true, true},
	{L"share", true, true},
	{L"sharepath", true, true},
	{L"accepteula", false, false} //non-documented PSExec command that we'll just silently eat
};

LPCWSTR EatWhiteSpace(LPCWSTR ptr)
{
	while(iswspace(*ptr) && *ptr)
		ptr++;

	return ptr;
}

bool SplitCommand(CString& restOfLine, LPCWSTR& paExecParams, LPCWSTR& appToRun)
{
	LPWSTR ptr = (LPWSTR)(LPCWSTR)restOfLine;
	while(true)
	{
		if((L'-' == *ptr) || (L'/' == *ptr))
		{
			if(NULL == paExecParams)
				paExecParams = ptr;

			ptr++;
			LPCWSTR startOfCmd = ptr;
			//skip to end of alpha chars which will signal the end of this command
			while( (iswalpha(*ptr) || (L'@' == *ptr)) && *ptr)
				ptr++;
			if(L'\0' == *ptr)
			{
				Log(L"Reached end of command before seeing expected parts", true);
				return false;
			}

			//ptr is beyond end of cmd
			size_t len = ptr - startOfCmd;

			bool bRecognized = false;
			//now see if this is a recognized command
			int i = 0;
			for(i = 0; i < sizeof(gSupportedCommands)/sizeof(gSupportedCommands[0]); i++)
			{
				if(wcslen(gSupportedCommands[i].cmd) != len)
					continue;
				CString lwrCmd = startOfCmd;
				lwrCmd.MakeLower();
				if(0 == wcsncmp(lwrCmd, gSupportedCommands[i].cmd, len))
				{
					bRecognized = true;
					break;
				}
			}

			if(false == bRecognized)
			{
				*ptr = L'\0';
				Log(StrFormat(L"%s is not a recognized option", startOfCmd), true);
				_ASSERT(0);
				return false;
			}				

			ptr = (LPWSTR)EatWhiteSpace(ptr);
			if(L'\0' == *ptr)
			{
				Log(L"Reached end of command before seeing expected parts", true);
				_ASSERT(0);
				return false;
			}

			if(gSupportedCommands[i].bCanHaveArgs)
			{
				if( (L'-' != *ptr) && (L'/' != *ptr))
				{
					//special handling for -i which may or may not have an argument, but if it does, it is numeric
					if(0 == wcscmp(gSupportedCommands[i].cmd, L"i"))
					{
						//wtodw(' ') and wtodw('0') both return 0
						if(L'0' != *ptr)
						{
							if(0 == wtodw(ptr))
								continue; //no argument
						}
					}
					
					bool bInQuote = false;
					while((!iswspace(*ptr) || (*ptr == L'"') || bInQuote) && *ptr)
					{
						if(*ptr == L'"')
							bInQuote = !bInQuote;
						ptr++;
					}
					//to end of arg now
					ptr = (LPWSTR)EatWhiteSpace(ptr);
					if((L'\0' == *ptr) && (gSupportedCommands[i].bMustHaveArgs))
					{
						Log(L"Reached end of command before seeing expected parts", true);
						_ASSERT(0);
						return false;
					}
				}
			}
		}
		else
		{
			//must have found the start of the app
			ptr[-1] = L'\0'; //terminate PAExec part
			appToRun = ptr;
			return true;			
		}
	}
}

bool ParseCommandLine(Settings& settings, LPCWSTR cmdLine)
{
	//Parse command line options into settings
	LPCWSTR cmdLinePastCompList = cmdLine;
	if(GetComputerList(settings, cmdLinePastCompList))
	{
		cmdLinePastCompList = EatWhiteSpace(cmdLinePastCompList);

		if(L'\0' == *cmdLinePastCompList)
			return false; //empty command line -- just show usage

		if((0 == wcsncmp(cmdLinePastCompList, L"/?", 2)) || (0 == wcsncmp(cmdLinePastCompList, L"-?", 2)))
			return false; //show usage

		//sitting on first character that could be program to run, or arguments.  Need to split the
		//rest of the string in half -- half for PAExec arguments, and the other half for the program
		//to run (and perhaps it's arguments as well)
		CString restOfLine = cmdLinePastCompList;
		LPCWSTR paExecParams = NULL;
		LPCWSTR pAppStart = NULL;
		if(false == SplitCommand(restOfLine, paExecParams, pAppStart))
			return false;

		CCmdLineParser cmdParser;
		//re-parse now that funky stuff has been skipped over
		cmdParser.Parse(paExecParams);

		if(cmdParser.HasKey(L"a"))
		{
			CString procList = cmdParser.GetVal(L"a");
			LPCWSTR pC = wcstok(procList.LockBuffer(), L",");
			while(NULL != pC)
			{
				DWORD num = wtodw(pC);
				if(0 != num)
					settings.allowedProcessors.push_back((WORD)num);
				pC = wcstok(NULL, L",");
			}
			if(settings.allowedProcessors.empty())
			{
				Log(L"-a specified without non-zero values", true);
				return false;
			}
		}

		if(cmdParser.HasKey(L"c"))
		{
			settings.bCopyFiles = true;
			if(cmdParser.HasKey(L"f"))
				settings.bForceCopy = true;
			else
				if(cmdParser.HasKey(L"v"))
					settings.bCopyIfNewerOrHigherVer = true;
		}

		if(cmdParser.HasKey(L"d"))
			settings.bDontWaitForTerminate = true;

		if(cmdParser.HasKey(L"i"))
		{
			settings.bInteractive = true;
			settings.sessionToInteractWith = (DWORD)-1; //default, meaning interactive session
			if(cmdParser.HasVal(L"i"))
				settings.sessionToInteractWith = wtodw(cmdParser.GetVal(L"i"));
		}

		if(cmdParser.HasKey(L"h"))
		{
			if(settings.bUseSystemAccount || settings.bRunLimited)
			{
				Log(L"Can't use -h and -l together", true);
				return false;
			}
			if(!settings.bUseSystemAccount) //ignore if -s already given
				settings.bRunElevated = true;
		}

		if(cmdParser.HasKey(L"l"))
		{
			if(settings.bUseSystemAccount || settings.bRunElevated)
			{
				Log(L"Can't use -h, -s, or -l together", true);
				return false;
			}
			settings.bRunLimited = true;
		}

		if(cmdParser.HasKey(L"n"))
		{
			if(settings.computerList.empty())
			{
				Log(L"Setting -n when no computers specified", true);
				return false;
			}
			settings.remoteCompConnectTimeoutSec = wtodw(cmdParser.GetVal(L"n"));
			if(0 == settings.remoteCompConnectTimeoutSec)
			{
				Log(L"Setting -n to 0", true);
				return false;
			}
		}

		bool bPasswordSet = false;
		if(cmdParser.HasKey(L"p"))
		{
			settings.password = cmdParser.GetVal(L"p");
			bPasswordSet = true;
			//empty passwords are supported
		}

		if(cmdParser.HasKey(L"p@"))
		{
			CString content, passFile = cmdParser.GetVal(L"p@");
			passFile.Trim(L"\""); //remove possible quotes
			ReadTextFile(passFile, content);
			wchar_t* pC = wcstok(content.LockBuffer(), L"\r\n");
			settings.password = pC;
			settings.password.Trim(L" ");
			content.UnlockBuffer();
			bPasswordSet = true;
			if(cmdParser.HasKey(L"p@d"))
				if(FALSE == DeleteFile(passFile))
					Log(L"Failed to delete p@d file", false);
		}

		if(cmdParser.HasKey(L"u"))
		{
			settings.user = cmdParser.GetVal(L"u");
			if(settings.user.IsEmpty())
			{
				Log(L"-u without user", true);
				return false;
			}

			if(false == bPasswordSet)
			{
#define PW_BUFF_LEN 500
				wchar_t pwBuffer[PW_BUFF_LEN] = {0};
				wprintf(L"Password: ");

				// Set the console mode to no-echo, not-line-buffered input
				DWORD mode;
				HANDLE ih = GetStdHandle( STD_INPUT_HANDLE  );
				GetConsoleMode(ih, &mode);
				SetConsoleMode(ih, (mode & ~ENABLE_ECHO_INPUT) | ENABLE_LINE_INPUT);
				DWORD read = 0;
				ReadConsole(ih, pwBuffer, PW_BUFF_LEN - 1, &read, NULL);
				pwBuffer[read] = L'\0';
				settings.password = pwBuffer;
				settings.password.TrimRight(L"\r\n");
				SetConsoleMode(ih, mode); // Restore the console mode
				wprintf(L"\r\n");
				if(gbStop)
					return false;
			}
		}

		if(cmdParser.HasKey(L"s"))
		{
			//It IS OK to use -u and -s (-u connects to server, then app is launched as -s)
			//if(FALSE == settings.user.IsEmpty())
			//{
			//	Log(L"Specified -s and -u");
			//	return false;
			//}
			if(settings.bRunLimited)
			{
				Log(L"Can't use -s and -l together", true);
				return false;
			}
			if(settings.bRunElevated)
				settings.bRunElevated = false; //ignore -h if -s given
			settings.bUseSystemAccount = true;
		}
		else
			if(cmdParser.HasKey(L"e"))
				settings.bDontLoadProfile = true;

		if(cmdParser.HasKey(L"w"))
		{
			settings.workingDir = cmdParser.GetVal(L"w");
			if(settings.workingDir.IsEmpty())
			{
				Log(L"-w without value", true);
				return false;
			}
		}

		if(cmdParser.HasKey(L"x"))
		{
			if(false == settings.bUseSystemAccount)
			{
				Log(L"Specified -x without -s", true);
				return false;
			}
			settings.bShowUIOnWinLogon = true;
		}

		//priority
		if(cmdParser.HasKey(L"low"))
			settings.priority = BELOW_NORMAL_PRIORITY_CLASS;
		if(cmdParser.HasKey(L"belownormal"))
			settings.priority = BELOW_NORMAL_PRIORITY_CLASS;
		if(cmdParser.HasKey(L"abovenormal"))
			settings.priority = ABOVE_NORMAL_PRIORITY_CLASS;
		if(cmdParser.HasKey(L"high"))
			settings.priority = HIGH_PRIORITY_CLASS;
		if(cmdParser.HasKey(L"realtime"))
			settings.priority = REALTIME_PRIORITY_CLASS;
		if(cmdParser.HasKey(L"background"))
			settings.priority = IDLE_PRIORITY_CLASS;

		if(cmdParser.HasKey(L"dfr"))
			settings.bDisableFileRedirection = true;

		if(cmdParser.HasKey(L"noname"))
			settings.bNoName = true;
		else if (cmdParser.HasKey(L"sname"))
			settings.serviceName = cmdParser.GetVal(L"sname");

		if (cmdParser.HasKey(L"share")) {
			settings.targetShare = cmdParser.GetVal(L"share");
			if (cmdParser.HasKey(L"sharepath")) {
				settings.targetSharePath = cmdParser.GetVal(L"sharepath");
			}
		}
			

		if(cmdParser.HasKey(L"csrc"))
		{
			if(false == settings.bCopyFiles)
			{
				Log(L"-csrc without -c", true);
				return false;
			}
			CString tmp = cmdParser.GetVal(L"csrc");
			if(tmp.IsEmpty())
			{
				Log(L"-csrc without value", true);
				return false;
			}
		
			//split tmp into directory and file
			LPWSTR cPtr = (LPWSTR)wcsrchr(tmp.LockBuffer(), L'\\');
			if(NULL != cPtr)
			{
				*cPtr = L'\0'; //truncate at filename
				cPtr++;
				FileInfo fi;
				fi.filenameOnly = cPtr;
				settings.srcFileInfos.push_back(fi);
			}
			tmp.UnlockBuffer();
			settings.srcDir = (LPCWSTR)tmp; //store folder
		}

		if(cmdParser.HasKey(L"clist"))
		{
			if(false == settings.srcFileInfos.empty())
			{
				Log(L"-csrc and -clist are not compatible", true);
				return false;
			}
			
			if(false == settings.bCopyFiles)
			{
				Log(L"-clist without -c", true);
				return false;
			}

			CString tmp = cmdParser.GetVal(L"clist");
			if(tmp.IsEmpty())
			{
				Log(L"-clist without value", true);
				return false;
			}
			tmp = ExpandToFullPath(tmp); //search on path if no path specified

			CString content;
			if(false == ReadTextFile(tmp, content))
				return false;

			//split tmp into directory and file
			LPWSTR cPtr = (LPWSTR)wcsrchr(tmp.LockBuffer(), L'\\');
			if(NULL != cPtr)
				*cPtr = L'\0'; //truncate at filename
			tmp.UnlockBuffer();
			settings.srcDir = (LPCWSTR)tmp; //store folder

			wchar_t* pC = wcstok(content.LockBuffer(), L"\r\n");
			while(NULL != pC)
			{
				CString s = pC;
				s.Trim();
				if(FALSE == s.IsEmpty())
				{
					FileInfo fi;
					fi.filenameOnly = s;
					settings.srcFileInfos.push_back(fi);
				}
				pC = wcstok(NULL, L"\r\n");
			}
			if(settings.srcFileInfos.empty())
			{
				Log(L"-clist file empty", true);
				return false;
			}
		}

		if(cmdParser.HasKey(L"cnodel"))
		{
			if(false == settings.bCopyFiles)
			{
				Log(L"-cnodel without -c", true);
				return false;
			}
			settings.bNoDelete = true;
		}

		if(cmdParser.HasKey(L"rlo"))
		{
			settings.remoteLogPath = cmdParser.GetVal(L"rlo");
			if(settings.remoteLogPath.IsEmpty())
			{
				Log(L"-rlo missing value", true);
				return false;
			}
		}

		if(cmdParser.HasKey(L"to"))
		{
			if(settings.bDontWaitForTerminate)
			{
				Log(L"-d and -to cannot be used together", true);
				return false;
			}
			settings.timeoutSeconds = wtodw(cmdParser.GetVal(L"to"));
			if(settings.timeoutSeconds == 0)
			{
				Log(L"invalid or missing value for -to", true);
				return false;
			}
		}

		//if(settings.bUseSystemAccount && settings.bRunLimited)
		//{
		//	Log(L"Can't run as System account, and run limited");
		//	return false;
		//}

		//if((FALSE == settings.user.IsEmpty()) && settings.bRunLimited)
		//{
		//	Log(L"Can't run limited and specify an account");
		//	return false;
		//}

		//program
		LPCWSTR pArgPtr = pAppStart;
		if(L'"' == *pAppStart)
		{
			pAppStart++;
			pArgPtr = wcschr(pAppStart, L'"');
			if(NULL == pArgPtr)
				return false;
			settings.app = CString(pAppStart).Left(pArgPtr - pAppStart);
			pArgPtr++; //get past quote we're on
		}
		else
		{
			//not quoted, so must be white space delimited
			while(!iswspace(*pArgPtr) && *pArgPtr)
				pArgPtr++;
			settings.app = CString(pAppStart).Left(pArgPtr - pAppStart);
		}

		if(settings.app.IsEmpty())
		{
			Log(L"No application specified", true);
			return false;
		}

		//arguments
		pArgPtr = EatWhiteSpace(pArgPtr);

		settings.appArgs = pArgPtr;

		if(settings.bCopyFiles) //if not copying files, it's OK if we can't find some
		{
			//now create destFileList based on app and srcFileList
			FileInfo fi;
			fi.filenameOnly = settings.app;
			settings.destFileInfos.push_back(fi); //index 0

			if(settings.srcFileInfos.empty())
			{
				//no source file specified (no -csrc or -clist) so just use settings.app, and resolve the path from it
				settings.srcFileInfos.push_back(fi);
			}

			_ASSERT(false == settings.srcFileInfos.empty());
			if(false == settings.srcFileInfos.empty())
			{
				std::vector<FileInfo>::iterator itr = settings.srcFileInfos.begin();  
				itr++; //skip over app (actual target executable) which we've already added to the list

				//copy rest of files (from -clist) if any
				while(itr != settings.srcFileInfos.end())
				{
					settings.destFileInfos.push_back(*itr);
					itr++;
				}
			}

			//read in source file information to verify all required source files exist
			if(false == settings.ResolveFilePaths())
			{
				Log(L"One or more source files were not found.  Try -csrc or -clist?", true);
				return false;
			}
		}

		return true;
	}

	return false;
}
