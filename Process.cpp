// Process.cpp: Executing requested process
//
// Copyright (c) Power Admin LLC, 2012 - 2013
//
// This code is provided "as is", with absolutely no warranty expressed
// or implied. Any use is at your own risk.
//
//////////////////////////////////////////////////////////////////////


#include "stdafx.h"
#include "PAExec.h"
#include <UserEnv.h>
#include <winsafer.h>
#include <Sddl.h>
#include <Psapi.h>
#include <WtsApi32.h>


#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


HANDLE GetLocalSystemProcessToken();
bool LimitRights(HANDLE& hUser);
bool ElevateUserToken(HANDLE& hEnvUser);

void Duplicate(HANDLE& h, LPCSTR file, int line)
{
	HANDLE hDupe = NULL;
	if(DuplicateTokenEx(h, MAXIMUM_ALLOWED, NULL, SecurityImpersonation, TokenPrimary, &hDupe))
	{
		CloseHandle(h);
		h = hDupe;
		hDupe = NULL;
	}
	else
	{
		DWORD gle = GetLastError();
		_ASSERT(0);
		Log(StrFormat(L"Error duplicating a user token (%S, %d)", file, line), GetLastError());
	}
}

#ifdef _DEBUG
void GrantAllPrivs(HANDLE h)
{
	Log(L"DEBUG: GrantAllPrivs", false);
	CString privs = L"SeCreateTokenPrivilege,SeAssignPrimaryTokenPrivilege,SeLockMemoryPrivilege,SeIncreaseQuotaPrivilege,SeMachineAccountPrivilege,"
					L"SeTcbPrivilege,SeSecurityPrivilege,SeTakeOwnershipPrivilege,SeLoadDriverPrivilege,SeSystemProfilePrivilege,SeSystemtimePrivilege,SeProfileSingleProcessPrivilege,"
					L"SeIncreaseBasePriorityPrivilege,SeCreatePagefilePrivilege,SeCreatePermanentPrivilege,SeBackupPrivilege,SeRestorePrivilege,SeShutdownPrivilege,SeDebugPrivilege,"
					L"SeAuditPrivilege,SeSystemEnvironmentPrivilege,SeChangeNotifyPrivilege,SeRemoteShutdownPrivilege,SeUndockPrivilege,SeSyncAgentPrivilege,SeEnableDelegationPrivilege,"
					L"SeManageVolumePrivilege,SeImpersonatePrivilege,SeCreateGlobalPrivilege,SeTrustedCredManAccessPrivilege,SeRelabelPrivilege,SeIncreaseWorkingSetPrivilege,"
					L"SeTimeZonePrivilege,SeCreateSymbolicLinkPrivilege";

	wchar_t* pC = wcstok(privs.LockBuffer(), L",");
	while(NULL != pC)
	{
		EnablePrivilege(pC, h); //needed to call CreateProcessAsUser
		pC = wcstok(NULL, L",");
	}
}
#endif


void GetUserDomain(LPCWSTR userIn, CString& user, CString& domain)
{
	//run as specified user
	CString tmp = userIn;
	LPCWSTR userStr = NULL, domainStr = NULL;
	if(NULL != wcschr(userIn, L'@'))
		userStr = userIn; //leave domain as NULL
	else
	{
		if(NULL != wcschr(userIn, L'\\'))
		{
			domainStr = wcstok(tmp.LockBuffer(), L"\\");
			userStr = wcstok(NULL, L"\\");
		}
		else
		{
			//no domain given
			userStr = userIn;
			domainStr = L".";
		}
	}
	user = userStr;
	domain = domainStr;
}


bool GetUserHandle(Settings& settings, BOOL& bLoadedProfile, PROFILEINFO& profile, HANDLE hCmdPipe)
{
	DWORD gle = 0;

	if(settings.bUseSystemAccount) 
	{
		if(BAD_HANDLE(settings.hUser)) //might already have hUser from a previous call
		{
			EnablePrivilege(SE_DEBUG_NAME); //helps with OpenProcess, required for GetLocalSystemProcessToken
			settings.hUser = GetLocalSystemProcessToken();
			if(BAD_HANDLE(settings.hUser))
			{
				Log(L"Not able to get Local System token", true);
				return false;
			}
			else
				Log(L"Got Local System handle", false);

			Duplicate(settings.hUser, __FILE__, __LINE__);
		}
		return true;
	}
	else
	{
		//not Local System, so either as specified user, or as current user
		if(FALSE == settings.user.IsEmpty())
		{
			CString user, domain;
			GetUserDomain(settings.user, user, domain);

			BOOL bLoggedIn = LogonUser(user, domain.IsEmpty() ? NULL : domain, settings.password, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_WINNT50, &settings.hUser);
			gle = GetLastError();
#ifdef _DEBUG
			Log(L"DEBUG: LogonUser", gle);
#endif
			if((FALSE == bLoggedIn) || BAD_HANDLE(settings.hUser))
			{
				Log(StrFormat(L"Error logging in as %s.", settings.user), gle);
				return false;
			}
			else
				Duplicate(settings.hUser, __FILE__, __LINE__); //gives max rights

			if(!BAD_HANDLE(settings.hUser) && (false == settings.bDontLoadProfile))
			{
				EnablePrivilege(SE_RESTORE_NAME);
				EnablePrivilege(SE_BACKUP_NAME);
				bLoadedProfile = LoadUserProfile(settings.hUser, &profile);
#ifdef _DEBUG
				gle = GetLastError();
				Log(L"DEBUG: LoadUserProfile", gle);
#endif
			}
			return true;
		}
		else
		{
			//run as current user
			
			if(NULL != hCmdPipe)
			{
				BOOL b = ImpersonateNamedPipeClient(hCmdPipe);
				DWORD gle = GetLastError();
				if(FALSE == b)
					Log(L"Failed to impersonate client user", gle);
				else
					Log(L"Impersonated caller", false);
			}

			HANDLE hThread = GetCurrentThread();
			BOOL bDupe = DuplicateHandle(GetCurrentProcess(), hThread, GetCurrentProcess(), &hThread, 0, TRUE, DUPLICATE_SAME_ACCESS);
			DWORD gle = GetLastError();

			BOOL bOpen = OpenThreadToken(hThread, TOKEN_DUPLICATE | TOKEN_QUERY, TRUE, &settings.hUser);
			gle = GetLastError();
			if(1008 == gle) //no thread token
			{
				bOpen = OpenProcessToken(GetCurrentProcess(), TOKEN_DUPLICATE | TOKEN_QUERY, &settings.hUser);
				gle = GetLastError();
			}

			if(FALSE == bOpen)
				Log(L"Failed to open current user token", GetLastError());

			Duplicate(settings.hUser, __FILE__, __LINE__); //gives max rights
			RevertToSelf();
			return !BAD_HANDLE(settings.hUser);
		}
	}
}

bool StartProcess(Settings& settings, HANDLE hCmdPipe)
{
	//Launching as one of:
	//1. System Account
	//2. Specified account (or limited account)
	//3. As current process

	DWORD gle = 0;

	BOOL bLoadedProfile = FALSE;
	PROFILEINFO profile = {0};
	profile.dwSize = sizeof(profile);
	profile.lpUserName = (LPWSTR)(LPCWSTR)settings.user;
	profile.dwFlags = PI_NOUI;

	if(false == GetUserHandle(settings, bLoadedProfile, profile, hCmdPipe))
		return false;

	PROCESS_INFORMATION pi = {0};
	STARTUPINFO si = {0};
	si.cb = sizeof(si);
	si.dwFlags = STARTF_USESHOWWINDOW;
	si.wShowWindow = SW_SHOW;
	if(!settings.bInteractive)
		si.wShowWindow = SW_HIDE;
	if(!BAD_HANDLE(settings.hStdErr))
	{
		si.hStdError = settings.hStdErr;
		si.hStdInput = settings.hStdIn;
		si.hStdOutput = settings.hStdOut;
		si.dwFlags |= STARTF_USESTDHANDLES;
#ifdef _DEBUG
		Log(L"DEBUG: Using redirected handles", false);
#endif
	}
#ifdef _DEBUG
	else
		Log(L"DEBUG: Not using redirected IO", false);
#endif

	CString path = StrFormat(L"\"%s\"", settings.app);
	if(FALSE == settings.appArgs.IsEmpty())
	{
		path += L" ";
		path += settings.appArgs;
	}

	LPCWSTR startingDir = NULL;
	if(FALSE == settings.workingDir.IsEmpty())
		startingDir = settings.workingDir;

	DWORD launchGLE = 0;

	CleanupInteractive ci = {0};

	if(settings.bInteractive || settings.bShowUIOnWinLogon)
	{
		BOOL b = PrepForInteractiveProcess(settings, &ci, settings.sessionToInteractWith);
		if(FALSE == b)
			Log(L"Failed to PrepForInteractiveProcess", true);

		if(NULL == si.lpDesktop)
			si.lpDesktop = L"WinSta0\\Default"; 
		if(settings.bShowUIOnWinLogon)
			si.lpDesktop = L"winsta0\\Winlogon";
		//Log(StrFormat(L"Using desktop: %s", si.lpDesktop), false);
		//http://blogs.msdn.com/b/winsdk/archive/2009/07/14/launching-an-interactive-process-from-windows-service-in-windows-vista-and-later.aspx
		//indicates desktop names are case sensitive
	}
#ifdef _DEBUG
	Log(StrFormat(L"DEBUG: PAExec using desktop %s", si.lpDesktop == NULL ? L"{default}" : si.lpDesktop), false);
#endif

	DWORD dwFlags = CREATE_SUSPENDED | CREATE_NEW_CONSOLE;

	LPVOID pEnvironment = NULL;
	VERIFY(CreateEnvironmentBlock(&pEnvironment, settings.hUser, TRUE));
	if(NULL != pEnvironment)
		dwFlags |= CREATE_UNICODE_ENVIRONMENT;
#ifdef _DEBUG
	gle = GetLastError();
	Log(L"DEBUG: CreateEnvironmentBlock", gle);
#endif

	if(settings.bDisableFileRedirection)
		DisableFileRedirection();

	if(settings.bRunLimited)
		if(false == LimitRights(settings.hUser))
			return false;

	if(settings.bRunElevated)
		if(false == ElevateUserToken(settings.hUser))
			return false;

	CString user, domain;
	GetUserDomain(settings.user, user, domain);

#ifdef _DEBUG
	Log(StrFormat(L"DEBUG: U:%s D:%s P:%s bP:%d Env:%s WD:%s",
		user, domain, settings.password, settings.bDontLoadProfile,
		pEnvironment ? L"true" : L"null", startingDir ? startingDir : L"null"), false);
#endif

	BOOL bLaunched = FALSE;

	if(settings.bUseSystemAccount)
	{
		Log(StrFormat(L"PAExec starting process [%s] as Local System", path), false);

		if(BAD_HANDLE(settings.hUser))
			Log(L"Have bad user handle", true);

		EnablePrivilege(SE_IMPERSONATE_NAME);
		BOOL bImpersonated = ImpersonateLoggedOnUser(settings.hUser);
		if(FALSE == bImpersonated)
		{
			Log(L"Failed to impersonate", GetLastError());
			_ASSERT(bImpersonated);
		}
		EnablePrivilege(SE_ASSIGNPRIMARYTOKEN_NAME);
		EnablePrivilege(SE_INCREASE_QUOTA_NAME);
		bLaunched = CreateProcessAsUser(settings.hUser, NULL, path.LockBuffer(), NULL, NULL, TRUE, dwFlags, pEnvironment, startingDir, &si, &pi);
		launchGLE = GetLastError();
		path.UnlockBuffer();

#ifdef _DEBUG
		if(0 != launchGLE)
			Log(StrFormat(L"Launch (launchGLE=%u) params: user=[x%X] path=[%s] flags=[x%X], pEnv=[%s], dir=[%s], stdin=[x%X], stdout=[x%X], stderr=[x%X]",
						launchGLE, (DWORD)settings.hUser, path, dwFlags, pEnvironment ? L"{env}" : L"{null}", startingDir ? startingDir : L"{null}", 
						(DWORD)si.hStdInput, (DWORD)si.hStdOutput, (DWORD)si.hStdError), false);
#endif

		RevertToSelf();
	}
	else
	{
		if(FALSE == settings.user.IsEmpty()) //launching as a specific user
		{
			Log(StrFormat(L"PAExec starting process [%s] as %s", path, settings.user), false);

			if(false == settings.bRunLimited)
			{
				bLaunched = CreateProcessWithLogonW(user, domain.IsEmpty() ? NULL : domain, settings.password, settings.bDontLoadProfile ? 0 : LOGON_WITH_PROFILE, NULL, path.LockBuffer(), dwFlags, pEnvironment, startingDir, &si, &pi);
				launchGLE = GetLastError();
				path.UnlockBuffer();

#ifdef _DEBUG
				if(0 != launchGLE) 
					Log(StrFormat(L"Launch (launchGLE=%u) params: user=[%s] domain=[%s] prof=[x%X] path=[%s] flags=[x%X], pEnv=[%s], dir=[%s], stdin=[x%X], stdout=[x%X], stderr=[x%X]",
									launchGLE, user, domain, settings.bDontLoadProfile ? 0 : LOGON_WITH_PROFILE, 
									path, dwFlags, pEnvironment ? L"{env}" : L"{null}", startingDir ? startingDir : L"{null}", 
									(DWORD)si.hStdInput, (DWORD)si.hStdOutput, (DWORD)si.hStdError), false);
#endif
			}
			else
				bLaunched = FALSE; //force to run with CreateProcessAsUser so rights can be limited

			//CreateProcessWithLogonW can't be called from LocalSystem on Win2003 and earlier, so LogonUser/CreateProcessAsUser must be used. Might as well try for everyone
			if((FALSE == bLaunched) && !BAD_HANDLE(settings.hUser))
			{
#ifdef _DEBUG
				Log(L"DEBUG: Failed CreateProcessWithLogonW - trying CreateProcessAsUser", GetLastError());
#endif
				EnablePrivilege(SE_ASSIGNPRIMARYTOKEN_NAME);
				EnablePrivilege(SE_INCREASE_QUOTA_NAME);
				EnablePrivilege(SE_IMPERSONATE_NAME);
				BOOL bImpersonated = ImpersonateLoggedOnUser(settings.hUser);
				if(FALSE == bImpersonated)
				{
					Log(L"Failed to impersonate", GetLastError());
					_ASSERT(bImpersonated);
				}

				bLaunched = CreateProcessAsUser(settings.hUser, NULL, path.LockBuffer(), NULL, NULL, TRUE, CREATE_SUSPENDED | CREATE_UNICODE_ENVIRONMENT | CREATE_NEW_CONSOLE, pEnvironment, startingDir, &si, &pi);
				if(0 == GetLastError())
					launchGLE = 0; //mark as successful, otherwise return our original error
				path.UnlockBuffer();
#ifdef _DEBUG
				if(0 != launchGLE)
					Log(StrFormat(L"Launch (launchGLE=%u) params: user=[x%X] path=[%s] pEnv=[%s], dir=[%s], stdin=[x%X], stdout=[x%X], stderr=[x%X]",
									launchGLE, (DWORD)settings.hUser, path, pEnvironment ? L"{env}" : L"{null}", startingDir ? startingDir : L"{null}", 
									(DWORD)si.hStdInput, (DWORD)si.hStdOutput, (DWORD)si.hStdError), false);
#endif
				RevertToSelf();
			}
		}
		else
		{
			Log(StrFormat(L"PAExec starting process [%s] as current user", path), false);

			EnablePrivilege(SE_ASSIGNPRIMARYTOKEN_NAME);
			EnablePrivilege(SE_INCREASE_QUOTA_NAME);
			EnablePrivilege(SE_IMPERSONATE_NAME);

			if(NULL != settings.hUser)
				bLaunched = CreateProcessAsUser(settings.hUser, NULL, path.LockBuffer(), NULL, NULL, TRUE, dwFlags, pEnvironment, startingDir, &si, &pi);
			if(FALSE == bLaunched)
				bLaunched = CreateProcess(NULL, path.LockBuffer(), NULL, NULL, TRUE, dwFlags, pEnvironment, startingDir, &si, &pi);
			launchGLE = GetLastError();
	
//#ifdef _DEBUG
			if(0 != launchGLE)
				Log(StrFormat(L"Launch (launchGLE=%u) params: path=[%s] user=[%s], pEnv=[%s], dir=[%s], stdin=[x%X], stdout=[x%X], stderr=[x%X]",
					launchGLE, path, settings.hUser ? L"{non-null}" : L"{null}", pEnvironment ? L"{env}" : L"{null}", startingDir ? startingDir : L"{null}", 
					(DWORD)si.hStdInput, (DWORD)si.hStdOutput, (DWORD)si.hStdError), false);
//#endif
			path.UnlockBuffer();
		}
	}

	if(bLaunched)
	{
		if(gbInService)
			Log(L"Successfully launched", false);

		settings.hProcess = pi.hProcess;
		settings.processID = pi.dwProcessId;

		if(false == settings.allowedProcessors.empty())
		{
			DWORD sysMask = 0, procMask = 0;
			VERIFY(GetProcessAffinityMask(pi.hProcess, &procMask, &sysMask));
			procMask = 0;
			for(std::vector<WORD>::iterator itr = settings.allowedProcessors.begin(); settings.allowedProcessors.end() != itr; itr++)
			{
				DWORD bit = 1;
				bit = bit << (*itr - 1);
				procMask |= bit & sysMask;
			}
			VERIFY(SetProcessAffinityMask(pi.hProcess, procMask));
		}

		VERIFY(SetPriorityClass(pi.hProcess, settings.priority));
		ResumeThread(pi.hThread);
		VERIFY(CloseHandle(pi.hThread));
	}
	else
	{
		Log(StrFormat(L"Failed to start %s.", path), launchGLE);
		if((ERROR_ELEVATION_REQUIRED == launchGLE) && (false == gbInService))
			Log(L"HINT: PAExec probably needs to be \"Run As Administrator\"", false);
	}

	if(ci.bPreped)
		CleanUpInteractiveProcess(&ci);

	if(settings.bDisableFileRedirection)
		RevertFileRedirection();

	if(NULL != pEnvironment)
		DestroyEnvironmentBlock(pEnvironment);
	pEnvironment = NULL;

	if(bLoadedProfile)
		UnloadUserProfile(settings.hUser, profile.hProfile);

	if(!BAD_HANDLE(settings.hUser))
	{
		CloseHandle(settings.hUser);
		settings.hUser = NULL;
	}

	return bLaunched ? true : false;
}



CString GetTokenUserSID(HANDLE hToken)
{
	DWORD tmp = 0;
	CString userName;
	DWORD sidNameSize = 64;
	std::vector<WCHAR> sidName;
	sidName.resize(sidNameSize);

	DWORD sidDomainSize = 64;
	std::vector<WCHAR> sidDomain;
	sidDomain.resize(sidNameSize);

	DWORD userTokenSize = 1024;
	std::vector<WCHAR> tokenUserBuf;
	tokenUserBuf.resize(userTokenSize);

	TOKEN_USER *userToken = (TOKEN_USER*)&tokenUserBuf.front();

	if(GetTokenInformation(hToken, TokenUser, userToken, userTokenSize, &tmp))
	{
		WCHAR *pSidString = NULL;
		if(ConvertSidToStringSid(userToken->User.Sid, &pSidString))
			userName = pSidString;
		if(NULL != pSidString)
			LocalFree(pSidString);
	}
	else
		_ASSERT(0);

	return userName;
}

HANDLE GetLocalSystemProcessToken()
{
	DWORD pids[1024*10] = {0}, cbNeeded = 0, cProcesses = 0;

	if ( !EnumProcesses(pids, sizeof(pids), &cbNeeded))
	{
		Log(L"Can't enumProcesses - Failed to get token for Local System.", true);
		return NULL;
	}

	// Calculate how many process identifiers were returned.
	cProcesses = cbNeeded / sizeof(DWORD);
	for(DWORD i = 0; i<cProcesses; ++i)
	{
		DWORD gle = 0;
		DWORD dwPid = pids[i];
		HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, FALSE, dwPid);
		if (hProcess)
		{
			HANDLE hToken = 0;
			if (OpenProcessToken(hProcess, TOKEN_QUERY | TOKEN_READ | TOKEN_IMPERSONATE | TOKEN_QUERY_SOURCE | TOKEN_DUPLICATE | TOKEN_ASSIGN_PRIMARY | TOKEN_EXECUTE, &hToken))
			{
				try
				{
					CString name = GetTokenUserSID(hToken);
					
					//const wchar_t arg[] = L"NT AUTHORITY\\";
					//if(0 == _wcsnicmp(name, arg, sizeof(arg)/sizeof(arg[0])-1))

					if(name == L"S-1-5-18") //Well known SID for Local System
					{
						CloseHandle(hProcess);
						return hToken;
					}
				}
				catch(...)
				{
					_ASSERT(0);
				}
			}
			else
				gle = GetLastError();
			CloseHandle(hToken);
		}
		else 
			gle = GetLastError();
		CloseHandle(hProcess);
	}
	Log(L"Failed to get token for Local System.", true);
	return NULL;
}

typedef BOOL (WINAPI *SaferCreateLevelProc)(DWORD dwScopeId, DWORD dwLevelId, DWORD OpenFlags, SAFER_LEVEL_HANDLE* pLevelHandle, LPVOID lpReserved);
typedef BOOL (WINAPI *SaferComputeTokenFromLevelProc)(SAFER_LEVEL_HANDLE LevelHandle, HANDLE InAccessToken, PHANDLE OutAccessToken, DWORD dwFlags, LPVOID lpReserved);
typedef BOOL (WINAPI *SaferCloseLevelProc)(SAFER_LEVEL_HANDLE hLevelHandle);

bool LimitRights(HANDLE& hUser)
{
	DWORD gle = 0;

	static SaferCreateLevelProc gSaferCreateLevel = NULL;
	static SaferComputeTokenFromLevelProc gSaferComputeTokenFromLevel = NULL;
	static SaferCloseLevelProc gSaferCloseLevel = NULL;

	if((NULL == gSaferCloseLevel) || (NULL == gSaferComputeTokenFromLevel) || (NULL == gSaferCreateLevel))
	{
		HMODULE hMod = LoadLibrary(L"advapi32.dll"); //GLOK
		if(NULL != hMod)
		{
			gSaferCreateLevel = (SaferCreateLevelProc)GetProcAddress(hMod, "SaferCreateLevel");
			gSaferComputeTokenFromLevel = (SaferComputeTokenFromLevelProc)GetProcAddress(hMod, "SaferComputeTokenFromLevel");
			gSaferCloseLevel = (SaferCloseLevelProc)GetProcAddress(hMod, "SaferCloseLevel");
		}
	}

	if((NULL == gSaferCloseLevel) || (NULL == gSaferComputeTokenFromLevel) || (NULL == gSaferCreateLevel))
	{
		Log(L"Safer... calls not supported on this OS -- can't limit rights", true);
		return false;
	}

	if(!BAD_HANDLE(hUser))
	{
		HANDLE hNew = NULL;
		SAFER_LEVEL_HANDLE safer = NULL;
		if(FALSE == gSaferCreateLevel(SAFER_SCOPEID_USER, SAFER_LEVELID_NORMALUSER, SAFER_LEVEL_OPEN, &safer, NULL))
		{
			gle = GetLastError();
			Log(L"Failed to limit rights (SaferCreateLevel).", gle);
			return false;
		}

		if(NULL != safer)
		{
			if(FALSE == gSaferComputeTokenFromLevel(safer, hUser, &hNew, 0, NULL))
			{
				gle = GetLastError();
				Log(L"Failed to limit rights (SaferComputeTokenFromLevel).", gle);
				VERIFY(gSaferCloseLevel(safer));
				return false;
			}
			VERIFY(gSaferCloseLevel(safer));
		}

		//if(BAD_HANDLE(hNew)) //try a second approach?
		//{
		//	if(FALSE == CreateRestrictedToken(hUser, DISABLE_MAX_PRIVILEGE, 0, 0, 0, 0, 0, 0, &hNew))
		//	{
		//		gle = GetLastError();
		//		Log(L"Failed to limit rights (CreateRestrictedToken).", gle);
		//	}
		//}

		if(!BAD_HANDLE(hNew))
		{
			VERIFY(CloseHandle(hUser));
			hUser = hNew;
			Duplicate(hUser, __FILE__, __LINE__);
			return true;
		}
	}

	Log(L"Don't have a good user -- can't limit rights", true);
	return false;
}


bool ElevateUserToken(HANDLE& hEnvUser)
{
	TOKEN_ELEVATION_TYPE tet;
	DWORD needed = 0;
	DWORD gle = 0;

	if(GetTokenInformation(hEnvUser, TokenElevationType, (LPVOID)&tet, sizeof(tet), &needed))
	{
		if(tet == TokenElevationTypeLimited)
		{
			//get the associated token, which is the full-admin token
			TOKEN_LINKED_TOKEN tlt = {0};
			if(GetTokenInformation(hEnvUser, TokenLinkedToken, (LPVOID)&tlt, sizeof(tlt), &needed))
			{
				Duplicate(tlt.LinkedToken, __FILE__, __LINE__);
				hEnvUser = tlt.LinkedToken;
				return true;
			}
			else
			{
				gle = GetLastError();
				Log(L"Failed to get elevated token", gle);
				return false;
			}
		}
		else
			return true;
	}
	else
	{
		//can't tell if it's elevated or not -- continue anyway

		gle = GetLastError();
		switch(gle)
		{ 
		case ERROR_INVALID_PARAMETER: //expected on 32-bit XP
		case ERROR_INVALID_FUNCTION: //expected on 64-bit XP
			break;
		default:
			Log(L"Can't query token to run elevated - continuing anyway", gle);
			break;
		}

		return true;
	}
}


