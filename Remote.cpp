#include "stdafx.h"
#include "PAExec.h"
#include <process.h>


volatile long gInProcessRequests = 0;

extern CString gLogPath;
CString GetRemoteServiceName(Settings& settings);
void SplitUserNameAndDomain(CString user, CString& username, CString& domain);
CString RemoveDomainFromUserName(CString user);

bool EstablishConnection(Settings& settings, LPCTSTR lpszRemote, LPCTSTR lpszResource, bool bConnect)
{
	DWORD rc = 0;
	if(0 == wcscmp(lpszRemote, L"."))
		return true; //already connected to self

	CString remoteResource = StrFormat(L"\\\\%s\\%s", lpszRemote, lpszResource);

	if((INVALID_HANDLE_VALUE == settings.hUserImpersonated) && (FALSE == settings.user.IsEmpty()))
	{
		CString user, domain;
		SplitUserNameAndDomain(settings.user, user, domain);
		bool success = ::LogonUserW(user, domain.IsEmpty() ? NULL : domain, settings.password, LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_WINNT50, &settings.hUserImpersonated);
		if (!success || BAD_HANDLE(settings.hUserImpersonated))
		{
			DWORD gle = GetLastError();
			Log(StrFormat(L"Failed to log on as remote user %s.", settings.user.GetString()), gle);
		}
	}

	if(bConnect)
	{
		//already connected??
		HANDLE hEnum = NULL;
		if(NO_ERROR == WNetOpenEnum(RESOURCE_CONNECTED, RESOURCETYPE_ANY, 0, NULL, &hEnum))
		{
			bool bConnected = false;
			BYTE buf[65536] = {0};
			DWORD count = (DWORD)-1;
			DWORD bufSize = sizeof(buf);
			WNetEnumResource(hEnum, &count, buf, &bufSize);
			for(DWORD i = 0; i < count; i++)
			{
				NETRESOURCE* pNR = (NETRESOURCE*)buf;
				if(0 == _wcsicmp(pNR[i].lpRemoteName, remoteResource))
				{
					bConnected = true;
					break;
				}					
			}
	
			WNetCloseEnum(hEnum);
			if(bConnected)
			{
				if(NULL != wcsstr(lpszResource, L"IPC$"))
					settings.bNeedToDetachFromIPC = false;
				else
					if(NULL != wcsstr(lpszResource, settings.targetShare))
						settings.bNeedToDetachFromAdmin = false;
					else
						_ASSERT(0);
				return true;
			}
		}

		NETRESOURCE nr = {0};
		nr.dwType = RESOURCETYPE_ANY;
		nr.lpLocalName = NULL;
		nr.lpRemoteName = (LPWSTR)(LPCWSTR)remoteResource;
		nr.lpProvider = NULL;

		//Establish connection (using username/pwd)
		rc = WNetAddConnection2(&nr, settings.password.IsEmpty() ? NULL : settings.password, settings.user.IsEmpty() ? NULL : settings.user, 0);
		if(NO_ERROR == rc)
		{
			if(NULL != wcsstr(lpszResource, L"IPC$"))
				settings.bNeedToDetachFromIPC = true;
			else
				if(NULL != wcsstr(lpszResource, settings.targetShare))
					settings.bNeedToDetachFromAdmin = true;
				else
					_ASSERT(0);
			return true;
		}
		else
		{
			Log(StrFormat(L"Failed to connect to %s.", remoteResource), rc);
			if(NULL != wcsstr(lpszResource, L"IPC$"))
				settings.bNeedToDetachFromIPC = false;
			else
				if(NULL != wcsstr(lpszResource, settings.targetShare))
					settings.bNeedToDetachFromAdmin = false;
				else
					_ASSERT(0);
			return false;
		}
	}
	else 
	{
		rc = WNetCancelConnection2(remoteResource, 0, FALSE);
		return true;
	}
}

void DeletePAExecFromRemote(LPCWSTR targetComputer, Settings& settings)
{
	CString dest = StrFormat(L"\\\\%s\\%s\\%s", targetComputer, settings.targetShare, GetRemoteServiceName(settings) + CString(L".exe"));
	if(0 == wcscmp(targetComputer, L"."))
	{
		//change to local system32 directory
		GetWindowsDirectory(dest.GetBuffer(_MAX_PATH * 2), _MAX_PATH * 2);
		dest.ReleaseBuffer();
		dest += L"\\";
		dest += GetRemoteServiceName(settings) + CString(L".exe");
	}

	if(NULL != settings.hUserImpersonated)
		ImpersonateLoggedOnUser(settings.hUserImpersonated);

	int tryCount = 0;
	while(FALSE == DeleteFile(dest))
	{
		DWORD gle = GetLastError();
		if(++tryCount < 70)
		{
			Sleep(100);
			continue;
		}
		Log(StrFormat(L"Failed to cleanup [%s] on %s.", dest, targetComputer), gle);
		break;
	}
}

bool CopyPAExecToRemote(Settings& settings, LPCWSTR targetComputer)
{
	CString remoteExeName = GetRemoteServiceName(settings);
	remoteExeName += L".exe";
	CString dest = StrFormat(L"\\\\%s\\%s\\%s", targetComputer, settings.targetShare, remoteExeName);

	wchar_t myPath[_MAX_PATH * 2] = {0};
	GetModuleFileName(NULL, myPath, ARRAYSIZE(myPath));

	if(0 == wcscmp(targetComputer, L"."))
	{
		//change to local Windows directory
		GetWindowsDirectory(dest.GetBuffer(_MAX_PATH * 2), _MAX_PATH * 2);
		dest.ReleaseBuffer();
		dest += L"\\";
		dest += remoteExeName;
	}

	if(0 == _wcsicmp(myPath, dest))
		return true; //don't need to copy -- running from there

	//try to use the credentials we were given (if any) so we don't try default credentials which might log in the Event Log
	if((FALSE == settings.user.IsEmpty()) && (false == settings.bNeedToDetachFromAdmin))
		EstablishConnection(settings, targetComputer, settings.targetShare, true);
	
	if(NULL != settings.hUserImpersonated)
	{
		BOOL b = ImpersonateLoggedOnUser(settings.hUserImpersonated);
		DWORD gle = GetLastError();
		if(FALSE == b)
			Log(StrFormat(L"Failed to impersonate [%s] - continuing anyway.", settings.user), gle);
	}

	//trying really hard above to not get failed logins in the Event Log but CopyFile seems intent on trying other connections besides what we've provided :(
	if((FALSE == CopyFile(myPath, dest, FALSE)) && (false == settings.bNeedToDetachFromAdmin))
	{
		DWORD gle = GetLastError();

		if (NULL != settings.hUserImpersonated)
			RevertToSelf();

		Log(StrFormat(L"Failed to copy [%s] to [%s] -- going to try to continue anyway.", myPath, dest), gle);
	}
	else
	{
		if (NULL != settings.hUserImpersonated)
			RevertToSelf();
		settings.bNeedToDeleteServiceFile = true;
	}

	return true;
}

//using unique service names so multiple remote requests (even from different source computers)
//can be happening at once
CString GetRemoteServiceName(Settings& settings)
{
	if(settings.bNoName)
		return L"PAExec";
	else if (!settings.serviceName.IsEmpty())
		return settings.serviceName;
	else
	{
		//Installed service will use a unique name so we can have multiple running at once
		CString name = L"PAExec-";
		wchar_t localComputerName[1000] = {0};
		DWORD len = ARRAYSIZE(localComputerName);
		GetComputerNameEx(ComputerNamePhysicalNetBIOS, localComputerName, &len);
		name += StrFormat(L"%u-%s", GetCurrentProcessId(), localComputerName);
		return name;
	}
}

void StopAndDeleteRemoteService(LPCWSTR remoteServer, Settings& settings)
{
	if(0 == wcscmp(remoteServer, L"."))
		remoteServer = NULL;

	//should already have a connection if one was needed
	if(NULL != settings.hUserImpersonated)
		ImpersonateLoggedOnUser(settings.hUserImpersonated);

	if(NULL != settings.hUserImpersonated)
		ImpersonateLoggedOnUser(settings.hUserImpersonated);

	SC_HANDLE hSCM = ::OpenSCManager(remoteServer, NULL, SC_MANAGER_ALL_ACCESS);
	DWORD gle = GetLastError();

	if (BAD_HANDLE(hSCM))
	{
		RevertToSelf();
		Log(StrFormat(L"Failed to connect to Service Control Manager on %s.  Can't cleanup PAExec.", remoteServer ? remoteServer : L"{local computer}"), gle);
		return;
	}

	// try to stop and delete now
	SC_HANDLE hService =::OpenService(hSCM, GetRemoteServiceName(settings), SERVICE_ALL_ACCESS);
	gle = GetLastError();
	if (!BAD_HANDLE(hService))
	{
		SERVICE_STATUS stat = {0};
		BOOL b = ControlService(hService, SERVICE_CONTROL_STOP, &stat);
		if(FALSE == b)
		{
			gle = GetLastError();
			Log(L"Failed to stop PAExec service.", gle);
		}

		int tryCount = 0;
		while(true)
		{
			SERVICE_STATUS_PROCESS ssp = {0};
			DWORD needed = 0;
			if(QueryServiceStatusEx(hService, SC_STATUS_PROCESS_INFO, (BYTE*)&ssp, sizeof(ssp), &needed))
				if(SERVICE_STOPPED == ssp.dwCurrentState)
				{
					if(DeleteService(hService))
						break;
					else
					{
						Log(L"Failed to delete PAExec service", GetLastError());
						break;
					}
				}

			tryCount++;
			if(tryCount > 300)
				break; //waited 30 seconds
			Sleep(100);
		}
		::CloseServiceHandle(hService);
	}
	::CloseServiceHandle(hSCM);

	RevertToSelf();
}


// Installs and starts the remote service on remote machine
bool InstallAndStartRemoteService(LPCWSTR remoteServer, Settings& settings)
{
	if(0 == wcscmp(remoteServer, L"."))
		remoteServer = NULL;

	//try to use given credentials (if any) so we don't try and connect using default user credentials 
	if((FALSE == settings.user.IsEmpty()) && ((false == settings.bNeedToDetachFromIPC) || (INVALID_HANDLE_VALUE == settings.hUserImpersonated)) )
		EstablishConnection(settings, remoteServer, L"IPC$", true);

	if(NULL != settings.hUserImpersonated)
		ImpersonateLoggedOnUser(settings.hUserImpersonated);

	SC_HANDLE hSCM = ::OpenSCManager(remoteServer, NULL, SC_MANAGER_ALL_ACCESS);
	DWORD gle = GetLastError();
	
	if(gbStop)
	{
		RevertToSelf();
		return false;
	}

	if (BAD_HANDLE(hSCM))
	{
		Log(StrFormat(L"Failed to connect to Service Control Manager on %s.", remoteServer ? remoteServer : L"{local computer}"), gle);
	}
	
	if(gbStop)
	{
		RevertToSelf();
		return false;
	}

	CString remoteServiceName = GetRemoteServiceName(settings);

	// Maybe it's already there and installed, let's try to run
	SC_HANDLE hService =::OpenService(hSCM, remoteServiceName, SERVICE_ALL_ACCESS);

	if (BAD_HANDLE(hService))
	{
		DWORD serviceType = SERVICE_WIN32_OWN_PROCESS;
		
		//as of Vista, services can no longer be interacted with
		//if( ((DWORD)-1 != settings.sessionToInteractWith) || (settings.bInteractive) )
		//	serviceType |= SERVICE_INTERACTIVE_PROCESS;

		CString svcExePath = StrFormat(L"%s\\%s.exe", settings.targetSharePath, remoteServiceName);
		if(NULL == remoteServer)
		{
			GetWindowsDirectory(svcExePath.GetBuffer(_MAX_PATH * 2), _MAX_PATH * 2);
			svcExePath.ReleaseBuffer();
			svcExePath += L"\\";
			svcExePath += remoteServiceName + L".exe";
		}
		svcExePath += L" -service"; //so it knows it is the service when it starts

		hService = ::CreateService(
						hSCM, remoteServiceName, remoteServiceName,
						SERVICE_ALL_ACCESS, 
						serviceType,
						SERVICE_DEMAND_START, SERVICE_ERROR_NORMAL,
						svcExePath,
						NULL, NULL, NULL, 
						NULL, NULL ); //using LocalSystem

		DWORD gle = GetLastError();
		if (BAD_HANDLE(hService))
		{
			::CloseServiceHandle(hSCM);
			Log(StrFormat(L"Failed to install service on %s", remoteServer ? remoteServer : L"{local computer}"), gle);
			RevertToSelf();
			return false;
		}
		settings.bNeedToDeleteService = true;
	}

	if(gbStop)
	{
		RevertToSelf();
		return false;
	}

	if(!StartService(hService, 0, NULL))
	{
		DWORD gle = GetLastError();
		if(ERROR_SERVICE_ALREADY_RUNNING != gle)
		{
			::CloseServiceHandle(hService);
			::CloseServiceHandle(hSCM);
			Log(StrFormat(L"Failed to start service on %s", remoteServer ? remoteServer : L"{local computer}"), gle);
			RevertToSelf();
			return false;
		}
		settings.bNeedToDeleteService = true;
	}

	::CloseServiceHandle(hService);
	::CloseServiceHandle(hSCM);

	RevertToSelf();

	return true;
}

bool SendSettings(LPCWSTR remoteServer, Settings& settings, HANDLE& hPipe, bool& bNeedToSendFile)
{
	RemMsg msg(MSGID_SETTINGS);
	settings.Serialize(msg, true);

	RemMsg response;
	if(false == SendRequest(remoteServer, hPipe, msg, response, settings))
		return false;	

	if(MSGID_RESP_SEND_FILES == response.m_msgID)
	{
		bNeedToSendFile = true;
		__int64 fileBits = 0;
		response >> fileBits; //each file is a bit indicating whether it should be copied or not
		__int64 testBit = 1;
		for(size_t i = 0; i < min(64, settings.srcFileInfos.size()); i++)
		{
			if(0 != (fileBits & testBit))
				settings.srcFileInfos[i].bCopyFile = true;
			testBit <<= 1;
		}
	}

	if(settings.bForceCopy)
		bNeedToSendFile = true;

	return true;
}


bool SendRequest(LPCWSTR remoteServer, HANDLE& hPipe, RemMsg& msgOut, RemMsg& msgReturned, Settings& settings)
{
	DWORD gle = 0;

	if(BAD_HANDLE(hPipe)) //need to connect?
	{
		CString remoteServiceName = GetRemoteServiceName(settings);
		CString pipeName = StrFormat(L"\\\\%s\\pipe\\%s", remoteServer, remoteServiceName + CString(".exe")); 
		int count = 0;
		while(true) 
		{ 
			hPipe = CreateFile( 
						pipeName,   // pipe name 
						GENERIC_READ |  // read and write access 
						GENERIC_WRITE, 
						0,              // no sharing 
						NULL,           // default security attributes
						OPEN_EXISTING,  // opens existing pipe 
						FILE_FLAG_OVERLAPPED, //using overlapped so we can poll gbStop flag too
						NULL);          // no template file 

			// Break if the pipe handle is valid. 
			if(!BAD_HANDLE(hPipe)) 
				break; 

			// Exit if an error other than ERROR_PIPE_BUSY occurs. 
			gle = GetLastError();
			if((ERROR_PIPE_BUSY != gle) && (ERROR_FILE_NOT_FOUND != gle))
			{
				Log(StrFormat(L"Failed to open communication channel to %s.", remoteServer), gle);
				return false;
			}

			count++;
			if(20 == count)
			{
				gle = GetLastError();
				Log(StrFormat(L"Timed out waiting for communication channel to %s.", remoteServer), gle);
				return false;
			}
			Sleep(1000);
		} 
	}

	//have a pipe so send
	// The pipe connected; change to message-read mode. 

	DWORD dwMode = PIPE_READMODE_MESSAGE; 
	BOOL fSuccess = SetNamedPipeHandleState( 
								hPipe,    // pipe handle 
								&dwMode,  // new pipe mode 
								NULL,     // don't set maximum bytes 
								NULL);    // don't set maximum time 
	if (!fSuccess) 
	{
		gle = GetLastError();
		Log(StrFormat(L"Failed to set communication channel to %s.", remoteServer), gle);
		CloseHandle(hPipe);
		hPipe = INVALID_HANDLE_VALUE;
		return false;
	}

	HANDLE hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);

	OVERLAPPED ol = {0};
	ol.hEvent = hEvent;

	DWORD totalLen = 0;
	const BYTE* pDataToSend = msgOut.GetDataToSend(totalLen);
	DWORD cbWritten = 0;
	// Send a message to the pipe server. 
	fSuccess = WriteFile( 
						hPipe,     // pipe handle 
						pDataToSend, // message 
						totalLen,  // message length 
						&cbWritten,// bytes written 
						&ol);     
	
	if (!fSuccess) 
	{
		gle = GetLastError();
		if(ERROR_IO_PENDING != gle)
		{
			Log(StrFormat(L"Error communicating with %s.", remoteServer), gle);
			CloseHandle(hPipe);
			hPipe = INVALID_HANDLE_VALUE;
			return false;
		}
	}

	while(true)
	{
		if(WAIT_OBJECT_0 == WaitForSingleObject(ol.hEvent, 0))
		{
			GetOverlappedResult(hPipe, &ol, &cbWritten, FALSE);
			break;
		}
		if(gbStop)
			return false;
		Sleep(100);
	}
	
	bool bFirstRead = true;
	do 
	{ 
		BYTE buffer[16384] = {0};
		
		ol.Offset = 0;
		ol.OffsetHigh = 0;

		DWORD cbRead = 0;
		// Read from the pipe. 
		fSuccess = ReadFile( 
			hPipe,			// pipe handle 
			buffer,			// buffer to receive reply 
			sizeof(buffer), // size of buffer 
			&cbRead,  // number of bytes read 
			&ol);    

		if(!fSuccess) 
		{
			gle = GetLastError();
			if(ERROR_IO_PENDING != gle)
			{
				Log(StrFormat(L"Error reading response from %s.", remoteServer), gle);
				CloseHandle(hPipe);
				hPipe = INVALID_HANDLE_VALUE;
				return false;
			}
		}

		while(true)
		{
			if(WAIT_OBJECT_0 == WaitForSingleObject(ol.hEvent, 0))
			{
				GetOverlappedResult(hPipe, &ol, &cbRead, FALSE);
				break;
			}
			if(gbStop)
				return false;
			Sleep(100);
		}

		if(bFirstRead)
		{
			if(cbRead < sizeof(WORD))
			{
				gle = GetLastError();
				Log(StrFormat(L"Received too little data from %s.", remoteServer), gle);
				CloseHandle(hPipe);
				hPipe = INVALID_HANDLE_VALUE;
				return false;
			}
			msgReturned.SetFromReceivedData(buffer, cbRead);
			bFirstRead = false;
		}
		else
			msgReturned.m_payload.insert(msgReturned.m_payload.end(), buffer, buffer + cbRead);
	} while( msgReturned.m_expectedLen != msgReturned.m_payload.size());

	CloseHandle(hEvent);
	return true; 
}

UINT WINAPI ListenOnRemotePipes(void* p)
{
	ListenParam* pLP = (ListenParam*)p;

	ConnectToRemotePipes(pLP, 10, 1000);

	InterlockedDecrement(&pLP->workerThreads);
	return 0;
}

void StartRemoteApp(LPCWSTR remoteServer, Settings& settings, HANDLE& hPipe, int& appExitCode)
{
	appExitCode = 0;

	RemMsg req(MSGID_START_APP), response;
	req << GetCurrentProcessId();
	wchar_t compName[MAX_COMPUTERNAME_LENGTH + 1] = {0};
	DWORD len = sizeof(compName)/sizeof(wchar_t);
	GetComputerName(compName, &len);
	req << compName;

	ListenParam lp;

	if((false == settings.bDontWaitForTerminate) && (false == settings.bInteractive) )
	{
		//ConnectToRemotePipes won't return until remote pipes are started and connected,
		//but they won't be created until MSGID_START_APP is called.  And SendRequest won't
		//return until the remote shuts down, so we have to launch a separate thread to listen
		//for the remote pipes so the main thread can continue forward and send the message
		lp.pSettings = &settings;
		lp.remoteServer = remoteServer;
		lp.machineName = compName;
		UINT ignored;
		InterlockedIncrement(&lp.workerThreads);
		HANDLE hThread = (HANDLE)_beginthreadex(NULL, 0, ListenOnRemotePipes, &lp, 0, &ignored);
		CloseHandle(hThread);
	}

	SendRequest(remoteServer, hPipe, req, response, settings);

	if(response.m_msgID == MSGID_OK)
	{
		response >> (DWORD&)appExitCode;		
		if(false == settings.bDontWaitForTerminate)
			Log(StrFormat(L"%s returned %i", settings.app, appExitCode), false);
		else
		{
			//exit code is really PID if not waiting
			Log(StrFormat(L"%s started with process ID %u", settings.app, (DWORD)appExitCode), false);
			appExitCode = 0; //value for returning from PAExec
		}
	}
	else
	{
		CString lastLog;
		response >> lastLog;
		Log(StrFormat(L"Remote app failed to start.  Returned error:\r\n  %s", lastLog), true);
		appExitCode = -9;
	}
	
	SetEvent(lp.hStop);
	int count = 0;
	while(0 < lp.workerThreads)
	{
		count++;
		if(gbStop)
			if(count > 6)
				break;
		Sleep(500);
	}
}

bool SendFilesToRemote(LPCWSTR remoteServer, Settings& settings, HANDLE& hPipe)
{
	int index = -1;
	for(std::vector<FileInfo>::iterator itr = settings.srcFileInfos.begin(); settings.srcFileInfos.end() != itr; itr++)
	{
		index++;
		if((*itr).bCopyFile)
		{
			CString src = (*itr).fullFilePath;
			_ASSERT(FALSE == src.IsEmpty());
			CString dest = StrFormat(L"\\\\%s\\%s\\PAExec_Move%u.dat", remoteServer, settings.targetShare, index);

			if(0 == wcscmp(remoteServer, L"."))
			{
				//change to local Windows directory
				GetWindowsDirectory(dest.GetBuffer(_MAX_PATH * 2), _MAX_PATH * 2);
				dest.ReleaseBuffer();
				dest += StrFormat(L"\\PAExec_Move%u.dat", index);
			}

			//make connection if we haven't already and credentials were given
			if ((FALSE == settings.user.IsEmpty()) && ((false == settings.bNeedToDetachFromAdmin) || (NULL == settings.hUserImpersonated)))
				EstablishConnection(settings, remoteServer, settings.targetShare, true);

			if(NULL != settings.hUserImpersonated)
				ImpersonateLoggedOnUser(settings.hUserImpersonated);

			if(FALSE == CopyFile(src, dest, FALSE))
			{
				DWORD gle = GetLastError();
				
				RevertToSelf();

				Log(StrFormat(L"Failed to copy [%s] to [%s].", src, dest), gle);
				return false;
			}
			RevertToSelf();
		}
	}

	RemMsg msg(MSGID_SENT_FILES); //tell service file is there so it can move it to the final location (potentially using local drive letters, etc)
	RemMsg response;
	bool b = SendRequest(remoteServer, hPipe, msg, response, settings);
	if(b)
	{
		if(response.m_msgID != MSGID_OK)
			b = false;
	}

	return b;
}


static bool							gRemoteSettingsMapCSInited = false;
static CRITICAL_SECTION				gRemoteSettingsMapCS;
static std::map<DWORD, Settings*>	gRemoteSettingsMap; //first is client's RemMsg::m_uniqueProcessID

//run at the remote service
void HandleMsg(RemMsg& msg, RemMsg& response, HANDLE hPipe)
{
	InterlockedIncrement(&gInProcessRequests);

	if(false == gRemoteSettingsMapCSInited)
	{
		gRemoteSettingsMapCSInited = true;
		InitializeCriticalSection(&gRemoteSettingsMapCS);
	}

#ifdef _DEBUG
	Log(StrFormat(L"DEBUG: PAExec service handling msg %u", msg.m_msgID), false);
#endif

	EnterCriticalSection(&gRemoteSettingsMapCS);
	Settings* pRemoteSettings = gRemoteSettingsMap[msg.m_uniqueProcessID];
	if(NULL == pRemoteSettings)
		pRemoteSettings = new Settings();
	LeaveCriticalSection(&gRemoteSettingsMapCS);

	switch(msg.m_msgID)
	{
	case MSGID_START_APP:
		{
			DWORD srcPID = 0;
			CString compName;
			msg >> srcPID;
			msg >> compName;

			if((false == pRemoteSettings->bDontWaitForTerminate) && (false == pRemoteSettings->bInteractive) )
			{
				if(false == CreateIOPipesInService(*pRemoteSettings, compName, srcPID))
				{
					DWORD gle = GetLastError();
					Log(StrFormat(L"Failed to open comm pipes for %s. ", pRemoteSettings->app), gle);
					response.m_msgID = MSGID_FAILED;

					if((false == pRemoteSettings->bNoDelete) && (false == pRemoteSettings->bDontWaitForTerminate))
					{
						for(std::vector<FileInfo>::iterator itr = pRemoteSettings->destFileInfos.begin(); pRemoteSettings->destFileInfos.end() != itr; itr++)
						{
							if((*itr).bCopyFile)
								DeleteFile((*itr).fullFilePath);
						}
					}
					break;
				}
			}
			else
			{
				Log(StrFormat(L"Not using redirected IO: DontWait=%u, Interactive=%u", pRemoteSettings->bDontWaitForTerminate, pRemoteSettings->bInteractive), false);
			}

			if(false == StartProcess(*pRemoteSettings, hPipe))
			{
				response.m_msgID = MSGID_FAILED;
				response << LastLog();
			}
			else
			{
				response.m_msgID = MSGID_OK;
				DWORD exitCode = 0;
				if(false == pRemoteSettings->bDontWaitForTerminate)
				{
					//doesn't work :(
//					if(!BAD_HANDLE(gpRemoteSettings->hStdIn))
//					{
//						Sleep(1000); //wait for app to start so we can disable local echo
//						if(AttachConsole(gpRemoteSettings->processID))
//						{
//							DWORD oldMode = 0;
//							GetConsoleMode(gpRemoteSettings->hStdIn, &oldMode);
//							oldMode = oldMode & ~ENABLE_ECHO_INPUT; //turn off echo
//							SetConsoleMode(gpRemoteSettings->hStdIn, oldMode);
//							FreeConsole();
//#ifdef _DEBUG
//							Log(L"DEBUG: Might have quieted echo...");
//#endif
//						}
//#ifdef _DEBUG
//						else
//						{
//							DWORD gle = GetLastError();
//							Log(L"DEBUG: Failed to quite echo in child.", gle);
//						}
//#endif
//					}

					//wait for app to shutdown
					DWORD waitMS = pRemoteSettings->timeoutSeconds * 1000;
					if(waitMS == 0)
						waitMS = INFINITE;

					DWORD ret = WaitForSingleObject(pRemoteSettings->hProcess, waitMS);
					switch(ret)
					{
					case WAIT_TIMEOUT:
						Log(L"PAExec timed out waiting for app to exit -- terminating app", true);
						TerminateProcess(pRemoteSettings->hProcess, (DWORD)-10);
						break;
					case WAIT_OBJECT_0: break;
					default: Log(L"PAExec error waiting for app to exit", GetLastError()); 
						break;
					}
					GetExitCodeProcess(pRemoteSettings->hProcess, &exitCode);
					Log(StrFormat(L"%s returned %d", pRemoteSettings->app, exitCode), false);

					if((false == pRemoteSettings->bNoDelete) && (false == pRemoteSettings->bDontWaitForTerminate))
					{
						for(std::vector<FileInfo>::iterator itr = pRemoteSettings->destFileInfos.begin(); pRemoteSettings->destFileInfos.end() != itr; itr++)
						{
							if((*itr).bCopyFile)
								DeleteFile((*itr).fullFilePath);
						}
					}
				}
				else
				{
					Log(StrFormat(L"Not waiting for %s to exit", pRemoteSettings->app), false);
					exitCode = pRemoteSettings->processID;
				}

				response << exitCode;
			}

			CloseHandle(pRemoteSettings->hProcess);
			pRemoteSettings->hProcess = NULL;

			//close pipes we redirected too also
			if(!BAD_HANDLE(pRemoteSettings->hStdErr))
			{
				CloseHandle(pRemoteSettings->hStdErr);
				pRemoteSettings->hStdErr = NULL;
			}

			if(!BAD_HANDLE(pRemoteSettings->hStdIn))
			{
				CloseHandle(pRemoteSettings->hStdIn);
				pRemoteSettings->hStdIn = NULL;
			}

			if(!BAD_HANDLE(pRemoteSettings->hStdOut))
			{
				CloseHandle(pRemoteSettings->hStdOut);
				pRemoteSettings->hStdOut = NULL;
			}
			break;
		}

	case MSGID_SETTINGS:
		//the first call from the remote PAExec
		pRemoteSettings->Serialize(msg, false);
		
		response.m_msgID = MSGID_OK;
		if(gLogPath.IsEmpty())
			gLogPath = pRemoteSettings->remoteLogPath;
		gbODS = pRemoteSettings->bODS; //now service will use whatever command was sent

		Log(StrFormat(L"-------------\r\nUser: '%s', LocalSystem: %s, Interactive: %s, Session: %i", pRemoteSettings->user, pRemoteSettings->bUseSystemAccount ? L"true" : L"false", pRemoteSettings->bInteractive ? L"true" : L"false", pRemoteSettings->sessionToInteractWith), false);

		//now see if the target file needs to be sent
		if(pRemoteSettings->bCopyFiles)
		{
			Log(L"Settings indicate files may need to be copied -- checking...", false);

			//below line also gets version information if possible
			pRemoteSettings->ResolveFilePaths(); //after logging is setup.  non-existent files will be marked as needing to be copied
			pRemoteSettings->app = pRemoteSettings->destFileInfos[0].fullFilePath;

			int index = 0;
			FILETIME zero = {0};
			for(std::vector<FileInfo>::iterator itr = pRemoteSettings->destFileInfos.begin(); pRemoteSettings->destFileInfos.end() != itr; itr++)
			{
				if(pRemoteSettings->bForceCopy)
					(*itr).bCopyFile = true;
				else
				{
					if(pRemoteSettings->bCopyIfNewerOrHigherVer)
					{
						//need to check date and/or version
						ULARGE_INTEGER destVer, srcVer;
						destVer.HighPart = (*itr).fileVersionMS;
						destVer.LowPart = (*itr).fileVersionLS;
						srcVer.HighPart = pRemoteSettings->srcFileInfos[index].fileVersionMS;
						srcVer.LowPart = pRemoteSettings->srcFileInfos[index].fileVersionLS;
						if(srcVer.QuadPart > destVer.QuadPart)
						{
							//source version is higher
							(*itr).bCopyFile = true;
						}
						else
						{
							//newer date?
							destVer.HighPart = (*itr).fileLastWrite.dwHighDateTime;
							destVer.LowPart= (*itr).fileLastWrite.dwLowDateTime;

							srcVer.HighPart = pRemoteSettings->srcFileInfos[index].fileLastWrite.dwHighDateTime;
							srcVer.LowPart= pRemoteSettings->srcFileInfos[index].fileLastWrite.dwLowDateTime;

							if(srcVer.QuadPart > destVer.QuadPart)
								(*itr).bCopyFile = true; //source version has a newer date
						}
					}
					else
					{
						//only copy if the target file doesn't exist
						ULARGE_INTEGER destVer; //file date
						destVer.HighPart = (*itr).fileLastWrite.dwHighDateTime;
						destVer.LowPart= (*itr).fileLastWrite.dwLowDateTime;
						if(0 == destVer.QuadPart)
							(*itr).bCopyFile = true;
					}
				}
			}

			__int64 sendFileBits = 0;  //a bit mask indicating whether a file needs to be copied or not
			__int64 startingBit = 1;
			//any files need to be copied?
			for(std::vector<FileInfo>::iterator itr = pRemoteSettings->destFileInfos.begin(); pRemoteSettings->destFileInfos.end() != itr; itr++)
			{
				if((*itr).bCopyFile)
				{
					response.m_msgID = MSGID_RESP_SEND_FILES;
					sendFileBits |= startingBit;
				}
				startingBit <<= 1;
			}
			response << sendFileBits;

			if(MSGID_RESP_SEND_FILES == response.m_msgID)
				Log(L"    ...need to copy files", false);
			else
				Log(L"    ...no files need to be copied", false);
		}
		break;

	case MSGID_SENT_FILES:
		{
			//copy from ADMIN$\PAExec_Move{index}.dat to final location
			response.m_msgID = MSGID_OK; //optimism

			for(size_t index = 0; index < pRemoteSettings->destFileInfos.size(); index++)
			{
				if(pRemoteSettings->destFileInfos[index].bCopyFile)
				{
					wchar_t srcPath[_MAX_PATH * 2] = {0};
					GetWindowsDirectory(srcPath, sizeof(srcPath)/sizeof(wchar_t));
					wcscat(srcPath, StrFormat(L"\\PAExec_Move%u.dat", index));
					if(FALSE == CopyFile(srcPath, pRemoteSettings->destFileInfos[index].fullFilePath, FALSE))
					{
						DWORD gle = GetLastError();
						if(0 != _waccess(srcPath, 0))
							Log(L"Source file does not exist", true);
						Log(StrFormat(L"Failed to move [%s] to [%s]", srcPath, pRemoteSettings->destFileInfos[index].fullFilePath), gle);
						response.m_msgID = MSGID_FAILED;
						break;
					}
					else
					{
						Log(StrFormat(L"Moved %s to %s", srcPath, pRemoteSettings->destFileInfos[index].fullFilePath), false);
					}
					DeleteFile(srcPath);
				}
			}
			break;
		}
	}

	//save it in the map for later use
	EnterCriticalSection(&gRemoteSettingsMapCS);
	gRemoteSettingsMap[msg.m_uniqueProcessID] = pRemoteSettings;
	LeaveCriticalSection(&gRemoteSettingsMapCS);

#ifdef _DEBUG
	Log(StrFormat(L"DEBUG: PAExec service finished handling msg %u", msg.m_msgID), false);
#endif

	InterlockedDecrement(&gInProcessRequests);
}


const BYTE* RemMsg::GetDataToSend(DWORD& totalLen)
{
	if(NULL != m_pBuff)
		delete [] m_pBuff;

	totalLen = 0;

	//build our send buffer
	_ASSERT(sizeof(m_msgID) == sizeof(WORD));
	totalLen = m_payload.size() + sizeof(WORD) + sizeof(DWORD) + sizeof(DWORD);
	if(MSGID_SETTINGS == m_msgID)
	{
		//send an extra DWORD of our random XOR value
		totalLen += sizeof(DWORD);
	}

	m_pBuff	= new BYTE[totalLen]; 
	BYTE* pPtr = m_pBuff;

	memcpy(pPtr, &m_msgID, sizeof(m_msgID));
	pPtr += sizeof(m_msgID);

	BYTE* pXORStart = NULL;
	DWORD xorVal = 0;
	rand_s((unsigned int*)&xorVal);
	if(MSGID_SETTINGS == m_msgID)
	{
		memcpy(pPtr, &xorVal, sizeof(xorVal));
		pPtr += sizeof(xorVal);
		pXORStart = pPtr;
	}

	memcpy(pPtr, &m_uniqueProcessID, sizeof(m_uniqueProcessID));
	pPtr += sizeof(m_uniqueProcessID);

	DWORD len = (DWORD)m_payload.size();
	memcpy(pPtr, &len, sizeof(len));
	pPtr += sizeof(len);

	if(0 != len)
	{
		memcpy(pPtr, &(*m_payload.begin()), m_payload.size());
		pPtr += m_payload.size();
	}

	if(NULL != pXORStart)
	{
		DWORD dataLen = pPtr - pXORStart;
		//flip the rest of the data
		for(DWORD i = 0; i < dataLen - (sizeof(DWORD) - 1); i++)
		{
			DWORD* pDW = (DWORD*)(pXORStart + i);
			*pDW ^= xorVal;
			xorVal += 3; //just being tricky :)
		}
	}

	return m_pBuff;
}

void RemMsg::SetFromReceivedData(BYTE* pData, DWORD dataLen)
{
	memcpy(&m_msgID, pData, sizeof(m_msgID));
	pData += sizeof(m_msgID);
	dataLen -= sizeof(m_msgID);
	
	if(MSGID_SETTINGS == m_msgID)
	{
		DWORD xorVal = 0;
		memcpy(&xorVal, pData, sizeof(xorVal));
		pData += sizeof(xorVal);
		dataLen -= sizeof(xorVal);
	
		//flip the rest of the data
		for(DWORD i = 0; i < dataLen - (sizeof(DWORD) - 1); i++)
		{
			DWORD* pDW = (DWORD*)(pData + i);
			*pDW ^= xorVal;
			xorVal += 3; //just being tricky :)
		}
	}

	m_uniqueProcessID = 0;
	memcpy(&m_uniqueProcessID, pData, sizeof(m_uniqueProcessID));
	pData += sizeof(m_uniqueProcessID);
	dataLen -= sizeof(m_uniqueProcessID);

	m_expectedLen = 0;
	memcpy(&m_expectedLen, pData, sizeof(m_expectedLen));
	pData += sizeof(m_expectedLen);
	dataLen -= sizeof(m_expectedLen);

	m_payload.clear();
	if(0 != dataLen)
		m_payload.insert(m_payload.end(), pData, pData + dataLen);
	m_bResetReadItr = true;
}

DWORD RemMsg::GetUniqueID()
{
	static DWORD id = 0;

	if(0 == id)
	{
		//a unique ID to identify this process on this machine to the service on the other side
		id = GetCurrentProcessId();
		wchar_t compNameBuffer[1024] = {0};
		DWORD buffLen = sizeof(compNameBuffer)/sizeof(wchar_t);
		GetComputerName(compNameBuffer, &buffLen);
		for(size_t s = 0; s < wcslen(compNameBuffer); s++)
		{
			DWORD* pDW = (DWORD*)compNameBuffer;
			id ^= pDW[s];
			s += 4; //sizeof DWORD
		}
	}

	return id;
}



