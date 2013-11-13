// ServiceImpl.cpp: Implementation for remote service
//
// Copyright (c) Power Admin LLC, 2012 - 2013
//
// This code is provided "as is", with absolutely no warranty expressed
// or implied. Any use is at your own risk.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "PAExec.h"
#include <process.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


__time64_t gLastClientContact = 0;
bool gbStop = false;
SERVICE_STATUS_HANDLE ghService = NULL;
SERVICE_STATUS gServiceStatus = {0};
static bool gbDevOnlyDebug = false;
bool gbInService = false;
extern volatile long gInProcessRequests;

void __cdecl StopServiceAsync(void*)
{
	while(false == gbStop)
	{
		Sleep(5000);
		if(0 == gInProcessRequests)
		{
			//wait 2 more seconds to see if it's still at zero to hope we didn't just catch it in between requests
			Sleep(2000);
			if(0 == gInProcessRequests)
			{
				Log(L"PAExec service stopping (async)", false);
				gbStop = true; //signals to other places in app to stop too
				if(NULL != ghService)
				{
					gServiceStatus.dwCurrentState = SERVICE_STOPPED;
					SetServiceStatus(ghService, &gServiceStatus);
				}
			}
		}
	}
}

void WINAPI ServiceControlHandler(DWORD dwControl)
{
	switch(dwControl)
	{
	case SERVICE_CONTROL_SHUTDOWN:
	case SERVICE_CONTROL_STOP:
		if(0 == gInProcessRequests)
		{
			gbStop = true;
			_ASSERT(NULL != ghService);
			if(NULL != ghService)
			{
				Log(L"PAExec service stopping", false);
				gServiceStatus.dwCurrentState = SERVICE_STOPPED;
				SetServiceStatus(ghService, &gServiceStatus);
			}
		}
		else
		{
			Log(L"PAExec received request to stop, but still have active requests.  Will stop later.", false);
			_beginthread(StopServiceAsync, 0, 0);
		}
		break;
	}
}

#define BUFF_SIZE_HINT 16384

UINT WINAPI PipeClientThreadProc(void* p)
{
	HANDLE hPipe = (HANDLE)p;

	bool bBail = false;

	while(true) 
	{ 
		RemMsg request, response;

		BYTE buff[BUFF_SIZE_HINT] = {0};

		// Read client requests from the pipe. 
		bool bFirstRead = true;
		BOOL fSuccess = FALSE;
		DWORD gle = 0;
		do 
		{ 
			DWORD cbRead = 0;
			// Read from the pipe. 
			fSuccess = ReadFile( 
				hPipe,			// pipe handle 
				buff,			// buffer to receive reply 
				sizeof(buff), // size of buffer 
				&cbRead,  // number of bytes read 
				NULL);    // not overlapped 

			if (!fSuccess && GetLastError() != ERROR_MORE_DATA) 
			{
				gle = GetLastError();
				Log(L"Error reading request from pipe -- stopping service", gle);
				bBail = true;
				break;
			}

			if(bFirstRead)
			{
				if(cbRead < sizeof(WORD))
				{
					gle = GetLastError();
					Log(L"Received too little data in request -- stopping service", gle);
					bBail = true;
					break;
				}
				request.SetFromReceivedData(buff, cbRead);
				bFirstRead = false;
			}
			else
				request.m_payload.insert(request.m_payload.end(), buff, buff + cbRead);
		} while (!fSuccess);  // repeat loop if ERROR_MORE_DATA 

		if(bBail)
			break;

		HandleMsg(request, response, hPipe);

		DWORD totalLen = 0;
		const BYTE* pDataToSend = response.GetDataToSend(totalLen);
		DWORD cbWritten = 0;
		// Send a message to the pipe server. 
		fSuccess = WriteFile( 
					hPipe,     // pipe handle 
					pDataToSend, // message 
					totalLen,  // message length 
					&cbWritten,// bytes written 
					NULL);     // not overlapped 
		if (!fSuccess || (cbWritten != totalLen)) 
		{
			gle = GetLastError();
			Log(L"Error sending data back -- stopping service.", gle);
			bBail = true;
			break;
		}
	}

	// Flush the pipe to allow the client to read the pipe's contents 
	// before disconnecting. Then disconnect the pipe, and close the 
	// handle to this pipe instance. 

	FlushFileBuffers(hPipe); 
	DisconnectNamedPipe(hPipe); 
	CloseHandle(hPipe); 

	if(bBail)
		ServiceControlHandler(SERVICE_CONTROL_STOP);

	return 0;
}

VOID WINAPI ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv)
{
	gbODS = true; //service always logs to DbgView at first

	Log(L"PAExec ServiceMain starting.", false);
	gServiceStatus.dwCurrentState = SERVICE_START_PENDING;
	gServiceStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP | SERVICE_ACCEPT_SHUTDOWN;
	gServiceStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS; //when this is set, the service name isn't validated since there is only one service running in the process

	if(false == gbDevOnlyDebug)
	{
		ghService = RegisterServiceCtrlHandler(L"PAExec", ServiceControlHandler);
		if(0 == ghService)
		{
			Log(L"RegisterServiceCtrlHandler failed in PAExec", GetLastError());
			return;
		}

		gServiceStatus.dwCurrentState = SERVICE_RUNNING;
		BOOL b = SetServiceStatus(ghService, &gServiceStatus);
		if(FALSE == b)
			Log(L"Failed to signal PAExec running.", GetLastError());
	}

	Log(L"PAExec service running.", false);

	gLastClientContact = _time64(NULL);

	//Pipe server
	//the name of the pipe will be the name of the executable, which is based on the caller, so it should always be unique
	wchar_t path[_MAX_PATH * 2] = {0};
	GetModuleFileName(NULL, path, ARRAYSIZE(path));
	wchar_t* pC = wcsrchr(path, L'\\');
	if(NULL != pC)
		pC++; //get past backslash
	CString pipename;
	pipename.Format(L"\\\\.\\pipe\\%s", pC);

	// The main loop creates an instance of the named pipe and 
	// then waits for a client to connect to it. When the client 
	// connects, a thread is created to handle communications 
	// with that client, and the loop is repeated. 
	while(false == gbStop)
	{ 
		HANDLE hPipe = CreateNamedPipe( 
							pipename,             // pipe name 
							PIPE_ACCESS_DUPLEX,       // read/write access 
							PIPE_TYPE_MESSAGE |       // message type pipe 
							PIPE_READMODE_MESSAGE |   // message-read mode 
							PIPE_WAIT,                // blocking mode 
							PIPE_UNLIMITED_INSTANCES, // max. instances  
							BUFF_SIZE_HINT,           // output buffer size 
							BUFF_SIZE_HINT,           // input buffer size 
							0,                        // client time-out 
							NULL);                    // default security attribute 

		if (BAD_HANDLE(hPipe)) 
		{
			CString msg;
			msg.Format(L"PAExec failed to create pipe %s.", pipename);
			Log(msg, GetLastError());
			gbStop = true;
			continue;
		}
		else
		{
			CString msg;
			msg.Format(L"PAExec created pipe %s.", pipename);
			Log(msg, false);
		}

		// Wait for the client to connect; if it succeeds, 
		// the function returns a nonzero value. If the function
		// returns zero, GetLastError returns ERROR_PIPE_CONNECTED. 
		BOOL fConnected = ConnectNamedPipe(hPipe, NULL) ? TRUE : (GetLastError() == ERROR_PIPE_CONNECTED); 

		if (fConnected) 
		{ 
			// Create a thread for this client. 
			HANDLE hThread = (HANDLE)_beginthreadex(NULL, 0, PipeClientThreadProc, hPipe, 0, NULL);
			if (hThread == NULL) 
			{
				gbStop = true;
				continue;
			}
			else 
				CloseHandle(hThread); 
		} 
		else 
			// The client could not connect, so close the pipe. 
			CloseHandle(hPipe); 
	}

	Log(L"PAExec exiting loop.", false);

	if(false == gbDevOnlyDebug)
	{
		gServiceStatus.dwCurrentState = SERVICE_STOPPED;
		SetServiceStatus(ghService, &gServiceStatus);
	}
}

DWORD StartLocalService(CCmdLineParser& cmdParser)
{
	gbDevOnlyDebug = cmdParser.HasKey(L"dbg");
	gbInService = true;

	SERVICE_TABLE_ENTRY st[] =
	{
		{ L"PAExec", ServiceMain }, //If the service is installed with the SERVICE_WIN32_OWN_PROCESS service type, this member is ignored, but cannot be NULL. This member can be an empty string ("").
		{ NULL, NULL }
	};

	if(false == gbDevOnlyDebug)
	{
		if (!::StartServiceCtrlDispatcher(st))
		{
			gbODS = true;
			Log(L"PAExec failed to start ServiceCtrlDispatcher.", GetLastError());
			return GetLastError();
		}
	}
	else
	{
		_ASSERT(0);
		ServiceMain(0, NULL);
	}

	//for some reason (probably because the service didn't or couldn't stop when requested, and then couldn't be deleted), we're seeing cases where the service definition 
	//still hangs around (ie a long list of PAExec-xx-yy in services.msc), so here we'll take an additional step and try to delete ourself
	SC_HANDLE sch = OpenSCManager(NULL, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS);
	if(NULL != sch)
	{
		DWORD myPID = GetCurrentProcessId();

		BYTE buffer[63 * 1024] = {0};	
		DWORD ignored;
		DWORD serviceCount = 0;
		DWORD resume = 0;
		if(EnumServicesStatusEx(sch, SC_ENUM_PROCESS_INFO, SERVICE_WIN32_OWN_PROCESS, SERVICE_STATE_ALL, buffer, sizeof(buffer), &ignored, &serviceCount, &resume, NULL))
		{
			ENUM_SERVICE_STATUS_PROCESS* pStruct = (ENUM_SERVICE_STATUS_PROCESS*)buffer;
			for(DWORD i = 0; i < serviceCount; i++)
			{
				if(pStruct[i].ServiceStatusProcess.dwProcessId == myPID)
				{
					SC_HANDLE hSvc = OpenService(sch, pStruct[i].lpServiceName, SC_MANAGER_ALL_ACCESS);
					if(NULL != hSvc)
					{
						//service should be marked as stopped if we are down here
						::DeleteService(hSvc);
						CloseServiceHandle(hSvc);
					}
					break;
				}
			}
		}
		CloseServiceHandle(sch);
	}

	return 0;
}


