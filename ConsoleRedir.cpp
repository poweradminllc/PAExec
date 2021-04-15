// ConsoleRedir.cpp: Redirection of Console I/O
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
#include <conio.h>

const DWORD gPIPE_TYPE = PIPE_TYPE_BYTE | PIPE_WAIT; //PIPE_TYPE_MESSAGE | PIPE_WAIT;

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


// Creates named pipes for stdout, stderr, stdin on the (remote) service side
bool CreateIOPipesInService(Settings& settings, LPCWSTR caller, DWORD pid)
{
	SECURITY_DESCRIPTOR SecDesc;
	InitializeSecurityDescriptor(&SecDesc, SECURITY_DESCRIPTOR_REVISION);
	SetSecurityDescriptorDacl(&SecDesc, TRUE, NULL, FALSE);

	SECURITY_ATTRIBUTES SecAttrib = {0};
	SecAttrib.nLength = sizeof(SECURITY_ATTRIBUTES);
	SecAttrib.lpSecurityDescriptor = &SecDesc;;
	SecAttrib.bInheritHandle = TRUE;

	DWORD gle = 0;
	CString pipeName;

	// Create StdOut pipe
	pipeName = StrFormat(L"\\\\.\\pipe\\PAExecOut%s%u", caller, pid);
	settings.hStdOut = CreateNamedPipe(pipeName,
		PIPE_ACCESS_OUTBOUND, 
		gPIPE_TYPE, 
		PIPE_UNLIMITED_INSTANCES,
		0,
		0,
		(DWORD)-1,
		&SecAttrib);
	gle = GetLastError();
	if(BAD_HANDLE(settings.hStdOut))
		Log(StrFormat(L"PAExec failed to create pipe %s.", pipeName), gle);

	// Create StdError pipe
	pipeName = StrFormat(L"\\\\.\\pipe\\PAExecErr%s%u", caller, pid);
	settings.hStdErr = CreateNamedPipe(pipeName,
		PIPE_ACCESS_OUTBOUND, 
		gPIPE_TYPE, 
		PIPE_UNLIMITED_INSTANCES,
		0,
		0,
		(DWORD)-1,
		&SecAttrib);
	gle = GetLastError();
	if(BAD_HANDLE(settings.hStdErr))
		Log(StrFormat(L"PAExec failed to create pipe %s.", pipeName), gle);

	// Create StdIn pipe
	pipeName = StrFormat(L"\\\\.\\pipe\\PAExecIn%s%u", caller, pid);
	settings.hStdIn = CreateNamedPipe(pipeName,
		PIPE_ACCESS_INBOUND, 
		gPIPE_TYPE, 
		PIPE_UNLIMITED_INSTANCES,
		0,
		0,
		(DWORD)-1,
		&SecAttrib);
	gle = GetLastError();
	if(BAD_HANDLE(settings.hStdIn))
		Log(StrFormat(L"PAExec failed to create pipe %s.", pipeName), gle);

	if (BAD_HANDLE(settings.hStdOut) ||
		BAD_HANDLE(settings.hStdErr) ||
		BAD_HANDLE(settings.hStdIn))
	{
		CloseHandle(settings.hStdOut);
		CloseHandle(settings.hStdErr);
		CloseHandle(settings.hStdIn);
		settings.hStdOut = NULL;
		settings.hStdErr = NULL;
		settings.hStdIn = NULL;

		Log(L"Error creating redirection pipes", true);
		return false;
	}

	// Waiting for client to connect to this pipe
	ConnectNamedPipe(settings.hStdOut, NULL );
	ConnectNamedPipe(settings.hStdErr, NULL );
	ConnectNamedPipe(settings.hStdIn, NULL );

	Log(L"DEBUG: Client connected to pipe", false);

	return true;
}

#define SIZEOF_BUFFER 256 //don't go much higher than 30000 or ERROR_NOT_ENOUGH_MEMORY is returned from ReadConsole

// Listens the remote stdout pipe
// Remote process will send its stdout to this pipe
UINT WINAPI ListenRemoteOutPipeThread(void* p)
{
	ListenParam* pLP = (ListenParam*)p;

	HANDLE hOutput = GetStdHandle( STD_OUTPUT_HANDLE );
	char szBuffer[SIZEOF_BUFFER] = {0};
	DWORD dwRead = 0;
	
	HANDLE hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);

	while(false == gbStop)
	{ 
		OVERLAPPED olR = {0};
		olR.hEvent = hEvent;
		if (!ReadFile( pLP->pSettings->hStdOut, szBuffer, SIZEOF_BUFFER - 1, &dwRead, &olR) || (dwRead == 0)) 
		{
			DWORD dwErr = GetLastError();
			if ( dwErr == ERROR_NO_DATA)
				break;
		}

		if(gbStop)
			break;

		HANDLE waits[2];
		waits[0] = pLP->hStop;
		waits[1] = hEvent;
		DWORD ret = WaitForMultipleObjects(2, waits, FALSE, INFINITE);
		if(ret == WAIT_OBJECT_0)
			break; //need to exit
		_ASSERT(ret == WAIT_OBJECT_0 + 1); //data in buffer now

		// Handle CLS command, just for fun :)
		switch( szBuffer[0] )
		{
		case 12: //cls
			{
				DWORD dwWritten = 0;
				COORD origin = {0,0};
				CONSOLE_SCREEN_BUFFER_INFO sbi = {0};

				if ( GetConsoleScreenBufferInfo( hOutput, &sbi ) )
				{
					FillConsoleOutputCharacter( 
						hOutput,
						_T(' '),
						sbi.dwSize.X * sbi.dwSize.Y,
						origin,
						&dwWritten );

					SetConsoleCursorPosition(
						hOutput,
						origin );
				}
			}
			continue;
			break;
		}

		szBuffer[ dwRead / sizeof(szBuffer[0]) ] = _T('\0');

		if(gbStop)
			break;

 		//before printing, see if this is output that should be suppressed
		EnterCriticalSection(&pLP->cs);
		bool bSuppress = false;
		for(std::vector<std::string>::iterator itr = pLP->inputSentToSuppressInOutput.begin(); pLP->inputSentToSuppressInOutput.end() != itr; itr++)
		{
			if(0 == strcmp(szBuffer, (*itr).c_str()))
			{
				bSuppress = true;
				pLP->inputSentToSuppressInOutput.erase(itr);
				break;
			}
		}
		LeaveCriticalSection(&pLP->cs);

		if(false == bSuppress)
		{
			// Send it to our stdout
			fprintf(stdout, "%s", szBuffer); 
			fflush(stdout); 
		}
	} 

	CloseHandle(hEvent);

	InterlockedDecrement(&pLP->workerThreads);

	return 0;
}

// Listens the remote stderr pipe
// Remote process will send its stderr to this pipe
UINT WINAPI ListenRemoteErrorPipeThread(void* p)
{
	//giThreadsWorking already incremented
	ListenParam* pLP = (ListenParam*)p;

	char szBuffer[SIZEOF_BUFFER];
	DWORD dwRead;

	HANDLE hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);

	while(false == gbStop)
	{ 
		OVERLAPPED olR = {0};
		olR.hEvent = hEvent;
		if(!ReadFile( pLP->pSettings->hStdErr, szBuffer, SIZEOF_BUFFER - 1, &dwRead, &olR) || (dwRead == 0)) 
		{
			DWORD dwErr = GetLastError();
			if ( dwErr == ERROR_NO_DATA)
				break;
		}

		if(gbStop)
			break;

		HANDLE waits[2];
		waits[0] = pLP->hStop;
		waits[1] = olR.hEvent;
		DWORD ret = WaitForMultipleObjects(2, waits, FALSE, INFINITE);
		if(ret == WAIT_OBJECT_0)
			break; //need to exit
		_ASSERT(ret == WAIT_OBJECT_0 + 1); //data in buffer now

		szBuffer[ dwRead / sizeof(szBuffer[0]) ] = _T('\0');

		// Write it to our stderr
		fprintf(stderr, "%s", szBuffer); 
		fflush(stderr);
	} 
	
	CloseHandle(hEvent);

	InterlockedDecrement(&pLP->workerThreads);

	return 0;
}

// Listens our console, and if the user types in something,
// we will pass it to the remote machine.
// ReadConsole return after pressing the ENTER
UINT WINAPI ListenRemoteStdInputPipeThread(void* p)
{
	//giThreadsWorking already incremented
	ListenParam* pLP = (ListenParam*)p;

	HANDLE hInput = GetStdHandle(STD_INPUT_HANDLE);
	char szInputBuffer[SIZEOF_BUFFER] = {0};
	DWORD nBytesRead = 0;
	DWORD nBytesWrote = 0;

	HANDLE hWritePipe = CreateEvent(NULL, TRUE, FALSE, NULL);
	HANDLE hReadEvt = CreateEvent(NULL, TRUE, FALSE, NULL);

	DWORD oldMode = 0;
	GetConsoleMode(hInput, &oldMode);
	//DWORD newMode = oldMode & ~(ENABLE_LINE_INPUT | ENABLE_ECHO_INPUT);
	//SetConsoleMode(hInput, newMode);

	bool bWaitForKeyPress = true;
	//detect if input redirected from file (in which case we don't want to wait for keyboard hits)
	
	//DWORD inputSize = GetFileSize(hInput, NULL);
	//if(INVALID_FILE_SIZE != inputSize)
	//	bWaitForKeyPress = false;
	DWORD fileType = GetFileType(hInput);
	if (FILE_TYPE_CHAR != fileType) 
		bWaitForKeyPress = false;

	while(false == gbStop)
	{
		if(bWaitForKeyPress)
		{
			bool bBail = false;
			while(0 == _kbhit())
			{
				if(WAIT_OBJECT_0 == WaitForSingleObject(pLP->hStop, 0))
				{
					bBail = true;
					break;
				}
				Sleep(100);
			}
			if(bBail)
				break;
		}

		nBytesRead = 0;

		if (FILE_TYPE_PIPE == fileType) 
		{
			OVERLAPPED olR = { 0 };
			olR.hEvent = hReadEvt;
			if (!ReadFile(hInput, szInputBuffer, SIZEOF_BUFFER - 1, &nBytesRead, &olR) || (nBytesRead == 0))
			{
				DWORD dwErr = GetLastError();
				if (dwErr == ERROR_NO_DATA)
					break;
			}

			if (gbStop)
				break;

			HANDLE waits[2];
			waits[0] = pLP->hStop;
			waits[1] = olR.hEvent;
			DWORD ret = WaitForMultipleObjects(2, waits, FALSE, INFINITE);
			if (ret == WAIT_OBJECT_0)
				break; //need to exit
			_ASSERT(ret == WAIT_OBJECT_0 + 1); //data in buffer now
			GetOverlappedResult(hInput, &olR, &nBytesRead, FALSE);
		}
		else 
		{
			//if ( !ReadConsole( hInput, szInputBuffer, SIZEOF_BUFFER, &nBytesRead, NULL ) ) -- returns UNICODE which is not what we want
			if (!ReadFile(hInput, szInputBuffer, SIZEOF_BUFFER - 1, &nBytesRead, NULL))
			{
				DWORD dwErr = GetLastError();
				if (dwErr == ERROR_NO_DATA)
					break;
			}
		}

		if(gbStop)
			break;

		if(bWaitForKeyPress)
		{
			//suppress the input from being printed in the output since it was already shown locally
			EnterCriticalSection(&pLP->cs);
			szInputBuffer[nBytesRead] = '\0';
			pLP->inputSentToSuppressInOutput.push_back(szInputBuffer);
			LeaveCriticalSection(&pLP->cs);
		}

		// Send it to remote process' stdin
		OVERLAPPED olW = {0};
		olW.hEvent = hWritePipe;

		if (!WriteFile( pLP->pSettings->hStdIn, szInputBuffer, nBytesRead, &nBytesWrote, &olW))
		{
			DWORD gle = GetLastError();
			break;
		}

		if(gbStop)
			break;
		 
		HANDLE waits[2];
		waits[0] = pLP->hStop;
		waits[1] = olW.hEvent;
		DWORD ret = WaitForMultipleObjects(2, waits, FALSE, INFINITE);
		if(ret == WAIT_OBJECT_0)
			break; //need to exit
		_ASSERT(ret == WAIT_OBJECT_0 + 1); //write finished

		FlushFileBuffers(pLP->pSettings->hStdIn);
	} 
	
	CloseHandle(hWritePipe);
	CloseHandle(hReadEvt);

	SetConsoleMode(hInput, oldMode);

	InterlockedDecrement(&pLP->workerThreads);

	return 0;
}


// Connects to the remote processes stdout, stderr and stdin named pipes
BOOL ConnectToRemotePipes(ListenParam* pLP, DWORD dwRetryCount, DWORD dwRetryTimeOut)
{
	SECURITY_DESCRIPTOR SecDesc;
	InitializeSecurityDescriptor(&SecDesc, SECURITY_DESCRIPTOR_REVISION);
	SetSecurityDescriptorDacl(&SecDesc, TRUE, NULL, FALSE);

	SECURITY_ATTRIBUTES SecAttrib = {0};
	SecAttrib.nLength = sizeof(SECURITY_ATTRIBUTES);
	SecAttrib.lpSecurityDescriptor = &SecDesc;;
	SecAttrib.bInheritHandle = TRUE;

	CString remoteOutPipeName = StrFormat(L"\\\\%s\\pipe\\PAExecOut%s%u", pLP->remoteServer, pLP->machineName, pLP->pid);
	CString remoteErrPipeName = StrFormat(L"\\\\%s\\pipe\\PAExecErr%s%u", pLP->remoteServer, pLP->machineName, pLP->pid);
	CString remoteInPipeName = StrFormat(L"\\\\%s\\pipe\\PAExecIn%s%u", pLP->remoteServer, pLP->machineName, pLP->pid);

	DWORD gle = 0;

	while((false == gbStop) && (dwRetryCount--))
	{
		// Connects to StdOut pipe
		if ( BAD_HANDLE(pLP->pSettings->hStdOut) )
			if (WaitNamedPipe(remoteOutPipeName, NULL))
			{
				pLP->pSettings->hStdOut = CreateFile( 
										remoteOutPipeName,
										GENERIC_READ, 
										0,
										&SecAttrib, 
										OPEN_EXISTING, 
										FILE_ATTRIBUTE_NORMAL, 
										NULL );
				gle = GetLastError();
			}


		// Connects to Error pipe
		if(BAD_HANDLE(pLP->pSettings->hStdErr) )
			if ( WaitNamedPipe(remoteErrPipeName, NULL ) )
				pLP->pSettings->hStdErr = CreateFile( 
										remoteErrPipeName,
										GENERIC_READ, 
										0,
										&SecAttrib, 
										OPEN_EXISTING, 
										FILE_ATTRIBUTE_NORMAL, 
										NULL );

		// Connects to StdIn pipe
		if(BAD_HANDLE(pLP->pSettings->hStdIn))
			if(WaitNamedPipe(remoteInPipeName, NULL ) )
				pLP->pSettings->hStdIn = CreateFile( 
										remoteInPipeName,
										GENERIC_WRITE, 
										0,
										&SecAttrib, 
										OPEN_EXISTING, 
										FILE_ATTRIBUTE_NORMAL, 
										NULL );

		if(	!BAD_HANDLE(pLP->pSettings->hStdErr) &&
			!BAD_HANDLE(pLP->pSettings->hStdIn) &&
			!BAD_HANDLE(pLP->pSettings->hStdOut) )
			break;

		// One of the pipes failed, try it again
		Sleep(dwRetryTimeOut);
	}

	if (BAD_HANDLE(pLP->pSettings->hStdErr) ||
		BAD_HANDLE(pLP->pSettings->hStdIn) ||
		BAD_HANDLE(pLP->pSettings->hStdOut) || 
		(WAIT_OBJECT_0 == WaitForSingleObject(pLP->hStop, 0)) ||
		gbStop)
	{
		if((gbStop) || (WAIT_OBJECT_0 == WaitForSingleObject(pLP->hStop, 0)))
			return false; 

		CloseHandle(pLP->pSettings->hStdErr);
		CloseHandle(pLP->pSettings->hStdIn);
		CloseHandle(pLP->pSettings->hStdOut);
		pLP->pSettings->hStdErr = NULL;
		pLP->pSettings->hStdIn = NULL;
		pLP->pSettings->hStdOut = NULL;
		Log(L"Failed to open remote pipes", true);
		return false;
	}

	//DWORD mode = 0;
	//if(gPIPE_TYPE & PIPE_TYPE_MESSAGE)
	//	mode |= PIPE_READMODE_MESSAGE;

	//BOOL b = SetNamedPipeHandleState(settings.hStdOut, &mode, NULL, NULL);
	//DWORD gle = GetLastError();
	//b &= SetNamedPipeHandleState(settings.hStdErr, &mode, NULL, NULL);
	//gle = GetLastError();
	//b &= SetNamedPipeHandleState(settings.hStdIn, &mode, NULL, NULL);
	//gle = GetLastError();
	//_ASSERT(b);

	UINT ignored;

	// StdOut
	InterlockedIncrement(&pLP->workerThreads);
	HANDLE h = (HANDLE)_beginthreadex(NULL, 0, ListenRemoteOutPipeThread, pLP, 0, &ignored);
	CloseHandle(h);

	// StdErr
	InterlockedIncrement(&pLP->workerThreads);
	h = (HANDLE)_beginthreadex(NULL, 0, ListenRemoteErrorPipeThread, pLP, 0, &ignored);
	CloseHandle(h);

	// StdIn
	InterlockedIncrement(&pLP->workerThreads);
	h = (HANDLE)_beginthreadex(NULL, 0, ListenRemoteStdInputPipeThread, pLP, 0, &ignored);
	CloseHandle(h);

#ifdef _DEBUG
	Log(L"DEBUG: Connected to remote pipes", false);
#endif

	return true;
}
