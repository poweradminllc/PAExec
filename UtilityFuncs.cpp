// UtilityFuncs.cpp: Helper functions
//
// Copyright (c) Power Admin LLC, 2012 - 2013
//
// This code is provided "as is", with absolutely no warranty expressed
// or implied. Any use is at your own risk.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "PAExec.h"

extern bool gbODS;
extern CString gLogPath;

static CString gLastLog;

CString LastLog()
{
	return gLastLog;
}

CString StrFormat(LPCTSTR pFormat, ...) //returns a formatted CString.  Inspired by .NET's String.Format
{
	CString result;
	DWORD size = max(4096, (DWORD)_tcslen(pFormat) + 4096);

	_ASSERT(NULL == wcsstr(pFormat, L"{0}")); //StrFormat2 should have been used?? GLOK

	try
	{
		while(true)
		{
			va_list pArg;
			va_start(pArg, pFormat);
			int res = _vsntprintf_s(result.GetBuffer(size), size, _TRUNCATE, pFormat, pArg);
			va_end(pArg);
			if(res >= 0)
			{
				result.ReleaseBuffer();
				return result;
			}
			else
			{
				result.ReleaseBuffer(1);
				size += 8192;
				if(size > (12 * 1024 * 1024))
				{
					_ASSERT(0);
					Log(StrFormat(L"STRING TOO LONG: %s", pFormat), true);
					CString s = L"<error - string too long -- "; //GLOK
					s += pFormat;
					s += L">"; //GLOK
					return s;
				}
			}
		}
	}
	catch(...)
	{
		_ASSERT(0);
		CString res = L"<string format exception: ["; //GLOK
		if(NULL != pFormat)
			res += pFormat;
		else
			res += L"(null)"; //GLOK
		res += L"]>"; //GLOK
		res.Replace(L'%', L'{'); //so no further formatting is attempted GLOK
		return res;
	}
}

void Log(LPCWSTR str, DWORD lastError)
{
	CString s = str;
	s += L" ";
	s += GetSystemErrorMessage(lastError);
	Log(s, 0 != lastError);
}

void Log(LPCWSTR str, bool bForceODS)
{
	CString s = str;
	s += L"\r\n";

	HANDLE oh = GetStdHandle( STD_OUTPUT_HANDLE );
	DWORD ignored;
	WriteConsole(oh, (LPCWSTR)s, wcslen(s), &ignored, NULL);

	if(gbODS || bForceODS)
		OutputDebugString(s);

	if(FALSE == gLogPath.IsEmpty())
	{
		HANDLE hf = CreateFile(gLogPath, GENERIC_WRITE, 0, 0, OPEN_ALWAYS, 0, 0);
		if(!BAD_HANDLE(hf))
		{
			DWORD ignored;
			ULARGE_INTEGER ui = {0};
			ui.LowPart = GetFileSize(hf, &ui.HighPart);
			if(0 == ui.QuadPart)
				WriteFile(hf, &UTF8_BOM, sizeof(UTF8_BOM), &ignored, 0);  
			SetFilePointer(hf, 0, 0, FILE_END);
			
			//now convert to UTF-8
			DWORD len = WideCharToMultiByte(CP_UTF8, 0, s, wcslen(s), NULL, 0, NULL, NULL);
			if(0 != len)
			{
				char* pBuff = new char[len + 5];
				WideCharToMultiByte(CP_UTF8, 0, s, wcslen(s), pBuff, len, NULL, NULL);
				pBuff[len] = '\0';
				WriteFile(hf, pBuff, len, &ignored, 0);
				delete [] pBuff;
			}
			FlushFileBuffers(hf);
			CloseHandle(hf);
		}
	}
	gLastLog = s;
}

DWORD wtodw(const wchar_t* num)
{
	if(NULL == num)
		return 0;
	//in VS2003 atoi/wtoi would convert an unsigned value even though it would overflow a signed value.  No more, and it causes trouble!
	__int64 res = _wtoi64(num);
	_ASSERT(res <= ULONG_MAX);
	return (DWORD)res;
}

CString GetSystemErrorMessage(DWORD lastErrorVal)
{
	CString osErr = _T("Unknown error value. ");
	const int errSize = 8192;
	HMODULE hMod = NULL;
	DWORD flags = FORMAT_MESSAGE_FROM_SYSTEM;
	DWORD len = FormatMessage(flags, hMod, lastErrorVal, 0, osErr.GetBuffer(errSize), errSize - 50, NULL);
	osErr.ReleaseBuffer();
	osErr.Replace(L"\r", L"");
	osErr.Replace(L"\n", L"");
	osErr += StrFormat(L" [Err=0x%0X, %u]", lastErrorVal, lastErrorVal);

	return osErr;
}


bool EnablePrivilege(LPCWSTR privilegeStr, HANDLE hToken /* = NULL */)
{
	TOKEN_PRIVILEGES  tp;         // token privileges
	LUID              luid;
	bool				bCloseToken = false;

	if(NULL == hToken)
	{
		if(!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken))
		{
			Log(StrFormat(L"Failed to open process to enable privilege %s", privilegeStr), false);
			return false;
		}
		bCloseToken = true;
	}

	if (!LookupPrivilegeValue(NULL, privilegeStr, &luid))
	{
		if(bCloseToken)
			CloseHandle (hToken);
		_ASSERT(0);
		Log(StrFormat(L"Could not find privilege %s", privilegeStr), false);
		return false;
	}

	ZeroMemory (&tp, sizeof (tp));
	tp.PrivilegeCount = 1;
	tp.Privileges[0].Luid = luid;
	tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;

	// Adjust Token privileges
	if (!AdjustTokenPrivileges (hToken, FALSE, &tp, sizeof(TOKEN_PRIVILEGES), NULL, NULL))
	{
		DWORD gle = GetLastError();
		Log(StrFormat(L"Failed to adjust token for privilege %s", privilegeStr), gle);
		if(bCloseToken)
			CloseHandle (hToken);
		_ASSERT(0);
		return false;
	}
	if(bCloseToken)
		CloseHandle (hToken);

	return true;
}



RemMsg& RemMsg::operator<<(LPCWSTR s)
{
	*this << (DWORD)wcslen(s);
	m_payload.insert(m_payload.end(), (BYTE*)s, (BYTE*)(s + wcslen(s)));
	return *this;
}

RemMsg& RemMsg::operator<<(bool b)
{
	m_payload.insert(m_payload.end(), (BYTE*)&b, ((BYTE*)&b) + sizeof(b));
	return *this;
}

RemMsg& RemMsg::operator<<(DWORD d)
{
	m_payload.insert(m_payload.end(), (BYTE*)&d, ((BYTE*)&d) + sizeof(d));
	return *this;
}

RemMsg& RemMsg::operator<<(__int64 d)
{
	m_payload.insert(m_payload.end(), (BYTE*)&d, ((BYTE*)&d) + sizeof(d));
	return *this;
}

RemMsg& RemMsg::operator<<(FILETIME d)
{
	m_payload.insert(m_payload.end(), (BYTE*)&d, ((BYTE*)&d) + sizeof(d));
	return *this;
}

RemMsg& RemMsg::operator>>(CString& s)
{
	if(m_bResetReadItr)
	{
		m_readItr = m_payload.begin();
		m_bResetReadItr = false;
	}

	s.Empty();
	DWORD len = 0;
	*this >> len; //this is len in wchar_t
	if(0 != len)
	{
		if(m_readItr >= m_payload.end())
			return *this;

		LPWSTR pStart = s.GetBuffer(len);
		memcpy(pStart, &(*m_readItr), len * sizeof(wchar_t));
		s.ReleaseBuffer(len);
		m_readItr += len * sizeof(wchar_t);
	}

	return *this;
}

RemMsg& RemMsg::operator>>(bool& b)
{
	if(m_bResetReadItr)
	{
		m_readItr = m_payload.begin();
		m_bResetReadItr = false;
	}

	if(m_readItr >= m_payload.end())
		return *this;

	memcpy(&b, &(*m_readItr), sizeof(b));
	m_readItr += sizeof(b);

	return *this;

}

RemMsg& RemMsg::operator>>(DWORD& d)
{
	if(m_bResetReadItr)
	{
		m_readItr = m_payload.begin();
		m_bResetReadItr = false;
	}

	if(m_readItr >= m_payload.end())
		return *this;

	memcpy(&d, &(*m_readItr), sizeof(d));
	m_readItr += sizeof(d);

	return *this;
}

RemMsg& RemMsg::operator>>(__int64& d)
{
	if(m_bResetReadItr)
	{
		m_readItr = m_payload.begin();
		m_bResetReadItr = false;
	}

	if(m_readItr >= m_payload.end())
		return *this;

	memcpy(&d, &(*m_readItr), sizeof(d));
	m_readItr += sizeof(d);

	return *this;
}


RemMsg& RemMsg::operator>>(FILETIME& d)
{
	if(m_bResetReadItr)
	{
		m_readItr = m_payload.begin();
		m_bResetReadItr = false;
	}

	if(m_readItr >= m_payload.end())
		return *this;

	memcpy(&d, &(*m_readItr), sizeof(d));
	m_readItr += sizeof(d);

	return *this;
}



typedef BOOL (WINAPI *Wow64DisableWow64FsRedirectionProc)(PVOID* OldValue);
typedef BOOL (WINAPI *Wow64RevertWow64FsRedirectionProc)(PVOID OldValue);

static HMODULE hKernel = NULL;
static Wow64DisableWow64FsRedirectionProc pDis = NULL;
static Wow64RevertWow64FsRedirectionProc pRev = NULL;
PVOID OldWow64RedirVal = NULL;

void DisableFileRedirection()
{
	if(NULL == hKernel)
		hKernel = LoadLibrary(_T("Kernel32.dll"));

	if((NULL != hKernel) && ((NULL == pDis) || (NULL == pRev)))
	{
		pDis = (Wow64DisableWow64FsRedirectionProc)GetProcAddress(hKernel, "Wow64DisableWow64FsRedirection");
		pRev = (Wow64RevertWow64FsRedirectionProc)GetProcAddress(hKernel, "Wow64RevertWow64FsRedirection");
	}

	if(NULL != pDis)
	{
		BOOL b = pDis(&OldWow64RedirVal);
		if(b)
			Log(L"Disabled WOW64 file system redirection", false);
		else
			Log(L"Failed to disable WOW64 file system redirection", GetLastError());
	}
	else
		Log(L"Failed to find Wow64DisableWow64FsRedirection API", true);
}

void RevertFileRedirection()
{
	if(NULL == hKernel)
		hKernel = LoadLibrary(_T("Kernel32.dll"));

	if((NULL != hKernel) && ((NULL == pDis) || (NULL == pRev)))
	{
		pDis = (Wow64DisableWow64FsRedirectionProc)GetProcAddress(hKernel, "Wow64DisableWow64FsRedirection");
		pRev = (Wow64RevertWow64FsRedirectionProc)GetProcAddress(hKernel, "Wow64RevertWow64FsRedirection");
	}

	if(NULL != pRev)
		pRev(OldWow64RedirVal);
}

CString ExpandToFullPath(LPCWSTR path)
{
	CString result = path;

	wchar_t found[_MAX_PATH * 2] = {0};
	LPWSTR ignored;
	if(0 != SearchPath(NULL, path, NULL, sizeof(found)/sizeof(wchar_t), found, &ignored))
		result = found;

	return result;
}


bool ReadTextFile(LPCWSTR fileName, CString& content)
{
	content.Empty();

	HANDLE hf = CreateFile(fileName, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, 0, 0);
	DWORD gle = GetLastError();
	if(!BAD_HANDLE(hf))
	{
		//read in UTF-8
		DWORD ignored;
		DWORD size = GetFileSize(hf, &ignored);
		BYTE* pOrig = NULL;
		BYTE* pBuff = pOrig = new BYTE[size + 1];
		DWORD read = 0;
		ReadFile(hf, pBuff, size, &read, 0);
		pBuff[read] = '\0';
		if(0 == memcmp(pBuff, UTF8_BOM, sizeof(UTF8_BOM)))
		{
			pBuff += sizeof(UTF8_BOM);
			read -=  sizeof(UTF8_BOM);
		}

		//now convert to unicode
		DWORD len = MultiByteToWideChar(CP_UTF8, 0, (LPCSTR)pBuff, read, NULL, 0);
		if(0 == len)
		{
			gle = GetLastError();
			_ASSERT(0);
			Log(StrFormat(L"Failed to translate UTF-8 file %s into Unicode.", fileName), gle);
			return false;
		}
		else
		{
			LPWSTR output = content.GetBuffer(len + 1);
			len = MultiByteToWideChar(CP_UTF8, 0, (LPCSTR)pBuff, read, output, len);
			content.ReleaseBuffer(len);
		}

		delete [] pOrig;
		pOrig = NULL;
		pBuff = NULL;
		CloseHandle(hf);
		
		return true;
	}
	else
	{
		Log(StrFormat(L"Failed to open text file %s.", fileName), gle);
		return false;
	}
}


void SplitUserNameAndDomain(CString user, CString& username, CString& domain)
{
	//defaults
	username = user;
	domain = L"";

	int idx = user.Find(L"\\");
	if (idx != -1)
	{
		domain = user.Left(idx);
		username = user.Right(user.GetLength() - (idx + 1));
	}
	else
	{
		idx = user.Find(L"@");
		if (idx != -1)
		{
			username = user.Left(idx);
			domain = user.Right(user.GetLength() - (idx + 1));
		}
	}
}

// Strips the domain name from an user name string, if given as domain\user or user@domain
CString RemoveDomainFromUserName(CString user)
{
	CString ret = user;
	int idx = user.Find(L"\\");
	if (idx != -1)
	{
		ret.Delete(0, idx + 1);
	}
	else
	{
		idx = user.Find(L"@");
		if (idx != -1)
			ret = ret.Left(idx);
	}
	return ret;
}

#ifdef _DEBUG

bool ReportAssert(LPCWSTR file, int line, LPCWSTR expr)
{
	CString msg = L"ASSERT: ";
	if(NULL != expr)
		msg += expr;
	msg += StrFormat(L" (%s, %d)", file, line);
	msg.Replace(L"\r", L"");
	msg.Replace(L"\n", L"");
	msg += L"\r\n";

	OutputDebugString(msg);

	if(gbInService)
		return true; //continue

	return 1 != _CrtDbgReportW(_CRT_ASSERT, _CRT_WIDE(__FILE__), __LINE__, NULL, expr);
}

#endif
