#pragma once

#include "resource.h"

#ifdef _DEBUG

#undef _ASSERT
#define _ASSERT(expr) \
	(void) ((!!(expr)) || \
	(ReportAssert(_CRT_WIDE(__FILE__), __LINE__, _CRT_WIDE(#expr))) || \
	(_CrtDbgBreak(), 0))

extern bool ReportAssert(LPCWSTR file, int line, LPCWSTR expr);

#endif
