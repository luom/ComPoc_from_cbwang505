// stdafx.h : ��׼ϵͳ�����ļ��İ����ļ���
// ���Ǿ���ʹ�õ��������ĵ�
// �ض�����Ŀ�İ����ļ�
//

#ifdef ARRAYSIZE
#define countof(a) (ARRAYSIZE(a)) // (sizeof((a))/(sizeof(*(a))))
#else
#define countof(a) (sizeof((a)) / (sizeof(*(a))))
#endif


#pragma once

#include "targetver.h"
#include <comdef.h>
#include <ComSvcs.h>
