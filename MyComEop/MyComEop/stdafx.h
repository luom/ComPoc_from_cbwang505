// stdafx.h : 标准系统包含文件的包含文件，
// 或是经常使用但不常更改的
// 特定于项目的包含文件
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
