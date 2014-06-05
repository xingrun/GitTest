
// stdafx.h : 标准系统包含文件的包含文件，
// 或是经常使用但不常更改的
// 特定于项目的包含文件

#pragma once

#ifndef _SECURE_ATL
#define _SECURE_ATL 1
#endif

#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN            // 从 Windows 头中排除极少使用的资料
#endif

#include "targetver.h"

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // 某些 CString 构造函数将是显式的

// 关闭 MFC 对某些常见但经常可放心忽略的警告消息的隐藏
#define _AFX_ALL_WARNINGS

#include <afxwin.h>         // MFC 核心组件和标准组件
#include <afxext.h>         // MFC 扩展


#include <afxdisp.h>        // MFC 自动化类



#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdtctl.h>           // MFC 对 Internet Explorer 4 公共控件的支持
#endif
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>             // MFC 对 Windows 公共控件的支持
#endif // _AFX_NO_AFXCMN_SUPPORT

#include <afxcontrolbars.h>     // 功能区和控件条的 MFC 支持



#include "Word2010/CApplication.h"	// word程序对象
#include "Word2010/CDocuments.h"	// 文档集对象
#include "Word2010/CDocument0.h"	// 单个docx对象
#include "Word2010/CSelection.h"	// 该对象代表窗口或窗格中的当前所选内容，使用率最高
#include "Word2010/CCell.h"	// 表格单元格
#include "Word2010/CCells.h"	// 表格单元格集合
#include "Word2010/CRange.h"	// 该对象代表文档中的一个连续范围
#include "Word2010/CTable0.h"	// 单个表格
#include "Word2010/CTables0.h"	// 表格集合
#include "Word2010/CFont0.h"	// 字体
#include "Word2010/CParagraphs.h"	// 段落集合
#include "Word2010/CParagraphFormat.h"	// 段落样式
#include "Word2010/CParagraph.h"	// 单个段落
#include "Word2010/CnlineShape.h"	// Inline图形对象集合
#include "Word2010/CnlineShapes.h"	// 单个Inline图像对象
#include "Word2010/CRow.h"	// 单个行
#include "Word2010/CRows.h"	// 行集合
#include "Word2010/CFields.h"
#include "Word2010/CPane0.h"	// 页眉等设置
#include "Word2010/CWindow0.h"// 页眉等设置
#include "Word2010/CView0.h"// 页眉等设置
#include "Word2010/CPageSetup.h"	// 页面设置

//#ifdef _UNICODE
#if defined _M_IX86
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
//#endif


