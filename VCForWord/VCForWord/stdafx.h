
// stdafx.h : ��׼ϵͳ�����ļ��İ����ļ���
// ���Ǿ���ʹ�õ��������ĵ�
// �ض�����Ŀ�İ����ļ�

#pragma once

#ifndef _SECURE_ATL
#define _SECURE_ATL 1
#endif

#ifndef VC_EXTRALEAN
#define VC_EXTRALEAN            // �� Windows ͷ���ų�����ʹ�õ�����
#endif

#include "targetver.h"

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // ĳЩ CString ���캯��������ʽ��

// �ر� MFC ��ĳЩ�����������ɷ��ĺ��Եľ�����Ϣ������
#define _AFX_ALL_WARNINGS

#include <afxwin.h>         // MFC ��������ͱ�׼���
#include <afxext.h>         // MFC ��չ


#include <afxdisp.h>        // MFC �Զ�����



#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdtctl.h>           // MFC �� Internet Explorer 4 �����ؼ���֧��
#endif
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>             // MFC �� Windows �����ؼ���֧��
#endif // _AFX_NO_AFXCMN_SUPPORT

#include <afxcontrolbars.h>     // �������Ϳؼ����� MFC ֧��



#include "Word2010/CApplication.h"	// word�������
#include "Word2010/CDocuments.h"	// �ĵ�������
#include "Word2010/CDocument0.h"	// ����docx����
#include "Word2010/CSelection.h"	// �ö�������ڻ򴰸��еĵ�ǰ��ѡ���ݣ�ʹ�������
#include "Word2010/CCell.h"	// ���Ԫ��
#include "Word2010/CCells.h"	// ���Ԫ�񼯺�
#include "Word2010/CRange.h"	// �ö�������ĵ��е�һ��������Χ
#include "Word2010/CTable0.h"	// �������
#include "Word2010/CTables0.h"	// ��񼯺�
#include "Word2010/CFont0.h"	// ����
#include "Word2010/CParagraphs.h"	// ���伯��
#include "Word2010/CParagraphFormat.h"	// ������ʽ
#include "Word2010/CParagraph.h"	// ��������
#include "Word2010/CnlineShape.h"	// Inlineͼ�ζ��󼯺�
#include "Word2010/CnlineShapes.h"	// ����Inlineͼ�����
#include "Word2010/CRow.h"	// ������
#include "Word2010/CRows.h"	// �м���
#include "Word2010/CFields.h"
#include "Word2010/CPane0.h"	// ҳü������
#include "Word2010/CWindow0.h"// ҳü������
#include "Word2010/CView0.h"// ҳü������
#include "Word2010/CPageSetup.h"	// ҳ������

//#ifdef _UNICODE
#if defined _M_IX86
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
//#endif


