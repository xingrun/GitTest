
// VCForWord.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CVCForWordApp:
// �йش����ʵ�֣������ VCForWord.cpp
//

class CVCForWordApp : public CWinApp
{
public:
	CVCForWordApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CVCForWordApp theApp;