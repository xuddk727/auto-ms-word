
// WordTestExe.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CWordTestExeApp:
// �йش����ʵ�֣������ WordTestExe.cpp
//

class CWordTestExeApp : public CWinAppEx
{
public:
	CWordTestExeApp();

// ��д
	public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CWordTestExeApp theApp;