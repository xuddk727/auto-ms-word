// WordExpert.cpp : ���� DLL Ӧ�ó���ĵ���������
//

#include "stdafx.h"
#include "../inc/Word.h"

extern "C" _declspec(dllexport) CWordBase* CreateWord(WORDEMODE w)
{
	CWordBase* word =NULL;
	switch (w)
	{
	case MODE_BASE:
	case MODE_FORMAT:
		word = new CWordFormat();
		break;
	default:break;
	}
	
	return word;
}






extern "C" _declspec(dllexport) void ReleaseResource(void* p)
{
	delete p;
}

