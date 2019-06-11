#pragma once
#include "msword.h"
#include <vector>
using namespace std;
typedef enum WordExpertMode
{
	MODE_BASE = 0,
	MODE_FORMAT,
	MODE_TEMPLATE
}WORDEMODE;


class CWordBase
{
public:
	CWordBase(void);
	virtual ~CWordBase(void);
	virtual bool CreateWord(BOOL bShow = TRUE);//����word��������ִ��
	void SetFont(int nBold,float fSize,DWORD dwColor,LPCTSTR lpName);
	//�������� nBold �Ƿ�Ӵ�, fSize �����С dwColor ����RGB��ɫ lpName ������
	virtual bool CreateText(CString strText);//����Ĭ����ʽ���ı�
	virtual CString GetLastError();
	virtual bool SaveWord(CString strFilePath);//����word
protected:
	_Application m_pApplication;
	Selection m_pSelection;
	_Document m_pDocument;
	bool m_bInit;
	CString m_strLastError;
	LPDISPATCH GetStyle(CComVariant vt);
	void CreateStyle(int nLevel);
};



class CWordFormat :public CWordBase
{
public:

	
	virtual bool CreateTitle(CString strTitle,int nCountID);//��������

	virtual bool CreateSection(CString strSection,int nLevel = 1);//˳������½ں��Լ��½ڱ���
	//nLevel �½�Ŀ¼����
	/*
		��һ��Ŀ¼�����ʽ     1.
		�ڶ���Ŀ¼�����ʽ     1.1
		������Ŀ¼�����ʽ     1.1.1
		���Ϊ˳����,��:
		1. aaa
		1.1 bbbb
		2. ccc
		2.1 dddd


	*/

	virtual bool CreateText(CString strText);//����������ʽ������
	
	virtual bool CreatePicture(CString strDescrption,CString strImgPath);//����ͼƬ
	virtual bool CreateTable(CString strDescrption,int nRow,int nColume,
		vector<CString>& vString);//������
};
  








extern "C" _declspec(dllexport) CWordBase* CreateWord(WORDEMODE w);
extern "C" _declspec(dllexport) void ReleaseResource(void* p);
