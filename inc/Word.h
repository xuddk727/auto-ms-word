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
	virtual bool CreateWord(BOOL bShow = TRUE);//创建word，必须先执行
	void SetFont(int nBold,float fSize,DWORD dwColor,LPCTSTR lpName);
	//设置字体 nBold 是否加粗, fSize 字体大小 dwColor 文字RGB颜色 lpName 字体名
	virtual bool CreateText(CString strText);//创建默认样式的文本
	virtual CString GetLastError();
	virtual bool SaveWord(CString strFilePath);//保存word
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

	
	virtual bool CreateTitle(CString strTitle,int nCountID);//插入大标题

	virtual bool CreateSection(CString strSection,int nLevel = 1);//顺序插入章节号以及章节标题
	//nLevel 章节目录级别
	/*
		第一级目录输出形式     1.
		第二级目录输出形式     1.1
		第三级目录输出形式     1.1.1
		编号为顺序编号,如:
		1. aaa
		1.1 bbbb
		2. ccc
		2.1 dddd


	*/

	virtual bool CreateText(CString strText);//创建正文样式的正文
	
	virtual bool CreatePicture(CString strDescrption,CString strImgPath);//插入图片
	virtual bool CreateTable(CString strDescrption,int nRow,int nColume,
		vector<CString>& vString);//插入表格
};
  








extern "C" _declspec(dllexport) CWordBase* CreateWord(WORDEMODE w);
extern "C" _declspec(dllexport) void ReleaseResource(void* p);
