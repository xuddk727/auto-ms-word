#include "StdAfx.h"
#include "../inc/Word.h"

//////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////
//CWordBase

CWordBase::CWordBase(void)
{
	CoInitialize(NULL);
	m_bInit = false;
	m_strLastError.Format(TEXT("未执行CreateWord初始化word!"));
}

CWordBase::~CWordBase(void)
{
	CoUninitialize();
}

void CWordBase::CreateStyle(int nLevel)
{
	ListGalleries lists = m_pApplication.GetListGalleries();
	ListGallery listex = lists.Item(3);
	ListTemplates temps = listex.GetListTemplates();
	COleVariant  tempID((short)1);
	ListTemplate temp = temps.Item(&tempID);
	ListLevels levels = temp.GetListLevels();
	ListLevel level1 = levels.Item((long)nLevel);
	switch(nLevel)
	{
	case 1:
		{
			level1.SetTrailingCharacter(0);
			level1.SetNumberFormat(TEXT("%1"));
			level1.SetNumberStyle(0);
			level1.SetLinkedStyle(TEXT("标题 1"));
		}
		break;
	case 2:
		{
			level1.SetTrailingCharacter(0);
			level1.SetNumberFormat(TEXT("%1.%2"));
			level1.SetNumberStyle(0);
			level1.SetLinkedStyle(TEXT("标题 2"));
		}break;
	case 3:
		{
			level1.SetTrailingCharacter(0);
			level1.SetNumberFormat(TEXT("%1.%2.%3"));
			level1.SetNumberStyle(0);
			level1.SetLinkedStyle(TEXT("标题 3"));
		}break;
	case 4:
		{

			level1.SetTrailingCharacter(0);
			level1.SetNumberFormat(TEXT("%1.%2.%3.%4"));
			level1.SetNumberStyle(0);
			level1.SetStartAt(1);
			level1.SetLinkedStyle(TEXT("标题 4"));
		}break;
	default:
		break;
	}
	Range rages = m_pSelection.GetRange();
	ListFormat listf = rages.GetListFormat();
	COleVariant   ContinuePreviousLists(VARIANT_FALSE);
	COleVariant  wdListApplyTo((long)0),wdWord10List((long)2),wdApplyLevel((long)9);
	listf.ApplyListTemplateWithLevel(temp,&ContinuePreviousLists,
		&wdListApplyTo,&wdWord10List,&wdApplyLevel);
}


LPDISPATCH CWordBase::GetStyle(CComVariant vt)
{
	try
	{
		if (!m_bInit)
			return NULL;
		CStyles styles =  m_pDocument.GetStyles();
		return styles.Item(&vt);

	}
	catch (_com_error* e)
	{
		m_strLastError.Format(TEXT("%s"),e->Description());
		return NULL;
	}
	return NULL;

}

void CWordBase::SetFont(int nBold,float fSize,DWORD dwColor,LPCTSTR lpName)
{
	_Font font = m_pSelection.GetFont();
	font.SetBold(nBold);
	font.SetSize(fSize);
	font.SetName(lpName);
	font.SetColor(dwColor);
}

bool CWordBase::CreateText( CString strText )
{
	try
	{
		if (!m_bInit)
			return false;
		m_pSelection.TypeText(strText.GetBuffer());
		m_pSelection.TypeParagraph();
		return true;

	}
	catch (_com_error* e)
	{
		m_strLastError.Format(TEXT("%s"),e->Description());
		return false;
	}
	return false;
}

bool CWordBase::CreateWord( BOOL bShow)
{
	try
	{
		COleVariant vTrue((short)TRUE),vFalse((short)FALSE);
		m_pApplication.CreateDispatch(_T("Word.Application"));
		m_pApplication.SetVisible(bShow);
		m_pApplication.SetWidth(776);
		m_pApplication.SetHeight(560);
		m_pApplication.SetWindowState(1);
		Documents docs = m_pApplication.GetDocuments(); 
		CComVariant tpl(_T("")),Visble,DocType(0),NewTemplate(false);
		docs.Add(&tpl,&NewTemplate,&DocType,&Visble);
		m_pDocument = m_pApplication.GetActiveDocument();
		m_pSelection = m_pApplication.GetSelection();
		m_bInit = true;
		m_strLastError = TEXT("");

		return true;
	}
	catch (_com_error* e)
	{
		m_strLastError.Format(TEXT("%s"),e->Description());
		return false;
	}
	return false;
}

CString CWordBase::GetLastError()
{
	return m_strLastError;
}

bool CWordBase::SaveWord( CString strFilePath )
{
	try
	{
		m_strLastError.Format(TEXT("未执行CreateWord初始化word!"));
		if (!m_bInit)
			return false;
		COleVariant	covOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR);
		COleVariant  wdFormatDocument((short)0);
		m_pDocument.SaveAs(COleVariant(strFilePath.GetBuffer(),VT_BSTR),
			wdFormatDocument,
			covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,
			covOptional,covOptional,covOptional,covOptional,covOptional,covOptional,covOptional);
		m_pDocument.Save();
		m_pDocument.ReleaseDispatch();
		m_pSelection.ReleaseDispatch();
		m_pApplication.ReleaseDispatch();
		m_bInit = false;
		m_strLastError = TEXT("已释放接口，无法继续。\r\n");
		return true;
	}
	catch (COleDispatchException* e)
	{
		m_strLastError.Format(TEXT("%s"),e->m_strDescription);
		return false;
	}
	return false;
}

//////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////




bool CWordFormat::CreateTitle( CString strTitle,int nCountID )
{
	try
	{
		m_strLastError.Format(TEXT("未执行CreateWord初始化word!"));
		if (!m_bInit)
			return false;
		CString str;
		_ParagraphFormat pf;
		SetFont(1,40,RGB(255,0,0),_T("宋体"));
		pf= m_pSelection.GetParagraphFormat();
		pf.SetAlignment(1);
		pf.SetLineSpacingRule(0);
		m_pSelection.SetParagraphFormat(pf);
		m_pSelection.TypeText(strTitle.GetBuffer());
		m_pSelection.TypeParagraph();

		SetFont(1,16,RGB(0,0,0),_T("黑体"));
		pf.SetAlignment(1);
		pf.SetLineSpacingRule(1);
		m_pSelection.SetParagraphFormat(pf);
		SYSTEMTIME st;
		GetLocalTime(&st);
		str.Format(TEXT("(%4d年%d月期 总第%d期)"),st.wYear,st.wMonth,nCountID);
		m_pSelection.TypeText(str);
		m_pSelection.TypeParagraph();

		SetFont(1,16,RGB(0,0,0),_T("宋体"));

		pf.SetAlignment(1);
		pf.SetLineSpacingRule(1);
		m_pSelection.SetParagraphFormat(pf);
		str.Format(TEXT("北京航天慧景科技有限公司         %4d年%2d月%2d日"),
			st.wYear,st.wMonth,st.wDay);
		m_pSelection.TypeText(str);
		m_pSelection.TypeParagraph();
		str.Format(TEXT("__________________________________________"));
		m_pSelection.TypeText(str);
		m_pSelection.TypeParagraph();
		return true;

	}
	catch (_com_error* e)
	{
		m_strLastError.Format(TEXT("%s"),e->Description());
		return false;
	}
	return false;
}





bool CWordFormat::CreateSection( CString strSection,int nLevel )
{
	try
	{
		m_strLastError.Format(TEXT("未执行CreateWord初始化word!"));
		if (!m_bInit)
			return false;
		CString strFormat;
		strFormat.Format(TEXT("标题 %d"),nLevel);
		_ParagraphFormat pf;
		SetFont(1,20,RGB(0,0,0),_T("宋体"));

		pf= m_pSelection.GetParagraphFormat();
		pf.SetAlignment(0);
		pf.SetLineSpacingRule(1);
		m_pSelection.SetParagraphFormat(pf);
		CreateStyle(nLevel);
		CComVariant styles(_T("")),styleset(strFormat);
		styles = m_pSelection.GetStyle();
		CComVariant vSetPath(GetStyle(styleset));
		m_pSelection.SetStyle(&vSetPath);
		
		m_pSelection.TypeText(strSection.GetBuffer());
		m_pSelection.TypeParagraph();
		m_pSelection.SetStyle(&styles);
		
		return true;
	}
	catch (_com_error* e)
	{
		m_strLastError.Format(TEXT("%s"),e->Description());
		return false;
	}
	return false;
}

bool CWordFormat::CreateText( CString strText )
{
	try
	{
		m_strLastError.Format(TEXT("未执行CreateWord初始化word!"));
		if (!m_bInit)
			return false;
		CString str;
		_ParagraphFormat pf;
		CComVariant styleset(_T("正文"));
		CComVariant vSetPath(GetStyle(styleset));
		m_pSelection.SetStyle(&vSetPath);
		
		SetFont(0,12,RGB(0,0,0),_T("宋体"));

		pf= m_pSelection.GetParagraphFormat();
		pf.SetAlignment(0);
		pf.SetLineSpacingRule(1);
		m_pSelection.SetParagraphFormat(pf);
		m_pSelection.TypeText(strText.GetBuffer());
		m_pSelection.TypeParagraph();
		return true;

	}
	catch (_com_error* e)
	{
		m_strLastError.Format(TEXT("%s"),e->Description());
		return false;
	}
	return false;
}


bool CWordFormat::CreatePicture( CString strDescrption,CString strImgPath )
{
	try
	{
		m_strLastError.Format(TEXT("未执行CreateWord初始化word!"));
		if (!m_bInit)
			return false;
		CString str;
		_ParagraphFormat pf;
		CComVariant styleset(_T("正文"));
		CComVariant vSetPath(GetStyle(styleset));
		m_pSelection.SetStyle(&vSetPath);
		
		SetFont(0,12,RGB(0,0,0),_T("宋体"));
		//m_pSelection.SetFont(font);
		pf= m_pSelection.GetParagraphFormat();
		pf.SetAlignment(1);
		m_pSelection.SetParagraphFormat(pf);
		CComVariant tpl(_T("")),Visble,DocType(0),NewTemplate(false);
		COleVariant vTrue((short)TRUE),vFalse((short)FALSE);
		InlineShapes inlineshapes;
		CComVariant varRang(m_pSelection.GetRange()),index(1);
		inlineshapes = m_pSelection.GetInlineShapes();
		inlineshapes.AddPicture(strImgPath.GetBuffer(),vFalse,vTrue,&varRang);
		m_pSelection.MoveRight(COleVariant((short)1),COleVariant(short(1)),COleVariant(short(0)));
		m_pSelection.TypeParagraph();

		pf.SetAlignment(1);
		m_pSelection.SetParagraphFormat(pf);
		CComVariant varLable(_T("图")),varTitle(strDescrption.GetBuffer()),
			varTitleAuto(_T("")),varPosition((long)1),varExclude((long)0);
	
		
		m_pSelection.InsertCaption(&varLable,&varTitle,&varTitleAuto,&varPosition,&varExclude);
		pf.SetAlignment(1);
		m_pSelection.SetParagraphFormat(pf);
		m_pSelection.MoveRight(COleVariant((short)1),COleVariant(short(1)),COleVariant(short(0)));
		m_pSelection.TypeParagraph();
		return true;

	}
	catch (_com_error* e)
	{
		m_strLastError.Format(TEXT("%s"),e->Description());
		return false;
	}
	return false;
}

bool CWordFormat::CreateTable( CString strDescrption,int nRow,int nColume, vector<CString>& vString )
{
	try
	{
		m_strLastError.Format(TEXT("未执行CreateWord初始化word!"));
		if (!m_bInit)
			return false;
		int nCount = vString.size();
		if (nCount != nRow * nColume)
		{
			m_strLastError.Format(TEXT("参数个数有误"));
			return false;
		}
		CString str;
		_ParagraphFormat pf;
		pf = m_pSelection.GetParagraphFormat();
		pf.SetAlignment(1);
		m_pSelection.SetParagraphFormat(pf);
		CComVariant styleset(_T("正文"));
		CComVariant vSetPath(GetStyle(styleset));
		m_pSelection.SetStyle(&vSetPath);
		SetFont(0,12,RGB(0,0,0),_T("宋体"));
		
		CComVariant varLable(_T("表")),varTitle(strDescrption.GetBuffer()),
			varTitleAuto(_T("")),varPosition((long)1),varExclude((long)0);

		
		//m_pSelection.InsertCaption(&varLable,&varTitle,&varTitleAuto,&varPosition,&varExclude);
		pf.SetAlignment(1);
		m_pSelection.SetParagraphFormat(pf);
		m_pSelection.MoveRight(COleVariant((short)1),COleVariant(short(1)),COleVariant(short(0)));
		m_pSelection.TypeParagraph();

		Tables tables = m_pDocument.GetTables();
		Range range = m_pSelection.GetRange();
		CComVariant v1((short)1);
		CComVariant v2((short)1);
		Table subTable = tables.Add( range, nRow, nColume, &v1, &v2 );
		Rows rows = subTable.GetRows();
		rows.SetAlignment( 1 );
		nCount = 0;
		for (int i = 0 ;i < nRow; i++)
		{
			for (int j = 0; j < nColume; j++)
			{
				m_pSelection.TypeText(vString[nCount++]);
				m_pSelection.MoveRight(COleVariant((short)1),
					COleVariant(short(1)),COleVariant(short(0)));
			}
		}
		m_pSelection.TypeText(TEXT(""));
		m_pSelection.TypeParagraph();


		return true;

	}
	catch (_com_error* e)
	{
		m_strLastError.Format(TEXT("%s"),e->Description());
		return false;
	}
	return false;
	return false;
}

