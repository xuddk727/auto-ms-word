
// WordTestExeDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "WordTestExe.h"
#include "WordTestExeDlg.h"
#include "Word.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

#define RTHROW(a) {if(a == FALSE)  throw a;}


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CWordTestExeDlg 对话框




CWordTestExeDlg::CWordTestExeDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CWordTestExeDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CWordTestExeDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CWordTestExeDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDC_BUTTON1, &CWordTestExeDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CWordTestExeDlg 消息处理程序

BOOL CWordTestExeDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CWordTestExeDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CWordTestExeDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CWordTestExeDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CWordTestExeDlg::OnBnClickedButton1()
{
	CString strTitle = TEXT("测试文件.doc");
	CFileDialog filedlg( FALSE, "doc", strTitle, OFN_FILEMUSTEXIST| OFN_HIDEREADONLY,"word文件|*.doc||", this );
	if( filedlg.DoModal() != IDOK )
		return;
	strTitle = filedlg.GetPathName();
	CWordBase* pBase = CreateWord(MODE_FORMAT);
	CWordFormat* pFormat = reinterpret_cast<CWordFormat*>(pBase);
	try
	{
		RTHROW(pFormat->CreateWord(TRUE));//生成的同时是否显示
		RTHROW(pFormat->CreateTitle("公司红头文件",1));
		RTHROW(pFormat->CreateSection("标题"));
		RTHROW(pFormat->CreateSection("子标题",2));
		RTHROW(pFormat->CreateText("正文段落1"));
		RTHROW(pFormat->CreateText("正文段落2"));
		vector<CString> vStrResult;
		vStrResult.push_back(TEXT("角度"));
		vStrResult.push_back(TEXT("二阶矩"));
		vStrResult.push_back(TEXT("对比度"));
		vStrResult.push_back(TEXT("相关性"));
		CString strTest;
		for (int i = 0;i < 4;i++)
		{
			
			strTest.Format(TEXT("%d"),i);
			vStrResult.push_back(strTest);
			
		}
		RTHROW(pFormat->CreateTable(TEXT("表格输入结果"),2,4,vStrResult));
		RTHROW(pFormat->SaveWord(strTitle));
	}
	catch (...)
	{
		MessageBox(pFormat->GetLastError());
	}
	
	
}
