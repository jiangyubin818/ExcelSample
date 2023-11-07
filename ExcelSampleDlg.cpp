/*************************************************************************
	> File Name:    ExcelSampleDlg.cpp
	> Author:       jiangyubin
	> Mail:         jiangyubin818@gmail.com 
    > QQ:           1327388399
	> Created Time: Mon 06 Nov 2023 10:10:50 PM PST
 ************************************************************************/

// ExcelSampleDlg.cpp: 实现文件
//

#include <string>

#include "pch.h"
#include "framework.h"
#include "ExcelSample.h"
#include "ExcelSampleDlg.h"
#include "afxdialogex.h"
#include "include/Excel.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CExcelSampleDlg 对话框



CExcelSampleDlg::CExcelSampleDlg(CWnd* pParent /*=nullptr*/)
	: CDialogEx(IDD_EXCELSAMPLE_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcelSampleDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CExcelSampleDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CExcelSampleDlg::OnBnClickedButton1)
	ON_BN_CLICKED(IDC_BUTTON2, &CExcelSampleDlg::OnBnClickedButton2)
END_MESSAGE_MAP()


// CExcelSampleDlg 消息处理程序

BOOL CExcelSampleDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CExcelSampleDlg::OnPaint()
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
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CExcelSampleDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void GetFileFromDir(CString& szDirPath, std::set<CString>& setFileList)
{
	setFileList.clear();

	const CString szXls = szDirPath + "\\*.xlsx";

	CFileFind finder;
	BOOL bWorking = finder.FindFile(szXls);
	while (bWorking)
	{
		bWorking = finder.FindNextFile();
		const CString strFile = finder.GetFileName();
		if (!strFile.IsEmpty())
		{
			setFileList.insert(strFile);
		}
	}
}

void CExcelSampleDlg::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码
	m_szFolderName.Empty();

	CFolderPickerDialog fd(NULL, 0, this, 0);
	if (IDOK == fd.DoModal())
	{
		m_szFolderName = fd.GetPathName();
		if (!m_szFolderName.IsEmpty())
		{
			GetFileFromDir(m_szFolderName, m_setFilesName);

			CString szName;
			szName.Format("Select Folder Name Is: %s\nTotal Excel Files Number Is:%2d", m_szFolderName, m_setFilesName.size());
			MessageBox(szName, MB_OK);
			return;
		}
	}
	AfxMessageBox(_T("Select Folder Failed!!!"), MB_OK);
}


void CExcelSampleDlg::OnBnClickedButton2()
{
	// TODO: 在此添加控件通知处理程序代码
	StartConvert();
}

void CExcelSampleDlg::StartConvert()
{
	if (m_setFilesName.empty())
	{
		AfxMessageBox(_T("Get Excel Files Failed!!!"));
		return;
	}

	Excel excl;
	const bool bInit = excl.initExcel();
	if (!bInit)
	{
		AfxMessageBox(_T("Excel Init Failed!!!"));
		excl.release();
		return;
	}

	for (const auto& szFile : m_setFilesName)
	{
		const CString& szFileName = m_szFolderName + "\\" + szFile;
		const std::string path = szFileName.GetString();
		if (!excl.open(path.c_str()))
		{
			AfxMessageBox(_T("Excel Open Failed!!!"));
			continue;
		}

		CString szPrompt;

		const auto nSheetCount = excl.getSheetCount();
		for (auto i = 0; i < nSheetCount; ++i)
		{
			const CString strSheetName = excl.getSheetName(i + 1);
			if (excl.loadSheet(strSheetName))
			{
				CString szCurrentSheet;
				szCurrentSheet.Format("[%04d]SheetName Is:	%s\nRow Number Is: %04d\nCol Number Is: %04d"
					, i + 1
					, strSheetName
					, excl.getRowCount()
					, excl.getColumnCount());
				if (!szCurrentSheet.IsEmpty())
				{
					szPrompt = szPrompt + szCurrentSheet + "\n";
				}
			}
		}

		szPrompt = "Files Name Is: " + szFile + "\n" + szPrompt;
		MessageBox(szPrompt, MB_OK);
		excl.close();
	}
	excl.release();
}
