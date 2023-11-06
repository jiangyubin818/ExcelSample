
// ExcelSampleDlg.h: 头文件
//

#pragma once

#include <set>

// CExcelSampleDlg 对话框
class CExcelSampleDlg : public CDialogEx
{
// 构造
public:
	CExcelSampleDlg(CWnd* pParent = nullptr);	// 标准构造函数

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_EXCELSAMPLE_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButton1();
	afx_msg void OnBnClickedButton2();

	void StartConvert();

private:
	CString				m_szFolderName;
	std::set<CString>	m_setFilesName;
};
