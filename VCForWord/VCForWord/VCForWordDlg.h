
// VCForWordDlg.h : 头文件
//

#pragma once

#define RTITableColumnSize 6
const CString RTITableFieldArray[RTITableColumnSize] = {_T("序号"), _T("属性名称"), _T("数据类型"),_T("数据长度"), _T("订购/发布"), _T("绑定属性/订购属性")};
const CString RTITableTitle = _T("表订购/发布属性列表");

// CVCForWordDlg 对话框
class CVCForWordDlg : public CDialogEx
{
// 构造
public:
	CVCForWordDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_VCFORWORD_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	void MakeRTITable(CDocument0& oDoc, CSelection& selection);
	CParagraph GetCurParagraph(CDocument0& curDoc);
	afx_msg void OnBnClickedButton1();
};
