
// VCForWordDlg.h : ͷ�ļ�
//

#pragma once

#define RTITableColumnSize 6
const CString RTITableFieldArray[RTITableColumnSize] = {_T("���"), _T("��������"), _T("��������"),_T("���ݳ���"), _T("����/����"), _T("������/��������")};
const CString RTITableTitle = _T("����/���������б�");

// CVCForWordDlg �Ի���
class CVCForWordDlg : public CDialogEx
{
// ����
public:
	CVCForWordDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_VCFORWORD_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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
