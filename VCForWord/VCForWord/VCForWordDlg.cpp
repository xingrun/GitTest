
// VCForWordDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "VCForWord.h"
#include "VCForWordDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CVCForWordDlg �Ի���




CVCForWordDlg::CVCForWordDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CVCForWordDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CVCForWordDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CVCForWordDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON1, &CVCForWordDlg::OnBnClickedButton1)
END_MESSAGE_MAP()


// CVCForWordDlg ��Ϣ�������

BOOL CVCForWordDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
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

	// ���ô˶Ի����ͼ�ꡣ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������


	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CVCForWordDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CVCForWordDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CVCForWordDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CVCForWordDlg::MakeRTITable(CDocument0& oDoc, CSelection& selection)
{
	int nRowSize = 10;
	nRowSize = nRowSize == 0 ? 2 : nRowSize;	// �����������
	// new paragraph
	selection.TypeParagraph();	// ����һ��

	// Add table title
	selection.TypeParagraph();	// ����һ��
	selection.TypeText(RTITableTitle);
	CParagraph lastPara = GetCurParagraph(oDoc);
	//lastPara.put_Alignment(wdAlignParagraphCenter);	// ����������Ҳ�ܴ˿���
	//selection.EndOf(COleVariant((short)wdParagraph), COleVariant((short)wdMove));
	lastPara.put_Alignment(1);	// ����������Ҳ�ܴ˿���
	selection.EndOf(COleVariant((short)4), COleVariant((short)0));

	// Add table
	CTables0 wordTables = oDoc.get_Tables();
	CRange wordRange = selection.get_Range();
	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	CTable0 wordTable = wordTables.Add(wordRange, nRowSize , RTITableColumnSize, covTrue, covFalse); // ��ӱ��
	wordRange = wordTable.get_Range();

	// Make header
	CCell cell;
	for ( int i=1; i<RTITableColumnSize+1; i++ )
	{
		cell = wordTable.Cell(1, i); // ����һ�е�i�е�Ԫ��
		cell.Select();
		selection.TypeText(RTITableFieldArray[i-1]);
	}
	//selection.EndOf(COleVariant((short)wdStory), COleVariant((short)wdMove)); // �������༭
	selection.EndOf(COleVariant((short)6), COleVariant((short)0)); // �������༭

	for ( int i=1; i<RTITableColumnSize+1; i++ )
	{
		cell = wordTable.Cell(2, i); // ����һ�е�i�е�Ԫ��
		cell.Select();
		selection.TypeText(_T("111.0"));
	}
	//selection.EndOf(COleVariant((short)wdStory), COleVariant((short)wdMove)); // �������༭
	selection.EndOf(COleVariant((short)6), COleVariant((short)0)); // �������༭

	// �ϲ���Ԫ����Ҫע����ǣ��ϲ�����ǰ�����е�Ԫ��ĺϲ��������޷���ȡ��������Ϣ
	CRows rows;
	rows = wordTable.get_Rows();	// ��ȡ������
	CRow row;
	row = rows.Item(3);	// ָ�������
	wordRange = row.get_Range();
	CCells cells;
	cells = wordRange.get_Cells();	// �õ��������е�Ԫ��
	cells.Merge();	// �ϲ�������Ϊһ��
	cell = wordTable.Cell( 4, 1 );
	cell.Merge( wordTable.Cell( 5, 1 ) );	//�ϲ���һ���еĵ������������

	// �����ƶ���ʽ.ͨ��Selection�����ķ����Թ���������,���ҵ��ƶ�,ʵ�ֶԹ��Ķ�λ����
	selection.MoveRight( COleVariant((short)1), COleVariant((short)1), COleVariant((short)0) );	// �����ƶ���굽��һ���ַ�
	selection.MoveDown( COleVariant((short)5), covOptional, covOptional );	// �����ƶ���굽��һ��
	COleVariant vUnit((short)RTITableColumnSize);	// ����ƶ���ʽΪ��	
	COleVariant vCount((short)3);
	selection.Move( &vUnit, &vCount );	// �ƶ�3��
	cells.Merge();	// �ϲ�������Ϊһ��
	COleVariant vEnd((short)5);
	selection.EndKey( &vEnd, &covOptional );	// ������ƶ�����β
}

CParagraph CVCForWordDlg::GetCurParagraph(CDocument0& curDoc)
{
	CParagraphs paragraphs = curDoc.get_Paragraphs();
	CParagraph lastPara = paragraphs.get_Last();
	return lastPara;
}


void CVCForWordDlg::OnBnClickedButton1()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	COleVariant	covZero((short)0),
		covTrue((short)TRUE), 
		covFalse((short)FALSE), 
		covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR), 
		covDocxType((short)0);
	
	// ����word����
	CApplication wordApp; // wordApp
	CDocuments docxs; // docxs
	CDocument0 docx, docx_active; // docx
	if ( !wordApp.CreateDispatch(_T("Word.Application")) ) // ʵ����wordApp�������г�ʼ��
	{
		AfxMessageBox(_T("����û�а�װword��Ʒ��"));
		return;
	}
	else
	{
		wordApp.put_Visible(FALSE);  // �����ĵ���ʼ���ɼ�

		CString wordVersion = wordApp.get_Version();	// ��õ�ǰword�İ汾������word2010Ϊ14.0,2013Ϊ15.0

		// ****************** ���һ��document ******************
		// �õ�docxs
		docxs = wordApp.get_Documents();	// ��������һ��
		// =====================================
		//LPDISPATCH disp = wordApp.get_Documents();
		//if ( NULL == disp )
		//	return;// FALSE;
		//docxs.AttachDispatch(disp);
		//if ( NULL == docxs.m_lpDispatch )
		//	return;// FALSE;
		// =====================================

		// ���һ��docx
		docx = docxs.Add(covOptional, covOptional, covOptional, covOptional);	// δ��ģ��ʱ���������������
		// =====================================
		// 2��δ��ģ��
		//docx.AttachDispatch(docxs.Add(covOptional, covOptional, covOptional, covOptional));
		// 3��ʹ��ģ��
		//CComVariant tpl(_T("")), Visble, DocxType(0), NewTemplate(false);
		//docx = docxs.Add(&tpl,&NewTemplate,&DocxType,&Visble);
		// =====================================
		if ( NULL == docx.m_lpDispatch )
			return;

		// ****************** ����ҳ�߾� ******************
		// ���ڴ����ĵ�����ҪCPageSetup.h
		docx_active = wordApp.get_ActiveDocument();
		CPageSetup oPageSetup = docx_active.get_PageSetup();
		// ����Ϊҳ�淽���ҳ�߾�
		if ( oPageSetup.get_Orientation() == 0 )	// ��Ϊ����������Ϊ��������wdOrientPortrait=0������wdOrientLandscape=1
		{
			oPageSetup.put_Orientation(1);	// ����
			// �����������ұ�࣬��λ羣����²������õ�ҳ�߾��ǡ����С�
			oPageSetup.put_TopMargin( (float) 72);	// ����ʱ72=2.54cm��Ĭ��ʱ90=3.17cm��10��0.35cm
			oPageSetup.put_BottomMargin( (float) 72);	// ����ʱ72=2.54cm��Ĭ��ʱ90=3.17cm��10��0.35cm
			oPageSetup.put_LeftMargin( (float) 54);	// ����ʱ54=1.9cm��Ĭ��ʱ72=2.54cm
			oPageSetup.put_RightMargin( (float) 54);	// ����ʱ54=1.9cm��Ĭ��ʱ72=2.54cm
		}
		//else	// ����Ϊ����
		//{
		//	oPageSetup.put_Orientation(0);
		//	// �����������ұ�࣬��λ羣����²������õ�ҳ�߾��ǡ����С�
		//	oPageSetup.put_TopMargin( (float) 72);	// ����ʱ72=2.54cm��Ĭ�ϻ���ͨʱ72=2.54cm��10��0.35cm
		//	oPageSetup.put_BottomMargin( (float) 72);	// ����ʱ72=2.54cm��Ĭ�ϻ���ͨʱ72=2.54cm��10��0.35cm
		//	oPageSetup.put_LeftMargin( (float) 54);	// ����ʱ54=1.9cm��Ĭ�ϻ���ͨʱ90=3.17cm
		//	oPageSetup.put_RightMargin( (float) 54);	// ����ʱ54=1.9cm��Ĭ�ϻ���ͨʱ90=3.17cm
		//}

		// ����һ��CSelection���󣬲�ʵ����
		CSelection wordSelection = wordApp.get_Selection();

		// ****************** �����ĵ����� ******************
		wordSelection.TypeText(_T("����������汨��"));
		wordSelection.HomeKey(COleVariant((short)5), COleVariant((short)1)); // wdLine=5�����ص�ǰ���ף���ѡ��ǰ��
		wordSelection.put_Style( COleVariant((short)-2) );// ����Ϊ������1����ʽ��wdStyleHeading1=-2
		// ����ѡ���������壬һ��Ҫ������ʽ�󣬷����ʽ�ᱻ��ʽ�ĸ���
		CFont0 font = wordSelection.get_Font();
		font.put_Name(_T("΢���ź�"));
		font.put_Size(16);	// ����ѡ����вſ����޸ģ���������HomeKey����
		// ��õ�ǰ���䣬�����ö��뷽ʽ
		CParagraph lastPara = GetCurParagraph(docx);
		lastPara.put_Alignment(1);	// wdAlignParagraphLeft=0, wdAlignParagraphCenter=1,wdAlignParagraphRight=2
		// ������ǰ����༭���ƶ���굽�����
		wordSelection.EndOf(COleVariant((short)4), COleVariant((short)0));	// wdParagraph=4,wdMove=0

		wordSelection.TypeParagraph(); // ����һ��
		COleVariant covTime(_T("yyyy-MM-dd:dddd"));	// ʱ���ʽ�ɵ���
		wordSelection.InsertDateTime(covTime, covFalse, covOptional, covOptional, covOptional);	// ���뵱ǰʱ��
		wordSelection.EndOf(COleVariant((short)4), COleVariant((short)0));	// ������ǰ����༭��wdParagraph=4��wdMove=0

		// ���ɱ��
		MakeRTITable( docx, wordSelection );

		// ����ΪΪ��ͬ�������ò�ͬ����Ͷ��뷽ʽʾ��
		wordSelection.TypeParagraph(); // ����һ��
		wordSelection.TypeText(_T("end of the story!"));
		wordSelection.HomeKey(COleVariant((short)5), COleVariant((short)1)); // wdLine=5�����ص�ǰ���ף���ѡ��ǰ��
		/*CFont0 */font = wordSelection.get_Font();
		font.put_Size(20);	// ����ѡ����вſ����޸ģ���������HomeKey����
		/*CParagraph */lastPara = GetCurParagraph(docx);
		lastPara.put_Alignment(3);	// �Ҷ���
		wordSelection.EndOf(COleVariant((short)4), COleVariant((short)0));	// ������ǰ����༭��wdParagraph=4��wdMove=0

		wordSelection.TypeParagraph(); // ����һ��
		wordSelection.TypeText(_T("Thanks for reading!"));
		wordSelection.HomeKey(COleVariant((short)5), COleVariant((short)1)); // wdLine=5�����ص�ǰ���ף���ѡ��ǰ��
		/*CFont0 */font = wordSelection.get_Font();
		font.put_Size(10);	// ����ѡ����вſ����޸ģ���������HomeKey����
		font.put_Name(_T("Times New Roman"));
		/*CParagraph */lastPara = GetCurParagraph(docx);
		lastPara.put_Alignment(1);	// ���ж���
		wordSelection.EndOf(COleVariant((short)4), COleVariant((short)0));	// ������ǰ����༭��wdParagraph=4��wdMove=0

		// �����ҳ�������ڻ�ҳ
		wordSelection.InsertBreak(covOptional);
		
		// ���빫ʽ��������
		CFields fields = wordSelection.get_Fields();
		COleVariant ofont = _variant_t(_T("Times New Roman"));
		COleVariant text = _variant_t(_T("EQ \\a \\ar \\co2 \\vs3 \\hs3(Axy,Bxy,A,B)"));	// ע��Ҫ����\\��һ��ת��󲻶ԣ�����
		fields.Add( wordSelection.get_Range(), covOptional, text, covFalse );
		wordSelection.HomeKey(COleVariant((short)5), COleVariant((short)1)); // wdLine=5�����ص�ǰ���ף���ѡ��ǰ��
		lastPara = GetCurParagraph(docx);
		lastPara.put_Alignment(0);	// �����
		wordSelection.EndOf(COleVariant((short)4), COleVariant((short)0));	// ������ǰ����༭��wdParagraph=4��wdMove=0

		// ��ȡӦ�õ�ǰDebug·��
		char fileName[MAX_PATH];
		GetModuleFileName(NULL, fileName, MAX_PATH);
		char dir[260];
		char dirver[100];
		_splitpath(fileName, dirver, dir, NULL, NULL);
		CString strAppPath = dirver;
		strAppPath += dir;
		//CString strAppPath = _T("D:\\");

		// ****************** ����ͼƬʾ�� ******************
		// ��ҪCWindow0.h, CPane0.h, CView0.h
		wordSelection.TypeParagraph();	// ����һ��
		CString strPicture = strAppPath + _T("\\��ͼ.jpg");
		CnlineShapes nLineShapes = wordSelection.get_InlineShapes();
		CnlineShape nLineshape = nLineShapes.AddPicture(strPicture, covFalse, covTrue, covOptional);

		// ****************** ����ҳüҳ�� ******************
		CWindow0 oWind = docx.get_ActiveWindow();
		CPane0 oPane = oWind.get_ActivePane();	// һ����CPane��ΪCPane0������
		CView0 oView = oPane.get_View();
		// =============== ����ҳü ===============
		oView.put_SeekView(9);	// wdSeekCurrentPageHeader=9
		/*CFont0 */font = wordSelection.get_Font();	// ����ѡ����������
		font.put_Name(_T("���Ŀ���"));
		font.put_Size(16);
		/*CParagraphFormat */lastPara = wordSelection.get_ParagraphFormat();	// Ĭ��Ϊ����
		lastPara.put_Alignment(1);	// wdAlignParagraphLeft=0, wdAlignParagraphCenter=1, wdAlignParagraphRight=2
		wordSelection.TypeText(_T("�����ѧ"));

		// =============== ����ҳ�ţ�����ҳ�� ===============
		oView.put_SeekView(10);	// wdSeekCurrentPageFooter=10
		/*CFont0 */font = wordSelection.get_Font();	// ����ѡ���������壬һ��Ҫ������ʽ�󣬷����ʽ�ᱻ��ʽ�ĸ���
		font.put_Name(_T("���Ŀ���"));
		font.put_Size(16);
		/*CParagraphFormat */lastPara = wordSelection.get_ParagraphFormat();	// Ĭ��Ϊ����
		lastPara.put_Alignment(1);	// wdAlignParagraphLeft=0, wdAlignParagraphCenter=1, wdAlignParagraphRight=2
		// ���ҳ��
		wordSelection.TypeText(_T("��ҳ ��ҳ"));
		wordSelection.MoveLeft( COleVariant((short)1), COleVariant((short)4), &covZero );
		/*CFields */fields = wordSelection.get_Fields();
		fields.Add( wordSelection.get_Range(), COleVariant((short)33), COleVariant("PAGE  "), &covTrue );	// ����ҳ���򣬵�ǰҳ��
		wordSelection.MoveRight( COleVariant((short)1), COleVariant((short)3), &covZero);
		fields.Add( wordSelection.get_Range(), COleVariant((short)26), COleVariant("NUMPAGES  "), &covTrue );	// ����ҳ������ҳ��
		oView.put_SeekView(0);	// �ر�ҳüҳ��,wdSeekMainDocument=0���ص������ĵ�

		// Word����ɼ�����ʾ����
		//wordApp.put_Visible(TRUE);

		// ����ɹ�
		CString strSavePath = strAppPath;
		strSavePath += _T("\\����.docx");
		docx.SaveAs(COleVariant(strSavePath), covOptional, covOptional, covOptional, covOptional,
			covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional,covOptional);

		// �˳�wordӦ��
		docx.Close(covFalse, covOptional, covOptional);
		wordApp.Quit(covOptional, covOptional, covOptional);
		wordApp.ReleaseDispatch();

		MessageBox(_T("���ɳɹ���"));
	}	
}
