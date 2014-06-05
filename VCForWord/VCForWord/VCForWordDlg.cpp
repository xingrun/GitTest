
// VCForWordDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "VCForWord.h"
#include "VCForWordDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
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

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CVCForWordDlg 对话框




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


// CVCForWordDlg 消息处理程序

BOOL CVCForWordDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CVCForWordDlg::OnPaint()
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
HCURSOR CVCForWordDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CVCForWordDlg::MakeRTITable(CDocument0& oDoc, CSelection& selection)
{
	int nRowSize = 10;
	nRowSize = nRowSize == 0 ? 2 : nRowSize;	// 表格至少两行
	// new paragraph
	selection.TypeParagraph();	// 新起一段

	// Add table title
	selection.TypeParagraph();	// 新起一段
	selection.TypeText(RTITableTitle);
	CParagraph lastPara = GetCurParagraph(oDoc);
	//lastPara.put_Alignment(wdAlignParagraphCenter);	// 下面表格内容也受此控制
	//selection.EndOf(COleVariant((short)wdParagraph), COleVariant((short)wdMove));
	lastPara.put_Alignment(1);	// 下面表格内容也受此控制
	selection.EndOf(COleVariant((short)4), COleVariant((short)0));

	// Add table
	CTables0 wordTables = oDoc.get_Tables();
	CRange wordRange = selection.get_Range();
	COleVariant covTrue((short)TRUE), covFalse((short)FALSE), covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
	CTable0 wordTable = wordTables.Add(wordRange, nRowSize , RTITableColumnSize, covTrue, covFalse); // 添加表格
	wordRange = wordTable.get_Range();

	// Make header
	CCell cell;
	for ( int i=1; i<RTITableColumnSize+1; i++ )
	{
		cell = wordTable.Cell(1, i); // 表格第一行第i列单元格
		cell.Select();
		selection.TypeText(RTITableFieldArray[i-1]);
	}
	//selection.EndOf(COleVariant((short)wdStory), COleVariant((short)wdMove)); // 结束表格编辑
	selection.EndOf(COleVariant((short)6), COleVariant((short)0)); // 结束表格编辑

	for ( int i=1; i<RTITableColumnSize+1; i++ )
	{
		cell = wordTable.Cell(2, i); // 表格第一行第i列单元格
		cell.Select();
		selection.TypeText(_T("111.0"));
	}
	//selection.EndOf(COleVariant((short)wdStory), COleVariant((short)wdMove)); // 结束表格编辑
	selection.EndOf(COleVariant((short)6), COleVariant((short)0)); // 结束表格编辑

	// 合并单元格，需要注意的是，合并整行前不能有单元格的合并，否则无法获取表格的行信息
	CRows rows;
	rows = wordTable.get_Rows();	// 获取表格的行
	CRow row;
	row = rows.Item(3);	// 指向第三行
	wordRange = row.get_Range();
	CCells cells;
	cells = wordRange.get_Cells();	// 得到该行所有单元格
	cells.Merge();	// 合并第三行为一列
	cell = wordTable.Cell( 4, 1 );
	cell.Merge( wordTable.Cell( 5, 1 ) );	//合并第一列中的第四行与第五行

	// 光标的移动方式.通过Selection类对象的方法对光标进行上下,左右等移动,实现对光标的定位功能
	selection.MoveRight( COleVariant((short)1), COleVariant((short)1), COleVariant((short)0) );	// 向右移动鼠标到下一个字符
	selection.MoveDown( COleVariant((short)5), covOptional, covOptional );	// 向下移动鼠标到下一行
	COleVariant vUnit((short)RTITableColumnSize);	// 光标移动方式为行	
	COleVariant vCount((short)3);
	selection.Move( &vUnit, &vCount );	// 移动3行
	cells.Merge();	// 合并第三行为一列
	COleVariant vEnd((short)5);
	selection.EndKey( &vEnd, &covOptional );	// 将光标移动到行尾
}

CParagraph CVCForWordDlg::GetCurParagraph(CDocument0& curDoc)
{
	CParagraphs paragraphs = curDoc.get_Paragraphs();
	CParagraph lastPara = paragraphs.get_Last();
	return lastPara;
}


void CVCForWordDlg::OnBnClickedButton1()
{
	// TODO: 在此添加控件通知处理程序代码
	COleVariant	covZero((short)0),
		covTrue((short)TRUE), 
		covFalse((short)FALSE), 
		covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR), 
		covDocxType((short)0);
	
	// 定义word变量
	CApplication wordApp; // wordApp
	CDocuments docxs; // docxs
	CDocument0 docx, docx_active; // docx
	if ( !wordApp.CreateDispatch(_T("Word.Application")) ) // 实例化wordApp，必须有初始化
	{
		AfxMessageBox(_T("本机没有安装word产品！"));
		return;
	}
	else
	{
		wordApp.put_Visible(FALSE);  // 设置文档开始不可见

		CString wordVersion = wordApp.get_Version();	// 获得当前word的版本，比如word2010为14.0,2013为15.0

		// ****************** 添加一个document ******************
		// 得到docxs
		docxs = wordApp.get_Documents();	// 或者下面一段
		// =====================================
		//LPDISPATCH disp = wordApp.get_Documents();
		//if ( NULL == disp )
		//	return;// FALSE;
		//docxs.AttachDispatch(disp);
		//if ( NULL == docxs.m_lpDispatch )
		//	return;// FALSE;
		// =====================================

		// 添加一个docx
		docx = docxs.Add(covOptional, covOptional, covOptional, covOptional);	// 未用模板时，或者下面段两种
		// =====================================
		// 2，未用模板
		//docx.AttachDispatch(docxs.Add(covOptional, covOptional, covOptional, covOptional));
		// 3，使用模板
		//CComVariant tpl(_T("")), Visble, DocxType(0), NewTemplate(false);
		//docx = docxs.Add(&tpl,&NewTemplate,&DocxType,&Visble);
		// =====================================
		if ( NULL == docx.m_lpDispatch )
			return;

		// ****************** 设置页边距 ******************
		// 放在创建文档后，需要CPageSetup.h
		docx_active = wordApp.get_ActiveDocument();
		CPageSetup oPageSetup = docx_active.get_PageSetup();
		// 设置为页面方向和页边距
		if ( oPageSetup.get_Orientation() == 0 )	// 若为纵向则设置为横向，纵向wdOrientPortrait=0，横向wdOrientLandscape=1
		{
			oPageSetup.put_Orientation(1);	// 横向
			// 设置上下左右变距，单位缇，以下参数设置的页边距是“适中”
			oPageSetup.put_TopMargin( (float) 72);	// 适中时72=2.54cm，默认时90=3.17cm；10≈0.35cm
			oPageSetup.put_BottomMargin( (float) 72);	// 适中时72=2.54cm，默认时90=3.17cm；10≈0.35cm
			oPageSetup.put_LeftMargin( (float) 54);	// 适中时54=1.9cm，默认时72=2.54cm
			oPageSetup.put_RightMargin( (float) 54);	// 适中时54=1.9cm，默认时72=2.54cm
		}
		//else	// 设置为纵向
		//{
		//	oPageSetup.put_Orientation(0);
		//	// 设置上下左右变距，单位缇，以下参数设置的页边距是“适中”
		//	oPageSetup.put_TopMargin( (float) 72);	// 适中时72=2.54cm，默认或普通时72=2.54cm；10≈0.35cm
		//	oPageSetup.put_BottomMargin( (float) 72);	// 适中时72=2.54cm，默认或普通时72=2.54cm；10≈0.35cm
		//	oPageSetup.put_LeftMargin( (float) 54);	// 适中时54=1.9cm，默认或普通时90=3.17cm
		//	oPageSetup.put_RightMargin( (float) 54);	// 适中时54=1.9cm，默认或普通时90=3.17cm
		//}

		// 声明一个CSelection对象，并实例化
		CSelection wordSelection = wordApp.get_Selection();

		// ****************** 设置文档内容 ******************
		wordSelection.TypeText(_T("虚拟试验仿真报表"));
		wordSelection.HomeKey(COleVariant((short)5), COleVariant((short)1)); // wdLine=5，返回当前行首，并选择当前行
		wordSelection.put_Style( COleVariant((short)-2) );// 设置为“标题1“样式，wdStyleHeading1=-2
		// 设置选择区域字体，一定要放在样式后，否则格式会被样式的覆盖
		CFont0 font = wordSelection.get_Font();
		font.put_Name(_T("微软雅黑"));
		font.put_Size(16);	// 必须选择该行才可以修改，即必须有HomeKey那行
		// 获得当前段落，并设置对齐方式
		CParagraph lastPara = GetCurParagraph(docx);
		lastPara.put_Alignment(1);	// wdAlignParagraphLeft=0, wdAlignParagraphCenter=1,wdAlignParagraphRight=2
		// 结束当前段落编辑，移动光标到段落后
		wordSelection.EndOf(COleVariant((short)4), COleVariant((short)0));	// wdParagraph=4,wdMove=0

		wordSelection.TypeParagraph(); // 新起一段
		COleVariant covTime(_T("yyyy-MM-dd:dddd"));	// 时间格式可调整
		wordSelection.InsertDateTime(covTime, covFalse, covOptional, covOptional, covOptional);	// 插入当前时间
		wordSelection.EndOf(COleVariant((short)4), COleVariant((short)0));	// 结束当前段落编辑，wdParagraph=4，wdMove=0

		// 生成表格
		MakeRTITable( docx, wordSelection );

		// 以下为为不同段落设置不同字体和对齐方式示例
		wordSelection.TypeParagraph(); // 新起一段
		wordSelection.TypeText(_T("end of the story!"));
		wordSelection.HomeKey(COleVariant((short)5), COleVariant((short)1)); // wdLine=5，返回当前行首，并选择当前行
		/*CFont0 */font = wordSelection.get_Font();
		font.put_Size(20);	// 必须选择该行才可以修改，即必须有HomeKey那行
		/*CParagraph */lastPara = GetCurParagraph(docx);
		lastPara.put_Alignment(3);	// 右对齐
		wordSelection.EndOf(COleVariant((short)4), COleVariant((short)0));	// 结束当前段落编辑，wdParagraph=4，wdMove=0

		wordSelection.TypeParagraph(); // 新起一段
		wordSelection.TypeText(_T("Thanks for reading!"));
		wordSelection.HomeKey(COleVariant((short)5), COleVariant((short)1)); // wdLine=5，返回当前行首，并选择当前行
		/*CFont0 */font = wordSelection.get_Font();
		font.put_Size(10);	// 必须选择该行才可以修改，即必须有HomeKey那行
		font.put_Name(_T("Times New Roman"));
		/*CParagraph */lastPara = GetCurParagraph(docx);
		lastPara.put_Alignment(1);	// 居中对齐
		wordSelection.EndOf(COleVariant((short)4), COleVariant((short)0));	// 结束当前段落编辑，wdParagraph=4，wdMove=0

		// 插入分页符，用于换页
		wordSelection.InsertBreak(covOptional);
		
		// 插入公式，操作域
		CFields fields = wordSelection.get_Fields();
		COleVariant ofont = _variant_t(_T("Times New Roman"));
		COleVariant text = _variant_t(_T("EQ \\a \\ar \\co2 \\vs3 \\hs3(Axy,Bxy,A,B)"));	// 注意要两个\\，一个转义后不对！！！
		fields.Add( wordSelection.get_Range(), covOptional, text, covFalse );
		wordSelection.HomeKey(COleVariant((short)5), COleVariant((short)1)); // wdLine=5，返回当前行首，并选择当前行
		lastPara = GetCurParagraph(docx);
		lastPara.put_Alignment(0);	// 左对齐
		wordSelection.EndOf(COleVariant((short)4), COleVariant((short)0));	// 结束当前段落编辑，wdParagraph=4，wdMove=0

		// 获取应用当前Debug路径
		char fileName[MAX_PATH];
		GetModuleFileName(NULL, fileName, MAX_PATH);
		char dir[260];
		char dirver[100];
		_splitpath(fileName, dirver, dir, NULL, NULL);
		CString strAppPath = dirver;
		strAppPath += dir;
		//CString strAppPath = _T("D:\\");

		// ****************** 插入图片示例 ******************
		// 需要CWindow0.h, CPane0.h, CView0.h
		wordSelection.TypeParagraph();	// 另起一段
		CString strPicture = strAppPath + _T("\\截图.jpg");
		CnlineShapes nLineShapes = wordSelection.get_InlineShapes();
		CnlineShape nLineshape = nLineShapes.AddPicture(strPicture, covFalse, covTrue, covOptional);

		// ****************** 设置页眉页脚 ******************
		CWindow0 oWind = docx.get_ActiveWindow();
		CPane0 oPane = oWind.get_ActivePane();	// 一定将CPane改为CPane0或其他
		CView0 oView = oPane.get_View();
		// =============== 设置页眉 ===============
		oView.put_SeekView(9);	// wdSeekCurrentPageHeader=9
		/*CFont0 */font = wordSelection.get_Font();	// 设置选择区域字体
		font.put_Name(_T("华文楷体"));
		font.put_Size(16);
		/*CParagraphFormat */lastPara = wordSelection.get_ParagraphFormat();	// 默认为居中
		lastPara.put_Alignment(1);	// wdAlignParagraphLeft=0, wdAlignParagraphCenter=1, wdAlignParagraphRight=2
		wordSelection.TypeText(_T("网络大学"));

		// =============== 设置页脚，包括页码 ===============
		oView.put_SeekView(10);	// wdSeekCurrentPageFooter=10
		/*CFont0 */font = wordSelection.get_Font();	// 设置选择区域字体，一定要放在样式后，否则格式会被样式的覆盖
		font.put_Name(_T("华文楷体"));
		font.put_Size(16);
		/*CParagraphFormat */lastPara = wordSelection.get_ParagraphFormat();	// 默认为居中
		lastPara.put_Alignment(1);	// wdAlignParagraphLeft=0, wdAlignParagraphCenter=1, wdAlignParagraphRight=2
		// 添加页码
		wordSelection.TypeText(_T("第页 共页"));
		wordSelection.MoveLeft( COleVariant((short)1), COleVariant((short)4), &covZero );
		/*CFields */fields = wordSelection.get_Fields();
		fields.Add( wordSelection.get_Range(), COleVariant((short)33), COleVariant("PAGE  "), &covTrue );	// 增加页码域，当前页码
		wordSelection.MoveRight( COleVariant((short)1), COleVariant((short)3), &covZero);
		fields.Add( wordSelection.get_Range(), COleVariant((short)26), COleVariant("NUMPAGES  "), &covTrue );	// 增加页码域，总页数
		oView.put_SeekView(0);	// 关闭页眉页脚,wdSeekMainDocument=0，回到主控文档

		// Word程序可见，显示报表
		//wordApp.put_Visible(TRUE);

		// 保存成果
		CString strSavePath = strAppPath;
		strSavePath += _T("\\报表.docx");
		docx.SaveAs(COleVariant(strSavePath), covOptional, covOptional, covOptional, covOptional,
			covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional, covOptional,covOptional);

		// 退出word应用
		docx.Close(covFalse, covOptional, covOptional);
		wordApp.Quit(covOptional, covOptional, covOptional);
		wordApp.ReleaseDispatch();

		MessageBox(_T("生成成功！"));
	}	
}
