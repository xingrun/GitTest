// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

#import "C:\\Program Files\\Microsoft Office\\Office14\\MSWORD.OLB" no_namespace
// CDocumentEvents2 ��װ��

class CDocumentEvents2 : public COleDispatchDriver
{
public:
	CDocumentEvents2(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CDocumentEvents2(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDocumentEvents2(const CDocumentEvents2& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// DocumentEvents2 ����
public:
	void New()
	{
		InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Open()
	{
		InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Close()
	{
		InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Sync(long SyncEventType)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, parms, SyncEventType);
	}
	void XMLAfterInsert(LPDISPATCH NewXMLNode, BOOL InUndoRedo)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BOOL ;
		InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NewXMLNode, InUndoRedo);
	}
	void XMLBeforeDelete(LPDISPATCH DeletedRange, LPDISPATCH OldXMLNode, BOOL InUndoRedo)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH VTS_BOOL ;
		InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, parms, DeletedRange, OldXMLNode, InUndoRedo);
	}
	void ContentControlAfterAdd(LPDISPATCH NewContentControl, BOOL InUndoRedo)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BOOL ;
		InvokeHelper(0xc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, NewContentControl, InUndoRedo);
	}
	void ContentControlBeforeDelete(LPDISPATCH OldContentControl, BOOL InUndoRedo)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BOOL ;
		InvokeHelper(0xd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, OldContentControl, InUndoRedo);
	}
	void ContentControlOnExit(LPDISPATCH ContentControl, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0xe, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ContentControl, Cancel);
	}
	void ContentControlOnEnter(LPDISPATCH ContentControl)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xf, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ContentControl);
	}
	void ContentControlBeforeStoreUpdate(LPDISPATCH ContentControl, BSTR * Content)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBSTR ;
		InvokeHelper(0x10, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ContentControl, Content);
	}
	void ContentControlBeforeContentUpdate(LPDISPATCH ContentControl, BSTR * Content)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBSTR ;
		InvokeHelper(0x11, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ContentControl, Content);
	}
	void BuildingBlockInsert(LPDISPATCH Range, LPCTSTR Name, LPCTSTR Category, LPCTSTR BlockType, LPCTSTR Template)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_BSTR VTS_BSTR VTS_BSTR ;
		InvokeHelper(0x12, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Range, Name, Category, BlockType, Template);
	}

	// DocumentEvents2 ����
public:

};
