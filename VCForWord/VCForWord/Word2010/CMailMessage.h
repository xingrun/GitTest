// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

#import "C:\\Program Files\\Microsoft Office\\Office14\\MSWORD.OLB" no_namespace
// CMailMessage ��װ��

class CMailMessage : public COleDispatchDriver
{
public:
	CMailMessage(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CMailMessage(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CMailMessage(const CMailMessage& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// MailMessage ����
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void CheckName()
	{
		InvokeHelper(0x14e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Delete()
	{
		InvokeHelper(0x14f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void DisplayMoveDialog()
	{
		InvokeHelper(0x150, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void DisplayProperties()
	{
		InvokeHelper(0x151, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void DisplaySelectNamesDialog()
	{
		InvokeHelper(0x152, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Forward()
	{
		InvokeHelper(0x153, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void GoToNext()
	{
		InvokeHelper(0x154, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void GoToPrevious()
	{
		InvokeHelper(0x155, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Reply()
	{
		InvokeHelper(0x156, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ReplyAll()
	{
		InvokeHelper(0x157, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ToggleHeader()
	{
		InvokeHelper(0x158, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// MailMessage ����
public:

};
