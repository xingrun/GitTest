// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

#import "C:\\Program Files\\Microsoft Office\\Office14\\MSWORD.OLB" no_namespace
// CApplicationEvents0 ��װ��

class CApplicationEvents0 : public COleDispatchDriver
{
public:
	CApplicationEvents0(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CApplicationEvents0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplicationEvents0(const CApplicationEvents0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// IApplicationEvents ����
public:
	void Startup()
	{
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Quit()
	{
		InvokeHelper(0x2, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void DocumentChange()
	{
		InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// IApplicationEvents ����
public:

};
