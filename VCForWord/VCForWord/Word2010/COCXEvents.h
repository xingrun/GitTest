// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

#import "C:\\Program Files\\Microsoft Office\\Office14\\MSWORD.OLB" no_namespace
// COCXEvents ��װ��

class COCXEvents : public COleDispatchDriver
{
public:
	COCXEvents(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	COCXEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COCXEvents(const COCXEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// OCXEvents ����
public:
	void GotFocus()
	{
		InvokeHelper(0x800100e0, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void LostFocus()
	{
		InvokeHelper(0x800100e1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// OCXEvents ����
public:

};
