// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

#import "C:\\Program Files\\Microsoft Office\\Office14\\MSWORD.OLB" no_namespace
// CBrowser ��װ��

class CBrowser : public COleDispatchDriver
{
public:
	CBrowser(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CBrowser(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CBrowser(const CBrowser& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// Browser ����
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
	long get_Target()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Target(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void Next()
	{
		InvokeHelper(0x65, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Previous()
	{
		InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// Browser ����
public:

};
