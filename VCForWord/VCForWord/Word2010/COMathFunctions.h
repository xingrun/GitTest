// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

#import "C:\\Program Files\\Microsoft Office\\Office14\\MSWORD.OLB" no_namespace
// COMathFunctions ��װ��

class COMathFunctions : public COleDispatchDriver
{
public:
	COMathFunctions(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	COMathFunctions(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COMathFunctions(const COMathFunctions& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// OMathFunctions ����
public:
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x65, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x66, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(long Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH Add(LPDISPATCH Range, long Type, VARIANT * NumArgs, VARIANT * NumCols)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_I4 VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x68, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, Type, NumArgs, NumCols);
		return result;
	}

	// OMathFunctions ����
public:

};
