// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

#import "C:\\Program Files\\Microsoft Office\\Office14\\MSWORD.OLB" no_namespace
// CSmartTags ��װ��

class CSmartTags : public COleDispatchDriver
{
public:
	CSmartTags(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CSmartTags(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSmartTags(const CSmartTags& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// SmartTags ����
public:
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
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
	LPDISPATCH Item(VARIANT * Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH Add(LPCTSTR Name, VARIANT * Range, VARIANT * Properties)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x5, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, Range, Properties);
		return result;
	}
	LPDISPATCH SmartTagsByType(LPCTSTR Name)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x3eb, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name);
		return result;
	}

	// SmartTags ����
public:

};
