// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

#import "C:\\Program Files\\Microsoft Office\\Office14\\MSWORD.OLB" no_namespace
// COMathAutoCorrectEntries ��װ��

class COMathAutoCorrectEntries : public COleDispatchDriver
{
public:
	COMathAutoCorrectEntries(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	COMathAutoCorrectEntries(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COMathAutoCorrectEntries(const COMathAutoCorrectEntries& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// OMathAutoCorrectEntries ����
public:
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
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(VARIANT * Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH Add(LPCTSTR Name, LPCTSTR Value)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR ;
		InvokeHelper(0xc8, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, Value);
		return result;
	}

	// OMathAutoCorrectEntries ����
public:

};
