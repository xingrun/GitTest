// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

#import "C:\\Program Files\\Microsoft Office\\Office14\\MSWORD.OLB" no_namespace
// CTwoInitialCapsException ��װ��

class CTwoInitialCapsException : public COleDispatchDriver
{
public:
	CTwoInitialCapsException(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CTwoInitialCapsException(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTwoInitialCapsException(const CTwoInitialCapsException& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// TwoInitialCapsException ����
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
	long get_Index()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void Delete()
	{
		InvokeHelper(0x65, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// TwoInitialCapsException ����
public:

};
