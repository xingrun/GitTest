// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

#import "C:\\Program Files\\Microsoft Office\\Office14\\MSWORD.OLB" no_namespace
// CReadabilityStatistic ��װ��

class CReadabilityStatistic : public COleDispatchDriver
{
public:
	CReadabilityStatistic(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CReadabilityStatistic(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CReadabilityStatistic(const CReadabilityStatistic& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// ReadabilityStatistic ����
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
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	float get_Value()
	{
		float result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}

	// ReadabilityStatistic ����
public:

};
