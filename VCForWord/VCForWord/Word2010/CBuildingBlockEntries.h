// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

//#import "C:\\Program Files\\Microsoft Office\\Office14\\MSWORD.OLB" no_namespace
// CBuildingBlockEntries ��װ��

class CBuildingBlockEntries : public COleDispatchDriver
{
public:
	CBuildingBlockEntries(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CBuildingBlockEntries(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CBuildingBlockEntries(const CBuildingBlockEntries& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// BuildingBlockEntries ����
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
	long get_Count()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(VARIANT * Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH Add(LPCTSTR Name, long Type, LPCTSTR Category, LPDISPATCH Range, VARIANT * Description, long InsertOptions)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_I4 VTS_BSTR VTS_DISPATCH VTS_PVARIANT VTS_I4 ;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, Type, Category, Range, Description, InsertOptions);
		return result;
	}

	// BuildingBlockEntries ����
public:

};
