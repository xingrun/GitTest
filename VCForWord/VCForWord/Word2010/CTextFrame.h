// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

#import "C:\\Program Files\\Microsoft Office\\Office14\\MSWORD.OLB" no_namespace
// CTextFrame ��װ��

class CTextFrame : public COleDispatchDriver
{
public:
	CTextFrame(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CTextFrame(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTextFrame(const CTextFrame& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// TextFrame ����
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x1f40, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x1f41, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	float get_MarginBottom()
	{
		float result;
		InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_MarginBottom(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x64, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_MarginLeft()
	{
		float result;
		InvokeHelper(0x65, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_MarginLeft(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x65, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_MarginRight()
	{
		float result;
		InvokeHelper(0x66, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_MarginRight(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x66, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_MarginTop()
	{
		float result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_MarginTop(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x67, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Orientation()
	{
		long result;
		InvokeHelper(0x68, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Orientation(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x68, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_TextRange()
	{
		LPDISPATCH result;
		InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ContainingRange()
	{
		LPDISPATCH result;
		InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Next()
	{
		LPDISPATCH result;
		InvokeHelper(0x1389, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void put_Next(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x1389, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Previous()
	{
		LPDISPATCH result;
		InvokeHelper(0x138a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void put_Previous(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x138a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_Overflowing()
	{
		BOOL result;
		InvokeHelper(0x138b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	long get_HasText()
	{
		long result;
		InvokeHelper(0x1390, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void BreakForwardLink()
	{
		InvokeHelper(0x138c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	BOOL ValidLinkTarget(LPDISPATCH TargetTextFrame)
	{
		BOOL result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x138e, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, TargetTextFrame);
		return result;
	}
	long get_AutoSize()
	{
		long result;
		InvokeHelper(0x1391, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_AutoSize(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1391, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_WordWrap()
	{
		long result;
		InvokeHelper(0x1392, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_WordWrap(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1392, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_VerticalAnchor()
	{
		long result;
		InvokeHelper(0x1393, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_VerticalAnchor(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1393, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_HorizontalAnchor()
	{
		long result;
		InvokeHelper(0x1394, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_HorizontalAnchor(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1394, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_PathFormat()
	{
		long result;
		InvokeHelper(0x1395, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_PathFormat(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1395, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_WarpFormat()
	{
		long result;
		InvokeHelper(0x1396, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_WarpFormat(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1396, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Column()
	{
		LPDISPATCH result;
		InvokeHelper(0x1397, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ThreeD()
	{
		LPDISPATCH result;
		InvokeHelper(0x1398, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_NoTextRotation()
	{
		long result;
		InvokeHelper(0x1399, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_NoTextRotation(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x1399, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void DeleteText()
	{
		InvokeHelper(0x139a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// TextFrame ����
public:

};
