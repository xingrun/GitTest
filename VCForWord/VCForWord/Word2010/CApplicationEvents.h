// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

#import "C:\\Program Files\\Microsoft Office\\Office14\\MSWORD.OLB" no_namespace
// CApplicationEvents ��װ��

class CApplicationEvents : public COleDispatchDriver
{
public:
	CApplicationEvents(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CApplicationEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplicationEvents(const CApplicationEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// ApplicationEvents ����
public:

	// ApplicationEvents ����
public:

};
