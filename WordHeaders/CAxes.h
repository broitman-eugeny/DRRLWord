// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CAxes wrapper class

class CAxes : public COleDispatchDriver
{
public:
	CAxes(){} // Calls COleDispatchDriver default constructor
	CAxes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAxes(const CAxes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// Axes methods
public:
	long get_Count()
	{
		long result;
		InvokeHelper(0x60020000, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(long Type, long AxisGroup)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_I4;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Type, AxisGroup);
		return result;
	}
	LPUNKNOWN _NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_METHOD, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH _Default(long Type, long AxisGroup)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_I4;
		InvokeHelper(0x60020005, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Type, AxisGroup);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// Axes properties
public:

};
