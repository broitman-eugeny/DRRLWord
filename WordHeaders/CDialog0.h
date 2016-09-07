// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CDialog0 wrapper class

class CDialog0 : public COleDispatchDriver
{
public:
	CDialog0(){} // Calls COleDispatchDriver default constructor
	CDialog0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDialog0(const CDialog0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// Dialog methods
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x7d03, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x7d04, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x7d05, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_DefaultTab()
	{
		long result;
		InvokeHelper(0x7d02, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_DefaultTab(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x7d02, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long Show(VARIANT * TimeOut)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT;
		InvokeHelper(0x150, DISPATCH_METHOD, VT_I4, (void*)&result, parms, TimeOut);
		return result;
	}
	long Display(VARIANT * TimeOut)
	{
		long result;
		static BYTE parms[] = VTS_PVARIANT;
		InvokeHelper(0x152, DISPATCH_METHOD, VT_I4, (void*)&result, parms, TimeOut);
		return result;
	}
	void Execute()
	{
		InvokeHelper(0x7d01, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Update()
	{
		InvokeHelper(0x12e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	CString get_CommandName()
	{
		CString result;
		InvokeHelper(0x7d06, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	long get_CommandBarId()
	{
		long result;
		InvokeHelper(0x7d07, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}

	// Dialog properties
public:

};
