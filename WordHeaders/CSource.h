// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CSource wrapper class

class CSource : public COleDispatchDriver
{
public:
	CSource(){} // Calls COleDispatchDriver default constructor
	CSource(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSource(const CSource& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// Source methods
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
	CString get_Tag()
	{
		CString result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get_Field(LPCTSTR Name)
	{
		CString result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x68, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms, Name);
		return result;
	}
	void put_Field(LPCTSTR Name, LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR VTS_BSTR;
		InvokeHelper(0x68, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, Name, newValue);
	}
	CString get_XML()
	{
		CString result;
		InvokeHelper(0x69, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	BOOL get_Cited()
	{
		BOOL result;
		InvokeHelper(0x6b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void Delete()
	{
		InvokeHelper(0x6a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// Source properties
public:

};
