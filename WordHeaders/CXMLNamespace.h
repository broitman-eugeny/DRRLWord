// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CXMLNamespace wrapper class

class CXMLNamespace : public COleDispatchDriver
{
public:
	CXMLNamespace(){} // Calls COleDispatchDriver default constructor
	CXMLNamespace(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CXMLNamespace(const CXMLNamespace& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// XMLNamespace methods
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
	CString get_URI()
	{
		CString result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get_Location(BOOL AllUsers)
	{
		CString result;
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms, AllUsers);
		return result;
	}
	void put_Location(BOOL AllUsers, LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BOOL VTS_BSTR;
		InvokeHelper(0x3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, AllUsers, newValue);
	}
	CString get_Alias(BOOL AllUsers)
	{
		CString result;
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms, AllUsers);
		return result;
	}
	void put_Alias(BOOL AllUsers, LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BOOL VTS_BSTR;
		InvokeHelper(0x4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, AllUsers, newValue);
	}
	LPDISPATCH get_XSLTransforms()
	{
		LPDISPATCH result;
		InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_DefaultTransform(BOOL AllUsers)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, AllUsers);
		return result;
	}
	void put_DefaultTransform(BOOL AllUsers, LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_BOOL VTS_DISPATCH;
		InvokeHelper(0x6, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, AllUsers, newValue);
	}
	void AttachToDocument(VARIANT * Document)
	{
		static BYTE parms[] = VTS_PVARIANT;
		InvokeHelper(0x64, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Document);
	}
	void Delete()
	{
		InvokeHelper(0x65, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// XMLNamespace properties
public:

};
