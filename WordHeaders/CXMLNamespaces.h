// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CXMLNamespaces wrapper class

class CXMLNamespaces : public COleDispatchDriver
{
public:
	CXMLNamespaces(){} // Calls COleDispatchDriver default constructor
	CXMLNamespaces(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CXMLNamespaces(const CXMLNamespaces& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// XMLNamespaces methods
public:
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
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
	LPDISPATCH Item(VARIANT * Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH Add(LPCTSTR Path, VARIANT * NamespaceURI, VARIANT * Alias, BOOL InstallForAllUsers)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_BOOL;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Path, NamespaceURI, Alias, InstallForAllUsers);
		return result;
	}
	void InstallManifest(LPCTSTR Path, BOOL InstallForAllUsers)
	{
		static BYTE parms[] = VTS_BSTR VTS_BOOL;
		InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Path, InstallForAllUsers);
	}

	// XMLNamespaces properties
public:

};
