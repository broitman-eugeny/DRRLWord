// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CEditor wrapper class

class CEditor : public COleDispatchDriver
{
public:
	CEditor(){} // Calls COleDispatchDriver default constructor
	CEditor(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CEditor(const CEditor& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// Editor methods
public:
	CString get_ID()
	{
		CString result;
		InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x65, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Range()
	{
		LPDISPATCH result;
		InvokeHelper(0x66, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_NextRange()
	{
		LPDISPATCH result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
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
	void Delete()
	{
		InvokeHelper(0x1f4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void DeleteAll()
	{
		InvokeHelper(0x1f5, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void SelectAll()
	{
		InvokeHelper(0x1f6, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// Editor properties
public:

};
