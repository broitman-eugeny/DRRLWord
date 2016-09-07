// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

//#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CListGallery wrapper class

class CListGallery : public COleDispatchDriver
{
public:
	CListGallery(){} // Calls COleDispatchDriver default constructor
	CListGallery(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CListGallery(const CListGallery& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// ListGallery methods
public:
	LPDISPATCH get_ListTemplates()
	{
		LPDISPATCH result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
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
	BOOL get_Modified(long Index)
	{
		BOOL result;
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x65, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, parms, Index);
		return result;
	}
	void Reset(long Index)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x64, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Index);
	}

	// ListGallery properties
public:

};
