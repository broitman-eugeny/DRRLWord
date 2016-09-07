// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CResearch wrapper class

class CResearch : public COleDispatchDriver
{
public:
	CResearch(){} // Calls COleDispatchDriver default constructor
	CResearch(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CResearch(const CResearch& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// Research methods
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
	VARIANT Query(LPCTSTR ServiceID, LPCTSTR QueryString, long QueryLanguage, BOOL UseSelection, BOOL LaunchQuery)
	{
		VARIANT result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR VTS_I4 VTS_BOOL VTS_BOOL;
		InvokeHelper(0x1f4, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, ServiceID, QueryString, QueryLanguage, UseSelection, LaunchQuery);
		return result;
	}
	VARIANT SetLanguagePair(long LanguageFrom, long LanguageTo)
	{
		VARIANT result;
		static BYTE parms[] = VTS_I4 VTS_I4;
		InvokeHelper(0x1f5, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, LanguageFrom, LanguageTo);
		return result;
	}
	BOOL IsResearchService(LPCTSTR ServiceID)
	{
		BOOL result;
		static BYTE parms[] = VTS_BSTR;
		InvokeHelper(0x1f6, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, ServiceID);
		return result;
	}

	// Research properties
public:

};
