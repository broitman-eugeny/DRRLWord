// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CSubdocuments wrapper class

class CSubdocuments : public COleDispatchDriver
{
public:
	CSubdocuments(){} // Calls COleDispatchDriver default constructor
	CSubdocuments(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSubdocuments(const CSubdocuments& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// Subdocuments methods
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
	long get_Count()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	BOOL get_Expanded()
	{
		BOOL result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_Expanded(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL;
		InvokeHelper(0x2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH Item(long Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH AddFromFile(VARIANT * Name, VARIANT * ConfirmConversions, VARIANT * ReadOnly, VARIANT * PasswordDocument, VARIANT * PasswordTemplate, VARIANT * Revert, VARIANT * WritePasswordDocument, VARIANT * WritePasswordTemplate)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
		InvokeHelper(0x64, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Name, ConfirmConversions, ReadOnly, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate);
		return result;
	}
	LPDISPATCH AddFromRange(LPDISPATCH Range)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range);
		return result;
	}
	void Merge(VARIANT * FirstSubdocument, VARIANT * LastSubdocument)
	{
		static BYTE parms[] = VTS_PVARIANT VTS_PVARIANT;
		InvokeHelper(0x66, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FirstSubdocument, LastSubdocument);
	}
	void Delete()
	{
		InvokeHelper(0x67, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Select()
	{
		InvokeHelper(0x68, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// Subdocuments properties
public:

};
