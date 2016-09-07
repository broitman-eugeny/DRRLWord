// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CDocumentEvents wrapper class

class CDocumentEvents : public COleDispatchDriver
{
public:
	CDocumentEvents(){} // Calls COleDispatchDriver default constructor
	CDocumentEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDocumentEvents(const CDocumentEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// DocumentEvents methods
public:
	void New()
	{
		InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Open()
	{
		InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Close()
	{
		InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// DocumentEvents properties
public:

};
