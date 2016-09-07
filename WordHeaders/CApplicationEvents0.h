// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CApplicationEvents0 wrapper class

class CApplicationEvents0 : public COleDispatchDriver
{
public:
	CApplicationEvents0(){} // Calls COleDispatchDriver default constructor
	CApplicationEvents0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplicationEvents0(const CApplicationEvents0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// IApplicationEvents methods
public:
	void Startup()
	{
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Quit()
	{
		InvokeHelper(0x2, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void DocumentChange()
	{
		InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// IApplicationEvents properties
public:

};
