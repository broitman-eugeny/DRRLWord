// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// COCXEvents wrapper class

class COCXEvents : public COleDispatchDriver
{
public:
	COCXEvents(){} // Calls COleDispatchDriver default constructor
	COCXEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	COCXEvents(const COCXEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// OCXEvents methods
public:
	void GotFocus()
	{
		InvokeHelper(0x800100e0, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void LostFocus()
	{
		InvokeHelper(0x800100e1, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// OCXEvents properties
public:

};
