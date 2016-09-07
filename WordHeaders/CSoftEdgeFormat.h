// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CSoftEdgeFormat wrapper class

class CSoftEdgeFormat : public COleDispatchDriver
{
public:
	CSoftEdgeFormat(){} // Calls COleDispatchDriver default constructor
	CSoftEdgeFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CSoftEdgeFormat(const CSoftEdgeFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// SoftEdgeFormat methods
public:
	long get_Type()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Type(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// SoftEdgeFormat properties
public:

};
