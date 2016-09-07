// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CGlowFormat wrapper class

class CGlowFormat : public COleDispatchDriver
{
public:
	CGlowFormat(){} // Calls COleDispatchDriver default constructor
	CGlowFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CGlowFormat(const CGlowFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// GlowFormat methods
public:
	float get_Radius()
	{
		float result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Radius(float newValue)
	{
		static BYTE parms[] = VTS_R4;
		InvokeHelper(0x1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Color()
	{
		LPDISPATCH result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// GlowFormat properties
public:

};
