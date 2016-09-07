// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CShapeNodes wrapper class

class CShapeNodes : public COleDispatchDriver
{
public:
	CShapeNodes(){} // Calls COleDispatchDriver default constructor
	CShapeNodes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CShapeNodes(const CShapeNodes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// ShapeNodes methods
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
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	void Delete(long Index)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Index);
	}
	LPDISPATCH Item(VARIANT * Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	void SetEditingType(long Index, long EditingType)
	{
		static BYTE parms[] = VTS_I4 VTS_I4;
		InvokeHelper(0xd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Index, EditingType);
	}
	void SetPosition(long Index, float X1, float Y1)
	{
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4;
		InvokeHelper(0xe, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Index, X1, Y1);
	}
	void SetSegmentType(long Index, long SegmentType)
	{
		static BYTE parms[] = VTS_I4 VTS_I4;
		InvokeHelper(0xf, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Index, SegmentType);
	}
	void Insert(long Index, long SegmentType, long EditingType, float X1, float Y1, float X2, float Y2, float X3, float Y3)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_R4;
		InvokeHelper(0xc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Index, SegmentType, EditingType, X1, Y1, X2, Y2, X3, Y3);
	}

	// ShapeNodes properties
public:

};
