// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CChartData wrapper class

class CChartData : public COleDispatchDriver
{
public:
	CChartData(){} // Calls COleDispatchDriver default constructor
	CChartData(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CChartData(const CChartData& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// ChartData methods
public:
	LPDISPATCH get_Workbook()
	{
		LPDISPATCH result;
		InvokeHelper(0x60020000, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void Activate()
	{
		InvokeHelper(0x60020001, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	BOOL get_IsLinked()
	{
		BOOL result;
		InvokeHelper(0x60020002, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void BreakLink()
	{
		InvokeHelper(0x60020003, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// ChartData properties
public:

};
