// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CTablesOfFigures wrapper class

class CTablesOfFigures : public COleDispatchDriver
{
public:
	CTablesOfFigures(){} // Calls COleDispatchDriver default constructor
	CTablesOfFigures(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CTablesOfFigures(const CTablesOfFigures& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// TablesOfFigures methods
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
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Format()
	{
		long result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Format(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x2, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH Item(long Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH AddOld(LPDISPATCH Range, VARIANT * Caption, VARIANT * IncludeLabel, VARIANT * UseHeadingStyles, VARIANT * UpperHeadingLevel, VARIANT * LowerHeadingLevel, VARIANT * UseFields, VARIANT * TableID, VARIANT * RightAlignPageNumbers, VARIANT * IncludePageNumbers, VARIANT * AddedStyles)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
		InvokeHelper(0x64, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, Caption, IncludeLabel, UseHeadingStyles, UpperHeadingLevel, LowerHeadingLevel, UseFields, TableID, RightAlignPageNumbers, IncludePageNumbers, AddedStyles);
		return result;
	}
	LPDISPATCH MarkEntry(LPDISPATCH Range, VARIANT * Entry, VARIANT * EntryAutoText, VARIANT * TableID, VARIANT * Level)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
		InvokeHelper(0x65, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, Entry, EntryAutoText, TableID, Level);
		return result;
	}
	LPDISPATCH Add(LPDISPATCH Range, VARIANT * Caption, VARIANT * IncludeLabel, VARIANT * UseHeadingStyles, VARIANT * UpperHeadingLevel, VARIANT * LowerHeadingLevel, VARIANT * UseFields, VARIANT * TableID, VARIANT * RightAlignPageNumbers, VARIANT * IncludePageNumbers, VARIANT * AddedStyles, VARIANT * UseHyperlinks, VARIANT * HidePageNumbersInWeb)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT;
		InvokeHelper(0x1bc, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Range, Caption, IncludeLabel, UseHeadingStyles, UpperHeadingLevel, LowerHeadingLevel, UseFields, TableID, RightAlignPageNumbers, IncludePageNumbers, AddedStyles, UseHyperlinks, HidePageNumbersInWeb);
		return result;
	}

	// TablesOfFigures properties
public:

};
