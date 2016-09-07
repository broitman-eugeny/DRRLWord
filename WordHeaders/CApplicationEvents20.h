// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CApplicationEvents20 wrapper class

class CApplicationEvents20 : public COleDispatchDriver
{
public:
	CApplicationEvents20(){} // Calls COleDispatchDriver default constructor
	CApplicationEvents20(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplicationEvents20(const CApplicationEvents20& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// IApplicationEvents2 methods
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
	void DocumentOpen(LPDISPATCH Doc)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc);
	}
	void DocumentBeforeClose(LPDISPATCH Doc, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, Cancel);
	}
	void DocumentBeforePrint(LPDISPATCH Doc, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0x7, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, Cancel);
	}
	void DocumentBeforeSave(LPDISPATCH Doc, BOOL * SaveAsUI, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL VTS_PBOOL;
		InvokeHelper(0x8, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, SaveAsUI, Cancel);
	}
	void NewDocument(LPDISPATCH Doc)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0x9, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc);
	}
	void WindowActivate(LPDISPATCH Doc, LPDISPATCH Wn)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH;
		InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, Wn);
	}
	void WindowDeactivate(LPDISPATCH Doc, LPDISPATCH Wn)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_DISPATCH;
		InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Doc, Wn);
	}
	void WindowSelectionChange(LPDISPATCH Sel)
	{
		static BYTE parms[] = VTS_DISPATCH;
		InvokeHelper(0xc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sel);
	}
	void WindowBeforeRightClick(LPDISPATCH Sel, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sel, Cancel);
	}
	void WindowBeforeDoubleClick(LPDISPATCH Sel, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL;
		InvokeHelper(0xe, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Sel, Cancel);
	}

	// IApplicationEvents2 properties
public:

};
