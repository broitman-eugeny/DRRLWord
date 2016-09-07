// Machine generated IDispatch wrapper class(es) created with Add Class from Typelib Wizard

#import "C:\\Program Files (x86)\\Microsoft Office\\Office12\\MSWORD.OLB" no_namespace
// CPictureFormat wrapper class

class CPictureFormat : public COleDispatchDriver
{
public:
	CPictureFormat(){} // Calls COleDispatchDriver default constructor
	CPictureFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CPictureFormat(const CPictureFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// Attributes
public:

	// Operations
public:


	// PictureFormat methods
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
	float get_Brightness()
	{
		float result;
		InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Brightness(float newValue)
	{
		static BYTE parms[] = VTS_R4;
		InvokeHelper(0x64, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_ColorType()
	{
		long result;
		InvokeHelper(0x65, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ColorType(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x65, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_Contrast()
	{
		float result;
		InvokeHelper(0x66, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Contrast(float newValue)
	{
		static BYTE parms[] = VTS_R4;
		InvokeHelper(0x66, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_CropBottom()
	{
		float result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_CropBottom(float newValue)
	{
		static BYTE parms[] = VTS_R4;
		InvokeHelper(0x67, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_CropLeft()
	{
		float result;
		InvokeHelper(0x68, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_CropLeft(float newValue)
	{
		static BYTE parms[] = VTS_R4;
		InvokeHelper(0x68, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_CropRight()
	{
		float result;
		InvokeHelper(0x69, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_CropRight(float newValue)
	{
		static BYTE parms[] = VTS_R4;
		InvokeHelper(0x69, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	float get_CropTop()
	{
		float result;
		InvokeHelper(0x6a, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_CropTop(float newValue)
	{
		static BYTE parms[] = VTS_R4;
		InvokeHelper(0x6a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_TransparencyColor()
	{
		long result;
		InvokeHelper(0x6b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_TransparencyColor(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x6b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_TransparentBackground()
	{
		long result;
		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_TransparentBackground(long newValue)
	{
		static BYTE parms[] = VTS_I4;
		InvokeHelper(0x6c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void IncrementBrightness(float Increment)
	{
		static BYTE parms[] = VTS_R4;
		InvokeHelper(0xa, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Increment);
	}
	void IncrementContrast(float Increment)
	{
		static BYTE parms[] = VTS_R4;
		InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Increment);
	}

	// PictureFormat properties
public:

};
