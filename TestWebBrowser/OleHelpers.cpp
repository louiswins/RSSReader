#include "stdafx.h"
#include "OleHelpers.h"

SafeArrayPtr WrapStringInArrayOfVariant(LPCWSTR string)
{
	SafeArrayPtr sfArray(SafeArrayCreateVector(VT_VARIANT, 0, 1));
	if (!sfArray)
	{
		return nullptr;
	}

	SafeArrayAccessor<VARIANT> data(sfArray.get());
	if (!data)
	{
		return nullptr;
	}

	data[0].vt = VT_BSTR;
	data[0].bstrVal = SysAllocString(string);
	bool success = !!data[0].bstrVal;

	if (!success)
	{
		return nullptr;
	}

	return sfArray;
}