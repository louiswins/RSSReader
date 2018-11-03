#pragma once
#include <Ole2.h>

struct OleInitializer
{
	OleInitializer() noexcept : hr(OleInitialize(NULL)) {}
	~OleInitializer() noexcept
	{
		if (SUCCEEDED(hr))
		{
			OleUninitialize();
		}
	}
	operator HRESULT() const noexcept { return hr; }
	HRESULT hr;
};