#pragma once
#include <memory>
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

template <typename T>
struct ComDeleter {
	void operator()(T* p) noexcept { if (p) { p->Release(); } }
};
template <typename T>
using GenericComPtr = std::unique_ptr<T, ComDeleter<T>>;

struct SafeArrayDeleter {
	void operator()(SAFEARRAY* sa) noexcept { if (sa) { SafeArrayDestroy(sa); } }
};
using SafeArrayPtr = std::unique_ptr<SAFEARRAY, SafeArrayDeleter>;
template <typename Element>
struct SafeArrayAccessor
{
	SafeArrayAccessor(SAFEARRAY* sa) noexcept : sa(sa)
	{
		if (FAILED(SafeArrayAccessData(sa, reinterpret_cast<void**>(&data))))
		{
			data = nullptr;
		}
	}
	~SafeArrayAccessor() noexcept
	{
		if (data) { SafeArrayUnaccessData(sa); }
	}
	explicit operator bool() const noexcept { return data; }
	Element* get() const noexcept { return data; }
	Element& operator[](int index) const noexcept { return data[index]; }
private:
	SAFEARRAY* sa;
	Element* data;
};

SafeArrayPtr WrapStringInArrayOfVariant(LPCWSTR string);
