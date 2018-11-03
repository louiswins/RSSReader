#include "stdafx.h"
#include "BrowserWindow.h"
#include <memory>


void BrowserWindow::DoPageAction(Action action)
{
	switch (action)
	{
	case BACK:
		m_webBrowser2->GoBack();
		break;
	case FORWARD:
		m_webBrowser2->GoForward();
		break;
	case HOME:
		m_webBrowser2->GoHome();
		break;
	case SEARCH:
		m_webBrowser2->GoSearch();
		break;
	case REFRESH:
		m_webBrowser2->Refresh();
		break;
	case STOP:
		m_webBrowser2->Stop();
		break;
	}
}


bool BrowserWindow::ShowWebPage(LPCWSTR webPageName)
{
	// Our URL must be passed to the IWebBrowser2's Navigate2() function as a BSTR. A BSTR is
	// like a pascal version of a double-byte character string. In other words, the first
	// unsigned short is a count of how many characters are in the string, and then this
	// is followed by those characters, each expressed as an unsigned short (rather than a
	// char). The string is not nul-terminated.
	//
	// What's more, our BSTR needs to be stuffed into a VARIANT struct, and that VARIANT struct is
	// then passed to Navigate2(). Why? The VARIANT struct makes it possible to define generic
	// 'datatypes' that can be used with all languages. Not all languages support things like
	// nul-terminated C strings. So, by using a VARIANT, whose first field tells what sort of
	// data (ie, string, float, etc) is in the VARIANT, COM interfaces can be used by just about
	// any language.
	_variant_t url = webPageName;

	// Call the Navigate2() function to actually display the page.
	return SUCCEEDED(m_webBrowser2->Navigate2(&url, nullptr, nullptr, nullptr, nullptr));
}


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

static SafeArrayPtr WrapStringInArrayOfVariant(LPCWSTR string)
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

bool BrowserWindow::ShowHTMLStr(LPCWSTR string)
{
	// Before we can get_Document(), we actually need to have some HTML page loaded in the browser. So,
	// let's load an empty HTML page. Then, once we have that empty page, we can get_Document() and
	// write() to stuff our HTML string into it.
	_variant_t myURL = L"about:blank";

	// Call the Navigate2() function to actually display the page.
	m_webBrowser2->Navigate2(&myURL, nullptr, nullptr, nullptr, nullptr);

	// Call the IWebBrowser2 object's get_Document so we can get its DISPATCH object.
	IDispatchPtr dispatch;
	if (FAILED(m_webBrowser2->get_Document(&dispatch)))
	{
		return false;
	}

	// We want to get a pointer to the IHTMLDocument2 object embedded within the IDispatch
	// object, so we can call some of the functions in the former's table.
	IHTMLDocument2Ptr htmlDoc2;
	if (FAILED(dispatch->QueryInterface(IID_PPV_ARGS(&htmlDoc2))))
	{
		return false;
	}

	// Our HTML must be in a SAFEARRAY of VARIANTs containing BSTRs.
	SafeArrayPtr sfArray = WrapStringInArrayOfVariant(string);
	if (!sfArray)
	{
		return false;
	}

	// Pass the string to write() to write our HTML into the document.
	htmlDoc2->write(sfArray.get());
	// Close the document. If we don't do this, subsequent calls to ShowHTMLStr
	// would append to the current contents of the page.
	htmlDoc2->close();

	return true;
}


void BrowserWindow::SetRect(RECT rc)
{
	m_webBrowser2->put_Left(rc.left);
	m_webBrowser2->put_Width(rc.right - rc.left);
	m_webBrowser2->put_Top(rc.top);
	m_webBrowser2->put_Height(rc.bottom - rc.top);
}


HRESULT BrowserWindow::Create(HWND parent, _In_opt_ LPCOLESTR wzApplicationName, _Outptr_ BrowserWindow** ppbh)
{
	if (!ppbh)
	{
		return E_POINTER;
	}
	*ppbh = nullptr;

	HRESULT hr;

	IOleObjectPtr oleObject;
	{
		IClassFactoryPtr classFactory;
		hr = CoGetClassObject(CLSID_WebBrowser, CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER, nullptr, IID_PPV_ARGS(&classFactory));
		if (FAILED(hr))
		{
			return hr;
		}
		if (!classFactory)
		{
			return E_FAIL;
		}
		hr = classFactory->CreateInstance(nullptr, IID_PPV_ARGS(&oleObject));
		if (FAILED(hr))
		{
			return hr;
		}
	}

	IWebBrowser2Ptr webBrowser2;
	hr = oleObject->QueryInterface(IID_PPV_ARGS(&webBrowser2));
	if (FAILED(hr))
	{
		return hr;
	}

	BrowserWindow* ret = new (std::nothrow) BrowserWindow(parent, webBrowser2);

	IOleClientSitePtr oleClientSite;
	oleClientSite.Attach(static_cast<IOleClientSite*>(ret)); // take ownership

	// Give the browser a pointer to its IOleClientSite
	hr = oleObject->SetClientSite(oleClientSite);
	if (FAILED(hr))
	{
		return hr;
	}

	if (wzApplicationName)
	{
		// We can now call the browser object's SetHostNames function. SetHostNames lets the browser object know our
		// application's name and the name of the document in which we're embedding the browser. (Since we have no
		// document name, we'll pass null for the latter). When the browser object is opened for editing, it displays
		// these names in its titlebar.
		oleObject->SetHostNames(wzApplicationName, nullptr);
	}

	RECT rect;
	GetClientRect(parent, &rect);

	// Let browser object know that it is embedded in an OLE container, set the display area of our browser
	// control the same as our window's size, and actually put the browser object into our window.
	hr = OleSetContainedObject(oleObject, TRUE);
	if (FAILED(hr))
	{
		return hr;
	}
	hr = oleObject->DoVerb(OLEIVERB_SHOW, nullptr, oleClientSite, -1, parent, &rect);
	if (FAILED(hr))
	{
		return hr;
	}

	*ppbh = ret;
	ret->AddRef();
	return S_OK;
}

BrowserWindow::BrowserWindow(HWND parent, IWebBrowser2* browserObj) : m_parent(parent), m_webBrowser2(browserObj), m_refcount(1) {}

void BrowserWindow::ShutDown()
{
	// Thank you https://stackoverflow.com/a/14652605!
	m_webBrowser2->put_Visible(VARIANT_FALSE);
	m_webBrowser2->Stop();
	m_webBrowser2->ExecWB(OLECMDID_CLOSE, OLECMDEXECOPT_DONTPROMPTUSER, 0, 0);

	IOleObjectPtr oleBrowser;
	m_webBrowser2->QueryInterface(IID_PPV_ARGS(&oleBrowser));
	oleBrowser->DoVerb(OLEIVERB_HIDE, nullptr, static_cast<IOleClientSite*>(this), 0, m_parent, nullptr);
	oleBrowser->Close(OLECLOSE_NOSAVE);
	OleSetContainedObject(oleBrowser, FALSE);
	oleBrowser->SetClientSite(nullptr);
	CoDisconnectObject(oleBrowser, 0);
}

HRESULT STDMETHODCALLTYPE BrowserWindow::QueryInterface(REFIID riid, LPVOID* ppvObject)
{
	if (riid == IID_IUnknown || riid == IID_IOleClientSite)
	{
		*ppvObject = (LPVOID)this;
	}
	else if (riid == IID_IOleInPlaceSite)
	{
		*ppvObject = (LPVOID)(IOleInPlaceSite*)this;
	}
	else if (riid == IID_IDocHostUIHandler)
	{
		*ppvObject = (LPVOID)(IDocHostUIHandler*)this;
	}
	else
	{
		*ppvObject = nullptr;
		return E_NOINTERFACE;
	}
	
	AddRef();
	return S_OK;
}

ULONG STDMETHODCALLTYPE BrowserWindow::AddRef()
{
	return InterlockedIncrement(&m_refcount);
}

ULONG STDMETHODCALLTYPE BrowserWindow::Release()
{
	auto ret = InterlockedDecrement(&m_refcount);
	if (ret == 0)
		delete this;
	return ret;
}

//// IDocHostUIHandler ////

// Called at initialization of the browser object UI. We can set various features of the browser object here.
HRESULT STDMETHODCALLTYPE BrowserWindow::GetHostInfo(DOCHOSTUIINFO* pInfo)
{
	pInfo->cbSize = sizeof(DOCHOSTUIINFO);

	// Set some flags. We don't want any 3D border. You can do other things like hide
	// the scroll bar (DOCHOSTUIFLAG_SCROLL_NO), display picture display (DOCHOSTUIFLAG_NOPICS),
	// disable any script running when the page is loaded (DOCHOSTUIFLAG_DISABLE_SCRIPT_INACTIVE),
	// open a site in a new browser window when the user clicks on some link (DOCHOSTUIFLAG_OPENNEWWIN),
	// and lots of other things. See the MSDN docs on the DOCHOSTUIINFO struct passed to us.
	pInfo->dwFlags = DOCHOSTUIFLAG_NO3DBORDER;

	// Set what happens when the user double-clicks on the object. Here we use the default.
	pInfo->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT;

	return S_OK;
}



//// IOleInPlaceSite ////

HRESULT STDMETHODCALLTYPE BrowserWindow::GetWindowContext(LPOLEINPLACEFRAME* lplpFrame, LPOLEINPLACEUIWINDOW* lplpDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
	// Give the browser the pointer to our IOleInPlaceFrame.
	this->QueryInterface(IID_PPV_ARGS(lplpFrame));

	// We have no OLEINPLACEUIWINDOW
	*lplpDoc = nullptr;

	// Fill in some other info for the browser
	lpFrameInfo->fMDIApp = FALSE;
	lpFrameInfo->hwndFrame = m_parent;
	lpFrameInfo->haccel = nullptr;
	lpFrameInfo->cAccelEntries = 0;

	// Give the browser the dimensions of where it can draw. We give it our entire window to fill.
	GetClientRect(m_parent, lprcPosRect);
	GetClientRect(m_parent, lprcClipRect);

	return S_OK;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::OnPosRectChange(LPCRECT lprcPosRect)
{
	IOleInPlaceObjectPtr inplace;
	if (SUCCEEDED(m_webBrowser2->QueryInterface(IID_PPV_ARGS(&inplace))))
	{
		// Confirm the rect we were given.
		inplace->SetObjectRects(lprcPosRect, lprcPosRect);
	}

	return S_OK;
}



//// IOleInPlaceFrame ////

HRESULT STDMETHODCALLTYPE BrowserWindow::GetWindow(HWND* lphwnd)
{
	*lphwnd = m_parent;
	return S_OK;
}
