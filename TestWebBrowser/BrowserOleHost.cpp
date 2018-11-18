#include "stdafx.h"
#include "BrowserOleHost.h"
#include <memory>


HRESULT BrowserOleHost::Create(HWND parent, _Outptr_ BrowserOleHost** ppboh)
{
	if (!ppboh)
	{
		return E_POINTER;
	}
	*ppboh = new (std::nothrow) BrowserOleHost(parent);
	return (*ppboh) ? S_OK : E_OUTOFMEMORY;
}

BrowserOleHost::BrowserOleHost(HWND parent) :
	m_parent(parent),
	m_refcount(1)
{}

bool BrowserOleHost::HandleMessage(LPMSG message)
{
	return m_activeObject && m_activeObject->TranslateAccelerator(message) != S_FALSE;
}

HRESULT STDMETHODCALLTYPE BrowserOleHost::QueryInterface(REFIID riid, LPVOID* ppvObject)
{
	if (riid == IID_IUnknown || riid == IID_IOleClientSite)
	{
		*ppvObject = (LPVOID)this;
	}
	else if (riid == IID_IOleInPlaceFrame)
	{
		*ppvObject = (LPVOID)(IOleInPlaceFrame*)this;
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

ULONG STDMETHODCALLTYPE BrowserOleHost::AddRef()
{
	return InterlockedIncrement(&m_refcount);
}

ULONG STDMETHODCALLTYPE BrowserOleHost::Release()
{
	auto ret = InterlockedDecrement(&m_refcount);
	if (ret == 0)
		delete this;
	return ret;
}

//// IDocHostUIHandler ////

// Called at initialization of the browser object UI. We can set various features of the browser object here.
HRESULT STDMETHODCALLTYPE BrowserOleHost::GetHostInfo(DOCHOSTUIINFO* pInfo)
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

HRESULT STDMETHODCALLTYPE BrowserOleHost::GetWindowContext(LPOLEINPLACEFRAME* lplpFrame, LPOLEINPLACEUIWINDOW* lplpDoc, LPRECT /*lprcPosRect*/, LPRECT /*lprcClipRect*/, LPOLEINPLACEFRAMEINFO lpFrameInfo)
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

	return S_OK;
}

HRESULT STDMETHODCALLTYPE BrowserOleHost::OnPosRectChange(LPCRECT lprcPosRect)
{
	IOleInPlaceObjectPtr inPlaceObject;
	if (SUCCEEDED(m_activeObject.QueryInterface(IID_PPV_ARGS(&inPlaceObject)))) {
		// Confirm the rect we were given.
		inPlaceObject->SetObjectRects(lprcPosRect, lprcPosRect);
	}

	return S_OK;
}



//// IOleInPlaceFrame ////

HRESULT STDMETHODCALLTYPE BrowserOleHost::GetWindow(HWND* lphwnd)
{
	*lphwnd = m_parent;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE BrowserOleHost::SetActiveObject(IOleInPlaceActiveObject* pActiveObject, LPCOLESTR /*pszObjName*/)
{
	m_activeObject = pActiveObject;
	return S_OK;
}
