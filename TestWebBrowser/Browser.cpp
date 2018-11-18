#include "stdafx.h"
#include "Browser.h"
#include "BrowserOleHost.h"
#include "OleHelpers.h"

LPCWSTR Browser::ClassName = L"Browser";

static const int TreeViewWidth = 150;

#define FAIL_IF_NOT(action) do { HRESULT hr = (action); if (FAILED(hr)) { return nullptr; } } while(0)
std::unique_ptr<Browser> Browser::Create(HINSTANCE hInst, LPCWSTR wzWindowTitle)
{
	std::unique_ptr<Browser> self(new (std::nothrow) Browser(hInst));
	if (!self || !self->DoCreateWindow(wzWindowTitle))
	{
		return nullptr;
	}
	{
		BrowserOleHost* window;
		FAIL_IF_NOT(BrowserOleHost::Create(self->m_hwnd, &window));
		self->m_browserOleHost.reset(window); // transfer ownership, window var is about to go out of scope
	}

	IOleObjectPtr browserOleObj;
	{
		IClassFactoryPtr classFactory;
		if (FAILED(CoGetClassObject(CLSID_WebBrowser, CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER, nullptr, IID_PPV_ARGS(&classFactory))) || !classFactory)
		{
			return nullptr;
		}
		FAIL_IF_NOT(classFactory->CreateInstance(nullptr, IID_PPV_ARGS(&browserOleObj)));
	}

	FAIL_IF_NOT(browserOleObj->QueryInterface(IID_PPV_ARGS(&self->m_webBrowser2)));

	// Give the browser a pointer to its IOleClientSite
	FAIL_IF_NOT(browserOleObj->SetClientSite(static_cast<IOleClientSite*>(self->m_browserOleHost.get())));

	if (wzWindowTitle)
	{
		// We can now call the browser object's SetHostNames function. SetHostNames lets the browser object know our
		// application's name and the name of the document in which we're embedding the browser. (Since we have no
		// document name, we'll pass null for the latter). When the browser object is opened for editing, it displays
		// these names in its titlebar.
		browserOleObj->SetHostNames(wzWindowTitle, nullptr);
	}

	// Let browser object know that it is embedded in an OLE container, set the display area of our browser
	// control the same as our window's size, and actually put the browser object into our window.
	FAIL_IF_NOT(OleSetContainedObject(browserOleObj, TRUE));
	FAIL_IF_NOT(browserOleObj->DoVerb(OLEIVERB_SHOW, nullptr, static_cast<IOleClientSite*>(self->m_browserOleHost.get()), 0, self->m_hwnd, nullptr));

	return self;
}
#undef FAIL_IF_NOT

Browser::Browser(HINSTANCE hInst) : m_hInst(hInst) {}

void Browser::ShowWindow(int nCmdShow)
{
	ShowWebPage(L"res://TestWebBrowser.exe/about.html");
	::ShowWindow(m_hwnd, nCmdShow);
}

bool Browser::HandleMessage(LPMSG message)
{
	return m_browserOleHost->HandleMessage(message);
}

Browser::~Browser()
{
	if (m_webBrowser2) {
		// Thank you https://stackoverflow.com/a/14652605!
		m_webBrowser2->put_Visible(VARIANT_FALSE);
		m_webBrowser2->Stop();
		m_webBrowser2->ExecWB(OLECMDID_CLOSE, OLECMDEXECOPT_DONTPROMPTUSER, 0, 0);

		IOleObjectPtr oleBrowser;
		m_webBrowser2->QueryInterface(IID_PPV_ARGS(&oleBrowser));
		oleBrowser->DoVerb(OLEIVERB_HIDE, nullptr, static_cast<IOleClientSite*>(m_browserOleHost.get()), 0, m_hwnd, nullptr);
		oleBrowser->Close(OLECLOSE_NOSAVE);
		OleSetContainedObject(oleBrowser, FALSE);
		oleBrowser->SetClientSite(nullptr);
		(void)CoDisconnectObject(oleBrowser, 0);
	}
}

void Browser::InitControls()
{
	NONCLIENTMETRICS metrics;
	metrics.cbSize = sizeof(metrics);
	SystemParametersInfo(SPI_GETNONCLIENTMETRICS, metrics.cbSize, &metrics, 0);
	m_guiFont.reset(CreateFontIndirect(&metrics.lfCaptionFont));

	RECT rc;
	GetClientRect(m_hwnd, &rc);

	HWND tv = CreateWindowEx(0, WC_TREEVIEW, nullptr,
		WS_CHILD | WS_VISIBLE | TVS_FULLROWSELECT,
		rc.left,
		rc.top,
		TreeViewWidth,
		(rc.bottom - rc.top),
		m_hwnd, nullptr, m_hInst, nullptr);

	// Load it up with some fake elements
	WCHAR name[128];
	TVINSERTSTRUCT tvis;
	tvis.hInsertAfter = TVI_LAST;
	tvis.item.mask = TVIF_TEXT | TVIF_STATE;
	tvis.item.pszText = name;
	tvis.item.stateMask = TVIS_EXPANDED | TVIS_BOLD;
	tvis.item.state = TVIS_EXPANDED | TVIS_BOLD;

	tvis.hParent = nullptr;
	wcscpy(name, L"The Old New Thing");
	HTREEITEM rootnode = TreeView_InsertItem(tv, &tvis);

	tvis.hParent = rootnode;
	tvis.item.state &= ~TVIS_BOLD;
	wcscpy(name, L"Entry #1");
	TreeView_InsertItem(tv, &tvis);
	wcscpy(name, L"Entry #2");
	TreeView_InsertItem(tv, &tvis);
	wcscpy(name, L"Entry #3");
	TreeView_InsertItem(tv, &tvis);

	tvis.hParent = nullptr;
	tvis.item.state |= TVIS_BOLD;
	wcscpy(name, L"xkcd What If?");
	rootnode = TreeView_InsertItem(tv, &tvis);

	tvis.hParent = rootnode;
	tvis.item.state &= ~TVIS_BOLD;
	wcscpy(name, L"Entry #1");
	TreeView_InsertItem(tv, &tvis);
	wcscpy(name, L"Entry #2");
	TreeView_InsertItem(tv, &tvis);
	wcscpy(name, L"Entry #3");
	TreeView_InsertItem(tv, &tvis);
	wcscpy(name, L"Entry #4");
	TreeView_InsertItem(tv, &tvis);

	TreeView_SetIndent(tv, 5);
}

void Browser::DoRegisterClass()
{
	WNDCLASS wc;
	wc.style = 0;
	wc.lpfnWndProc = Browser::s_WndProc;
	wc.cbClsExtra = 0;
	wc.cbWndExtra = 0;
	wc.hInstance = m_hInst;
	wc.hIcon = NULL;
	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
	wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
	wc.lpszMenuName = NULL;
	wc.lpszClassName = ClassName;
	RegisterClass(&wc);
}

HWND Browser::DoCreateWindow(LPCWSTR wzWindowTitle)
{
	DoRegisterClass();
	return CreateWindowEx(0, ClassName, wzWindowTitle, WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, m_hInst, this);
}

LRESULT CALLBACK Browser::s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	Browser* self;
	if (uMsg == WM_NCCREATE) {
		LPCREATESTRUCT lpcs = reinterpret_cast<LPCREATESTRUCT>(lParam);
		self = reinterpret_cast<Browser*>(lpcs->lpCreateParams);
		self->m_hwnd = hwnd;
		SetWindowLongPtr(hwnd, GWLP_USERDATA, reinterpret_cast<LPARAM>(self));
	} else {
		self = reinterpret_cast<Browser*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));
	}
	if (self) {
		return self->WndProc(uMsg, wParam, lParam);
	} else {
		return DefWindowProc(hwnd, uMsg, wParam, lParam);
	}
}

LRESULT Browser::WndProc(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg) {
	case WM_CREATE:
		InitControls();
		return 0;
	case WM_NCDESTROY:
	{
		LRESULT lres = DefWindowProc(m_hwnd, uMsg, wParam, lParam);
		SetWindowLongPtr(m_hwnd, GWLP_USERDATA, 0);
		PostQuitMessage(0);
		return lres;
	}
	case WM_SIZE:
		OnSize(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
		return 0;
	}

	return DefWindowProc(m_hwnd, uMsg, wParam, lParam);
}

void Browser::OnSize(int cx, int cy)
{
	m_webBrowser2->put_Left(TreeViewWidth);
	m_webBrowser2->put_Width(cx - TreeViewWidth);
	m_webBrowser2->put_Top(0);
	m_webBrowser2->put_Height(cy);
}

bool Browser::ShowWebPage(LPCWSTR webPageName)
{
	_variant_t url = webPageName;
	return SUCCEEDED(m_webBrowser2->Navigate2(&url, nullptr, nullptr, nullptr, nullptr));
}

bool Browser::ShowHTMLStr(LPCWSTR string)
{
	// Before we can get_Document(), we actually need to have some HTML page loaded in the browser. So,
	// let's load an empty HTML page. Then, once we have that empty page, we can get_Document() and
	// write() to stuff our HTML string into it.
	ShowWebPage(L"about:blank");

	IDispatchPtr dispatch;
	if (FAILED(m_webBrowser2->get_Document(&dispatch)))
	{
		return false;
	}
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

	htmlDoc2->write(sfArray.get());
	// Close the document. If we don't do this, subsequent calls to ShowHTMLStr
	// would append to the current contents of the page.
	htmlDoc2->close();

	return true;
}
