#include "stdafx.h"
#include "Browser.h"
#include "BrowserWindow.h"

LPCWSTR Browser::ClassName = L"Browser";

static const int TreeViewWidth = 150;


Browser* Browser::Create(HINSTANCE hInst, LPCWSTR wzWindowTitle)
{
	auto self = new (std::nothrow) Browser(hInst);
	if (self && self->DoCreateWindow(wzWindowTitle))
	{
		return self;
	}
	delete self;
	return nullptr;
}

Browser::Browser(HINSTANCE hInst) : m_hInst(hInst) {}

void Browser::Show(int nCmdShow)
{
	ShowWindow(m_hwnd, nCmdShow);
}

Browser::~Browser()
{
	if (m_browserWindow)
	{
		m_browserWindow->ShutDown();
		m_browserWindow->Release();
	}
}

bool Browser::InitControls()
{
	HRESULT hr = BrowserWindow::Create(m_hwnd, L"Test Browser", &m_browserWindow);
	if (FAILED(hr))
	{
		return false;
	}

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
	int indent = TreeView_GetIndent(tv);

	return m_browserWindow->ShowWebPage(L"res://TestWebBrowser.exe/about.html");
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
	LRESULT lres;

	switch (uMsg) {
	case WM_CREATE:
		return InitControls() ? 0 : -1;
	case WM_NCDESTROY:
		lres = DefWindowProc(m_hwnd, uMsg, wParam, lParam);
		SetWindowLongPtr(m_hwnd, GWLP_USERDATA, 0);
		delete this;
		PostQuitMessage(0);
		return lres;
	case WM_SIZE:
		OnSize(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
		return 0;
	}

	return DefWindowProc(m_hwnd, uMsg, wParam, lParam);
}

void Browser::OnSize(int cx, int cy)
{
	RECT rc;
	rc.left = TreeViewWidth;
	rc.right = cx;
	rc.top = 0;
	rc.bottom = cy;
	m_browserWindow->SetRect(rc);
}
