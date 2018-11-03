#include "stdafx.h"
#include "Browser.h"
#include "BrowserWindow.h"

LPCWSTR Browser::ClassName = L"Browser";


Browser* Browser::Create(HINSTANCE hInst, LPCWSTR wzWindowTitle)
{
	auto self = new (std::nothrow) Browser();
	if (self && self->DoCreateWindow(hInst, wzWindowTitle))
	{
		return self;
	}
	delete self;
	return nullptr;
}

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
	return m_browserWindow->ShowWebPage(L"https://bing.com");
}

void Browser::DoRegisterClass(HINSTANCE hInst)
{
	WNDCLASS wc;
	wc.style = 0;
	wc.lpfnWndProc = Browser::s_WndProc;
	wc.cbClsExtra = 0;
	wc.cbWndExtra = 0;
	wc.hInstance = hInst;
	wc.hIcon = NULL;
	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
	wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
	wc.lpszMenuName = NULL;
	wc.lpszClassName = ClassName;
	RegisterClass(&wc);
}

HWND Browser::DoCreateWindow(HINSTANCE hInst, LPCWSTR wzWindowTitle)
{
	DoRegisterClass(hInst);
	return CreateWindowEx(0, ClassName, wzWindowTitle, WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, NULL, NULL, hInst, this);
}

LRESULT CALLBACK Browser::s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	Browser* self;
	if (uMsg == WM_NCCREATE) {
		LPCREATESTRUCT lpcs = reinterpret_cast<LPCREATESTRUCT>(lParam);
		self = reinterpret_cast<Browser*>(lpcs->lpCreateParams);
		self->m_hwnd = hwnd;
		if (!self->InitControls())
		{
			return FALSE;
		}
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
	rc.left = 100;
	rc.right = cx;
	rc.top = 0;
	rc.bottom = cy;
	m_browserWindow->SetRect(rc);
}