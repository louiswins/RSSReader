#include "stdafx.h"
#include "Browser.h"
#include "BrowserWindow.h"

LPCWSTR Browser::ClassName = L"Browser";

enum Buttons
{
	BackButton = 1,
	ForwardButton,
	StopButton,
	RefreshButton,
};

static const int ButtonSize = 50;
static const int ButtonMargin = 5;


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

	RECT rc;
	GetClientRect(m_hwnd, &rc);
	auto MakeButton = [&rc, this](LPCWSTR text, Buttons button) {
		CreateWindowEx(0, L"button", text, WS_CHILDWINDOW | WS_TABSTOP | WS_VISIBLE | BS_PUSHBUTTON,
			rc.left + ButtonMargin,
			rc.top + ButtonMargin,
			ButtonSize,
			ButtonSize,
			m_hwnd, (HMENU)button, m_hInst, nullptr);
		rc.top += ButtonSize + ButtonMargin;
	};
	MakeButton(L"<", BackButton);
	MakeButton(L">", ForwardButton);
	MakeButton(L"C", RefreshButton);
	MakeButton(L"#", StopButton);

	return m_browserWindow->ShowWebPage(L"https://bing.com");
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
	case WM_COMMAND:
		OnCommand((int)LOWORD(wParam), (HWND)lParam, (UINT)HIWORD(wParam));
		return 0;
	}

	return DefWindowProc(m_hwnd, uMsg, wParam, lParam);
}

void Browser::OnSize(int cx, int cy)
{
	RECT rc;
	rc.left = ButtonSize + 2*ButtonMargin;
	rc.right = cx;
	rc.top = 0;
	rc.bottom = cy;
	m_browserWindow->SetRect(rc);
}

void Browser::OnCommand(int controlId, HWND hwnd, UINT message)
{
	switch (controlId)
	{
	case BackButton:
		m_browserWindow->DoPageAction(BrowserWindow::BACK);
		return;
	case ForwardButton:
		m_browserWindow->DoPageAction(BrowserWindow::FORWARD);
		return;
	case StopButton:
		m_browserWindow->DoPageAction(BrowserWindow::STOP);
		return;
	case RefreshButton:
		m_browserWindow->DoPageAction(BrowserWindow::REFRESH);
		return;
	}
}
