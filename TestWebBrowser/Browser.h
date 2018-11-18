#pragma once
#include <memory>
#include <Windows.h>
#include "OleHelpers.h"

class BrowserOleHost;

// Browser is loosely based on Raymond Chen's C++ scratch program:
// https://blogs.msdn.microsoft.com/oldnewthing/20050422-08/?p=35813/
class Browser
{
public:
	static std::unique_ptr<Browser> Create(HINSTANCE hInst, LPCWSTR wzWindowTitle);

	void ShowWindow(int nCmdShow);
	bool HandleMessage(LPMSG message);

	bool ShowWebPage(LPCWSTR webPageName);
	bool ShowHTMLStr(LPCWSTR string);

	~Browser();
private:
	void OnSize(int cx, int cy);

	Browser(HINSTANCE hInst);
	void InitControls();
	HWND DoCreateWindow(LPCWSTR wzWindowTitle);
	void DoRegisterClass();
	LRESULT WndProc(UINT uMsg, WPARAM wParam, LPARAM lParam);
	static LRESULT CALLBACK s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	static LPCWSTR ClassName;

	HINSTANCE m_hInst;
	HWND m_hwnd;
	GenericComPtr<BrowserOleHost> m_browserOleHost;
	IWebBrowser2Ptr m_webBrowser2;
	struct FontDeleter {
		void operator()(HFONT f) { if (f) { DeleteObject(f); } }
	};
	std::unique_ptr<HFONT__, FontDeleter> m_guiFont;
};
