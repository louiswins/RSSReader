#pragma once
#include <memory>
#include <Windows.h>

class BrowserWindow;

// Browser is loosely based on Raymond Chen's C++ scratch program:
// https://blogs.msdn.microsoft.com/oldnewthing/20050422-08/?p=35813/
class Browser
{
public:
	static Browser* Create(HINSTANCE hInst, LPCWSTR wzWindowTitle);

	void Show(int nCmdShow);

	~Browser();
private:
	void OnSize(int cx, int cy);
	void OnCommand(int controlId, HWND hwnd, UINT action);

	Browser(HINSTANCE hInst);
	bool InitControls();
	HWND DoCreateWindow(LPCWSTR wzWindowTitle);
	void DoRegisterClass();
	LRESULT WndProc(UINT uMsg, WPARAM wParam, LPARAM lParam);
	static LRESULT CALLBACK s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	static LPCWSTR ClassName;

	HINSTANCE m_hInst;
	HWND m_hwnd;
	BrowserWindow* m_browserWindow;
	struct FontDeleter {
		void operator()(HFONT f) { if (f) { DeleteObject(f); } }
	};
	std::unique_ptr<HFONT__, FontDeleter> m_guiFont;
};
