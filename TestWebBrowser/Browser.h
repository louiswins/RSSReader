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

	bool InitControls();
	HWND DoCreateWindow(HINSTANCE hInst, LPCWSTR wzWindowTitle);
	void DoRegisterClass(HINSTANCE hInst);
	LRESULT WndProc(UINT uMsg, WPARAM wParam, LPARAM lParam);
	static LRESULT CALLBACK s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	static LPCWSTR ClassName;

	HWND m_hwnd;
	BrowserWindow* m_browserWindow;
};
