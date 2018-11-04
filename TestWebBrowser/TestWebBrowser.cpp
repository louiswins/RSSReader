#include "stdafx.h"
#include "Browser.h"
#include "OleHelpers.h"


int CALLBACK wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int nCmdShow)
{
	OleInitializer initOle;
	// Initialize the OLE interface. We do this once-only.
	if (FAILED(initOle))
	{
		MessageBox(0, L"Can't open OLE!", L"ERROR", MB_OK);
		return -1;
	}

	INITCOMMONCONTROLSEX icc;
	icc.dwSize = sizeof(icc);
	icc.dwICC = ICC_WIN95_CLASSES;
	InitCommonControlsEx(&icc);

	auto browser = Browser::Create(hInstance, L"My Web Browser");
	if (!browser)
	{
		MessageBox(0, L"Unable to create web browser", L"ERROR", MB_OK);
		return -1;
	}
	browser->Show(nCmdShow);
	
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	return 0;
}
