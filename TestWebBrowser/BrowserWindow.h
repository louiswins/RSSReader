#pragma once
#include <comdef.h>
#include <exdisp.h>
#include <memory>
#include <mshtml.h>
#include <mshtmhst.h>

// BrowserWindow is loosely based on this article:
// https://www.codeguru.com/cpp/i-n/ieprogram/article.php/c4379/Display-a-Web-Page-in-a-Plain-C-Win32-Application.htm
class BrowserWindow : public IOleClientSite, public IOleInPlaceFrame, public IOleInPlaceSite, public IDocHostUIHandler
{
public:
	static HRESULT Create(HWND parent, _In_opt_ LPCOLESTR wzApplicationName, _Outptr_ BrowserWindow** ppbh);
	void ShutDown();
	void SetRect(RECT rc);

	enum Action {
		BACK = 0,
		FORWARD = 1,
		HOME = 2,
		SEARCH = 3,
		REFRESH = 4,
		STOP = 5,
	};
	void DoPageAction(Action action);
	bool ShowWebPage(LPCWSTR webPageName);
	bool ShowHTMLStr(LPCWSTR string);

	// IUnknown
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, LPVOID* ppvObj) override;
	ULONG STDMETHODCALLTYPE AddRef() override;
	ULONG STDMETHODCALLTYPE Release() override;

	// IOleClientSite
	HRESULT STDMETHODCALLTYPE SaveObject() override;
	HRESULT STDMETHODCALLTYPE GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, IMoniker** ppmk) override;
	HRESULT STDMETHODCALLTYPE GetContainer(LPOLECONTAINER* ppContainer) override;
	HRESULT STDMETHODCALLTYPE ShowObject() override;
	HRESULT STDMETHODCALLTYPE OnShowWindow(BOOL fShow) override;
	HRESULT STDMETHODCALLTYPE RequestNewObjectLayout() override;

	// IOleInPlaceFrame
	HRESULT STDMETHODCALLTYPE GetWindow(HWND* lphwnd) override;
	HRESULT STDMETHODCALLTYPE ContextSensitiveHelp(BOOL fEnterMode) override;
	HRESULT STDMETHODCALLTYPE GetBorder(LPRECT lprectBorder) override;
	HRESULT STDMETHODCALLTYPE RequestBorderSpace(LPCBORDERWIDTHS pborderwidths) override;
	HRESULT STDMETHODCALLTYPE SetBorderSpace(LPCBORDERWIDTHS pborderwidths) override;
	HRESULT STDMETHODCALLTYPE SetActiveObject(IOleInPlaceActiveObject* pActiveObject, LPCOLESTR pszObjName) override;
	HRESULT STDMETHODCALLTYPE InsertMenus(HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths) override;
	HRESULT STDMETHODCALLTYPE SetMenu(HMENU hmenuShared, HOLEMENU holemenu, HWND hwndActiveObject) override;
	HRESULT STDMETHODCALLTYPE RemoveMenus(HMENU hmenuShared) override;
	HRESULT STDMETHODCALLTYPE SetStatusText(LPCOLESTR pszStatusText) override;
	HRESULT STDMETHODCALLTYPE EnableModeless(BOOL fEnable) override;
	HRESULT STDMETHODCALLTYPE TranslateAccelerator(LPMSG lpmsg, WORD wID) override;

	// IOleInPlaceSite
	HRESULT STDMETHODCALLTYPE CanInPlaceActivate() override;
	HRESULT STDMETHODCALLTYPE OnInPlaceActivate() override;
	HRESULT STDMETHODCALLTYPE OnUIActivate() override;
	HRESULT STDMETHODCALLTYPE GetWindowContext(LPOLEINPLACEFRAME* lplpFrame, LPOLEINPLACEUIWINDOW* lplpDoc, LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo) override;
	HRESULT STDMETHODCALLTYPE Scroll(SIZE scrollExtent) override;
	HRESULT STDMETHODCALLTYPE OnUIDeactivate(BOOL fUndoable) override;
	HRESULT STDMETHODCALLTYPE OnInPlaceDeactivate() override;
	HRESULT STDMETHODCALLTYPE DiscardUndoState() override;
	HRESULT STDMETHODCALLTYPE DeactivateAndUndo() override;
	HRESULT STDMETHODCALLTYPE OnPosRectChange(LPCRECT lprcPosRect) override;

	// IDocHostUIHandler
	HRESULT STDMETHODCALLTYPE ShowContextMenu(DWORD dwID, POINT* ppt, IUnknown* pcmdtReserved, IDispatch* pdispReserved) override;
	HRESULT STDMETHODCALLTYPE GetHostInfo(DOCHOSTUIINFO* pInfo) override;
	HRESULT STDMETHODCALLTYPE ShowUI(DWORD dwID, IOleInPlaceActiveObject* pActiveObject, IOleCommandTarget* pCommandTarget, IOleInPlaceFrame* pFrame, IOleInPlaceUIWindow* pDoc) override;
	HRESULT STDMETHODCALLTYPE HideUI() override;
	HRESULT STDMETHODCALLTYPE UpdateUI() override;
	HRESULT STDMETHODCALLTYPE OnDocWindowActivate(BOOL fActivate) override;
	HRESULT STDMETHODCALLTYPE OnFrameWindowActivate(BOOL fActivate) override;
	HRESULT STDMETHODCALLTYPE ResizeBorder(LPCRECT prcBorder, IOleInPlaceUIWindow* pUIWindow, BOOL fRameWindow) override;
	HRESULT STDMETHODCALLTYPE TranslateAccelerator(LPMSG lpMsg, const GUID* pguidCmdGroup, DWORD nCmdID) override;
	HRESULT STDMETHODCALLTYPE GetOptionKeyPath(LPOLESTR* pchKey, DWORD dw) override;
	HRESULT STDMETHODCALLTYPE GetDropTarget(IDropTarget* pDropTarget, IDropTarget** ppDropTarget) override;
	HRESULT STDMETHODCALLTYPE GetExternal(IDispatch** ppDispatch) override;
	HRESULT STDMETHODCALLTYPE TranslateUrl(DWORD dwTranslate, OLECHAR* pchURLIn, OLECHAR** ppchURLOut) override;
	HRESULT STDMETHODCALLTYPE FilterDataObject(IDataObject* pDO, IDataObject** ppDORet) override;


private:
	BrowserWindow(HWND parent, IWebBrowser2* browserObj);

	HWND m_parent;
	IWebBrowser2Ptr m_webBrowser2;
	ULONG m_refcount;
};
