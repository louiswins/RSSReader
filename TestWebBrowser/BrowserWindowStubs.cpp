#include "stdafx.h"
#include "BrowserWindow.h"



//// IDocHostUIHandler ////

// Called when the browser object is about to display its context menu.
HRESULT STDMETHODCALLTYPE BrowserWindow::ShowContextMenu(DWORD /*dwID*/, POINT* /*ppt*/, IUnknown* /*pcmdtReserved*/, IDispatch* /*pdispReserved*/)
{
	// If desired, we can pop up your own custom context menu here.
	// Return S_OK to tell the browser not to display its default context menu,
	// or return S_FALSE to let the browser show its default context menu.
	return S_FALSE;
}

// Called when the browser object shows its UI. This allows us to replace its menus and toolbars by creating our
// own and displaying them here.
HRESULT STDMETHODCALLTYPE BrowserWindow::ShowUI(DWORD /*dwID*/, IOleInPlaceActiveObject* /*pActiveObject*/, IOleCommandTarget* /*pCommandTarget*/, IOleInPlaceFrame* /*pFrame*/, IOleInPlaceUIWindow* /*pDoc*/)
{
	// We've already got our own UI in place so just return S_OK to tell the browser
	// not to display its menus/toolbars. Otherwise we'd return S_FALSE to let it do
	// that.
	return S_OK;
}

// Called when browser object hides its UI. This allows us to hide any menus/toolbars we created in ShowUI.
HRESULT STDMETHODCALLTYPE BrowserWindow::HideUI()
{
	return S_OK;
}

// Called when the browser object wants to notify us that the command state has changed. We should update any
// controls we have that are dependent upon our embedded object, such as "Back", "Forward", "Stop", or "Home"
// buttons.
HRESULT STDMETHODCALLTYPE BrowserWindow::UpdateUI()
{
	// We update our UI in our window message loop so we don't do anything here.
	return S_OK;
}

// Called from the browser object's IOleInPlaceActiveObject object's OnDocWindowActivate() function.
// This informs off of when the object is getting/losing the focus.
HRESULT STDMETHODCALLTYPE BrowserWindow::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

// Called from the browser object's IOleInPlaceActiveObject object's OnFrameWindowActivate() function.
HRESULT STDMETHODCALLTYPE BrowserWindow::OnFrameWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

// Called from the browser object's IOleInPlaceActiveObject object's ResizeBorder() function.
HRESULT STDMETHODCALLTYPE BrowserWindow::ResizeBorder(LPCRECT /*prcBorder*/, IOleInPlaceUIWindow* /*pUIWindow*/, BOOL /*fRameWindow*/)
{
	return S_OK;
}

// Called from the browser object's TranslateAccelerator routines to translate key strokes to commands.
HRESULT STDMETHODCALLTYPE BrowserWindow::TranslateAccelerator(LPMSG /*lpMsg*/, const GUID* /*pguidCmdGroup*/, DWORD /*nCmdID*/)
{
	// We don't intercept any keystrokes, so we do nothing here. But for example, if we wanted to
	// override the TAB key, perhaps do something with it ourselves, and then tell the browser
	// not to do anything with this keystroke, we'd do:
	//
	//	if (pMsg && pMsg->message == WM_KEYDOWN && pMsg->wParam == VK_TAB)
	//	{
	//		// Here we do something as a result of a TAB key press.
	//
	//		// Tell the browser not to do anything with it.
	//		return S_FALSE;
	//	}
	//
	//	// Otherwise, let the browser do something with this message.
	//	return S_OK;

	// For our example, we want to make sure that the user can invoke some key to popup the context
	// menu, so we'll tell it to ignore all messages.
	return S_OK;
}

// Called by the browser object to find where the host wishes the browser to get its options in the registry.
// We can use this to prevent the browser from using its default settings in the registry, by telling it to use
// some other registry key we've setup with the options we want.
HRESULT STDMETHODCALLTYPE BrowserWindow::GetOptionKeyPath(LPOLESTR* /*pchKey*/, DWORD /*dw*/)
{
	// Let the browser use its default registry settings.
	return S_FALSE;
}

// Called by the browser object when it is used as a drop target. We can supply our own IDropTarget object,
// IDropTarget functions, and IDropTarget VTable if we want to determine what happens when someone drags and
// drops some object on our embedded browser object.
HRESULT STDMETHODCALLTYPE BrowserWindow::GetDropTarget(IDropTarget* /*pDropTarget*/, IDropTarget** ppDropTarget)
{
	// Return our IDropTarget object associated with this IDocHostUIHandler object.
	// For our purposes, we don't need an IDropTarget object, so we'll tell whomever is calling
	// us that we don't have one.
	*ppDropTarget = nullptr;
	return S_FALSE;
}

// Called by the browser when it wants a pointer to our IDispatch object. This object allows us to expose
// our own automation interface (ie, our own COM objects) to other entities that are running within the
// context of the browser so they can call our functions if they want. An example could be a javascript
// running in the URL we display could call our IDispatch functions. We'd write them so that any args passed
// to them would use the generic datatypes like a BSTR for utmost flexibility.
HRESULT STDMETHODCALLTYPE BrowserWindow::GetExternal(IDispatch** ppDispatch)
{
	// Return our IDispatch object associated with this IDocHostUIHandler object.
	// For our purposes, we don't need an IDispatch object, so we'll tell whomever is calling
	// us that we don't have one.
	*ppDispatch = nullptr;
	return S_FALSE;
}

// Called by the browser object to give us an opportunity to modify the URL to be loaded.
HRESULT STDMETHODCALLTYPE BrowserWindow::TranslateUrl(DWORD /*dwTranslate*/, OLECHAR* /*pchURLIn*/, OLECHAR** ppchURLOut)
{
	// We don't need to modify the URL.
	*ppchURLOut = nullptr;
	return S_FALSE;
}

// Called by the browser when it does cut/paste to the clipboard. This allows us to block certain clipboard
// formats or support additional clipboard formats.
HRESULT STDMETHODCALLTYPE BrowserWindow::FilterDataObject(IDataObject* /*pDO*/, IDataObject** ppDORet)
{
	// Return our IDataObject object associated with this IDocHostUIHandler object.
	// But for our purposes, we don't need an IDataObject object, so we'll tell whomever is calling
	// us that we don't have one.
	*ppDORet = nullptr;
	return S_FALSE;
}




//// IOleClientSite ////
// We give the browser object a pointer to our IOleClientSite object when we call OleCreate() or DoVerb().

HRESULT STDMETHODCALLTYPE BrowserWindow::SaveObject()
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::GetMoniker(DWORD /*dwAssign*/, DWORD /*dwWhichMoniker*/, IMoniker** /*ppmk*/)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::GetContainer(LPOLECONTAINER* ppContainer)
{
	// Tell the browser that we are a simple object and don't support a container
	*ppContainer = nullptr;
	return E_NOINTERFACE;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::ShowObject()
{
	return NOERROR;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::OnShowWindow(BOOL /*fShow*/)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::RequestNewObjectLayout()
{
	return E_NOTIMPL;
}




//// IOleInPlaceSite ////

HRESULT STDMETHODCALLTYPE BrowserWindow::CanInPlaceActivate()
{
	// Tell the browser we can in place activate
	return S_OK;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::OnInPlaceActivate()
{
	// Tell the browser we did it ok
	return S_OK;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::OnUIActivate()
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::Scroll(SIZE /*scrollExtent*/)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::OnUIDeactivate(BOOL /*fUndoable*/)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::OnInPlaceDeactivate()
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::DiscardUndoState()
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::DeactivateAndUndo()
{
	return E_NOTIMPL;
}



//// IOleInPlaceFrame ////

HRESULT STDMETHODCALLTYPE BrowserWindow::ContextSensitiveHelp(BOOL /*fEnterMode*/)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::GetBorder(LPRECT /*lprectBorder*/)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::RequestBorderSpace(LPCBORDERWIDTHS /*pborderwidths*/)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::SetBorderSpace(LPCBORDERWIDTHS /*pborderwidths*/)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::SetActiveObject(IOleInPlaceActiveObject* /*pActiveObject*/, LPCOLESTR /*pszObjName*/)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::InsertMenus(HMENU /*hmenuShared*/, LPOLEMENUGROUPWIDTHS /*lpMenuWidths*/)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::SetMenu(HMENU /*hmenuShared*/, HOLEMENU /*holemenu*/, HWND /*hwndActiveObject*/)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::RemoveMenus(HMENU /*hmenuShared*/)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::SetStatusText(LPCOLESTR /*pszStatusText*/)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::EnableModeless(BOOL /*fEnable*/)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE BrowserWindow::TranslateAccelerator(LPMSG /*lpmsg*/, WORD /*wID*/)
{
	return E_NOTIMPL;
}
