# WebView2

https://learn.microsoft.com/en-us/microsoft-edge/webview2/?form=MA13LH

The Microsoft Edge WebView2 control allows you to embed web technologies (HTML, CSS, and JavaScript) in your native apps. The WebView2 control uses Microsoft Edge as the rendering engine to display the web content in native apps.

With WebView2, you can embed web code in different parts of your native app, or build all of the native app within a single WebView2 instance.

Main classes for WebView2: Environment, Controller, and Core
07/14/2022
The CoreWebView2Environment, CoreWebView2Controller, and CoreWebView2 classes (or equivalent interfaces) work together so your app can host a WebView2 browser control and access its browser features. These three large classes expose a wide range of APIs that your host app can access to provide many categories of browser-related features for your users.

The CoreWebView2Environment class represents a group of WebView2 controls that share the same WebView2 browser process, user data folder, and renderer. From this CoreWebView2Environment class, you create pairs of CoreWebView2Controller and CoreWebView2 instances.
The CoreWebView2Controller class is responsible for hosting-related functionality such as window focus, visibility, size, and input, where your app hosts the WebView2 control.
The CoreWebView2 class is for the web-specific parts of the WebView2 control, including networking, navigation, script, and parsing and rendering HTML.

# ICoreWebView2Environment interface

Represents the WebView2 Environment.

WebViews created from an environment run on the browser process specified with environment parameters and objects created from an environment should be used in the same environment. Using it in different environments are not guaranteed to be compatible and may fail.

| Name       | Description |
| ---------- | ----------- |
| [add_NewBrowserVersionAvailable](#add_newbrowserversionavailable) | Add an event handler for the NewBrowserVersionAvailable event. |
| [CreateCoreWebView2Controller](#createcorewebview2controller) | Asynchronously create a new WebView. |
| [CreateWebResourceResponse](#createwebresourceresponse) | Create a new web resource response object. |
| [get_BrowserVersionString](#get_browserversionstring) | The browser version info of the current ICoreWebView2Environment, including channel name if it is not the WebView2 Runtime. |
| [remove_NewBrowserVersionAvailable](#remove_newbrowserversionavailable) | Remove an event handler previously added with `add_NewBrowserVersionAvailable`. |

---

# ICoreWebView2Controller

The owner of the `CoreWebView2` object that provides support for resizing, showing and hiding, focusing, and other functionality related to windowing and composition.

| Name       | Description |
| ---------- | ----------- |
| [add_AcceleratorKeyPressed](#add_acceleratorkeypressed) | Adds an event handler for the AcceleratorKeyPressed event. |
| [add_GotFocus](#add_gotfocus) | Adds an event handler for the GotFocus event. |
| [add_LostFocus](#add_lostfocus) | Adds an event handler for the LostFocus event. |
| [add_MoveFocusRequested](#add_movefocusrequested) | Adds an event handler for the MoveFocusRequested event. |
| [add_ZoomFactorChanged](#add_zoomfactorchanged) | Adds an event handler for the ZoomFactorChanged event. |
| [Close](#close) | Closes the WebView and cleans up the underlying browser instance. |
| [get_Bounds](#get_bounds) | The WebView bounds. |
| [get_CoreWebView2](#get_corewebview2) | Gets the CoreWebView2 associated with this CoreWebView2Controller. |
| [get_IsVisible](#get_isvisible) | The IsVisible property determines whether to show or hide the WebView2. |
| [get_ParentWindow](#get_parentwindow) | The parent window provided by the app that this WebView is using to render content. |
| [get_ZoomFactor](#get_zoomfactor) | The zoom factor for the WebView. |
| [MoveFocus](#movefocus) | Moves focus into WebView. |
| [NotifyParentWindowPositionChanged](#notifyparentwindowpositionchanged) | This is a notification separate from Bounds that tells WebView that the main WebView parent (or any ancestor) HWND moved. |
| [put_Bounds](#put_bounds) | Sets the Bounds property. |
| [put_IsVisible](#put_isvisible) | Sets the IsVisible property. |
| [put_ParentWindow](#put_parentwindow) | Sets the parent window for the WebView. |
| [put_ZoomFactor](#put_zoomfactor) | Sets the ZoomFactor property. |
| [remove_AcceleratorKeyPressed](#remove_acceleratorkeypressed) | Removes an event handler previously added with add_AcceleratorKeyPressed. |
| [remove_GotFocus](#remove_gotfocus) | Removes an event handler previously added with add_GotFocus. |
| [remove_LostFocus](#remove_lostfocus) | Removes an event handler previously added with add_LostFocus. |
| [remove_MoveFocusRequested](#remove_movefocusrequested) | Removes an event handler previously added with add_MoveFocusRequested. |
| [remove_ZoomFactorChanged](#remove_zoomfactorchanged) | Remove an event handler previously added with add_ZoomFactorChanged. |
| [SetBoundsAndZoomFactor](#setboundsandzoomfactor) | Updates Bounds and ZoomFactor properties at the same time. |

---

## add_NewBrowserVersionAvailable

Add an event handler for the NewBrowserVersionAvailable event.

`NewBrowserVersionAvailable` runs when a newer version of the WebView2 Runtime is installed and available using WebView2. To use the newer version of the browser you must create a new environment and WebView. The event only runs for new version from the same WebView2 Runtime from which the code is running. When not running with installed WebView2 Runtime, no event is run.

Because a user data folder is only able to be used by one browser process at a time, if you want to use the same user data folder in the WebView using the new version of the browser, you must close the environment and instance of WebView that are using the older version of the browser first. Or simply prompt the user to restart the app.

```
public HRESULT add_NewBrowserVersionAvailable(ICoreWebView2NewBrowserVersionAvailableEventHandler * eventHandler, EventRegistrationToken * token)
```
```
// After the environment is successfully created,
    // register a handler for the NewBrowserVersionAvailable event.
    // This handler tells when there is a new Edge version available on the machine.
    CHECK_FAILURE(m_webViewEnvironment->add_NewBrowserVersionAvailable(
        Callback<ICoreWebView2NewBrowserVersionAvailableEventHandler>(
            [this](ICoreWebView2Environment* sender, IUnknown* args) -> HRESULT
            {
                // Don't block the event handler with a message box
                RunAsync(
                    [this]
                    {
                        std::wstring message =
                            L"We detected there is a new version for the browser.";
                        if (m_webView)
                        {
                            message += L"Do you want to restart the app? \n\n";
                            message +=
                                L"Click No if you only want to re-create the webviews. \n";
                            message += L"Click Cancel for no action. \n";
                        }
                        int response = MessageBox(
                            m_mainWindow, message.c_str(), L"New available version",
                            m_webView ? MB_YESNOCANCEL : MB_OK);

                        if (response == IDYES)
                        {
                            RestartApp();
                        }
                        else if (response == IDNO)
                        {
                            ReinitializeWebViewWithNewBrowser();
                        }
                        else
                        {
                            // do nothing
                        }
                    });

                return S_OK;
            })
            .Get(),
        nullptr));
```
---

## CreateCoreWebView2Controller

Asynchronously create a new WebView.

```
public HRESULT CreateCoreWebView2Controller(HWND parentWindow, ICoreWebView2CreateCoreWebView2ControllerCompletedHandler * handler)
```

`parentWindow` is the `HWND` in which the WebView should be displayed and from which receive input. The WebView adds a child window to the provided window before this function returns. Z-order and other things impacted by sibling window order are affected accordingly. If you want to move the WebView to a different parent after it has been created, you must call put_ParentWindow to update tooltip positions, accessibility trees, and such.

HWND_MESSAGE is a valid parameter for `parentWindow` for an invisible WebView for Windows 8 and above. In this case the window will never become visible. You are not able to reparent the window after you have created the WebView. This is not supported in Windows 7 or below. Passing this parameter in Windows 7 or below will return ERROR_INVALID_WINDOW_HANDLE in the controller callback.

It is recommended that the app set Application User Model ID for the process or the app window. If none is set, during WebView creation a generated Application User Model ID is set to root window of `parentWindow`.

```
// Create or recreate the WebView and its environment.
void AppWindow::InitializeWebView()
{
    // To ensure browser switches get applied correctly, we need to close
    // the existing WebView. This will result in a new browser process
    // getting created which will apply the browser switches.
    CloseWebView();
    m_dcompDevice = nullptr;
    m_wincompCompositor = nullptr;
    LPCWSTR subFolder = nullptr;

    if (m_creationModeId == IDM_CREATION_MODE_VISUAL_DCOMP ||
        m_creationModeId == IDM_CREATION_MODE_TARGET_DCOMP)
    {
        HRESULT hr = DCompositionCreateDevice2(nullptr, IID_PPV_ARGS(&m_dcompDevice));
        if (!SUCCEEDED(hr))
        {
            MessageBox(
                m_mainWindow,
                L"Attempting to create WebView using DComp Visual is not supported.\r\n"
                "DComp device creation failed.\r\n"
                "Current OS may not support DComp.",
                L"Create with Windowless DComp Visual Failed", MB_OK);
            return;
        }
    }
    else if (m_creationModeId == IDM_CREATION_MODE_VISUAL_WINCOMP)
    {
        HRESULT hr = TryCreateDispatcherQueue();
        if (!SUCCEEDED(hr))
        {
            MessageBox(
                m_mainWindow,
                L"Attempting to create WebView using WinComp Visual is not supported.\r\n"
                "WinComp compositor creation failed.\r\n"
                "Current OS may not support WinComp.",
                L"Create with Windowless WinComp Visual Failed", MB_OK);
            return;
        }
        m_wincompCompositor = winrtComp::Compositor();
    }

    std::wstring args;
    args.append(L"--enable-features=ThirdPartyStoragePartitioning,PartitionedCookies");
    auto options = Microsoft::WRL::Make<CoreWebView2EnvironmentOptions>();
    options->put_AdditionalBrowserArguments(args.c_str());
    CHECK_FAILURE(
        options->put_AllowSingleSignOnUsingOSPrimaryAccount(m_AADSSOEnabled ? TRUE : FALSE));
    CHECK_FAILURE(options->put_ExclusiveUserDataFolderAccess(
        m_ExclusiveUserDataFolderAccess ? TRUE : FALSE));
    if (!m_language.empty())
        CHECK_FAILURE(options->put_Language(m_language.c_str()));
    CHECK_FAILURE(options->put_IsCustomCrashReportingEnabled(
        m_CustomCrashReportingEnabled ? TRUE : FALSE));

    Microsoft::WRL::ComPtr<ICoreWebView2EnvironmentOptions4> options4;
    if (options.As(&options4) == S_OK)
    {
        const WCHAR* allowedOrigins[1] = {L"https://*.example.com"};

        auto customSchemeRegistration =
            Microsoft::WRL::Make<CoreWebView2CustomSchemeRegistration>(L"custom-scheme");
        customSchemeRegistration->SetAllowedOrigins(1, allowedOrigins);
        auto customSchemeRegistration2 =
            Microsoft::WRL::Make<CoreWebView2CustomSchemeRegistration>(L"wv2rocks");
        customSchemeRegistration2->put_TreatAsSecure(TRUE);
        customSchemeRegistration2->SetAllowedOrigins(1, allowedOrigins);
        customSchemeRegistration2->put_HasAuthorityComponent(TRUE);
        auto customSchemeRegistration3 =
            Microsoft::WRL::Make<CoreWebView2CustomSchemeRegistration>(
                L"custom-scheme-not-in-allowed-origins");
        ICoreWebView2CustomSchemeRegistration* registrations[3] = {
            customSchemeRegistration.Get(), customSchemeRegistration2.Get(),
            customSchemeRegistration3.Get()};
        options4->SetCustomSchemeRegistrations(
            2, static_cast<ICoreWebView2CustomSchemeRegistration**>(registrations));
    }

    Microsoft::WRL::ComPtr<ICoreWebView2EnvironmentOptions5> options5;
    if (options.As(&options5) == S_OK)
    {
        CHECK_FAILURE(
            options5->put_EnableTrackingPrevention(m_TrackingPreventionEnabled ? TRUE : FALSE));
    }

    Microsoft::WRL::ComPtr<ICoreWebView2EnvironmentOptions6> options6;
    if (options.As(&options6) == S_OK)
    {
        CHECK_FAILURE(options6->put_AreBrowserExtensionsEnabled(TRUE));
    }

    Microsoft::WRL::ComPtr<ICoreWebView2EnvironmentOptions8> options8;
    if (options.As(&options8) == S_OK)
    {
        COREWEBVIEW2_SCROLLBAR_STYLE style = COREWEBVIEW2_SCROLLBAR_STYLE_FLUENT_OVERLAY;
        CHECK_FAILURE(options8->put_ScrollBarStyle(style));
    }

    HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(
        subFolder, m_userDataFolder.c_str(), options.Get(),
        Callback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler>(
            this, &AppWindow::OnCreateEnvironmentCompleted)
            .Get());
    if (!SUCCEEDED(hr))
    {
        switch (hr)
        {
        case HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND):
        {
            MessageBox(
                m_mainWindow,
                L"Couldn't find Edge WebView2 Runtime. "
                "Do you have a version installed?",
                nullptr, MB_OK);
        }
        break;
        case HRESULT_FROM_WIN32(ERROR_FILE_EXISTS):
        {
            MessageBox(
                m_mainWindow,
                L"User data folder cannot be created because a file with the same name already "
                L"exists.",
                nullptr, MB_OK);
        }
        break;
        case E_ACCESSDENIED:
        {
            MessageBox(
                m_mainWindow, L"Unable to create user data folder, Access Denied.", nullptr,
                MB_OK);
        }
        break;
        case E_FAIL:
        {
            MessageBox(m_mainWindow, L"Edge runtime unable to start", nullptr, MB_OK);
        }
        break;
        default:
        {
            ShowFailure(hr, L"Failed to create WebView2 environment");
        }
        }
    }
}
// This is the callback passed to CreateWebViewEnvironmentWithOptions.
// Here we simply create the WebView.
HRESULT AppWindow::OnCreateEnvironmentCompleted(
    HRESULT result, ICoreWebView2Environment* environment)
{
    if (result != S_OK)
    {
        ShowFailure(result, L"Failed to create environment object.");
        return S_OK;
    }
    m_webViewEnvironment = environment;

    if (m_webviewOption.entry == WebViewCreateEntry::EVER_FROM_CREATE_WITH_OPTION_MENU ||
        m_creationModeId == IDM_CREATION_MODE_HOST_INPUT_PROCESSING
    )
    {
        return CreateControllerWithOptions();
    }
    auto webViewEnvironment3 = m_webViewEnvironment.try_query<ICoreWebView2Environment3>();

    if (webViewEnvironment3 && (m_dcompDevice || m_wincompCompositor))
    {
        CHECK_FAILURE(webViewEnvironment3->CreateCoreWebView2CompositionController(
            m_mainWindow,
            Callback<ICoreWebView2CreateCoreWebView2CompositionControllerCompletedHandler>(
                [this](
                    HRESULT result,
                    ICoreWebView2CompositionController* compositionController) -> HRESULT
                {
                    auto controller =
                        wil::com_ptr<ICoreWebView2CompositionController>(compositionController)
                            .query<ICoreWebView2Controller>();
                    return OnCreateCoreWebView2ControllerCompleted(result, controller.get());
                })
                .Get()));
    }
    else
    {
        CHECK_FAILURE(m_webViewEnvironment->CreateCoreWebView2Controller(
            m_mainWindow, Callback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler>(
                              this, &AppWindow::OnCreateCoreWebView2ControllerCompleted)
                              .Get()));
    }

    return S_OK;
}
```
It is recommended that the app handles restart manager messages, to gracefully restart it in the case when the app is using the WebView2 Runtime from a certain installation and that installation is being uninstalled. For example, if a user installs a version of the WebView2 Runtime and opts to use another version of the WebView2 Runtime for testing the app, and then uninstalls the 1st version of the WebView2 Runtime without closing the app, the app restarts to allow un-installation to succeed.

```
case WM_QUERYENDSESSION:
    {
        // yes, we can shut down
        // Register how we might be restarted
        RegisterApplicationRestart(L"--restore", RESTART_NO_CRASH | RESTART_NO_HANG);
        *result = TRUE;
        return true;
    }
    break;
    case WM_ENDSESSION:
    {
        if (wParam == TRUE)
        {
            // save app state and exit.
            PostQuitMessage(0);
            return true;
        }
    }
    break;
```
The app should retry `CreateCoreWebView2Controller` upon failure, unless the error code is `HRESULT_FROM_WIN32(ERROR_INVALID_STATE)`. When the app retries `CreateCoreWebView2Controller` upon failure, it is recommended that the app restarts from creating a new WebView2 Environment. If a WebView2 Runtime update happens, the version associated with a WebView2 Environment may have been removed and causing the object to no longer work. Creating a new WebView2 Environment works since it uses the latest version.

WebView creation fails with `HRESULT_FROM_WIN32(ERROR_INVALID_STATE)` if a running instance using the same user data folder exists, and the Environment objects have different `EnvironmentOptions`. For example, if a WebView was created with one language, an attempt to create a WebView with a different language using the same user data folder will fail.

The creation will fail with `E_ABORT` if `parentWindow` is destroyed before the creation is finished. If this is caused by a call to `DestroyWindow`, the creation completed handler will be invoked before `DestroyWindow` returns, so you can use this to cancel creation and clean up resources synchronously when quitting a thread.

In rare cases the creation can fail with `E_UNEXPECTED` if runtime does not have permissions to the user data folder.

---

## CreateWebResourceResponse

Create a new web resource response object.

```
public HRESULT CreateWebResourceResponse(IStream * content, int statusCode, LPCWSTR reasonPhrase, LPCWSTR headers, ICoreWebView2WebResourceResponse ** response)
```

The `headers` parameter is the raw response header string delimited by newline. It is also possible to create this object with null headers string and then use the `ICoreWebView2HttpResponseHeaders` to construct the headers line by line. For more information about other parameters, navigate to `ICoreWebView2WebResourceResponse`.

```
if (m_blockImages)
        {
            CHECK_FEATURE_RETURN_EMPTY(m_webView2_22);
            m_webView2_22->AddWebResourceRequestedFilterWithRequestSourceKinds(
                L"*", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE,
                COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS_DOCUMENT);
            CHECK_FAILURE(m_webView->add_WebResourceRequested(
                Callback<ICoreWebView2WebResourceRequestedEventHandler>(
                    [this](
                        ICoreWebView2* sender, ICoreWebView2WebResourceRequestedEventArgs* args)
                    {
                        COREWEBVIEW2_WEB_RESOURCE_CONTEXT resourceContext;
                        CHECK_FAILURE(args->get_ResourceContext(&resourceContext));
                        // Ensure that the type is image
                        if (resourceContext != COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE)
                        {
                            return E_INVALIDARG;
                        }
                        // Override the response with an empty one to block the image.
                        // If put_Response is not called, the request will
                        // continue as normal.
                        wil::com_ptr<ICoreWebView2WebResourceResponse> response;
                        wil::com_ptr<ICoreWebView2Environment> environment;
                        wil::com_ptr<ICoreWebView2_2> webview2;
                        CHECK_FAILURE(m_webView->QueryInterface(IID_PPV_ARGS(&webview2)));
                        CHECK_FAILURE(webview2->get_Environment(&environment));
                        CHECK_FAILURE(environment->CreateWebResourceResponse(
                            nullptr, 403 /*NoContent*/, L"Blocked", L"Content-Type: image/jpeg",
                            &response));
                        CHECK_FAILURE(args->put_Response(response.get()));
                        return S_OK;
                    })
                    .Get(),
                &m_webResourceRequestedTokenForImageBlocking));
        }
        else
        {
            CHECK_FAILURE(m_webView->remove_WebResourceRequested(
                m_webResourceRequestedTokenForImageBlocking));
        }
```
```
if (m_replaceImages)
        {
            CHECK_FEATURE_RETURN_EMPTY(m_webView2_22);
            m_webView2_22->AddWebResourceRequestedFilterWithRequestSourceKinds(
                L"*", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE,
                COREWEBVIEW2_WEB_RESOURCE_REQUEST_SOURCE_KINDS_DOCUMENT);
            CHECK_FAILURE(m_webView->add_WebResourceRequested(
                Callback<ICoreWebView2WebResourceRequestedEventHandler>(
                    [this](
                        ICoreWebView2* sender, ICoreWebView2WebResourceRequestedEventArgs* args)
                    {
                        COREWEBVIEW2_WEB_RESOURCE_CONTEXT resourceContext;
                        CHECK_FAILURE(args->get_ResourceContext(&resourceContext));
                        // Ensure that the type is image
                        if (resourceContext != COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE)
                        {
                            return E_INVALIDARG;
                        }
                        // Override the response with an another image.
                        // If put_Response is not called, the request will
                        // continue as normal.
                        // It's not required for this scenario, but generally you should examine
                        // relevant HTTP request headers just like an HTTP server would do when
                        // producing a response stream.
                        wil::com_ptr<IStream> stream;
                        CHECK_FAILURE(SHCreateStreamOnFileEx(
                            L"assets/EdgeWebView2-80.jpg", STGM_READ, FILE_ATTRIBUTE_NORMAL,
                            FALSE, nullptr, &stream));
                        wil::com_ptr<ICoreWebView2WebResourceResponse> response;
                        wil::com_ptr<ICoreWebView2Environment> environment;
                        wil::com_ptr<ICoreWebView2_2> webview2;
                        CHECK_FAILURE(m_webView->QueryInterface(IID_PPV_ARGS(&webview2)));
                        CHECK_FAILURE(webview2->get_Environment(&environment));
                        CHECK_FAILURE(environment->CreateWebResourceResponse(
                            stream.get(), 200, L"OK", L"Content-Type: image/jpeg", &response));
                        CHECK_FAILURE(args->put_Response(response.get()));
                        return S_OK;
                    })
                    .Get(),
                &m_webResourceRequestedTokenForImageReplacing));
        }
        else
        {
            CHECK_FAILURE(m_webView->remove_WebResourceRequested(
                m_webResourceRequestedTokenForImageReplacing));
        }
```
---

## get_BrowserVersionString

The browser version info of the current `ICoreWebView2Environment`, including channel name if it is not the WebView2 Runtime.
```
public HRESULT get_BrowserVersionString(LPWSTR * versionInfo)
```
It matches the format of the `GetAvailableCoreWebView2BrowserVersionString` API. Channel names are `beta`, `dev`, and `canary`.

The caller must free the returned string with `CoTaskMemFree`. See API Conventions.

```
wil::unique_cotaskmem_string version_info;
        m_webViewEnvironment->get_BrowserVersionString(&version_info);
        MessageBox(
            m_mainWindow, version_info.get(), L"Browser Version Info After WebView Creation",
            MB_OK);
```
---

## remove_NewBrowserVersionAvailable

Remove an event handler previously added with `add_NewBrowserVersionAvailable`.

```
public HRESULT remove_NewBrowserVersionAvailable(EventRegistrationToken token)
```
---

## add_AcceleratorKeyPressed

Adds an event handler for the AcceleratorKeyPressed event.
```
public HRESULT add_AcceleratorKeyPressed(ICoreWebView2AcceleratorKeyPressedEventHandler * eventHandler, EventRegistrationToken * token)
```
AcceleratorKeyPressed runs when an accelerator key or key combo is pressed or released while the WebView is focused. A key is considered an accelerator if either of the following conditions are true.

- Ctrl or Alt is currently being held.

- The pressed key does not map to a character.

A few specific keys are never considered accelerators, such as Shift. The Escape key is always considered an accelerator.

Auto-repeated key events caused by holding the key down also triggers this event. Filter out the auto-repeated key events by verifying the KeyEventLParam or PhysicalKeyStatus event args.

In windowed mode, the event handler is run synchronously. Until you run Handled() on the event args or the event handler returns, the browser process is blocked and outgoing cross-process COM requests fail with RPC_E_CANTCALLOUT_ININPUTSYNCCALL. All CoreWebView2 API methods work, however.

In windowless mode, the event handler is run asynchronously. Further input do not reach the browser until the event handler returns or Handled() is run, but the browser process is not blocked, and outgoing COM requests work normally.

It is recommended to run Handled(TRUE) as early as are able to know that you want to handle the accelerator key.

```
// Register a handler for the AcceleratorKeyPressed event.
    CHECK_FAILURE(m_controller->add_AcceleratorKeyPressed(
        Callback<ICoreWebView2AcceleratorKeyPressedEventHandler>(
            [this](
                ICoreWebView2Controller* sender,
                ICoreWebView2AcceleratorKeyPressedEventArgs* args) -> HRESULT {
                COREWEBVIEW2_KEY_EVENT_KIND kind;
                CHECK_FAILURE(args->get_KeyEventKind(&kind));
                // We only care about key down events.
                if (kind == COREWEBVIEW2_KEY_EVENT_KIND_KEY_DOWN ||
                    kind == COREWEBVIEW2_KEY_EVENT_KIND_SYSTEM_KEY_DOWN)
                {
                    UINT key;
                    CHECK_FAILURE(args->get_VirtualKey(&key));
                    // Check if the key is one we want to handle.
                    std::function<void()> action = m_appWindow->GetAcceleratorKeyFunction(key);
                    if (action)
                    {
                        // Keep the browser from handling this key, whether it's autorepeated or
                        // not.
                        CHECK_FAILURE(args->put_Handled(TRUE));

                        // Filter out autorepeated keys.
                        COREWEBVIEW2_PHYSICAL_KEY_STATUS status;
                        CHECK_FAILURE(args->get_PhysicalKeyStatus(&status));
                        if (!status.WasKeyDown)
                        {
                            // Perform the action asynchronously to avoid blocking the
                            // browser process's event queue.
                            m_appWindow->RunAsync(action);
                        }
                    }
                }
                return S_OK;
            })
            .Get(),
        &m_acceleratorKeyPressedToken));
```
---

## add_GotFocus

Adds an event handler for the GotFocus event.
```
public HRESULT add_GotFocus(ICoreWebView2FocusChangedEventHandler * eventHandler, EventRegistrationToken * token)
```
GotFocus runs when WebView has focus.

---

## add_LostFocus

Adds an event handler for the LostFocus event.
```
public HRESULT add_LostFocus(ICoreWebView2FocusChangedEventHandler * eventHandler, EventRegistrationToken * token)
```
LostFocus runs when WebView loses focus. In the case where MoveFocusRequested event is run, the focus is still on WebView when MoveFocusRequested event runs. LostFocus only runs afterwards when code of the app or default action of MoveFocusRequested event set focus away from WebView.

## add_MoveFocusRequested

Adds an event handler for the MoveFocusRequested event.
```
public HRESULT add_MoveFocusRequested(ICoreWebView2MoveFocusRequestedEventHandler * eventHandler, EventRegistrationToken * token)
```
MoveFocusRequested runs when user tries to tab out of the WebView. The focus of the WebView has not changed when this event is run.

```
// Register a handler for the MoveFocusRequested event.
    // This event will be fired when the user tabs out of the webview.
    // The handler will focus another window in the app, depending on which
    // direction the focus is being shifted.
    CHECK_FAILURE(m_controller->add_MoveFocusRequested(
        Callback<ICoreWebView2MoveFocusRequestedEventHandler>(
            [this](
                ICoreWebView2Controller* sender,
                ICoreWebView2MoveFocusRequestedEventArgs* args) -> HRESULT {
                if (!g_autoTabHandle)
                {
                    COREWEBVIEW2_MOVE_FOCUS_REASON reason;
                    CHECK_FAILURE(args->get_Reason(&reason));

                    if (reason == COREWEBVIEW2_MOVE_FOCUS_REASON_NEXT)
                    {
                        TabForwards(-1);
                    }
                    else if (reason == COREWEBVIEW2_MOVE_FOCUS_REASON_PREVIOUS)
                    {
                        TabBackwards(m_tabbableWindows.size());
                    }
                    CHECK_FAILURE(args->put_Handled(TRUE));
                }
                return S_OK;
            })
            .Get(),
        &m_moveFocusRequestedToken));
```
---

## add_ZoomFactorChanged

Adds an event handler for the ZoomFactorChanged event.
```
public HRESULT add_ZoomFactorChanged(ICoreWebView2ZoomFactorChangedEventHandler * eventHandler, EventRegistrationToken * token)
```
ZoomFactorChanged runs when the ZoomFactor property of the WebView changes. The event may run because the ZoomFactor property was modified, or due to the user manually modifying the zoom. When it is modified using the ZoomFactor property, the internal zoom factor is updated immediately and no ZoomFactorChanged event is triggered. WebView associates the last used zoom factor for each site. It is possible for the zoom factor to change when navigating to a different page. When the zoom factor changes due to a navigation change, the ZoomFactorChanged event runs right after the ContentLoading event.

```
// Register a handler for the ZoomFactorChanged event.
    // This handler just announces the new level of zoom on the window's title bar.
    CHECK_FAILURE(m_controller->add_ZoomFactorChanged(
        Callback<ICoreWebView2ZoomFactorChangedEventHandler>(
            [this](ICoreWebView2Controller* sender, IUnknown* args) -> HRESULT {
                double zoomFactor;
                CHECK_FAILURE(sender->get_ZoomFactor(&zoomFactor));

                UpdateDocumentTitle(m_appWindow, L" (Zoom: ", zoomFactor);
                return S_OK;
            })
        .Get(),
                &m_zoomFactorChangedToken));
```
---

## Close

Closes the WebView and cleans up the underlying browser instance.
```
public HRESULT Close()
```
Cleaning up the browser instance releases the resources powering the WebView. The browser instance is shut down if no other WebViews are using it.

After running Close, most methods will fail and event handlers stop running. Specifically, the WebView releases the associated references to any associated event handlers when Close is run.

Close is implicitly run when the CoreWebView2Controller loses the final reference and is destructed. But it is best practice to explicitly run Close to avoid any accidental cycle of references between the WebView and the app code. Specifically, if you capture a reference to the WebView in an event handler you create a reference cycle between the WebView and the event handler. Run Close to break the cycle by releasing all event handlers. But to avoid the situation, it is best to both explicitly run Close on the WebView and to not capture a reference to the WebView to ensure the WebView is cleaned up correctly. Close is synchronous and won't trigger the beforeunload event.

```
// Close the WebView and deinitialize related state. This doesn't close the app window.
bool AppWindow::CloseWebView(bool cleanupUserDataFolder)
{
    if (auto file = GetComponent<FileComponent>())
    {
        if (file->IsPrintToPdfInProgress())
        {
            int selection = MessageBox(
                m_mainWindow, L"Print to PDF is in progress. Continue closing?",
                L"Print to PDF", MB_YESNO);
            if (selection == IDNO)
            {
                return false;
            }
        }
    }
    // 1. Delete components.
    DeleteAllComponents();

    // 2. If cleanup needed and BrowserProcessExited event interface available,
    // register to cleanup upon browser exit.
    wil::com_ptr<ICoreWebView2Environment5> environment5;
    if (m_webViewEnvironment)
    {
        environment5 = m_webViewEnvironment.try_query<ICoreWebView2Environment5>();
    }
    if (cleanupUserDataFolder && environment5)
    {
        // Before closing the WebView, register a handler with code to run once the
        // browser process and associated processes are terminated.
        CHECK_FAILURE(environment5->add_BrowserProcessExited(
            Callback<ICoreWebView2BrowserProcessExitedEventHandler>(
                [environment5, this](
                    ICoreWebView2Environment* sender,
                    ICoreWebView2BrowserProcessExitedEventArgs* args)
                {
                    COREWEBVIEW2_BROWSER_PROCESS_EXIT_KIND kind;
                    UINT32 pid;
                    CHECK_FAILURE(args->get_BrowserProcessExitKind(&kind));
                    CHECK_FAILURE(args->get_BrowserProcessId(&pid));

                    // If a new WebView is created from this CoreWebView2Environment after
                    // the browser has exited but before our handler gets to run, a new
                    // browser process will be created and lock the user data folder
                    // again. Do not attempt to cleanup the user data folder in these
                    // cases. We check the PID of the exited browser process against the
                    // PID of the browser process to which our last CoreWebView2 attached.
                    if (pid == m_newestBrowserPid)
                    {
                        // Watch for graceful browser process exit. Let ProcessFailed event
                        // handler take care of failed browser process termination.
                        if (kind == COREWEBVIEW2_BROWSER_PROCESS_EXIT_KIND_NORMAL)
                        {
                            CHECK_FAILURE(environment5->remove_BrowserProcessExited(
                                m_browserExitedEventToken));
                            // Release the environment only after the handler is invoked.
                            // Otherwise, there will be no environment to raise the event when
                            // the collection of WebView2 Runtime processes exit.
                            m_webViewEnvironment = nullptr;
                            RunAsync([this]() { CleanupUserDataFolder(); });
                        }
                    }
                    else
                    {
                        // The exiting process is not the last in use. Do not attempt cleanup
                        // as we might still have a webview open over the user data folder.
                        // Do not block from event handler.
                        AsyncMessageBox(
                            L"A new browser process prevented cleanup of the user data folder.",
                            L"Cleanup User Data Folder");
                    }

                    return S_OK;
                })
                .Get(),
            &m_browserExitedEventToken));
    }

    // 3. Close the webview.
    if (m_controller)
    {
        m_controller->Close();
        m_controller = nullptr;
        m_webView = nullptr;
        m_webView3 = nullptr;
    }

    // 4. If BrowserProcessExited event interface is not available, release
    // environment and proceed to cleanup immediately. If the interface is
    // available, release environment only if not waiting for the event.
    if (!environment5)
    {
        m_webViewEnvironment = nullptr;
        if (cleanupUserDataFolder)
        {
            CleanupUserDataFolder();
        }
    }
    else if (!cleanupUserDataFolder)
    {
        // Release the environment object here only if no cleanup is needed.
        // If cleanup is needed, the environment object release is deferred
        // until the browser process exits, otherwise the handler for the
        // BrowserProcessExited event will not be called.
        m_webViewEnvironment = nullptr;
    }

    // reset profile name
    m_profileName = L"";
    m_documentTitle = L"";
    return true;
}
```
---

## get_Bounds

The WebView bounds.
```
public HRESULT get_Bounds(RECT * bounds)
```
Bounds are relative to the parent HWND. The app has two ways to position a WebView.

Create a child HWND that is the WebView parent HWND. Position the window where the WebView should be. Use (0, 0) for the top-left corner (the offset) of the Bounds of the WebView.

Use the top-most window of the app as the WebView parent HWND. For example, to position WebView correctly in the app, set the top-left corner of the Bound of the WebView.

The values of Bounds are limited by the coordinate space of the host.

---
