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

# ICoreWebView2 interface

`WebView2` enables you to host web content using the latest Microsoft Edge browser and web technology.

| Name       | Description |
| ---------- | ----------- |
| [add_ContainsFullScreenElementChanged](#add_containsfullscreenelementchanged) | Adds an event handler for the AcceleratorKeyPressed event. |
| [add_ContentLoading](#add_contentloading) | Add an event handler for the ContentLoading event. |
| [add_DocumentTitleChanged](#add_documenttitlechanged) | Add an event handler for the DocumentTitleChanged event. |
| [add_FrameNavigationCompleted](#add_framenavigationcompleted) | Add an event handler for the FrameNavigationCompleted event. |
| [add_FrameNavigationStarting](#add_framenavigationstarting) | Add an event handler for the FrameNavigationStarting event. |
| [add_HistoryChanged](#add_historychanged) | Add an event handler for the HistoryChanged event. |
| [add_NavigationCompleted](#add_navigationcompleted) | Add an event handler for the NavigationCompleted event. |
| [add_NavigationStarting](#add_navigationstarting) | Add an event handler for the NavigationStarting event. |
| [add_NewWindowRequested](#add_newWindowrequested) | Add an event handler for the NewWindowRequested event. |
| [add_PermissionRequested](#add_permissionrequested) | Add an event handler for the PermissionRequested event. |
| [add_ProcessFailed](#add_processfailed) | Add an event handler for the ProcessFailed event. |
| [add_ScriptDialogOpening](#add_scriptdialogopening) | Add an event handler for the ScriptDialogOpening event. |
| [add_SourceChanged](#add_sourcechanged) | Add an event handler for the SourceChanged event. |
| [add_WebMessageReceived](#add_webmessagereceived) | Add an event handler for the WebMessageReceived event. |
| [add_WebResourceRequested](#add_webresourcerequested) | Add an event handler for the WebResourceRequested event. |
| [add_WindowCloseRequested](#add_windowcloserequested) | Add an event handler for the WindowCloseRequested event. |
| [AddHostObjectToScript](#addhostobjecttoscript) | Add the provided host object to script running in the WebView with the specified name. |
| [AddScriptToExecuteOnDocumentCreated](#addscripttoexecuteondocumentcreated) | Add the provided JavaScript to a list of scripts that should be run after the global object has been created, but before the HTML document has been parsed and before any other script included by the HTML document is run. |
| [AddWebResourceRequestedFilter](#addwebresourcerequestedfilter) | Warning: This method is deprecated and does not behave as expected for iframes. |
| [CallDevToolsProtocolMethod](#calldevtoolsprotocolmethod) | Runs an asynchronous DevToolsProtocol method. |
| [CapturePreview](#capturepreview) | Capture an image of what WebView is displaying. |
| [ExecuteScript](#executescript) | Run JavaScript code from the javascript parameter in the current top-level document rendered in the WebView. |
| [get_BrowserProcessId](#get_browserprocessid) | The process ID of the browser process that hosts the WebView. |
| [get_CanGoBack](#get_cangoback) | TRUE if the WebView is able to navigate to a previous page in the navigation history. |
| [get_CanGoForward](#get_cangoforward) | TRUE if the WebView is able to navigate to a next page in the navigation history. |
| [get_ContainsFullScreenElement](#get_containsfullscreenelement) | TRUE if the WebView is able to navigate to a next page in the navigation history. |
| [get_DocumentTitle](#get_documenttitle) | The title for the current top-level document. |
| [get_Settings](#get_settings) | The ICoreWebView2Settings object contains various modifiable settings for the running WebView. |
| [get_Source](#get_source) | The ICoreWebView2Settings object contains various modifiable settings for the running WebView. |
| [GetDevToolsProtocolEventReceiver](#getdevtoolsprotocoleventreceiver) | Get a DevTools Protocol event receiver that allows you to subscribe to a DevTools Protocol event. |
| [GoBack](#goBbck) | Navigates the WebView to the previous page in the navigation history. |
| [GoForward](#goforward) | Navigates the WebView to the next page in the navigation history. |
| [Navigate](#navigate) | Cause a navigation of the top-level document to run to the specified URI. |
| [NavigateToString](#navigatetostring) | Initiates a navigation to htmlContent as source HTML of a new document. |
| [OpenDevToolsWindow](#opendevtoolswindow) | Opens the DevTools window for the current document in the WebView. |
| [PostWebMessageAsJson](#postwebmessageAsjson) | Post the specified webMessage to the top level document in this WebView. |
| [PostWebMessageAsString](#postwebmessageasstring) | Posts a message that is a simple string rather than a JSON string representation of a JavaScript object. |
| [remove_ContainsFullScreenElementChanged](#remove_containsfullscreenelementchanged) | Remove an event handler previously added with add_ContainsFullScreenElementChanged. |
| [remove_ContentLoading](#remove_contentloading) | Remove an event handler previously added with add_ContentLoading. |
| [remove_DocumentTitleChanged](#remove_documenttitlechanged) | Remove an event handler previously added with add_DocumentTitleChanged. |
| [remove_FrameNavigationCompleted](#remove_framenavigationcompleted) | Remove an event handler previously added with add_FrameNavigationCompleted. |
| [remove_FrameNavigationStarting](#remove_framenavigationstarting) | Remove an event handler previously added with add_FrameNavigationStarting. |
| [remove_HistoryChanged](#remove_HistoryChanged) | Remove an event handler previously added with add_HistoryChanged. |
| [remove_NavigationCompleted](#remove_navigationcompleted) | Remove an event handler previously added with add_NavigationCompleted. |
| [remove_NavigationCompleted](#remove_navigationcompleted) | Remove an event handler previously added with add_NavigationCompleted. |
| [remove_NavigationStarting](#remove_navigationstarting) | Remove an event handler previously added with add_NavigationStarting. |
| [remove_NewWindowRequested](#remove_newwindowrequested) | Remove an event handler previously added with add_NewWindowRequested. |
| [remove_PermissionRequested](#remove_permissionrequested) | Remove an event handler previously added with add_PermissionRequested. |
| [remove_ProcessFailed](#remove_processfailed) | Remove an event handler previously added with add_ProcessFailed. |
| [remove_ScriptDialogOpening](#remove_scriptdialogopening) | Remove an event handler previously added with add_ScriptDialogOpening. |
| [remove_SourceChanged](#remove_sourcechanged) | Remove an event handler previously added with add_SourceChanged. |
| [remove_WebMessageReceived](#remove_webmessagereceived) | Remove an event handler previously added with add_WebMessageReceived. |
| [remove_WebResourceRequested](#remove_webresourceeequested) | Remove an event handler previously added with add_WebResourceRequested. |
| [remove_WindowCloseRequested](#remove_windowclosetequested) | Remove an event handler previously added with add_WindowCloseRequested. |
| [RemoveHostObjectFromScript](#removehostobjectfromscript) | Remove the host object specified by the name so that it is no longer accessible from JavaScript code in the WebView. |
| [RemoveScriptToExecuteOnDocumentCreated](#removescripttoexecuteondocumentcreated) | Remove the corresponding JavaScript added using AddScriptToExecuteOnDocumentCreated with the specified script ID. |
| [RemoveWebResourceRequestedFilter](#removewebresourcerequestedfilter) | Warning: This method and AddWebResourceRequestedFilter are deprecated. |
| [Stop](stop) | Stop all navigations and pending resource fetches. Does not stop scripts. |

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

## get_CoreWebView2

Gets the CoreWebView2 associated with this CoreWebView2Controller.
```
public HRESULT get_CoreWebView2(ICoreWebView2 ** coreWebView2)
```

## get_IsVisible

The IsVisible property determines whether to show or hide the WebView2.
```
public HRESULT get_IsVisible(BOOL * isVisible)
```
If IsVisible is set to FALSE, the WebView2 is transparent and is not rendered. However, this does not affect the window containing the WebView2 (the HWND parameter that was passed to CreateCoreWebView2Controller). If you want that window to disappear too, run ShowWindow on it directly in addition to modifying the IsVisible property. WebView2 as a child window does not get window messages when the top window is minimized or restored. For performance reasons, developers should set the IsVisible property of the WebView to FALSE when the app window is minimized and back to TRUE when the app window is restored. The app window does this by handling SIZE_MINIMIZED and SIZE_RESTORED command upon receiving WM_SIZE message.

There are CPU and memory benefits when the page is hidden. For instance, Chromium has code that throttles activities on the page like animations and some tasks are run less frequently. Similarly, WebView2 will purge some caches to reduce memory usage.

```
void ViewComponent::ToggleVisibility()
{
    BOOL visible;
    m_controller->get_IsVisible(&visible);
    m_isVisible = !visible;
    m_controller->put_IsVisible(m_isVisible);
}
```
---

## get_ParentWindow

The parent window provided by the app that this WebView is using to render content.
```
public HRESULT get_ParentWindow(HWND * parentWindow)
```
This API initially returns the window passed into CreateCoreWebView2Controller.

---

## get_ZoomFactor

The zoom factor for the WebView.
```
public HRESULT get_ZoomFactor(double * zoomFactor)
```
Note
Changing zoom factor may cause window.innerWidth, window.innerHeight, both, and page layout to change. A zoom factor that is applied by the host by running ZoomFactor becomes the new default zoom for the WebView. The zoom factor applies across navigations and is the zoom factor WebView is returned to when the user chooses Ctrl+0. When the zoom factor is changed by the user (resulting in the app receiving ZoomFactorChanged), that zoom applies only for the current page. Any user applied zoom is only for the current page and is reset on a navigation. Specifying a zoomFactor less than or equal to 0 is not allowed. WebView also has an internal supported zoom factor range. When a specified zoom factor is out of that range, it is normalized to be within the range, and a ZoomFactorChanged event is triggered for the real applied zoom factor. When the range normalization happens, the ZoomFactor property reports the zoom factor specified during the previous modification of the ZoomFactor property until the ZoomFactorChanged event is received after WebView applies the normalized zoom factor.

---

## MoveFocus

Moves focus into WebView.
```
public HRESULT MoveFocus(COREWEBVIEW2_MOVE_FOCUS_REASON reason)
```
WebView gets focus and focus is set to correspondent element in the page hosted in the WebView. For Programmatic reason, focus is set to previously focused element or the default element if no previously focused element exists. For Next reason, focus is set to the first element. For Previous reason, focus is set to the last element. WebView changes focus through user interaction including selecting into a WebView or Tab into it. For tabbing, the app runs MoveFocus with Next or Previous to align with Tab and Shift+Tab respectively when it decides the WebView is the next element that may exist in a tab. Or, the app runs IsDialogMessage as part of the associated message loop to allow the platform to auto handle tabbing. The platform rotates through all windows with WS_TABSTOP. When the WebView gets focus from IsDialogMessage, it is internally put the focus on the first or last element for tab and Shift+Tab respectively.

```
while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            // Calling IsDialogMessage handles Tab traversal automatically. If the
            // app wants the platform to auto handle tab, then call IsDialogMessage
            // before calling TranslateMessage/DispatchMessage. If the app wants to
            // handle tabbing itself, then skip calling IsDialogMessage and call
            // TranslateMessage/DispatchMessage directly.
            if (!g_autoTabHandle || !IsDialogMessage(GetAncestor(msg.hwnd, GA_ROOT), &msg))
            {
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
        }
    }
```
```
if (wParam == VK_TAB)
        {
            // Find out if the window is one we've customized for tab handling
            for (size_t i = 0; i < m_tabbableWindows.size(); i++)
            {
                if (m_tabbableWindows[i].first == hWnd)
                {
                    if (GetKeyState(VK_SHIFT) < 0)
                    {
                        TabBackwards(i);
                    }
                    else
                    {
                        TabForwards(i);
                    }
                    return true;
                }
            }
        }
```
```
void ControlComponent::TabForwards(size_t currentIndex)
{
    // Find first enabled window after the active one
    for (size_t i = currentIndex + 1; i < m_tabbableWindows.size(); i++)
    {
        HWND hwnd = m_tabbableWindows.at(i).first;
        if (IsWindowEnabled(hwnd))
        {
            SetFocus(hwnd);
            return;
        }
    }
    // If this is the last enabled window, tab forwards into the WebView.
    m_controller->MoveFocus(COREWEBVIEW2_MOVE_FOCUS_REASON_NEXT);
}

void ControlComponent::TabBackwards(size_t currentIndex)
{
    // Find first enabled window before the active one
    for (size_t i = currentIndex - 1; i >= 0 && i < m_tabbableWindows.size(); i--)
    {
        HWND hwnd = m_tabbableWindows.at(i).first;
        if (IsWindowEnabled(hwnd))
        {
            SetFocus(hwnd);
            return;
        }
    }
    // If this is the last enabled window, tab forwards into the WebView.
    CHECK_FAILURE(m_controller->MoveFocus(COREWEBVIEW2_MOVE_FOCUS_REASON_PREVIOUS));
}
```

## NotifyParentWindowPositionChanged

This is a notification separate from Bounds that tells WebView that the main WebView parent (or any ancestor) HWND moved.
```
public HRESULT NotifyParentWindowPositionChanged()
```
This is needed for accessibility and certain dialogs in WebView to work correctly.

```
if (message == WM_MOVE || message == WM_MOVING)
    {
        m_controller->NotifyParentWindowPositionChanged();
        return true;
    }
```
---

## put_Bounds

Sets the Bounds property.
```
public HRESULT put_Bounds(RECT bounds)
```
```
// Update the bounds of the WebView window to fit available space.
void ViewComponent::ResizeWebView()
{
    SIZE webViewSize = {
            LONG((m_webViewBounds.right - m_webViewBounds.left) * m_webViewRatio * m_webViewScale),
            LONG((m_webViewBounds.bottom - m_webViewBounds.top) * m_webViewRatio * m_webViewScale) };

    RECT desiredBounds = m_webViewBounds;
    desiredBounds.bottom = LONG(
        webViewSize.cy + m_webViewBounds.top);
    desiredBounds.right = LONG(
        webViewSize.cx + m_webViewBounds.left);

    m_controller->put_Bounds(desiredBounds);
    if (m_compositionController)
    {
        POINT webViewOffset = {m_webViewBounds.left, m_webViewBounds.top};

        if (m_dcompDevice)
        {
            CHECK_FAILURE(m_dcompRootVisual->SetOffsetX(float(webViewOffset.x)));
            CHECK_FAILURE(m_dcompRootVisual->SetOffsetY(float(webViewOffset.y)));
            CHECK_FAILURE(m_dcompRootVisual->SetClip(
                {0, 0, float(webViewSize.cx), float(webViewSize.cy)}));
            CHECK_FAILURE(m_dcompDevice->Commit());
        }
        else if (m_wincompCompositor)
        {
            if (m_wincompRootVisual != nullptr)
            {
                numerics::float2 size = {static_cast<float>(webViewSize.cx),
                                         static_cast<float>(webViewSize.cy)};
                m_wincompRootVisual.Size(size);

                numerics::float3 offset = {static_cast<float>(webViewOffset.x),
                                           static_cast<float>(webViewOffset.y), 0.0f};
                m_wincompRootVisual.Offset(offset);

                winrtComp::IInsetClip insetClip = m_wincompCompositor.CreateInsetClip();
                m_wincompRootVisual.Clip(insetClip.as<winrtComp::CompositionClip>());
            }
        }
    }
}
```

## put_IsVisible

Sets the IsVisible property.
```
public HRESULT put_IsVisible(BOOL isVisible)
```
```
if (message == WM_SIZE)
    {
        if (wParam == SIZE_MINIMIZED)
        {
            // Hide the webview when the app window is minimized.
            m_controller->put_IsVisible(FALSE);
            Suspend();
        }
        else if (wParam == SIZE_RESTORED)
        {
            // When the app window is restored, show the webview
            // (unless the user has toggle visibility off).
            if (m_isVisible)
            {
                Resume();
                m_controller->put_IsVisible(TRUE);
            }
        }
    }
```
---

## put_ParentWindow

Sets the parent window for the WebView.
```
public HRESULT put_ParentWindow(HWND parentWindow)
```
This causes the WebView to re-parent the main WebView window to the newly provided window.

---

## put_ZoomFactor

Sets the ZoomFactor property.
```
public HRESULT put_ZoomFactor(double zoomFactor)
```
---

## remove_AcceleratorKeyPressed

Removes an event handler previously added with add_AcceleratorKeyPressed.
```
public HRESULT remove_AcceleratorKeyPressed(EventRegistrationToken token)
```
---

## remove_GotFocus

Removes an event handler previously added with add_GotFocus.
```
public HRESULT remove_GotFocus(EventRegistrationToken token)
```
---

## remove_LostFocus

Removes an event handler previously added with add_LostFocus.
```
public HRESULT remove_LostFocus(EventRegistrationToken token)
```
---

## remove_MoveFocusRequested

Removes an event handler previously added with add_MoveFocusRequested.
```
public HRESULT remove_MoveFocusRequested(EventRegistrationToken token)
```
---

## remove_ZoomFactorChanged

Remove an event handler previously added with add_ZoomFactorChanged.
```
public HRESULT remove_ZoomFactorChanged(EventRegistrationToken token)
```
---

## SetBoundsAndZoomFactor

Updates Bounds and ZoomFactor properties at the same time.
```
public HRESULT SetBoundsAndZoomFactor(RECT bounds, double zoomFactor)
```
This operation is atomic from the perspective of the host. After returning from this function, the Bounds and ZoomFactor properties are both updated if the function is successful, or neither is updated if the function fails. If Bounds and ZoomFactor are both updated by the same scale (for example, Bounds and ZoomFactor are both doubled), then the page does not display a change in window.innerWidth or window.innerHeight and the WebView renders the content at the new size and zoom without intermediate renderings. This function also updates just one of ZoomFactor or Bounds by passing in the new value for one and the current value for the other.

```
void ViewComponent::SetScale(float scale)
{
    RECT bounds;
    CHECK_FAILURE(m_controller->get_Bounds(&bounds));
    double scaleChange = scale / m_webViewScale;

    bounds.bottom = LONG(
        (bounds.bottom - bounds.top) * scaleChange + bounds.top);
    bounds.right = LONG(
        (bounds.right - bounds.left) * scaleChange + bounds.left);

    m_webViewScale = scale;
    m_controller->SetBoundsAndZoomFactor(bounds, scale);
}
```
---

## add_ContainsFullScreenElementChanged

Add an event handler for the ContainsFullScreenElementChanged event.
```
public HRESULT add_ContainsFullScreenElementChanged(ICoreWebView2ContainsFullScreenElementChangedEventHandler * eventHandler, EventRegistrationToken * token)
```
ContainsFullScreenElementChanged triggers when the ContainsFullScreenElement property changes. An HTML element inside the WebView may enter fullscreen to the size of the WebView or leave fullscreen. This event is useful when, for example, a video element requests to go fullscreen. The listener of ContainsFullScreenElementChanged may resize the WebView in response.

```
// Register a handler for the ContainsFullScreenChanged event.
    CHECK_FAILURE(m_webView->add_ContainsFullScreenElementChanged(
        Callback<ICoreWebView2ContainsFullScreenElementChangedEventHandler>(
            [this](ICoreWebView2* sender, IUnknown* args) -> HRESULT
            {
                CHECK_FAILURE(
                    sender->get_ContainsFullScreenElement(&m_containsFullscreenElement));
                if (m_containsFullscreenElement)
                {
                    EnterFullScreen();
                }
                else
                {
                    ExitFullScreen();
                }
                return S_OK;
            })
            .Get(),
        nullptr));
```

## add_ContentLoading

Add an event handler for the ContentLoading event.
```
public HRESULT add_ContentLoading(ICoreWebView2ContentLoadingEventHandler * eventHandler, EventRegistrationToken * token)
```
ContentLoading triggers before any content is loaded, including scripts added with AddScriptToExecuteOnDocumentCreated. ContentLoading does not trigger if a same page navigation occurs (such as through fragment navigations or history.pushState navigations). This operation follows the NavigationStarting and SourceChanged events and precedes the HistoryChanged and NavigationCompleted events.

---

## add_DocumentTitleChanged

Add an event handler for the DocumentTitleChanged event.
```
public HRESULT add_DocumentTitleChanged(ICoreWebView2DocumentTitleChangedEventHandler * eventHandler, EventRegistrationToken * token)
```
DocumentTitleChanged runs when the DocumentTitle property of the WebView changes and may run before or after the NavigationCompleted event.
```
// Register a handler for the DocumentTitleChanged event.
    // This handler just announces the new title on the window's title bar.
    CHECK_FAILURE(m_webView->add_DocumentTitleChanged(
        Callback<ICoreWebView2DocumentTitleChangedEventHandler>(
            [this](ICoreWebView2* sender, IUnknown* args) -> HRESULT {
                wil::unique_cotaskmem_string title;
                CHECK_FAILURE(sender->get_DocumentTitle(&title));
                m_appWindow->SetDocumentTitle(title.get());
                return S_OK;
            })
            .Get(),
        &m_documentTitleChangedToken));
```
---

## add_FrameNavigationCompleted

Add an event handler for the FrameNavigationCompleted event.
```
public HRESULT add_FrameNavigationCompleted(ICoreWebView2NavigationCompletedEventHandler * eventHandler, EventRegistrationToken * token)
```
FrameNavigationCompleted triggers when a child frame has completely loaded (concurrently when body.onload has triggered) or loading stopped with error.

```
// Register a handler for the FrameNavigationCompleted event.
    // Check whether the navigation succeeded, and if not, do something.
    CHECK_FAILURE(m_webView->add_FrameNavigationCompleted(
        Callback<ICoreWebView2NavigationCompletedEventHandler>(
            [this](ICoreWebView2* sender, ICoreWebView2NavigationCompletedEventArgs* args)
                -> HRESULT {
                BOOL success;
                CHECK_FAILURE(args->get_IsSuccess(&success));
                if (!success)
                {
                    COREWEBVIEW2_WEB_ERROR_STATUS webErrorStatus;
                    CHECK_FAILURE(args->get_WebErrorStatus(&webErrorStatus));
                    // The web page can cancel its own iframe loads, so we'll ignore that.
                    if (webErrorStatus != COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED)
                    {
                        m_appWindow->AsyncMessageBox(
                            L"Iframe navigation failed: "
                                + WebErrorStatusToString(webErrorStatus),
                            L"Navigation Failure");
                    }
                }
                return S_OK;
            })
            .Get(),
        &m_frameNavigationCompletedToken));
```
---

## add_FrameNavigationStarting

Add an event handler for the FrameNavigationStarting event.
```
public HRESULT add_FrameNavigationStarting(ICoreWebView2NavigationStartingEventHandler * eventHandler, EventRegistrationToken * token)
```
FrameNavigationStarting triggers when a child frame in the WebView requests permission to navigate to a different URI. Redirects trigger this operation as well, and the navigation id is the same as the original one.

Navigations will be blocked until all FrameNavigationStarting event handlers return.

```
// Register a handler for the FrameNavigationStarting event.
    // This handler will prevent a frame from navigating to a blocked domain.
    CHECK_FAILURE(m_webView->add_FrameNavigationStarting(
        Callback<ICoreWebView2NavigationStartingEventHandler>(
            [this](ICoreWebView2* sender, ICoreWebView2NavigationStartingEventArgs* args)
                -> HRESULT
            {
                wil::unique_cotaskmem_string uri;
                CHECK_FAILURE(args->get_Uri(&uri));

                if (ShouldBlockUri(uri.get()))
                {
                    CHECK_FAILURE(args->put_Cancel(true));
                }
                return S_OK;
            })
            .Get(),
        &m_frameNavigationStartingToken));
```
---

## add_HistoryChanged

Add an event handler for the HistoryChanged event.
```
public HRESULT add_HistoryChanged(ICoreWebView2HistoryChangedEventHandler * eventHandler, EventRegistrationToken * token)
```
HistoryChanged is raised for changes to joint session history, which consists of top-level and manual frame navigations. Use HistoryChanged to verify that the CanGoBack or CanGoForward value has changed. HistoryChanged also runs for using GoBack or GoForward. HistoryChanged runs after SourceChanged and ContentLoading. CanGoBack is false for navigations initiated through ICoreWebView2Frame APIs if there has not yet been a user gesture.

```
// Register a handler for the HistoryChanged event.
    // Update the Back, Forward buttons.
    CHECK_FAILURE(m_webView->add_HistoryChanged(
        Callback<ICoreWebView2HistoryChangedEventHandler>(
            [this](ICoreWebView2* sender, IUnknown* args) -> HRESULT {
                BOOL canGoBack;
                BOOL canGoForward;
                sender->get_CanGoBack(&canGoBack);
                sender->get_CanGoForward(&canGoForward);
                m_toolbar->SetItemEnabled(Toolbar::Item_BackButton, canGoBack);
                m_toolbar->SetItemEnabled(Toolbar::Item_ForwardButton, canGoForward);

                return S_OK;
            })
            .Get(),
        &m_historyChangedToken));
```
---

## add_NavigationCompleted

Add an event handler for the NavigationCompleted event.
```
public HRESULT add_NavigationCompleted(ICoreWebView2NavigationCompletedEventHandler * eventHandler, EventRegistrationToken * token)
```
NavigationCompleted runs when the WebView has completely loaded (concurrently when body.onload runs) or loading stopped with error.

```
// Register a handler for the NavigationCompleted event.
    // Check whether the navigation succeeded, and if not, do something.
    // Also update the Cancel buttons.
    CHECK_FAILURE(m_webView->add_NavigationCompleted(
        Callback<ICoreWebView2NavigationCompletedEventHandler>(
            [this](ICoreWebView2* sender, ICoreWebView2NavigationCompletedEventArgs* args)
                -> HRESULT {
                BOOL success;
                CHECK_FAILURE(args->get_IsSuccess(&success));
                if (!success)
                {
                    COREWEBVIEW2_WEB_ERROR_STATUS webErrorStatus;
                    CHECK_FAILURE(args->get_WebErrorStatus(&webErrorStatus));
                    if (webErrorStatus == COREWEBVIEW2_WEB_ERROR_STATUS_DISCONNECTED)
                    {
                        // Do something here if you want to handle a specific error case.
                        // In most cases this isn't necessary, because the WebView will
                        // display its own error page automatically.
                    }
                }
                m_toolbar->SetItemEnabled(Toolbar::Item_CancelButton, false);
                m_toolbar->SetItemEnabled(Toolbar::Item_ReloadButton, true);
                return S_OK;
            })
            .Get(),
        &m_navigationCompletedToken));
```
---

## add_NavigationStarting

Add an event handler for the NavigationStarting event.
```
public HRESULT add_NavigationStarting(ICoreWebView2NavigationStartingEventHandler * eventHandler, EventRegistrationToken * token)
```
NavigationStarting runs when the WebView main frame is requesting permission to navigate to a different URI. Redirects trigger this operation as well, and the navigation id is the same as the original one.

Navigations will be blocked until all NavigationStarting event handlers return.

```
// Register a handler for the NavigationStarting event.
    // This handler will check the domain being navigated to, and if the domain
    // matches a list of blocked sites, it will cancel the navigation and
    // possibly display a warning page.  It will also disable JavaScript on
    // selected websites.
    CHECK_FAILURE(m_webView->add_NavigationStarting(
        Callback<ICoreWebView2NavigationStartingEventHandler>(
            [this](ICoreWebView2* sender, ICoreWebView2NavigationStartingEventArgs* args)
                -> HRESULT
            {
                wil::unique_cotaskmem_string uri;
                CHECK_FAILURE(args->get_Uri(&uri));

                if (ShouldBlockUri(uri.get()))
                {
                    CHECK_FAILURE(args->put_Cancel(true));

                    // If the user clicked a link to navigate, show a warning page.
                    BOOL userInitiated;
                    CHECK_FAILURE(args->get_IsUserInitiated(&userInitiated));
                    static const PCWSTR htmlContent =
                        L"<h1>Domain Blocked</h1>"
                        L"<p>You've attempted to navigate to a domain in the blocked "
                        L"sites list. Press back to return to the previous page.</p>";
                    CHECK_FAILURE(sender->NavigateToString(htmlContent));
                }
                // Changes to settings will apply at the next navigation, which includes the
                // navigation after a NavigationStarting event.  We can use this to change
                // settings according to what site we're visiting.
                if (ShouldBlockScriptForUri(uri.get()))
                {
                    m_settings->put_IsScriptEnabled(FALSE);
                }
                else
                {
                    m_settings->put_IsScriptEnabled(m_isScriptEnabled);
                }
                if (m_settings2)
                {
                    static const PCWSTR url_compare_example = L"fourthcoffee.com";
                    wil::unique_bstr domain = GetDomainOfUri(uri.get());
                    const wchar_t* domains = domain.get();

                    if (wcscmp(url_compare_example, domains) == 0)
                    {
                        SetUserAgent(L"example_navigation_ua");
                    }
                }
                // [NavigationKind]
                wil::com_ptr<ICoreWebView2NavigationStartingEventArgs3> args3;
                if (SUCCEEDED(args->QueryInterface(IID_PPV_ARGS(&args3))))
                {
                    COREWEBVIEW2_NAVIGATION_KIND kind =
                        COREWEBVIEW2_NAVIGATION_KIND_NEW_DOCUMENT;
                    CHECK_FAILURE(args3->get_NavigationKind(&kind));
                }
                // ! [NavigationKind]
                return S_OK;
            })
            .Get(),
        &m_navigationStartingToken));
```
---

## add_NewWindowRequested

Add an event handler for the NewWindowRequested event.
```
public HRESULT add_NewWindowRequested(ICoreWebView2NewWindowRequestedEventHandler * eventHandler, EventRegistrationToken * token)
```
NewWindowRequested runs when content inside the WebView requests to open a new window, such as through window.open. The app can pass a target WebView that is considered the opened window or mark the event as Handled, in which case WebView2 does not open a window. If either Handled or NewWindow properties are not set, the target content will be opened on a popup window.

If a deferral is not taken on the event args, scripts that resulted in the new window that are requested are blocked until the event handler returns. If a deferral is taken, then scripts are blocked until the deferral is completed or new window is set.

For more details and considerations on the target WebView to be supplied at the opened window, see ICoreWebView2NewWindowRequestedEventArgs::put_NewWindow.

```
// Register a handler for the NewWindowRequested event.
    // This handler will defer the event, create a new app window, and then once the
    // new window is ready, it'll provide that new window's WebView as the response to
    // the request.
    CHECK_FAILURE(m_webView->add_NewWindowRequested(
        Callback<ICoreWebView2NewWindowRequestedEventHandler>(
            [this](ICoreWebView2* sender, ICoreWebView2NewWindowRequestedEventArgs* args)
            {
                if (!m_shouldHandleNewWindowRequest)
                {
                    args->put_Handled(FALSE);
                    return S_OK;
                }
                wil::com_ptr<ICoreWebView2NewWindowRequestedEventArgs> args_as_comptr = args;
                auto args3 =
                    args_as_comptr.try_query<ICoreWebView2NewWindowRequestedEventArgs3>();
                if (args3)
                {
                    wil::com_ptr<ICoreWebView2FrameInfo> frame_info;
                    CHECK_FAILURE(args3->get_OriginalSourceFrameInfo(&frame_info));
                    wil::unique_cotaskmem_string source;
                    CHECK_FAILURE(frame_info->get_Source(&source));
                    // The host can decide how to open based on source frame info,
                    // such as URI.
                    static const wchar_t* browser_launching_domain = L"www.example.com";
                    wil::unique_bstr source_domain = GetDomainOfUri(source.get());
                    const wchar_t* source_domain_as_wchar = source_domain.get();
                    if (source_domain_as_wchar &&
                        wcscmp(browser_launching_domain, source_domain_as_wchar) == 0)
                    {
                        // Open the URI in the default browser.
                        wil::unique_cotaskmem_string target_uri;
                        CHECK_FAILURE(args->get_Uri(&target_uri));
                        ShellExecute(
                            nullptr, L"open", target_uri.get(), nullptr, nullptr,
                            SW_SHOWNORMAL);
                        CHECK_FAILURE(args->put_Handled(TRUE));
                        return S_OK;
                    }
                }

                wil::com_ptr<ICoreWebView2Deferral> deferral;
                CHECK_FAILURE(args->GetDeferral(&deferral));
                AppWindow* newAppWindow;

                wil::com_ptr<ICoreWebView2WindowFeatures> windowFeatures;
                CHECK_FAILURE(args->get_WindowFeatures(&windowFeatures));

                RECT windowRect = {0};
                UINT32 left = 0;
                UINT32 top = 0;
                UINT32 height = 0;
                UINT32 width = 0;
                BOOL shouldHaveToolbar = true;

                BOOL hasPosition = FALSE;
                BOOL hasSize = FALSE;
                CHECK_FAILURE(windowFeatures->get_HasPosition(&hasPosition));
                CHECK_FAILURE(windowFeatures->get_HasSize(&hasSize));

                bool useDefaultWindow = true;

                if (!!hasPosition && !!hasSize)
                {
                    CHECK_FAILURE(windowFeatures->get_Left(&left));
                    CHECK_FAILURE(windowFeatures->get_Top(&top));
                    CHECK_FAILURE(windowFeatures->get_Height(&height));
                    CHECK_FAILURE(windowFeatures->get_Width(&width));
                    useDefaultWindow = false;
                }
                CHECK_FAILURE(windowFeatures->get_ShouldDisplayToolbar(&shouldHaveToolbar));

                windowRect.left = left;
                windowRect.right =
                    left + (width < s_minNewWindowSize ? s_minNewWindowSize : width);
                windowRect.top = top;
                windowRect.bottom =
                    top + (height < s_minNewWindowSize ? s_minNewWindowSize : height);

                // passing "none" as uri as its a noinitialnavigation
                if (!useDefaultWindow)
                {
                    newAppWindow = new AppWindow(
                        m_creationModeId, GetWebViewOption(), L"none", m_userDataFolder, false,
                        nullptr, true, windowRect, !!shouldHaveToolbar);
                }
                else
                {
                    newAppWindow = new AppWindow(m_creationModeId, GetWebViewOption(), L"none");
                }
                newAppWindow->m_isPopupWindow = true;
                newAppWindow->m_onWebViewFirstInitialized = [args, deferral, newAppWindow]()
                {
                    CHECK_FAILURE(args->put_NewWindow(newAppWindow->m_webView.get()));
                    CHECK_FAILURE(args->put_Handled(TRUE));
                    CHECK_FAILURE(deferral->Complete());
                };
                return S_OK;
            })
            .Get(),
        nullptr));
```
---
