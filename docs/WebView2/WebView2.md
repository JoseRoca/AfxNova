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
---

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
---
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
---

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

## add_PermissionRequested

Add an event handler for the PermissionRequested event.
```
public HRESULT add_PermissionRequested(ICoreWebView2PermissionRequestedEventHandler * eventHandler, EventRegistrationToken * token)
```
PermissionRequested runs when content in a WebView requests permission to access some privileged resources.

If a deferral is not taken on the event args, the subsequent scripts are blocked until the event handler returns. If a deferral is taken, the scripts are blocked until the deferral is completed.

```
// Register a handler for the PermissionRequested event.
    // This handler prompts the user to allow or deny the request, and remembers
    // the user's choice for later.
    CHECK_FAILURE(m_webView->add_PermissionRequested(
        Callback<ICoreWebView2PermissionRequestedEventHandler>(
            this, &SettingsComponent::OnPermissionRequested)
            .Get(),
        &m_permissionRequestedToken));
```
```
HRESULT SettingsComponent::OnPermissionRequested(
    ICoreWebView2* sender, ICoreWebView2PermissionRequestedEventArgs* args)
{
    // Obtain a deferral for the event so that the CoreWebView2
    // doesn't examine the properties we set on the event args until
    // after we call the Complete method asynchronously later.
    wil::com_ptr<ICoreWebView2Deferral> deferral;
    CHECK_FAILURE(args->GetDeferral(&deferral));

    // Do not save state to the profile so that the PermissionRequested event is
    // always raised and the app is in control of all permission requests. In
    // this example, the app listens to all requests and caches permission on
    // its own to decide whether to show custom UI to the user.
    wil::com_ptr<ICoreWebView2PermissionRequestedEventArgs3> extended_args;
    CHECK_FAILURE(args->QueryInterface(IID_PPV_ARGS(&extended_args)));
    CHECK_FAILURE(extended_args->put_SavesInProfile(FALSE));

    // Do the rest asynchronously, to avoid calling MessageBox in an event handler.
    m_appWindow->RunAsync(
        [this, deferral, args = wil::com_ptr<ICoreWebView2PermissionRequestedEventArgs>(args)]
        {
            wil::unique_cotaskmem_string uri;
            COREWEBVIEW2_PERMISSION_KIND kind = COREWEBVIEW2_PERMISSION_KIND_UNKNOWN_PERMISSION;
            BOOL userInitiated = FALSE;
            CHECK_FAILURE(args->get_Uri(&uri));
            CHECK_FAILURE(args->get_PermissionKind(&kind));
            CHECK_FAILURE(args->get_IsUserInitiated(&userInitiated));

            COREWEBVIEW2_PERMISSION_STATE state = COREWEBVIEW2_PERMISSION_STATE_DEFAULT;

            auto cached_key = std::make_tuple(std::wstring(uri.get()), kind, userInitiated);
            auto cached_permission = m_cached_permissions.find(cached_key);
            if (cached_permission != m_cached_permissions.end())
            {
                state =
                    (cached_permission->second ? COREWEBVIEW2_PERMISSION_STATE_ALLOW
                                               : COREWEBVIEW2_PERMISSION_STATE_DENY);
            }
            else
            {
                std::wstring message = L"An iframe has requested device permission for ";
                message += PermissionKindToString(kind);
                message += L" to the website at ";
                message += uri.get();
                message += L"?\n\n";
                message += L"Do you want to grant permission?\n";
                message +=
                    (userInitiated ? L"This request came from a user gesture."
                                   : L"This request did not come from a user gesture.");

                int response = MessageBox(
                    nullptr, message.c_str(), L"Permission Request",
                    MB_YESNOCANCEL | MB_ICONWARNING);
                switch (response)
                {
                case IDYES:
                    m_cached_permissions[cached_key] = true;
                    state = COREWEBVIEW2_PERMISSION_STATE_ALLOW;
                    break;
                case IDNO:
                    m_cached_permissions[cached_key] = false;
                    state = COREWEBVIEW2_PERMISSION_STATE_DENY;
                    break;
                default:
                    state = COREWEBVIEW2_PERMISSION_STATE_DEFAULT;
                    break;
                }
            }
            CHECK_FAILURE(args->put_State(state));
            CHECK_FAILURE(deferral->Complete());
        });
    return S_OK;
}
```
---

## add_ProcessFailed

Add an event handler for the ProcessFailed event.
```
public HRESULT add_ProcessFailed(ICoreWebView2ProcessFailedEventHandler * eventHandler, EventRegistrationToken * token)
```
ProcessFailed runs when any of the processes in the WebView2 Process Group encounters one of the following conditions:

| Condition  | Details |
| ---------- | ----------- |
| Unexpected exit | The process indicated by the event args has exited unexpectedly (usually due to a crash). The failure might or might not be recoverable and some failures are auto-recoverable. |
| Unresponsiveness | The process indicated by the event args has become unresponsive to user input. This is only reported for renderer processes, and will run every few seconds until the process becomes responsive again. |

Note
When the failing process is the browser process, a ICoreWebView2Environment5::BrowserProcessExited event will run too.

Your application can use ICoreWebView2ProcessFailedEventArgs and ICoreWebView2ProcessFailedEventArgs2 to identify which condition and process the event is for, and to collect diagnostics and handle recovery if necessary. For more details about which cases need to be handled by your application, see COREWEBVIEW2_PROCESS_FAILED_KIND.

```
// Register a handler for the ProcessFailed event.
    // This handler checks the failure kind and tries to:
    //   * Recreate the webview for browser failure and render unresponsive.
    //   * Reload the webview for render failure.
    //   * Reload the webview for frame-only render failure impacting app content.
    //   * Log information about the failure for other failures.
    CHECK_FAILURE(m_webView->add_ProcessFailed(
        Callback<ICoreWebView2ProcessFailedEventHandler>(
            [this](ICoreWebView2* sender, ICoreWebView2ProcessFailedEventArgs* argsRaw)
                -> HRESULT {
                wil::com_ptr<ICoreWebView2ProcessFailedEventArgs> args = argsRaw;
                COREWEBVIEW2_PROCESS_FAILED_KIND kind;
                CHECK_FAILURE(args->get_ProcessFailedKind(&kind));
                if (kind == COREWEBVIEW2_PROCESS_FAILED_KIND_BROWSER_PROCESS_EXITED)
                {
                    // Do not run a message loop from within the event handler
                    // as that could lead to reentrancy and leave the event
                    // handler in stack indefinitely. Instead, schedule the
                    // appropriate work to take place after completion of the
                    // event handler.
                    ScheduleReinitIfSelectedByUser(
                        L"Browser process exited unexpectedly.  Recreate webview?",
                        L"Browser process exited");
                }
                else if (kind == COREWEBVIEW2_PROCESS_FAILED_KIND_RENDER_PROCESS_UNRESPONSIVE)
                {
                    ScheduleReinitIfSelectedByUser(
                        L"Browser render process has stopped responding.  Recreate webview?",
                        L"Web page unresponsive");
                }
                else if (kind == COREWEBVIEW2_PROCESS_FAILED_KIND_RENDER_PROCESS_EXITED)
                {
                    // Reloading the page will start a new render process if
                    // needed.
                    ScheduleReloadIfSelectedByUser(
                        L"Browser render process exited unexpectedly. Reload page?",
                        L"Render process exited");
                }
                // Check the runtime event args implements the newer interface.
                auto args2 = args.try_query<ICoreWebView2ProcessFailedEventArgs2>();
                if (!args2)
                {
                    return S_OK;
                }
                if (kind == COREWEBVIEW2_PROCESS_FAILED_KIND_FRAME_RENDER_PROCESS_EXITED)
                {
                    // A frame-only renderer has exited unexpectedly. Check if
                    // reload is needed.
                    wil::com_ptr<ICoreWebView2FrameInfoCollection> frameInfos;
                    wil::com_ptr<ICoreWebView2FrameInfoCollectionIterator> iterator;
                    CHECK_FAILURE(args2->get_FrameInfosForFailedProcess(&frameInfos));
                    CHECK_FAILURE(frameInfos->GetIterator(&iterator));

                    BOOL hasCurrent = FALSE;
                    while (SUCCEEDED(iterator->get_HasCurrent(&hasCurrent)) && hasCurrent)
                    {
                        wil::com_ptr<ICoreWebView2FrameInfo> frameInfo;
                        CHECK_FAILURE(iterator->GetCurrent(&frameInfo));

                        wil::unique_cotaskmem_string nameRaw;
                        wil::unique_cotaskmem_string sourceRaw;
                        CHECK_FAILURE(frameInfo->get_Name(&nameRaw));
                        CHECK_FAILURE(frameInfo->get_Source(&sourceRaw));
                        if (IsAppContentUri(sourceRaw.get()))
                        {
                            ScheduleReloadIfSelectedByUser(
                                L"Browser render process for app frame exited unexpectedly. "
                                L"Reload page?",
                                L"App content frame unresponsive");
                            break;
                        }

                        BOOL hasNext = FALSE;
                        CHECK_FAILURE(iterator->MoveNext(&hasNext));
                    }
                }
                else
                {
                    // Show the process failure details. Apps can collect info for their logging
                    // purposes.
                    COREWEBVIEW2_PROCESS_FAILED_REASON reason;
                    wil::unique_cotaskmem_string processDescription;
                    int exitCode;
                    wil::unique_cotaskmem_string failedModule;

                    CHECK_FAILURE(args2->get_Reason(&reason));
                    CHECK_FAILURE(args2->get_ProcessDescription(&processDescription));
                    CHECK_FAILURE(args2->get_ExitCode(&exitCode));

                    auto argFailedModule =
                        args.try_query<ICoreWebView2ProcessFailedEventArgs3>();
                    if (argFailedModule)
                    {
                        CHECK_FAILURE(
                            argFailedModule->get_FailureSourceModulePath(&failedModule));
                    }

                    std::wstringstream message;
                    message << L"Kind: " << ProcessFailedKindToString(kind) << L"\n"
                            << L"Reason: " << ProcessFailedReasonToString(reason) << L"\n"
                            << L"Exit code: " << exitCode << L"\n"
                            << L"Process description: " << processDescription.get() << std::endl
                            << (failedModule ? L"Failed module: " : L"")
                            << (failedModule ? failedModule.get() : L"");
                    m_appWindow->AsyncMessageBox( std::move(message.str()), L"Child process failed");
                }
                return S_OK;
            })
            .Get(),
        &m_processFailedToken));
```
---

## add_ScriptDialogOpening

Add an event handler for the ScriptDialogOpening event.
```
public HRESULT add_ScriptDialogOpening(ICoreWebView2ScriptDialogOpeningEventHandler * eventHandler, EventRegistrationToken * token)
```
ScriptDialogOpening runs when a JavaScript dialog (alert, confirm, prompt, or beforeunload) displays for the webview. This event only triggers if the ICoreWebView2Settings::AreDefaultScriptDialogsEnabled property is set to FALSE. The ScriptDialogOpening event suppresses dialogs or replaces default dialogs with custom dialogs.

If a deferral is not taken on the event args, the subsequent scripts are blocked until the event handler returns. If a deferral is taken, the scripts are blocked until the deferral is completed.

```
// Register a handler for the ScriptDialogOpening event.
    // This handler will set up a custom prompt dialog for the user.  Because
    // running a message loop inside of an event handler causes problems, we
    // defer the event and handle it asynchronously.
    CHECK_FAILURE(m_webView->add_ScriptDialogOpening(
        Callback<ICoreWebView2ScriptDialogOpeningEventHandler>(
            [this](ICoreWebView2* sender, ICoreWebView2ScriptDialogOpeningEventArgs* args)
                -> HRESULT
            {
                AppWindow* appWindow = m_appWindow;
                wil::com_ptr<ICoreWebView2ScriptDialogOpeningEventArgs> eventArgs = args;
                wil::com_ptr<ICoreWebView2Deferral> deferral;
                CHECK_FAILURE(args->GetDeferral(&deferral));
                appWindow->RunAsync(
                    [appWindow, eventArgs, deferral]
                    {
                        wil::unique_cotaskmem_string uri;
                        COREWEBVIEW2_SCRIPT_DIALOG_KIND type;
                        wil::unique_cotaskmem_string message;
                        wil::unique_cotaskmem_string defaultText;

                        CHECK_FAILURE(eventArgs->get_Uri(&uri));
                        CHECK_FAILURE(eventArgs->get_Kind(&type));
                        CHECK_FAILURE(eventArgs->get_Message(&message));
                        CHECK_FAILURE(eventArgs->get_DefaultText(&defaultText));

                        std::wstring promptString =
                            std::wstring(L"The page at '") + uri.get() + L"' says:";
                        TextInputDialog dialog(
                            appWindow->GetMainWindow(), L"Script Dialog", promptString.c_str(),
                            message.get(), defaultText.get(),
                            /* readonly */ type != COREWEBVIEW2_SCRIPT_DIALOG_KIND_PROMPT);
                        if (dialog.confirmed)
                        {
                            CHECK_FAILURE(eventArgs->put_ResultText(dialog.input.c_str()));
                            CHECK_FAILURE(eventArgs->Accept());
                        }
                        deferral->Complete();
                    });
                return S_OK;
            })
            .Get(),
        &m_scriptDialogOpeningToken));
```
---

## add_SourceChanged

Add an event handler for the SourceChanged event.
```
public HRESULT add_SourceChanged(ICoreWebView2SourceChangedEventHandler * eventHandler, EventRegistrationToken * token)
```
SourceChanged triggers when the Source property changes. SourceChanged runs when navigating to a different site or fragment navigations. It does not trigger for other types of navigations such as page refreshes or history.pushState with the same URL as the current page. SourceChanged runs before ContentLoading for navigation to a new document.

```
// Register a handler for the SourceChanged event.
    // This handler will read the webview's source URI and update
    // the app's address bar.
    CHECK_FAILURE(m_webView->add_SourceChanged(
        Callback<ICoreWebView2SourceChangedEventHandler>(
            [this](ICoreWebView2* sender, ICoreWebView2SourceChangedEventArgs* args)
                -> HRESULT {
                wil::unique_cotaskmem_string uri;
                sender->get_Source(&uri);
                if (wcscmp(uri.get(), L"about:blank") == 0)
                {
                    uri = wil::make_cotaskmem_string(L"");
                }
                SetWindowText(GetAddressBar(), uri.get());

                return S_OK;
            })
            .Get(),
        &m_sourceChangedToken));
```
---

## add_WebMessageReceived

Add an event handler for the WebMessageReceived event.
```
public HRESULT add_WebMessageReceived(ICoreWebView2WebMessageReceivedEventHandler * handler, EventRegistrationToken * token)
```
WebMessageReceived runs when the ICoreWebView2Settings::IsWebMessageEnabled setting is set and the top-level document of the WebView runs window.chrome.webview.postMessage. The postMessage function is void postMessage(object) where object is any object supported by JSON conversion.

```
window.chrome.webview.addEventListener('message', arg => {
            if ("SetColor" in arg.data) {
                document.getElementById("colorable").style.color = arg.data.SetColor;
            }
            if ("WindowBounds" in arg.data) {
                document.getElementById("window-bounds").value = arg.data.WindowBounds;
            }
        });

        function SetTitleText() {
            let titleText = document.getElementById("title-text");
            window.chrome.webview.postMessage(`SetTitleText ${titleText.value}`);
        }
        function GetWindowBounds() {
            window.chrome.webview.postMessage("GetWindowBounds");
        }
```
When the page calls postMessage, the object parameter is converted to a JSON string and is posted asynchronously to the host process. This will result in the handler's Invoke method being called with the JSON string as a parameter.

```
// Setup the web message received event handler before navigating to
    // ensure we don't miss any messages.
    CHECK_FAILURE(m_webView->add_WebMessageReceived(
        Microsoft::WRL::Callback<ICoreWebView2WebMessageReceivedEventHandler>(
            [this](ICoreWebView2* sender, ICoreWebView2WebMessageReceivedEventArgs* args)
    {
        wil::unique_cotaskmem_string uri;
        CHECK_FAILURE(args->get_Source(&uri));

        // Always validate that the origin of the message is what you expect.
        if (uri.get() != m_sampleUri)
        {
            // Ignore messages from untrusted sources.
            return S_OK;
        }
        wil::unique_cotaskmem_string messageRaw;
        HRESULT hr = args->TryGetWebMessageAsString(&messageRaw);
        if (hr == E_INVALIDARG)
        {
            // Was not a string message. Ignore.
            return S_OK;
        }
        // Any other problems are fatal.
        CHECK_FAILURE(hr);
        std::wstring message = messageRaw.get();

        if (message.compare(0, 13, L"SetTitleText ") == 0)
        {
            m_appWindow->SetDocumentTitle(message.substr(13).c_str());
        }
        else if (message.compare(L"GetWindowBounds") == 0)
        {
            RECT bounds = m_appWindow->GetWindowBounds();
            std::wstring reply =
                L"{\"WindowBounds\":\"Left:" + std::to_wstring(bounds.left)
                + L"\\nTop:" + std::to_wstring(bounds.top)
                + L"\\nRight:" + std::to_wstring(bounds.right)
                + L"\\nBottom:" + std::to_wstring(bounds.bottom)
                + L"\"}";
            CHECK_FAILURE(sender->PostWebMessageAsJson(reply.c_str()));
        }
        else
        {
            // Ignore unrecognized messages, but log for further investigation
            // since it suggests a mismatch between the web content and the host.
            OutputDebugString(
                std::wstring(L"Unexpected message from main page:" + message).c_str());
        }
        return S_OK;
    }).Get(), &m_webMessageReceivedToken));
```

If the same page calls postMessage multiple times, the corresponding WebMessageReceived events are guaranteed to be fired in the same order. However, if multiple frames call postMessage, there is no guaranteed order. In addition, WebMessageReceived events caused by calls to postMessage are not guaranteed to be sequenced with events caused by DOM APIs. For example, if the page runs

```
chrome.webview.postMessage("message");
window.open();
```

then the NewWindowRequested event might be fired before the WebMessageReceived event. If you need the WebMessageReceived event to happen before anything else, then in the WebMessageReceived handler you can post a message back to the page and have the page wait until it receives that message before continuing.

---

## add_WebResourceRequested

Add an event handler for the WebResourceRequested event.
```
public HRESULT add_WebResourceRequested(ICoreWebView2WebResourceRequestedEventHandler * eventHandler, EventRegistrationToken * token)
```
WebResourceRequested runs when the WebView is performing a URL request to a matching URL and resource context and source kind filter that was added with AddWebResourceRequestedFilterWithRequestSourceKinds. At least one filter must be added for the event to run.

The web resource requested may be blocked until the event handler returns if a deferral is not taken on the event args. If a deferral is taken, then the web resource requested is blocked until the deferral is completed.

If this event is subscribed in the add_NewWindowRequested handler it should be called after the new window is set. For more details see ICoreWebView2NewWindowRequestedEventArgs::put_NewWindow.

This event is by default raised for file, http, and https URI schemes. This is also raised for registered custom URI schemes. For more details see ICoreWebView2CustomSchemeRegistration.

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

## add_WindowCloseRequested

Add an event handler for the WindowCloseRequested event.
```
public HRESULT add_WindowCloseRequested(ICoreWebView2WindowCloseRequestedEventHandler * eventHandler, EventRegistrationToken * token)
```
WindowCloseRequested triggers when content inside the WebView requested to close the window, such as after window.close is run. The app should close the WebView and related app window if that makes sense to the app. After the first window.close() call, this event may not fire for any immediate back to back window.close() calls.

```
// Register a handler for the WindowCloseRequested event.
    // This handler will close the app window if it is not the main window.
    CHECK_FAILURE(m_webView->add_WindowCloseRequested(
        Callback<ICoreWebView2WindowCloseRequestedEventHandler>(
            [this](ICoreWebView2* sender, IUnknown* args)
            {
                if (m_isPopupWindow)
                {
                    CloseAppWindow();
                }
                return S_OK;
            })
            .Get(),
        nullptr));
```
---

## AddHostObjectToScript

Add the provided host object to script running in the WebView with the specified name.
```
public HRESULT AddHostObjectToScript(LPCWSTR name, VARIANT * object)
```
Host objects are exposed as host object proxies using window.chrome.webview.hostObjects.{name}. Host object proxies are promises and resolves to an object representing the host object. The promise is rejected if the app has not added an object with the name. When JavaScript code access a property or method of the object, a promise is return, which resolves to the value returned from the host for the property or method, or rejected in case of error, for example, no property or method on the object or parameters are not valid.

Note
While simple types, IDispatch and array are supported, and IUnknown objects that also implement IDispatch are treated as IDispatch, generic IUnknown, VT_DECIMAL, or VT_RECORD variant is not supported. Remote JavaScript objects like callback functions are represented as an VT_DISPATCH``VARIANT with the object implementing IDispatch. The JavaScript callback method may be invoked using DISPID_VALUE for the DISPID. Such callback method invocations will return immediately and will not wait for the JavaScript function to run and so will not provide the return value of the JavaScript function. Nested arrays are supported up to a depth of 3. Arrays of by reference types are not supported. VT_EMPTY and VT_NULL are mapped into JavaScript as null. In JavaScript, null and undefined are mapped to VT_EMPTY.

Additionally, all host objects are exposed as window.chrome.webview.hostObjects.sync.{name}. Here the host objects are exposed as synchronous host object proxies. These are not promises and function runtimes or property access synchronously block running script waiting to communicate cross process for the host code to run. Accordingly the result may have reliability issues and it is recommended that you use the promise-based asynchronous window.chrome.webview.hostObjects.{name} API.

Synchronous host object proxies and asynchronous host object proxies may both use a proxy to the same host object. Remote changes made by one proxy propagates to any other proxy of that same host object whether the other proxies and synchronous or asynchronous.

While JavaScript is blocked on a synchronous run to native code, that native code is unable to run back to JavaScript. Attempts to do so fail with HRESULT_FROM_WIN32(ERROR_POSSIBLE_DEADLOCK).

Host object proxies are JavaScript Proxy objects that intercept all property get, property set, and method invocations. Properties or methods that are a part of the Function or Object prototype are run locally. Additionally any property or method in the chrome.webview.hostObjects.options.forceLocalProperties array are also run locally. This defaults to including optional methods that have meaning in JavaScript like toJSON and Symbol.toPrimitive. Add more to the array as required.

The chrome.webview.hostObjects.cleanupSome method performs a best effort garbage collection on host object proxies.

The chrome.webview.hostObjects.options object provides the ability to change some functionality of host objects.

| Options property | Details |
| ---------------- | ------- |
| `forceLocalProperties` | This is an array of host object property names that will be run locally, instead of being called on the native host object. This defaults to then, toJSON, Symbol.toString, and Symbol.toPrimitive. You can add other properties to specify that they should be run locally on the javascript host object proxy. |
| `log` | This is a callback that will be called with debug information. For example, you can set this to console.log.bind(console) to have it print debug information to the console to help when troubleshooting host object usage. By default this is null. |
| `shouldSerializeDates` | By default this is false, and javascript Date objects will be sent to host objects as a string using JSON.stringify. You can set this property to true to have Date objects properly serialize as a VT_DATE when sending to the native host object, and have VT_DATE properties and return values create a javascript Date object. |
| `defaultSyncProxy` | When calling a method on a synchronous proxy, the result should also be a synchronous proxy. But in some cases, the sync/async context is lost (for example, when providing to native code a reference to a function, and then calling that function in native code). In these cases, the proxy will be asynchronous, unless this property is set. |
| `forceAsyncMethodMatches` | This is an array of regular expressions. When calling a method on a synchronous proxy, the method call will be performed asynchronously if the method name matches a string or regular expression in this array. Setting this value to Async will make any method that ends with Async be an asynchronous method call. If an async method doesn't match here and isn't forced to be asynchronous, the method will be invoked synchronously, blocking execution of the calling JavaScript and then returning the resolution of the promise, rather than returning a promise. |
| `ignoreMemberNotFoundError` | By default, an exception is thrown when attempting to get the value of a proxy property that doesn't exist on the corresponding native class. Setting this property to true switches the behavior to match Chakra WinRT projection (and general JavaScript) behavior of returning undefined with no error. |
| `shouldPassTypedArraysAsArrays` | By default, typed arrays are passed to the host as IDispatch. To instead pass typed arrays to the host as array, set this to true. |

Host object proxies additionally have the following methods which run locally.

| Method name | Details |
| ----------- | ------- |
| `applyHostFunction, getHostProperty, setHostProperty` | Perform a method invocation, property get, or property set on the host object. Use the methods to explicitly force a method or property to run remotely if a conflicting local method or property exists. For instance, proxy.toString() runs the local toString method on the proxy object. But proxy.applyHostFunction('toString') runs toString on the host proxied object instead. |
| `getLocalProperty, setLocalProperty` | Perform property get, or property set locally. Use the methods to force getting or setting a property on the host object proxy rather than on the host object it represents. For instance, proxy.unknownProperty gets the property named unknownProperty from the host proxied object. But proxy.getLocalProperty('unknownProperty') gets the value of the property unknownProperty on the proxy object. |
| `sync` | Asynchronous host object proxies expose a sync method which returns a promise for a synchronous host object proxy for the same host object. For example, chrome.webview.hostObjects.sample.methodCall() returns an asynchronous host object proxy. Use the sync method to obtain a synchronous host object proxy instead: const syncProxy = await chrome.webview.hostObjects.sample.methodCall().sync(). |
| `async` | Synchronous host object proxies expose an async method which blocks and returns an asynchronous host object proxy for the same host object. For example, chrome.webview.hostObjects.sync.sample.methodCall() returns a synchronous host object proxy. Running the async method on this blocks and then returns an asynchronous host object proxy for the same host object: const asyncProxy = chrome.webview.hostObjects.sync.sample.methodCall().async(). |
| `then` | Asynchronous host object proxies have a then method. Allows proxies to be awaitable. then returns a promise that resolves with a representation of the host object. If the proxy represents a JavaScript literal, a copy of that is returned locally. If the proxy represents a function, a non-awaitable proxy is returned. If the proxy represents a JavaScript object with a mix of literal properties and function properties, the a copy of the object is returned with some properties as host object proxies. |

All other property and method invocations (other than the above Remote object proxy methods, forceLocalProperties list, and properties on Function and Object prototypes) are run remotely. Asynchronous host object proxies return a promise representing asynchronous completion of remotely invoking the method, or getting the property. The promise resolves after the remote operations complete and the promises resolve to the resulting value of the operation. Synchronous host object proxies work similarly, but block running JavaScript and wait for the remote operation to complete.

Setting a property on an asynchronous host object proxy works slightly differently. The set returns immediately and the return value is the value that is set. This is a requirement of the JavaScript Proxy object. If you need to asynchronously wait for the property set to complete, use the setHostProperty method which returns a promise as described above. Synchronous object property set property synchronously blocks until the property is set.

For example, suppose you have a COM object with the following interface.

```
[uuid(3a14c9c0-bc3e-453f-a314-4ce4a0ec81d8), object, local]
    interface IHostObjectSample : IUnknown
    {
        // Demonstrate basic method call with some parameters and a return value.
        HRESULT MethodWithParametersAndReturnValue([in] BSTR stringParameter, [in] INT integerParameter, [out, retval] BSTR* stringResult);

        // Demonstrate getting and setting a property.
        [propget] HRESULT Property([out, retval] BSTR* stringResult);
        [propput] HRESULT Property([in] BSTR stringValue);

        [propget] HRESULT IndexedProperty(INT index, [out, retval] BSTR * stringResult);
        [propput] HRESULT IndexedProperty(INT index, [in] BSTR stringValue);

        // Demonstrate native calling back into JavaScript.
        HRESULT CallCallbackAsynchronously([in] IDispatch* callbackParameter);

        // Demonstrate a property which uses Date types
        [propget] HRESULT DateProperty([out, retval] DATE * dateResult);
        [propput] HRESULT DateProperty([in] DATE dateValue);

        // Creates a date object on the native side and sets the DateProperty to it.
        HRESULT CreateNativeDate();

    };
```

Add an instance of this interface into your JavaScript with AddHostObjectToScript. In this case, name it sample.

```
VARIANT remoteObjectAsVariant = {};
m_hostObject.query_to<IDispatch>(&remoteObjectAsVariant.pdispVal);
remoteObjectAsVariant.vt = VT_DISPATCH;

// We can call AddHostObjectToScript multiple times in a row without
// calling RemoveHostObject first. This will replace the previous object
// with the new object. In our case this is the same object and everything
// is fine.
CHECK_FAILURE(
   m_webView->AddHostObjectToScript(L"sample", &remoteObjectAsVariant));
remoteObjectAsVariant.pdispVal->Release();
```

In the HTML document, use the COM object using chrome.webview.hostObjects.sample. Note that CoreWebView2.AddHostObjectToScript only applies to the top-level document and not to frames. To add host objects to frames use CoreWebView2Frame.AddHostObjectToScript.

```
document.getElementById("getPropertyAsyncButton").addEventListener("click", async () => {
      const propertyValue = await chrome.webview.hostObjects.sample.property;
      document.getElementById("getPropertyAsyncOutput").textContent = propertyValue;
    });

    document.getElementById("getPropertySyncButton").addEventListener("click", () => {
      const propertyValue = chrome.webview.hostObjects.sync.sample.property;
      document.getElementById("getPropertySyncOutput").textContent = propertyValue;
    });

    document.getElementById("setPropertyAsyncButton").addEventListener("click", async () => {
      const propertyValue = document.getElementById("setPropertyAsyncInput").value;
      // The following line will work but it will return immediately before the property value has actually been set.
      // If you need to set the property and wait for the property to change value, use the setHostProperty function.
      chrome.webview.hostObjects.sample.property = propertyValue;
      document.getElementById("setPropertyAsyncOutput").textContent = "Set";
    });

    document.getElementById("setPropertyExplicitAsyncButton").addEventListener("click", async () => {
      const propertyValue = document.getElementById("setPropertyExplicitAsyncInput").value;
      // If you care about waiting until the property has actually changed value use the setHostProperty function.
      await chrome.webview.hostObjects.sample.setHostProperty("property", propertyValue);
      document.getElementById("setPropertyExplicitAsyncOutput").textContent = "Set";
    });

    document.getElementById("setPropertySyncButton").addEventListener("click", () => {
      const propertyValue = document.getElementById("setPropertySyncInput").value;
      chrome.webview.hostObjects.sync.sample.property = propertyValue;
      document.getElementById("setPropertySyncOutput").textContent = "Set";
    });

    document.getElementById("getIndexedPropertyAsyncButton").addEventListener("click", async () => {
      const index = parseInt(document.getElementById("getIndexedPropertyAsyncParam").value);
      const resultValue = await chrome.webview.hostObjects.sample.IndexedProperty[index];
      document.getElementById("getIndexedPropertyAsyncOutput").textContent = resultValue;
    });
    document.getElementById("setIndexedPropertyAsyncButton").addEventListener("click", async () => {
      const index = parseInt(document.getElementById("setIndexedPropertyAsyncParam1").value);
      const value = document.getElementById("setIndexedPropertyAsyncParam2").value;;
      chrome.webview.hostObjects.sample.IndexedProperty[index] = value;
      document.getElementById("setIndexedPropertyAsyncOutput").textContent = "Set";
    });
    document.getElementById("invokeMethodAsyncButton").addEventListener("click", async () => {
      const paramValue1 = document.getElementById("invokeMethodAsyncParam1").value;
      const paramValue2 = parseInt(document.getElementById("invokeMethodAsyncParam2").value);
      const resultValue = await chrome.webview.hostObjects.sample.MethodWithParametersAndReturnValue(paramValue1, paramValue2);
      document.getElementById("invokeMethodAsyncOutput").textContent = resultValue;
    });

    document.getElementById("invokeMethodSyncButton").addEventListener("click", () => {
      const paramValue1 = document.getElementById("invokeMethodSyncParam1").value;
      const paramValue2 = parseInt(document.getElementById("invokeMethodSyncParam2").value);
      const resultValue = chrome.webview.hostObjects.sync.sample.MethodWithParametersAndReturnValue(paramValue1, paramValue2);
      document.getElementById("invokeMethodSyncOutput").textContent = resultValue;
    });

    let callbackCount = 0;
    document.getElementById("invokeCallbackButton").addEventListener("click", async () => {
      chrome.webview.hostObjects.sample.CallCallbackAsynchronously(() => {
        document.getElementById("invokeCallbackOutput").textContent = "Native object called the callback " + (++callbackCount) + " time(s).";
      });
    });

    // Date property
    document.getElementById("setDateButton").addEventListener("click", () => {
      chrome.webview.hostObjects.options.shouldSerializeDates = true;
      chrome.webview.hostObjects.sync.sample.dateProperty = new Date();
      document.getElementById("dateOutput").textContent = "sample.dateProperty: " + chrome.webview.hostObjects.sync.sample.dateProperty;
    });
    document.getElementById("createRemoteDateButton").addEventListener("click", () => {
      chrome.webview.hostObjects.sync.sample.createNativeDate();
      document.getElementById("dateOutput").textContent = "sample.dateProperty: " + chrome.webview.hostObjects.sync.sample.dateProperty;
    });
```

Exposing host objects to script has security risk. For more information about best practices, navigate to Best practices for developing secure WebView2 applications.

---

## AddScriptToExecuteOnDocumentCreated

Add the provided JavaScript to a list of scripts that should be run after the global object has been created, but before the HTML document has been parsed and before any other script included by the HTML document is run.
```
public HRESULT AddScriptToExecuteOnDocumentCreated(LPCWSTR javaScript, ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler * handler)
```
This method injects a script that runs on all top-level document and child frame page navigations. This method runs asynchronously, and you must wait for the completion handler to finish before the injected script is ready to run. When this method completes, the Invoke method of the handler is run with the id of the injected script. id is a string. To remove the injected script, use RemoveScriptToExecuteOnDocumentCreated.

If the method is run in add_NewWindowRequested handler it should be called before the new window is set. If called after setting the NewWindow property, the initial script may or may not apply to the initial navigation and may only apply to the subsequent navigation. For more details see ICoreWebView2NewWindowRequestedEventArgs::put_NewWindow.

Note
If an HTML document is running in a sandbox of some kind using sandbox properties or the Content-Security-Policy HTTP header affects the script that runs. For example, if the allow-modals keyword is not set then requests to run the alert function are ignored.

```
// Prompt the user for some script and register it to execute whenever a new page loads.
void ScriptComponent::AddInitializeScript()
{
    TextInputDialog dialog(
        m_appWindow->GetMainWindow(),
        L"Add Initialize Script",
        L"Initialization Script:",
        L"Enter the JavaScript code to run as the initialization script that "
            L"runs before any script in the HTML document.",
    // This example script stops child frames from opening new windows.  Because
    // the initialization script runs before any script in the HTML document, we
    // can trust the results of our checks on window.parent and window.top.
        L"if (window.parent !== window.top) {\r\n"
        L"    delete window.open;\r\n"
        L"}");
    if (dialog.confirmed)
    {
        m_webView->AddScriptToExecuteOnDocumentCreated(
            dialog.input.c_str(),
            Callback<ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler>(
                [this](HRESULT error, PCWSTR id) -> HRESULT
        {
            m_lastInitializeScriptId = id;
            m_appWindow->AsyncMessageBox(
                m_lastInitializeScriptId, L"AddScriptToExecuteOnDocumentCreated Id");
            return S_OK;
        }).Get());

    }
}
```
---

## AddWebResourceRequestedFilter

Warning: This method is deprecated and does not behave as expected for iframes.
```
public HRESULT AddWebResourceRequestedFilter(LPCWSTR const uri, COREWEBVIEW2_WEB_RESOURCE_CONTEXT const resourceContext)
```
It will cause the WebResourceRequested event to fire only for the main frame and its same-origin iframes. Please use AddWebResourceRequestedFilterWithRequestSourceKinds instead, which will let the event to fire for all iframes on the document.

Adds a URI and resource context filter for the WebResourceRequested event. A web resource request with a resource context that matches this filter's resource context and a URI that matches this filter's URI wildcard string will be raised via the WebResourceRequested event.

The uri parameter value is a wildcard string matched against the URI of the web resource request. This is a glob style wildcard string in which a * matches zero or more characters and a ? matches exactly one character. These wildcard characters can be escaped using a backslash just before the wildcard character in order to represent the literal * or ?.

The matching occurs over the URI as a whole string and not limiting wildcard matches to particular parts of the URI. The wildcard filter is compared to the URI after the URI has been normalized, any URI fragment has been removed, and non-ASCII hostnames have been converted to punycode.

Specifying a nullptr for the uri is equivalent to an empty string which matches no URIs.

For more information about resource context filters, navigate to COREWEBVIEW2_WEB_RESOURCE_CONTEXT.

| URI Filter String | Request URI | Match | Notes |
| ----------------- | ----------- | ----- | ----- |
| * | https://contoso.com/a/b/c | Yes | A single * will match all URIs |
| *://contoso.com/* | https://contoso.com/a/b/c | Yes | Matches everything in contoso.com across all schemes |
| *://contoso.com/* | https://example.com/?https://contoso.com/ | Yes | But also matches a URI with just the same text anywhere in the URI |
| example | https://contoso.com/example | No | The filter does not perform partial matches |
| *example | https://contoso.com/example | Yes | The filter matches across URI parts |
| *example | https://contoso.com/path/?example | Yes | The filter matches across URI parts |
| *example | https://contoso.com/path/?query#example | No | The filter is matched against the URI with no fragment |
| *example | https://example | No | The URI is normalized before filter matching so the actual URI used for comparison is https://example/ |
| *example/ | https://example | Yes | Just like above, but this time the filter ends with a / just like the normalized URI |
| https://xn--qei.example/ | https://&#x2764;.example/ | Yes | Non-ASCII hostnames are normalized to punycode before wildcard comparison |
| https://&#x2764;.example/ | https://xn--qei.example/ | No | Non-ASCII hostnames are normalized to punycode before wildcard comparison |

---

## CallDevToolsProtocolMethod

Runs an asynchronous DevToolsProtocol method.
```
public HRESULT CallDevToolsProtocolMethod(LPCWSTR methodName, LPCWSTR parametersAsJson, ICoreWebView2CallDevToolsProtocolMethodCompletedHandler * handler)
```
For more information about available methods, navigate to DevTools Protocol Viewer . The methodName parameter is the full name of the method in the {domain}.{method} format. The parametersAsJson parameter is a JSON formatted string containing the parameters for the corresponding method. The Invoke method of the handler is run when the method asynchronously completes. Invoke is run with the return object of the method as a JSON string. This function returns E_INVALIDARG if the methodName is unknown or the parametersAsJson has an error. In the case of such an error, the returnObjectAsJson parameter of the handler will include information about the error. Note even though WebView2 dispatches the CDP messages in the order called, CDP method calls may be processed out of order. If you require CDP methods to run in a particular order, you should wait for the previous method's completed handler to run before calling the next method. If the method is to run in add_NewWindowRequested handler it should be called before the new window is set if the cdp message should affect the initial navigation. If called after setting the NewWindow property, the cdp messages may or may not apply to the initial navigation and may only apply to the subsequent navigation. For more details see ICoreWebView2NewWindowRequestedEventArgs::put_NewWindow.

```
// Prompt the user for the name and parameters of a CDP method, then call it.
void ScriptComponent::CallCdpMethod()
{
    TextInputDialog dialog(
        m_appWindow->GetMainWindow(),
        L"Call CDP Method",
        L"CDP method name:",
        L"Enter the CDP method name to call, followed by a space,\r\n"
            L"followed by the parameters in JSON format.",
        L"Runtime.evaluate {\"expression\":\"alert(\\\"test\\\")\"}");
    if (dialog.confirmed)
    {
        size_t delimiterPos = dialog.input.find(L' ');
        std::wstring methodName = dialog.input.substr(0, delimiterPos);
        std::wstring methodParams =
            (delimiterPos < dialog.input.size()
                ? dialog.input.substr(delimiterPos + 1)
                : L"{}");

        m_webView->CallDevToolsProtocolMethod(
            methodName.c_str(), methodParams.c_str(),
            Callback<ICoreWebView2CallDevToolsProtocolMethodCompletedHandler>(
                this, &ScriptComponent::CDPMethodCallback)
                .Get());
    }
}
```
---

## CapturePreview

Capture an image of what WebView is displaying.
```
public HRESULT CapturePreview(COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT imageFormat, IStream * imageStream, ICoreWebView2CapturePreviewCompletedHandler * handler)
```
Specify the format of the image with the imageFormat parameter. The resulting image binary data is written to the provided imageStream parameter. When CapturePreview finishes writing to the stream, the Invoke method on the provided handler parameter is run. This method fails if called before the first ContentLoading event. For example if this is called in the NavigationStarting event for the first navigation it will fail. For subsequent navigations, the method may not fail, but will not capture an image of a given webpage until the ContentLoading event has been fired for it. Any call to this method prior to that will result in a capture of the page being navigated away from.

```
// Show the user a file selection dialog, then save a screenshot of the WebView
// to the selected file.
void FileComponent::SaveScreenshot()
{
    WCHAR defaultName[MAX_PATH] = L"WebView2_Screenshot.png";
    OPENFILENAME openFileName = CreateOpenFileName(defaultName, L"PNG File\0*.png\0");
    if (GetSaveFileName(&openFileName))
    {
        wil::com_ptr<IStream> stream;
        CHECK_FAILURE(SHCreateStreamOnFileEx(
            defaultName, STGM_READWRITE | STGM_CREATE, FILE_ATTRIBUTE_NORMAL, TRUE, nullptr,
            &stream));

        CHECK_FAILURE(m_webView->CapturePreview(
            COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT_PNG, stream.get(),
            Callback<ICoreWebView2CapturePreviewCompletedHandler>(
                [appWindow{m_appWindow}](HRESULT error_code) -> HRESULT {
                    CHECK_FAILURE(error_code);
                    appWindow->AsyncMessageBox(L"Preview Captured", L"Preview Captured");
                    return S_OK;
                })
                .Get()));
    }
}
```
---

## ExecuteScript

Run JavaScript code from the javascript parameter in the current top-level document rendered in the WebView.
```
public HRESULT ExecuteScript(LPCWSTR javaScript, ICoreWebView2ExecuteScriptCompletedHandler * handler)
```
The result of evaluating the provided JavaScript is used in this parameter. The result value is a JSON encoded string. If the result is undefined, contains a reference cycle, or otherwise is not able to be encoded into JSON, then the result is considered to be null, which is encoded in JSON as the string "null".

Note
A function that has no explicit return value returns undefined. If the script that was run throws an unhandled exception, then the result is also "null". This method is applied asynchronously. If the method is run after the NavigationStarting event during a navigation, the script runs in the new document when loading it, around the time ContentLoading is run. This operation executes the script even if ICoreWebView2Settings::IsScriptEnabled is set to FALSE.

```
// Prompt the user for some script and then execute it.
void ScriptComponent::InjectScript()
{
    TextInputDialog dialog(
        m_appWindow->GetMainWindow(),
        L"Inject Script",
        L"Enter script code:",
        L"Enter the JavaScript code to run in the webview.",
        L"window.getComputedStyle(document.body).backgroundColor");
    if (dialog.confirmed)
    {
        m_webView->ExecuteScript(dialog.input.c_str(),
            Callback<ICoreWebView2ExecuteScriptCompletedHandler>(
                [appWindow = m_appWindow](HRESULT error, PCWSTR result) -> HRESULT
        {
            if (error != S_OK) {
                ShowFailure(error, L"ExecuteScript failed");
            }
            appWindow->AsyncMessageBox(result, L"ExecuteScript Result");
            return S_OK;
        }).Get());
    }
}
```
---

## get_BrowserProcessId

The process ID of the browser process that hosts the WebView.
```
public HRESULT get_BrowserProcessId(UINT32 * value)
```

## get_CanGoBack

TRUE if the WebView is able to navigate to a previous page in the navigation history.
```
public HRESULT get_CanGoBack(BOOL * canGoBack)
```
If CanGoBack changes value, the HistoryChanged event runs.

---

## get_CanGoForward

TRUE if the WebView is able to navigate to a next page in the navigation history.
```
public HRESULT get_CanGoForward(BOOL * canGoForward)
```
If CanGoForward changes value, the HistoryChanged event runs.

---

## get_ContainsFullScreenElement

Indicates if the WebView contains a fullscreen HTML element.
```
public HRESULT get_ContainsFullScreenElement(BOOL * containsFullScreenElement)
```
---

## get_DocumentTitle

The title for the current top-level document.
```
public HRESULT get_DocumentTitle(LPWSTR * title)
```
If the document has no explicit title or is otherwise empty, a default that may or may not match the URI of the document is used.

The caller must free the returned string with CoTaskMemFree. See API Conventions.

---

## get_Settings

The ICoreWebView2Settings object contains various modifiable settings for the running WebView.
```
public HRESULT get_Settings(ICoreWebView2Settings ** settings)
```

## get_Source

The URI of the current top level document.
```
public HRESULT get_Source(LPWSTR * uri)
```
This value potentially changes as a part of the SourceChanged event that runs for some cases such as navigating to a different site or fragment navigations. It remains the same for other types of navigations such as page refreshes or history.pushState with the same URL as the current page.

The caller must free the returned string with CoTaskMemFree. See API Conventions.

```
// Register a handler for the SourceChanged event.
    // This handler will read the webview's source URI and update
    // the app's address bar.
    CHECK_FAILURE(m_webView->add_SourceChanged(
        Callback<ICoreWebView2SourceChangedEventHandler>(
            [this](ICoreWebView2* sender, ICoreWebView2SourceChangedEventArgs* args)
                -> HRESULT {
                wil::unique_cotaskmem_string uri;
                sender->get_Source(&uri);
                if (wcscmp(uri.get(), L"about:blank") == 0)
                {
                    uri = wil::make_cotaskmem_string(L"");
                }
                SetWindowText(GetAddressBar(), uri.get());

                return S_OK;
            })
            .Get(),
        &m_sourceChangedToken));
```
---

## GetDevToolsProtocolEventReceiver

Get a DevTools Protocol event receiver that allows you to subscribe to a DevTools Protocol event.
```
public HRESULT GetDevToolsProtocolEventReceiver(LPCWSTR eventName, ICoreWebView2DevToolsProtocolEventReceiver ** receiver)
```
The eventName parameter is the full name of the event in the format {domain}.{event}. For more information about DevTools Protocol events description and event args, navigate to DevTools Protocol Viewer.

```
// Prompt the user to name a CDP event, and then subscribe to that event.
void ScriptComponent::SubscribeToCdpEvent()
{
    TextInputDialog dialog(
        m_appWindow->GetMainWindow(),
        L"Subscribe to CDP Event",
        L"CDP event name:",
        L"Enter the name of the CDP event to subscribe to.\r\n"
            L"You may also have to call the \"enable\" method of the\r\n"
            L"event's domain to receive events (for example \"Log.enable\").\r\n",
        L"Log.entryAdded");
    if (dialog.confirmed)
    {
        std::wstring eventName = dialog.input;
        wil::com_ptr<ICoreWebView2DevToolsProtocolEventReceiver> receiver;
        CHECK_FAILURE(
            m_webView->GetDevToolsProtocolEventReceiver(eventName.c_str(), &receiver));

        // If we are already subscribed to this event, unsubscribe first.
        auto preexistingToken = m_devToolsProtocolEventReceivedTokenMap.find(eventName);
        if (preexistingToken != m_devToolsProtocolEventReceivedTokenMap.end())
        {
            CHECK_FAILURE(receiver->remove_DevToolsProtocolEventReceived(
                preexistingToken->second));
        }

        CHECK_FAILURE(receiver->add_DevToolsProtocolEventReceived(
            Callback<ICoreWebView2DevToolsProtocolEventReceivedEventHandler>(
                [this, eventName](
                    ICoreWebView2* sender,
                    ICoreWebView2DevToolsProtocolEventReceivedEventArgs* args) -> HRESULT
                {
                    wil::unique_cotaskmem_string parameterObjectAsJson;
                    CHECK_FAILURE(args->get_ParameterObjectAsJson(&parameterObjectAsJson));
                    std::wstring title = eventName;
                    std::wstring details = parameterObjectAsJson.get();
                    wil::com_ptr<ICoreWebView2DevToolsProtocolEventReceivedEventArgs2> args2;
                    if (SUCCEEDED(args->QueryInterface(IID_PPV_ARGS(&args2))))
                    {
                        wil::unique_cotaskmem_string sessionId;
                        CHECK_FAILURE(args2->get_SessionId(&sessionId));
                        if (sessionId.get() && *sessionId.get())
                        {
                            title = eventName + L" (session:" + sessionId.get() + L")";
                            std::wstring targetId = m_devToolsSessionMap[sessionId.get()];
                            std::wstring targetLabel = m_devToolsTargetLabelMap[targetId];
                            details = L"From " + targetLabel + L" (session:" + sessionId.get() +
                                      L")\r\n" + details;
                        }
                    }
                    m_appWindow->AsyncMessageBox(details, L"CDP Event Fired: " + title);
                    return S_OK;
                })
                .Get(),
            &m_devToolsProtocolEventReceivedTokenMap[eventName]));
    }
}
```

## GoBack

Navigates the WebView to the previous page in the navigation history.
```
public HRESULT GoBack()
```

## GoForward

Navigates the WebView to the next page in the navigation history.
```
public HRESULT GoForward()
```

## Navigate

Cause a navigation of the top-level document to run to the specified URI.
```
public HRESULT Navigate(LPCWSTR uri)
```
For more information, navigate to Navigation events.

Note
This operation starts a navigation and the corresponding NavigationStarting event triggers sometime after Navigate runs.

```
void ControlComponent::NavigateToAddressBar()
{
    int length = GetWindowTextLength(GetAddressBar());
    std::wstring uri(length, 0);
    PWSTR buffer = const_cast<PWSTR>(uri.data());
    GetWindowText(GetAddressBar(), buffer, length + 1);

    HRESULT hr = m_webView->Navigate(uri.c_str());
    if (hr == E_INVALIDARG)
    {
        // An invalid URI was provided.
        if (uri.find(L' ') == std::wstring::npos
            && uri.find(L'.') != std::wstring::npos)
        {
            // If it contains a dot and no spaces, try tacking http:// on the front.
            hr = m_webView->Navigate((L"http://" + uri).c_str());
        }
        else
        {
            // Otherwise treat it as a web search.
            std::wstring urlEscaped(2048, ' ');
            DWORD dwEscaped = (DWORD)urlEscaped.length();
            UrlEscapeW(uri.c_str(), &urlEscaped[0], &dwEscaped, URL_ESCAPE_ASCII_URI_COMPONENT);
            hr = m_webView->Navigate(
                (L"https://bing.com/search?q=" +
                 std::regex_replace(urlEscaped, std::wregex(L"(?:%20)+"), L"+"))
                    .c_str());
        }
    }
    if (hr != E_INVALIDARG) {
        CHECK_FAILURE(hr);
    }
}
```
---

@@ NavigateToString

Initiates a navigation to htmlContent as source HTML of a new document.
```
public HRESULT NavigateToString(LPCWSTR htmlContent)
```
The htmlContent parameter may not be larger than 2 MB (2 * 1024 * 1024 bytes) in total size. The origin of the new page is about:blank.

```
SetVirtualHostNameToFolderMapping(
    L"appassets.example", L"assets",
    COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND_DENY);

WCHAR c_navString[] = LR"
<head><link rel='stylesheet' href ='http://appassets.example/wv2.css'/></head>
<body>
  <img src='http://appassets.example/wv2.png' />
  <p><a href='http://appassets.example/winrt_test.txt'>Click me</a></p>
</body>";
m_webView->NavigateToString(c_navString);
```
```
static const PCWSTR htmlContent =
                        L"<h1>Domain Blocked</h1>"
                        L"<p>You've attempted to navigate to a domain in the blocked "
                        L"sites list. Press back to return to the previous page.</p>";
                    CHECK_FAILURE(sender->NavigateToString(htmlContent));
```
---

## OpenDevToolsWindow

Opens the DevTools window for the current document in the WebView.
```
public HRESULT OpenDevToolsWindow()
```
Does nothing if run when the DevTools window is already open.

---

## PostWebMessageAsJson

Post the specified webMessage to the top level document in this WebView.
```
public HRESULT PostWebMessageAsJson(LPCWSTR webMessageAsJson)
```
The main page receives the message by subscribing to the message event of the window.chrome.webview of the page document.
```
window.chrome.webview.addEventListener('message', handler)
window.chrome.webview.removeEventListener('message', handler)
```
The event args is an instance of MessageEvent. The ICoreWebView2Settings::IsWebMessageEnabled setting must be TRUE or the web message will not be sent. The data property of the event arg is the webMessage string parameter parsed as a JSON string into a JavaScript object. The source property of the event arg is a reference to the window.chrome.webview object. For information about sending messages from the HTML document in the WebView to the host, navigate to add_WebMessageReceived. The message is delivered asynchronously. If a navigation occurs before the message is posted to the page, the message is discarded.

```
// Setup the web message received event handler before navigating to
    // ensure we don't miss any messages.
    CHECK_FAILURE(m_webView->add_WebMessageReceived(
        Microsoft::WRL::Callback<ICoreWebView2WebMessageReceivedEventHandler>(
            [this](ICoreWebView2* sender, ICoreWebView2WebMessageReceivedEventArgs* args)
    {
        wil::unique_cotaskmem_string uri;
        CHECK_FAILURE(args->get_Source(&uri));

        // Always validate that the origin of the message is what you expect.
        if (uri.get() != m_sampleUri)
        {
            // Ignore messages from untrusted sources.
            return S_OK;
        }
        wil::unique_cotaskmem_string messageRaw;
        HRESULT hr = args->TryGetWebMessageAsString(&messageRaw);
        if (hr == E_INVALIDARG)
        {
            // Was not a string message. Ignore.
            return S_OK;
        }
        // Any other problems are fatal.
        CHECK_FAILURE(hr);
        std::wstring message = messageRaw.get();

        if (message.compare(0, 13, L"SetTitleText ") == 0)
        {
            m_appWindow->SetDocumentTitle(message.substr(13).c_str());
        }
        else if (message.compare(L"GetWindowBounds") == 0)
        {
            RECT bounds = m_appWindow->GetWindowBounds();
            std::wstring reply =
                L"{\"WindowBounds\":\"Left:" + std::to_wstring(bounds.left)
                + L"\\nTop:" + std::to_wstring(bounds.top)
                + L"\\nRight:" + std::to_wstring(bounds.right)
                + L"\\nBottom:" + std::to_wstring(bounds.bottom)
                + L"\"}";
            CHECK_FAILURE(sender->PostWebMessageAsJson(reply.c_str()));
        }
        else
        {
            // Ignore unrecognized messages, but log for further investigation
            // since it suggests a mismatch between the web content and the host.
            OutputDebugString(
                std::wstring(L"Unexpected message from main page:" + message).c_str());
        }
        return S_OK;
    }).Get(), &m_webMessageReceivedToken));
```
---

## PostWebMessageAsString

Posts a message that is a simple string rather than a JSON string representation of a JavaScript object.
```
public HRESULT PostWebMessageAsString(LPCWSTR webMessageAsString)
```
This behaves in exactly the same manner as PostWebMessageAsJson, but the data property of the event arg of the window.chrome.webview message is a string with the same value as webMessageAsString. Use this instead of PostWebMessageAsJson if you want to communicate using simple strings rather than JSON objects.

---

## Reload

Reload the current page.
```
public HRESULT Reload()
```
This is similar to navigating to the URI of current top level document including all navigation events firing and respecting any entries in the HTTP cache. But, the back or forward history are not modified.

---

## remove_ContainsFullScreenElementChanged

Remove an event handler previously added with add_ContainsFullScreenElementChanged.
```
public HRESULT remove_ContainsFullScreenElementChanged(EventRegistrationToken token)
```
---

## remove_ContentLoading

Remove an event handler previously added with add_ContentLoading.
```
public HRESULT remove_ContentLoading(EventRegistrationToken token)
```
---

# remove_DocumentTitleChanged

Remove an event handler previously added with add_DocumentTitleChanged.
```
public HRESULT remove_DocumentTitleChanged(EventRegistrationToken token)
```
---

## remove_FrameNavigationCompleted
Remove an event handler previously added with add_FrameNavigationCompleted.
```
public HRESULT remove_FrameNavigationCompleted(EventRegistrationToken token)
```
---

## remove_FrameNavigationStarting

Remove an event handler previously added with add_FrameNavigationStarting.
```
public HRESULT remove_FrameNavigationStarting(EventRegistrationToken token)
```

## remove_HistoryChanged

Remove an event handler previously added with add_HistoryChanged.
```
public HRESULT remove_HistoryChanged(EventRegistrationToken token)
```

## remove_NavigationCompleted

Remove an event handler previously added with add_NavigationCompleted.
```
public HRESULT remove_NavigationCompleted(EventRegistrationToken token)
```
---

## remove_NavigationStarting

Remove an event handler previously added with add_NavigationStarting.
```
public HRESULT remove_NavigationStarting(EventRegistrationToken token)
```
---

## remove_NewWindowRequested

Remove an event handler previously added with add_NewWindowRequested.
```
public HRESULT remove_NewWindowRequested(EventRegistrationToken token)
```
---

## remove_PermissionRequested

Remove an event handler previously added with add_PermissionRequested.
```
public HRESULT remove_PermissionRequested(EventRegistrationToken token)
```
---

## remove_ProcessFailed

Remove an event handler previously added with add_ProcessFailed.
```
public HRESULT remove_ProcessFailed(EventRegistrationToken token)
```

## remove_ScriptDialogOpening

Remove an event handler previously added with add_ScriptDialogOpening.
```
public HRESULT remove_ScriptDialogOpening(EventRegistrationToken token)
```
---

## remove_SourceChanged

Remove an event handler previously added with add_SourceChanged.
```
public HRESULT remove_SourceChanged(EventRegistrationToken token)
```
---

## remove_WebMessageReceived

Remove an event handler previously added with add_WebMessageReceived.
```
public HRESULT remove_WebMessageReceived(EventRegistrationToken token)
```
---

## remove_WebResourceRequested

Remove an event handler previously added with add_WebResourceRequested.
```
public HRESULT remove_WebResourceRequested(EventRegistrationToken token)
```
---

## remove_WindowCloseRequested

Remove an event handler previously added with add_WindowCloseRequested.
```
public HRESULT remove_WindowCloseRequested(EventRegistrationToken token)
```
---

## RemoveHostObjectFromScript

Remove the host object specified by the name so that it is no longer accessible from JavaScript code in the WebView.
```
public HRESULT RemoveHostObjectFromScript(LPCWSTR name)
```
While new access attempts are denied, if the object is already obtained by JavaScript code in the WebView, the JavaScript code continues to have access to that object. Run this method for a name that is already removed or never added fails.

---

## RemoveScriptToExecuteOnDocumentCreated

Remove the corresponding JavaScript added using AddScriptToExecuteOnDocumentCreated with the specified script ID.
```
public HRESULT RemoveScriptToExecuteOnDocumentCreated(LPCWSTR id)
```
The script ID should be the one returned by the AddScriptToExecuteOnDocumentCreated. Both use AddScriptToExecuteOnDocumentCreated and this method in NewWindowRequested event handler at the same time sometimes causes trouble. Since invalid scripts will be ignored, the script IDs you got may not be valid anymore.

---

## RemoveWebResourceRequestedFilter

Warning: This method and AddWebResourceRequestedFilter are deprecated.
```
public HRESULT RemoveWebResourceRequestedFilter(LPCWSTR const uri, COREWEBVIEW2_WEB_RESOURCE_CONTEXT const resourceContext)
```
Please use AddWebResourceRequestedFilterWithRequestSourceKinds and RemoveWebResourceRequestedFilterWithRequestSourceKinds instead.

Removes a matching WebResource filter that was previously added for the WebResourceRequested event. If the same filter was added multiple times, then it must be removed as many times as it was added for the removal to be effective. Returns E_INVALIDARG for a filter that was never added.

---

## Stop

Stop all navigations and pending resource fetches. Does not stop scripts.
```
public HRESULT Stop()
```
---

Minimal bring-up checklist
Create environment (evergreen or with options): Call CreateCoreWebView2Environment[WithOptions|WithDetails], then in the completed handler, create the controller. Your AfxWebView2.bi already resolves these loader exports dynamically from WebView2Loader.dll, so you can stay dependency-light

Create controller and get core: In the environment-completed callback, call CreateCoreWebView2Controller(hwnd), then get ICoreWebView2 via get_CoreWebView2. Remember: controller owns the WebView; use put_Bounds on resize.

Enable messaging early: Before navigating, add_WebMessageReceived; handle JSON via your reader. JS posts with window.chrome.webview.postMessage(obj), host replies with PostWebMessageAsJson. This keeps your JSON include front and center.

Navigate and verify: Use Navigate or NavigateToString for a first-page sanity check, then ExecuteScript to confirm round-trips. ExecuteScript returns a JSON-encoded result string in the completion handler

Size and visibility: On WM_SIZE, update bounds; minimize → put_IsVisible(FALSE), restore → TRUE. You’ll avoid needless CPU/GPU while hidden.

Tear-down explicitly: Call controller->Close() when done; it releases handlers and shuts down the browser instance if no others remain

Callback classes to implement first
Environment creation: ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler (Invoke(HRESULT, ICoreWebView2Environment*))

Controller creation: ICoreWebView2CreateCoreWebView2ControllerCompletedHandler (Invoke(HRESULT, ICoreWebView2Controller*)).

Messaging bridge: ICoreWebView2WebMessageReceivedEventHandler (Invoke(ICoreWebView2, ICoreWebView2WebMessageReceivedEventArgs))

Navigation status (optional but useful): ICoreWebView2NavigationStartingEventHandler, ICoreWebView2NavigationCompletedEventHandler for diagnostics and gating behavior.

Title/source updates (quality-of-life): ICoreWebView2DocumentTitleChangedEventHandler, ICoreWebView2SourceChangedEventHandler.

Your .bi already defines the interface IIDs and abstract method signatures; you just need to supply the vtables/refcount and wire AddRef/Release/QueryInterface correctly

Lifetime, threading, and tokens
Apartment model: Initialize COM as STA on the UI thread and keep all WebView2 calls on that thread; callbacks come back there.

Event tokens: Store EventRegistrationToken values returned by add_* and remove them during cleanup to avoid dangling subscriptions.

Controller ownership and Close():
Close is synchronous; it releases associated event handlers and lets the browser instance shut down when no longer in use. Call it explicitly to break potential reference cycles.

Visibility and movement: Use put_IsVisible on minimize/restore and NotifyParentWindowPositionChanged on WM_MOVE/WM_MOVING to keep accessibility/tooltips right.

Smoke tests to prove the bridge
Hello DOM: NavigateToString(L"<html><body><h1>FB + WV2</h1></body></html>"), then add_DocumentTitleChanged to confirm title updates.

JS → native:
In page JS: window.chrome.webview.postMessage({ ping: 1 }). In your handler, parse webMessageAsJson and log “ping” with your reader; reply with PostWebMessageAsJson using your writer.

Native → JS: PostWebMessageAsJson({ hello:"from native" }) and in JS, window.chrome.webview.addEventListener('message', e => console.log(e.data.hello)).

ExecuteScript round-trip: ExecuteScript("JSON.stringify({ sum: 2+2 })"), then decode resultObjectAsJson; you should see {"sum":4} as the JSON string

Gotchas you can dodge
Parent HWND lifetime: If parentWindow is destroyed before controller creation finishes, the creation callback returns E_ABORT; structure your shutdown to handle that path cleanly

Message ordering footnote: Multiple frames posting messages aren’t strictly ordered with other events; if you need sequencing, round-trip an ACK from native back to JS before continuing

Runtime presence: If the loader can’t find the runtime (ERROR_FILE_NOT_FOUND), surface a friendly prompt; your loader shim already makes this check trivial to branch on1

COM event handler class in FreeBASIC

' ============================================================================
' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler implementation
' ============================================================================
Type EnvCompletedHandler
    lpVtbl As Any Ptr  ' pointer to vtable
    refCount As ULong

    ' --- IUnknown methods ---
    Declare Function QueryInterface(ByVal this As EnvCompletedHandler Ptr, _
                                    ByVal riid As REFIID, _
                                    ByVal ppv As Any Ptr Ptr) As HRESULT
    Declare Function AddRef(ByVal this As EnvCompletedHandler Ptr) As ULong
    Declare Function Release(ByVal this As EnvCompletedHandler Ptr) As ULong

    ' --- Invoke method from this interface ---
    Declare Function Invoke(ByVal this As EnvCompletedHandler Ptr, _
                            ByVal result As HRESULT, _
                            ByVal env As ICoreWebView2Environment Ptr) As HRESULT
End Type

' ----------------------------------------------------------------------------
' Forward declarations
' ----------------------------------------------------------------------------
Function EnvCompletedHandler_QueryInterface( _
    ByVal this As EnvCompletedHandler Ptr, _
    ByVal riid As REFIID, _
    ByVal ppv As Any Ptr Ptr) As HRESULT

Function EnvCompletedHandler_AddRef(ByVal this As EnvCompletedHandler Ptr) As ULong
Function EnvCompletedHandler_Release(ByVal this As EnvCompletedHandler Ptr) As ULong

Function EnvCompletedHandler_Invoke( _
    ByVal this As EnvCompletedHandler Ptr, _
    ByVal result As HRESULT, _
    ByVal env As ICoreWebView2Environment Ptr) As HRESULT

' ----------------------------------------------------------------------------
' VTable for EnvCompletedHandler
' ----------------------------------------------------------------------------
Dim Shared EnvCompletedHandler_Vtbl As Const ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandlerVtbl = ( _
    @EnvCompletedHandler_QueryInterface, _
    @EnvCompletedHandler_AddRef, _
    @EnvCompletedHandler_Release, _
    @EnvCompletedHandler_Invoke _
)

' ----------------------------------------------------------------------------
' Constructor: allocate + hook vtable + refcount
' ----------------------------------------------------------------------------
Function New_EnvCompletedHandler() As EnvCompletedHandler Ptr
    Dim h As EnvCompletedHandler Ptr = Callocate(1, SizeOf(EnvCompletedHandler))
    h->lpVtbl = @EnvCompletedHandler_Vtbl
    h->refCount = 1
    Return h
End Function

```
' ============================================================================
' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler implementation
' ============================================================================
Type EnvCompletedHandler
    lpVtbl As Any Ptr  ' pointer to vtable
    refCount As ULong

    ' --- IUnknown methods ---
    Declare Function QueryInterface(ByVal this As EnvCompletedHandler Ptr, _
                                    ByVal riid As REFIID, _
                                    ByVal ppv As Any Ptr Ptr) As HRESULT
    Declare Function AddRef(ByVal this As EnvCompletedHandler Ptr) As ULong
    Declare Function Release(ByVal this As EnvCompletedHandler Ptr) As ULong

    ' --- Invoke method from this interface ---
    Declare Function Invoke(ByVal this As EnvCompletedHandler Ptr, _
                            ByVal result As HRESULT, _
                            ByVal env As ICoreWebView2Environment Ptr) As HRESULT
End Type

' ----------------------------------------------------------------------------
' Forward declarations
' ----------------------------------------------------------------------------
Function EnvCompletedHandler_QueryInterface( _
    ByVal this As EnvCompletedHandler Ptr, _
    ByVal riid As REFIID, _
    ByVal ppv As Any Ptr Ptr) As HRESULT

Function EnvCompletedHandler_AddRef(ByVal this As EnvCompletedHandler Ptr) As ULong
Function EnvCompletedHandler_Release(ByVal this As EnvCompletedHandler Ptr) As ULong

Function EnvCompletedHandler_Invoke( _
    ByVal this As EnvCompletedHandler Ptr, _
    ByVal result As HRESULT, _
    ByVal env As ICoreWebView2Environment Ptr) As HRESULT

' ----------------------------------------------------------------------------
' VTable for EnvCompletedHandler
' ----------------------------------------------------------------------------
Dim Shared EnvCompletedHandler_Vtbl As Const ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandlerVtbl = ( _
    @EnvCompletedHandler_QueryInterface, _
    @EnvCompletedHandler_AddRef, _
    @EnvCompletedHandler_Release, _
    @EnvCompletedHandler_Invoke _
)

' ----------------------------------------------------------------------------
' Constructor: allocate + hook vtable + refcount
' ----------------------------------------------------------------------------
Function New_EnvCompletedHandler() As EnvCompletedHandler Ptr
    Dim h As EnvCompletedHandler Ptr = Callocate(1, SizeOf(EnvCompletedHandler))
    h->lpVtbl = @EnvCompletedHandler_Vtbl
    h->refCount = 1
    Return h
End Function

' ============================================================================
' IUnknown implementation
' ============================================================================
Function EnvCompletedHandler_QueryInterface( _
    ByVal this As EnvCompletedHandler Ptr, _
    ByVal riid As REFIID, _
    ByVal ppv As Any Ptr Ptr) As HRESULT

    If ppv = 0 Then Return E_POINTER
    *ppv = 0

    ' Compare riid to IID_IUnknown or IID_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
    If IsEqualIID(riid, @IID_IUnknown) OrElse _
       IsEqualIID(riid, @IID_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler) Then
        EnvCompletedHandler_AddRef(this)
        *ppv = this
        Return S_OK
    End If
    Return E_NOINTERFACE
End Function

Function EnvCompletedHandler_AddRef(ByVal this As EnvCompletedHandler Ptr) As ULong
    this->refCount += 1
    Return this->refCount
End Function

Function EnvCompletedHandler_Release(ByVal this As EnvCompletedHandler Ptr) As ULong
    this->refCount -= 1
    Dim c As ULong = this->refCount
    If c = 0 Then Deallocate(this)
    Return c
End Function

' ============================================================================
' Invoke: called when the environment has been created
' ============================================================================
Function EnvCompletedHandler_Invoke( _
    ByVal this As EnvCompletedHandler Ptr, _
    ByVal result As HRESULT, _
    ByVal env As ICoreWebView2Environment Ptr) As HRESULT

    If FAILED(result) Or env = 0 Then
        Print "Environment creation failed: "; Hex(result)
        Return result
    End If

    ' From here you can call env->CreateCoreWebView2Controller(...) with your own
    ' ControllerCompletedHandler instance.
    Print "Environment ready!"
    Return S_OK
End Function
```

How to adapt for other handlers
Change the type/interface — e.g., for controller creation, base your type/vtable on ICoreWebView2CreateCoreWebView2ControllerCompletedHandlerVtbl and adjust Invoke parameters.

Keep the IUnknown trio identical — QueryInterface, AddRef, Release differ only in the IID you accept.

Put your work inside Invoke — for message events, you’d retrieve the message via get_WebMessageAsJson, parse it with your reader, and respond.

Use the same constructor pattern — one New_XxxHandler() per class.

Here’s a side‑by‑side FreeBASIC template for:

EnvironmentCompletedHandler

ControllerCompletedHandler

WebMessageReceivedHandler

NavigationCompletedHandler

```
' ============================================================================
' 1) EnvironmentCompletedHandler
' ============================================================================
Type EnvironmentCompletedHandler
    lpVtbl As Any Ptr
    refCount As ULong
End Type

Declare Function Env_QueryInterface(ByVal this As EnvironmentCompletedHandler Ptr, _
                                    ByVal riid As REFIID, _
                                    ByVal ppv As Any Ptr Ptr) As HRESULT
Declare Function Env_AddRef(ByVal this As EnvironmentCompletedHandler Ptr) As ULong
Declare Function Env_Release(ByVal this As EnvironmentCompletedHandler Ptr) As ULong
Declare Function Env_Invoke(ByVal this As EnvironmentCompletedHandler Ptr, _
                            ByVal result As HRESULT, _
                            ByVal env As ICoreWebView2Environment Ptr) As HRESULT

Dim Shared Env_Vtbl As Const ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandlerVtbl = ( _
    @Env_QueryInterface, @Env_AddRef, @Env_Release, @Env_Invoke _
)

Function New_EnvHandler() As EnvironmentCompletedHandler Ptr
    Dim h = Callocate(1, SizeOf(EnvironmentCompletedHandler))
    h->lpVtbl = @Env_Vtbl : h->refCount = 1
    Return h
End Function

Function Env_QueryInterface(...) As HRESULT
    ' check IID_IUnknown / IID_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
End Function
Function Env_AddRef(...) As ULong : this->refCount += 1 : Return this->refCount : End Function
Function Env_Release(...) As ULong : this->refCount -= 1 : If this->refCount=0 Then Deallocate(this) : Return this->refCount : End Function
Function Env_Invoke(...) As HRESULT
    Print "Env ready" : Return S_OK
End Function

' ============================================================================
' 2) ControllerCompletedHandler
' ============================================================================
Type ControllerCompletedHandler
    lpVtbl As Any Ptr
    refCount As ULong
End Type

Declare Function Ctrl_QueryInterface(ByVal this As ControllerCompletedHandler Ptr, _
                                     ByVal riid As REFIID, _
                                     ByVal ppv As Any Ptr Ptr) As HRESULT
Declare Function Ctrl_AddRef(ByVal this As ControllerCompletedHandler Ptr) As ULong
Declare Function Ctrl_Release(ByVal this As ControllerCompletedHandler Ptr) As ULong
Declare Function Ctrl_Invoke(ByVal this As ControllerCompletedHandler Ptr, _
                             ByVal result As HRESULT, _
                             ByVal controller As ICoreWebView2Controller Ptr) As HRESULT

Dim Shared Ctrl_Vtbl As Const ICoreWebView2CreateCoreWebView2ControllerCompletedHandlerVtbl = ( _
    @Ctrl_QueryInterface, @Ctrl_AddRef, @Ctrl_Release, @Ctrl_Invoke _
)

Function New_CtrlHandler() As ControllerCompletedHandler Ptr
    Dim h = Callocate(1, SizeOf(ControllerCompletedHandler))
    h->lpVtbl = @Ctrl_Vtbl : h->refCount = 1
    Return h
End Function

Function Ctrl_QueryInterface(...) As HRESULT
    ' check IID_IUnknown / IID_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
End Function
Function Ctrl_AddRef(...) As ULong : this->refCount += 1 : Return this->refCount : End Function
Function Ctrl_Release(...) As ULong : this->refCount -= 1 : If this->refCount=0 Then Deallocate(this) : Return this->refCount : End Function
Function Ctrl_Invoke(...) As HRESULT
    Print "Controller ready"
    ' store controller, get CoreWebView2, hook up events
    Return S_OK
End Function

' ============================================================================
' 3) WebMessageReceivedHandler
' ============================================================================
Type MessageReceivedHandler
    lpVtbl As Any Ptr
    refCount As ULong
End Type

Declare Function Msg_QueryInterface(ByVal this As MessageReceivedHandler Ptr, _
                                    ByVal riid As REFIID, _
                                    ByVal ppv As Any Ptr Ptr) As HRESULT
Declare Function Msg_AddRef(ByVal this As MessageReceivedHandler Ptr) As ULong
Declare Function Msg_Release(ByVal this As MessageReceivedHandler Ptr) As ULong
Declare Function Msg_Invoke(ByVal this As MessageReceivedHandler Ptr, _
                            ByVal sender As ICoreWebView2 Ptr, _
                            ByVal args As ICoreWebView2WebMessageReceivedEventArgs Ptr) As HRESULT

Dim Shared Msg_Vtbl As Const ICoreWebView2WebMessageReceivedEventHandlerVtbl = ( _
    @Msg_QueryInterface, @Msg_AddRef, @Msg_Release, @Msg_Invoke _
)

Function New_MsgHandler() As MessageReceivedHandler Ptr
    Dim h = Callocate(1, SizeOf(MessageReceivedHandler))
    h->lpVtbl = @Msg_Vtbl : h->refCount = 1
    Return h
End Function

Function Msg_QueryInterface(...) As HRESULT
    ' check IID_IUnknown / IID_ICoreWebView2WebMessageReceivedEventHandler
End Function
Function Msg_AddRef(...) As ULong : this->refCount += 1 : Return this->refCount : End Function
Function Msg_Release(...) As ULong : this->refCount -= 1 : If this->refCount=0 Then Deallocate(this) : Return this->refCount : End Function
Function Msg_Invoke(...) As HRESULT
    ' extract JSON string from args, parse, respond
    Return S_OK
End Function

' ============================================================================
' 4) NavigationCompletedHandler
' ============================================================================
Type NavigationCompletedHandler
    lpVtbl As Any Ptr
    refCount As ULong
End Type

Declare Function Nav_QueryInterface(ByVal this As NavigationCompletedHandler Ptr, _
                                    ByVal riid As REFIID, _
                                    ByVal ppv As Any Ptr Ptr) As HRESULT
Declare Function Nav_AddRef(ByVal this As NavigationCompletedHandler Ptr) As ULong
Declare Function Nav_Release(ByVal this As NavigationCompletedHandler Ptr) As ULong
Declare Function Nav_Invoke(ByVal this As NavigationCompletedHandler Ptr, _
                            ByVal sender As ICoreWebView2 Ptr, _
                            ByVal args As ICoreWebView2NavigationCompletedEventArgs Ptr) As HRESULT

Dim Shared Nav_Vtbl As Const ICoreWebView2NavigationCompletedEventHandlerVtbl = ( _
    @Nav_QueryInterface, @Nav_AddRef, @Nav_Release, @Nav_Invoke _
)

Function New_NavHandler() As NavigationCompletedHandler Ptr
    Dim h = Callocate(1, SizeOf(NavigationCompletedHandler))
    h->lpVtbl = @Nav_Vtbl : h->refCount = 1
    Return h
End Function

Function Nav_QueryInterface(...) As HRESULT
    ' check IID_IUnknown / IID_ICoreWebView2NavigationCompletedEventHandler
End Function
Function Nav_AddRef(...) As ULong : this->refCount += 1 : Return this->refCount : End Function
Function Nav_Release(...) As ULong : this->refCount -= 1 : If this->refCount=0 Then Deallocate(this) : Return this->refCount : End Function
Function Nav_Invoke(...) As HRESULT
    ' check navigation status, maybe inject JS bridge
    Return S_OK
End Function

```

Why this helps
With these stubs in place, all COM ceremony is out of the way. You just:

Swap the vtable type and Invoke signature per interface.

Fill the Invoke body with your logic.

Call New_XxxHandler() wherever WebView2 expects that callback.

If you like, I can next illustrate the glue code that chains them: Environment → Controller → Event wires, so your app launches straight into a loaded webview. That would make this a complete drop‑in starter.

Here’s a minimal bootstrap showing the flow: create environment → create controller → hook events → run.

```
' ============================================================================
' Bootstrap for WebView2 in FreeBASIC using our handler templates
' ============================================================================

Sub InitWebView2()
    ' Kick off the async environment creation.
    ' Pass our Env handler; its Invoke will call for a Controller.
    Dim hr As HRESULT
    hr = CreateCoreWebView2EnvironmentWithOptions( _
             0, 0, 0, _
             New_EnvHandler())
    If FAILED(hr) Then
        Print "CreateCoreWebView2EnvironmentWithOptions failed: "; Hex(hr)
    End If
End Sub

' Example Env_Invoke chaining to Controller creation
Function Env_Invoke( _
    ByVal this As EnvironmentCompletedHandler Ptr, _
    ByVal result As HRESULT, _
    ByVal env As ICoreWebView2Environment Ptr) As HRESULT

    If FAILED(result) Or env = 0 Then
        Print "Environment creation failed: "; Hex(result)
        Return result
    End If

    Print "Environment ready"
    ' Create a WebView2 controller for our window (replace hWnd with yours)
    env->lpVtbl->CreateCoreWebView2Controller(env, hWnd, New_CtrlHandler())
    Return S_OK
End Function

' Example Ctrl_Invoke chaining to event hookup
Function Ctrl_Invoke( _
    ByVal this As ControllerCompletedHandler Ptr, _
    ByVal result As HRESULT, _
    ByVal controller As ICoreWebView2Controller Ptr) As HRESULT

    If FAILED(result) Or controller = 0 Then
        Print "Controller creation failed: "; Hex(result)
        Return result
    End If

    Print "Controller ready"
    Dim webview As ICoreWebView2 Ptr
    controller->lpVtbl->get_CoreWebView2(controller, @webview)

    ' Hook WebMessageReceived
    Dim token As EventRegistrationToken
    webview->lpVtbl->add_WebMessageReceived(webview, New_MsgHandler(), @token)

    ' Hook NavigationCompleted
    Dim navToken As EventRegistrationToken
    webview->lpVtbl->add_NavigationCompleted(webview, New_NavHandler(), @navToken)

    ' Navigate somewhere to test
    webview->lpVtbl->Navigate(webview, WStr("https://example.com"))

    Return S_OK
End Function
```

How it plays out at runtime

InitWebView2 kicks off environment creation.

Env_Invoke fires → spins up a controller.

Ctrl_Invoke gets you the WebView2 core → attaches your message + navigation events.

From here, your message handler can parse JSON from the web content and send responses back with PostWebMessageAsString/JSON.

Want me to drop in a tiny HTML/JS test page too, so you can instantly see the WebMessageReceived plumbing working? That way, you can fire a button click in the browser and watch it print in your FreeBASIC console.

Here’s a minimal embedded test page (save as test.html or feed to NavigateToString):

```
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>WebView2 → FreeBASIC Bridge</title>
<style>
body { font-family: sans-serif; padding: 2em; }
button { padding: 0.5em 1em; }
</style>
<script>
function sendMessage() {
    const payload = { time: new Date().toISOString(), text: "Hello from JS" };
    window.chrome.webview.postMessage(JSON.stringify(payload));
}
</script>
</head>
<body>
<h1>WebMessage Test</h1>
<p>Click the button to send a JSON payload to the host.</p>
<button onclick="sendMessage()">Send Message</button>
</body>
</html>
```
Hooking it into your handler
Inside your Msg_Invoke:

Function Msg_Invoke( _
    ByVal this As MessageReceivedHandler Ptr, _
    ByVal sender As ICoreWebView2 Ptr, _
    ByVal args As ICoreWebView2WebMessageReceivedEventArgs Ptr) As HRESULT

    Dim raw As AFX_BSTR
    args->lpVtbl->get_WebMessageAsJson(args, @raw)
    Print "Message from JS: " & *raw  ' or parse with your JSON lib
    CoTaskMemFree(raw)
    Return S_OK
End Function
```

Serve or embed test.html.

Launch your bootstrap code — it navigates to the page.

Click the button → postMessage fires → your Msg_Invoke prints the payload.

With that, you’ll see the full handshake: HTML/JS → WebView2 → COM callback → FreeBASIC console.

When you’re ready, we can expand the bridge to handle two‑way chatter so your FB side can call ExecuteScript to update the page or trigger DOM events. That’s where it starts feeling like magic.
