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

---

## add_NewBrowserVersionAvailable

Add an event handler for the NewBrowserVersionAvailable event.

`NewBrowserVersionAvailable` runs when a newer version of the WebView2 Runtime is installed and available using WebView2. To use the newer version of the browser you must create a new environment and WebView. The event only runs for new version from the same WebView2 Runtime from which the code is running. When not running with installed WebView2 Runtime, no event is run.

Because a user data folder is only able to be used by one browser process at a time, if you want to use the same user data folder in the WebView using the new version of the browser, you must close the environment and instance of WebView that are using the older version of the browser first. Or simply prompt the user to restart the app.

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
