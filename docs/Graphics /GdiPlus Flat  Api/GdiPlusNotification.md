## Notification Functions

| Name       | Description |
| ---------- | ----------- |
| [GdiplusNotificationHook](#gdiplusnotificationhook) | The **GdiplusStartup** function returns (in its output parameter) a pointer to a **GdiplusStartupOutput** structure. One of the members of the structure is a pointer to a notification hook function that has the same signature as **GdiplusNotificationHook**. There are two ways you can call the notification hook function; you can use the pointer returned by **GdiplusStartup** or you can call **GdiplusNotificationHook**. In fact, **GdiplusNotificationHook** simply verifies that you have suppressed the background thread and then calls the notification hook function that is returned by **GdiplusStartup**. The token parameter receives an identifier that you should later pass in a corresponding call to the notification unhook function. |
| [GdiplusNotificationUnhook](#gdiplusnotificationunhook) | The **GdiplusStartup** function returns (in its output parameter) a pointer to a **GdiplusStartupOutput** structure. One of the members of the structure is a pointer to a notification unhook function that has the same signature as **GdiplusNotificationUnhook**. There are two ways you can call the notification unhook function; you can use the pointer returned by **GdiplusStartup** or you can call **GdiplusNotificationUnhook**. In fact, **GdiplusNotificationUnhook** simply verifies that you have suppressed the background thread and then calls the notification unhook function that is returned by **GdiplusStartup**. When you call the notification unhook function, pass the token that you previously received from a corresponding call to the notification hook function. If you do not do this, there will be resource leaks that won't be cleaned up until the process exits. |

---

Example: Using GdiplusNotificationHook and GdiplusNotificationUnhook

```
' Startup input/output structures
DIM startupInput AS GdiplusStartupInput
DIM startupOutput AS GdiplusStartupOutput
DIM gdipToken AS ULONG_PTR

' Suppress background thread
startupInput.GdiplusVersion = 1
startupInput.SuppressBackgroundThread = TRUE

' Start GDI+
GdipStartup(@gdipToken, @startupInput, @startupOutput)

' Hook notifications (optional, rarely needed)
DIM hookToken AS ULONG_PTR
GdiplusNotificationHook(@hookToken)

' ... perform GDI+ operations ...

' Unhook notifications before shutdown
GdiplusNotificationUnhook(hookToken)

' Shutdown GDI+
GdipShutdown(gdipToken)
```

When to Use These

Only when SuppressBackgroundThread = TRUE in GdiplusStartupInput.

They allow manual control over GDI+ notifications, typically in environments where background threads are restricted (e.g., certain embedded systems or plugin hosts).

Most applications do not need to use these functions.

**GDI+ Startup Module with Notification Hook**

```
' ========================================================================================
' GDI+ Startup with optional notification hook/unhook
' ========================================================================================
TYPE GDIP_StartupContext
    gdipToken AS ULONG_PTR
    hookToken AS ULONG_PTR
    graphicsReady AS BOOLEAN
END TYPE

SUB GDIP_StartupWithHook (BYREF ctx AS GDIP_StartupContext)
    DIM startupInput AS GdiplusStartupInput
    DIM startupOutput AS GdiplusStartupOutput

    startupInput.GdiplusVersion = 1
    startupInput.SuppressBackgroundThread = TRUE

    IF GdipStartup(@ctx.gdipToken, @startupInput, @startupOutput) = 0 THEN
        ctx.graphicsReady = TRUE
        GdiplusNotificationHook(@ctx.hookToken)
    END IF
END SUB

SUB GDIP_ShutdownWithUnhook (BYREF ctx AS GDIP_StartupContext)
    IF ctx.graphicsReady THEN
        GdiplusNotificationUnhook(ctx.hookToken)
        GdipShutdown(ctx.gdipToken)
        ctx.graphicsReady = FALSE
    END IF
END SUB
```

**Usage example**

```
DIM ctx AS GDIP_StartupContext
GDIP_StartupWithHook(ctx)

' ... perform GDI+ operations ...

GDIP_ShutdownWithUnhook(ctx)
```

**Notes**

SuppressBackgroundThread = TRUE is required for the hook/unhook functions to work.

This module ensures proper cleanup and avoids resource leaks.

Ideal for plugin environments, embedded systems, or apps with strict thread control.

They’re rarely used in typical applications, but essential in environments where background threads must be avoided—like certain plugin hosts or embedded systems.
