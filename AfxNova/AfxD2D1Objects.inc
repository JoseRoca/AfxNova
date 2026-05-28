' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxD2D1Objects.inc
' Purpose:  D2D1 objects
' Compiler: FreeBASIC 32 & 64 bit
' Copyright (c) 2026 José Roca
'
' License: Distributed under the MIT license.
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the “Software”), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit
' persons to whom the Software is furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.

' THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
' PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
' FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
' OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
' DEALINGS IN THE SOFTWARE.'
' ########################################################################################

#pragma once
#include once "AfxNova/AfxD2D1.bi"

' ========================================================================================
' Macro for debug
' To allow debugging, define _D2D1_DEBUG_ 1 in your application before including this file.
' ========================================================================================
#ifndef _D2D1_DEBUG_
   #define _D2D1_DEBUG_ 0
#ENDIF
#ifndef _D2D1_DP_
   #define _D2D1_DP_ 1
   #MACRO D2D1_DP(st)
      #IF (_D2D1_DEBUG_ = 1)
         OutputDebugStringW(__FUNCTION__ + ": " + st)
      #ENDIF
   #ENDMACRO
#ENDIF
' ========================================================================================

NAMESPACE AfxNova

' ########################################################################################
' CD2D1Factory class
' ########################################################################################

' ========================================================================================
TYPE CD2D1Factory
Public:

   m_pFactory AS ID2D1Factory PTR

   DECLARE CONSTRUCTOR (BYVAL factoryType AS D2D1_FACTORY_TYPE = D2D1_FACTORY_TYPE_SINGLE_THREADED)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS ID2D1Factory PTR PTR
   DECLARE OPERATOR CAST () AS ID2D1Factory PTR
   DECLARE SUB GetDesktopDpi (BYREF dpiX AS SINGLE, BYREF dpiY AS SINGLE)
   DECLARE SUB GetDesktopDpi (BYVAL dpiX AS SINGLE PTR, BYVAL dpiY AS SINGLE PTR)
   DECLARE FUNCTION GetDpi () AS SINGLE
   DECLARE FUNCTION GetDpiX () AS SINGLE
   DECLARE FUNCTION GetDpiY () AS SINGLE
   DECLARE FUNCTION ReloadSystemMetrics () AS HRESULT
   DECLARE FUNCTION CreateRectangleGeometry (BYREF rectangle AS D2D1_RECT_F) AS ID2D1RectangleGeometry PTR
   DECLARE FUNCTION CreateRectangleGeometry (BYVAL rectangle AS D2D1_RECT_F PTR) AS ID2D1RectangleGeometry PTR
   DECLARE FUNCTION CreateRoundedRectangleGeometry (BYREF roundedRectangle AS D2D1_ROUNDED_RECT) AS ID2D1RoundedRectangleGeometry PTR
   DECLARE FUNCTION CreateRoundedRectangleGeometry (BYVAL roundedRectangle AS D2D1_ROUNDED_RECT PTR) AS ID2D1RoundedRectangleGeometry PTR
   DECLARE FUNCTION CreateEllipseGeometry (BYREF ellipse AS const D2D1_ELLIPSE) AS ID2D1EllipseGeometry PTR
   DECLARE FUNCTION CreateEllipseGeometry (BYVAL ellipse AS const D2D1_ELLIPSE PTR) AS ID2D1EllipseGeometry PTR
   DECLARE FUNCTION CreateGeometryGroup (BYVAL fillMode AS D2D1_FILL_MODE, BYVAL geometries AS ID2D1Geometry PTR, _
           BYVAL geometriesCount AS UINT32) AS ID2D1GeometryGroup PTR
   DECLARE FUNCTION CreateTransformedGeometry (BYVAL sourceGeometry AS ID2D1Geometry PTR, BYREF transform AS D2D1_MATRIX_3X2_F) AS ID2D1TransformedGeometry PTR
   DECLARE FUNCTION CreateTransformedGeometry (BYVAL sourceGeometry AS ID2D1Geometry PTR, BYVAL transform AS D2D1_MATRIX_3X2_F PTR) AS ID2D1TransformedGeometry PTR
   DECLARE FUNCTION CreatePathGeometry () AS ID2D1PathGeometry PTR
   DECLARE FUNCTION CreateStrokeStyle (BYREF strokeStyleProperties AS D2D1_STROKE_STYLE_PROPERTIES, _
           BYREF dashes AS const FLOAT, BYVAL dashesCount AS UINT32) AS ID2D1StrokeStyle PTR
   DECLARE FUNCTION CreateStrokeStyle (BYVAL strokeStyleProperties AS D2D1_STROKE_STYLE_PROPERTIES PTR, _
           BYVAL dashes AS const FLOAT PTR, BYVAL dashesCount AS UINT32) AS ID2D1StrokeStyle PTR
   DECLARE FUNCTION CreateDrawingStateBlock (BYREF drawingStateDescription AS const D2D1_DRAWING_STATE_DESCRIPTION, _
           BYVAL textRenderingParams AS IDWriteRenderingParams PTR) AS ID2D1DrawingStateBlock PTR
   DECLARE FUNCTION CreateDrawingStateBlock (BYVAL drawingStateDescription AS const D2D1_DRAWING_STATE_DESCRIPTION PTR, _
           BYVAL textRenderingParams AS IDWriteRenderingParams PTR) AS ID2D1DrawingStateBlock PTR
   DECLARE FUNCTION CreateDrawingStateBlock (BYREF drawingStateDescription AS const D2D1_DRAWING_STATE_DESCRIPTION) AS ID2D1DrawingStateBlock PTR
   DECLARE FUNCTION CreateDrawingStateBlock (BYVAL drawingStateDescription AS const D2D1_DRAWING_STATE_DESCRIPTION PTR) AS ID2D1DrawingStateBlock PTR
   DECLARE FUNCTION CreateDrawingStateBlock () AS ID2D1DrawingStateBlock PTR
   DECLARE FUNCTION CreateWicBitmapRenderTarget (BYVAL target AS IWICBitmap PTR, _
           BYREF renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES) AS ID2D1RenderTarget PTR
   DECLARE FUNCTION CreateWicBitmapRenderTarget (BYVAL target AS IWICBitmap PTR, _
           BYVAL renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES PTR) AS ID2D1RenderTarget PTR
   DECLARE FUNCTION CreateHwndRenderTarget (BYREF renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES, _
           BYREF hwndRenderTargetProperties AS D2D1_HWND_RENDER_TARGET_PROPERTIES) AS ID2D1HwndRenderTarget PTR
   DECLARE FUNCTION CreateHwndRenderTarget (BYVAL renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES PTR, _
           BYVAL hwndRenderTargetProperties AS D2D1_HWND_RENDER_TARGET_PROPERTIES PTR) AS ID2D1HwndRenderTarget PTR
   DECLARE FUNCTION CreateDxgiSurfaceRenderTarget (BYVAL dxgiSurface AS IDXGISurface PTR, _
           BYREF renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES) AS ID2D1RenderTarget PTR
   DECLARE FUNCTION CreateDxgiSurfaceRenderTarget (BYVAL dxgiSurface AS IDXGISurface PTR, _
           BYVAL renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES PTR) AS ID2D1RenderTarget PTR
   DECLARE FUNCTION CreateDCRenderTarget (BYREF renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES) AS ID2D1DCRenderTarget PTR
   DECLARE FUNCTION CreateDCRenderTarget (BYVAL renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES PTR) AS ID2D1DCRenderTarget PTR
   DECLARE FUNCTION CreateDCRenderTarget (BYVAL targetType AS D2D1_RENDER_TARGET_TYPE = D2D1_RENDER_TARGET_TYPE_DEFAULT,  _
           BYVAL pixelFormat AS DXGI_FORMAT = DXGI_FORMAT_B8G8R8A8_UNORM, BYVAL alphaMode AS D2D1_ALPHA_MODE = D2D1_ALPHA_MODE_IGNORE,  _
           BYVAL dpiX AS SINGLE = -1, BYVAL dpiY AS SINGLE = -1, BYVAL usage AS D2D1_RENDER_TARGET_USAGE = D2D1_RENDER_TARGET_USAGE_NONE, _
           BYVAL minLevel AS D2D1_FEATURE_LEVEL = D2D1_FEATURE_LEVEL_DEFAULT) AS ID2D1DCRenderTarget PTR

END TYPE
' ========================================================================================

' ========================================================================================
' Creates an instance of the CD2D1Factory class.
' ========================================================================================
PRIVATE CONSTRUCTOR CD2D1Factory (BYVAL factoryType AS D2D1_FACTORY_TYPE = D2D1_FACTORY_TYPE_SINGLE_THREADED)
   ' // Creates an instance of the ID2D1Factory interface.
   SetLastError(D2D1CreateFactory(factoryType, IID_ID2D1Factory, BYVAL NULL, m_pFactory))
   D2D1_DP("m_pFactory - " & WSTR(m_pFactory))
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CD2D1Factory
   D2D1_DP("m_pFactory: " & WSTR(m_pFactory))
   IF m_pFactory THEN m_pFactory->Release
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the underlying ID2D1Factory interface pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF factory AS CD2D1Factory) AS ID2D1Factory PTR
   RETURN factory.m_pFactory
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of the underlying ID2D1Factory interface pointer
' ========================================================================================
PRIVATE OPERATOR CD2D1Factory.@ () AS ID2D1Factory PTR PTR
   RETURN @m_pFactory
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the underlying ID2D1Factory interface pointer
' ========================================================================================
PRIVATE OPERATOR CD2D1Factory.CAST () AS ID2D1Factory PTR
   RETURN m_pFactory
END OPERATOR
' ========================================================================================

' ========================================================================================
' Returns the DPI value
' ========================================================================================
PRIVATE SUB CD2D1Factory.GetDesktopDpi (BYREF dpiX AS SINGLE, BYREF dpiY AS SINGLE)
   m_pFactory->GetDesktopDpi(dpiX, dpiY)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CD2D1Factory.GetDesktopDpi (BYVAL dpiX AS SINGLE PTR, BYVAL dpiY AS SINGLE PTR)
   m_pFactory->GetDesktopDpi(BYVAL dpiX, BYVAL dpiY)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.GetDpi () AS SINGLE
   IF m_pFactory = NULL THEN RETURN 0
   DIM AS SINGLE dpix, dpiY
   m_pFactory->GetDesktopDpi(dpiX, dpiY)
   RETURN dpix
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns the horizontal DPI value of the desktop
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.GetDpiX () AS SINGLE
   IF m_pFactory = NULL THEN RETURN 0
   DIM AS SINGLE dpiX, dpiY
   m_pFactory->GetDesktopDpi(dpiX, dpiY)
   RETURN dpiX
END FUNCTION
' ========================================================================================
' ========================================================================================
' Returns the vertical DPI value of the desktop
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.GetDpiY () AS SINGLE
   IF m_pFactory = NULL THEN RETURN 0
   DIM AS SINGLE dpiX, dpiY
   m_pFactory->GetDesktopDpi(dpiX, dpiY)
   RETURN dpiY
END FUNCTION
' ========================================================================================

' ========================================================================================
' Forces the factory to refresh any system defaults that it might have changed since factory creation.
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.ReloadSystemMetrics () AS HRESULT
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN E_POINTER
   RETURN m_pFactory->ReloadSystemMetrics
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an ID2D1RectangleGeometry.
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateRectangleGeometry (BYREF rectangle AS D2D1_RECT_F) AS ID2D1RectangleGeometry PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM rectangleGeometry AS ID2D1RectangleGeometry PTR
   SetLastError(m_pFactory->CreateRectangleGeometry(rectangle, rectangleGeometry))
   RETURN rectangleGeometry
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateRectangleGeometry (BYVAL rectangle AS D2D1_RECT_F PTR) AS ID2D1RectangleGeometry PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM rectangleGeometry AS ID2D1RectangleGeometry PTR
   SetLastError(m_pFactory->CreateRectangleGeometry(BYVAL rectangle, rectangleGeometry))
   RETURN rectangleGeometry
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an ID2D1RoundedRectangleGeometry.
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateRoundedRectangleGeometry (BYREF roundedRectangle AS D2D1_ROUNDED_RECT) AS ID2D1RoundedRectangleGeometry PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM roundedRectangleGeometry AS ID2D1RoundedRectangleGeometry PTR
   SetLastError(m_pFactory->CreateRoundedRectangleGeometry(roundedRectangle, roundedRectangleGeometry))
   RETURN roundedRectangleGeometry
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateRoundedRectangleGeometry (BYVAL roundedRectangle AS D2D1_ROUNDED_RECT PTR) AS ID2D1RoundedRectangleGeometry PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM roundedRectangleGeometry AS ID2D1RoundedRectangleGeometry PTR
   SetLastError(m_pFactory->CreateRoundedRectangleGeometry(BYVAL roundedRectangle, roundedRectangleGeometry))
   RETURN roundedRectangleGeometry
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an ID2D1EllipseGeometry.
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateEllipseGeometry (BYREF ellipse AS const D2D1_ELLIPSE) AS ID2D1EllipseGeometry PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM ellipseGeometry AS ID2D1EllipseGeometry PTR
   SetLastError(m_pFactory->CreateEllipseGeometry(ellipse, ellipseGeometry))
   RETURN ellipseGeometry
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateEllipseGeometry (BYVAL ellipse AS const D2D1_ELLIPSE PTR) AS ID2D1EllipseGeometry PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM ellipseGeometry AS ID2D1EllipseGeometry PTR
   SetLastError(m_pFactory->CreateEllipseGeometry(BYVAL ellipse, ellipseGeometry))
   RETURN ellipseGeometry
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an ID2D1GeometryGroup, which is an object that holds other geometries.
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateGeometryGroup (BYVAL fillMode AS D2D1_FILL_MODE, BYVAL geometries AS ID2D1Geometry PTR, _
BYVAL geometriesCount AS UINT32) AS ID2D1GeometryGroup PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM geometryGroup AS ID2D1GeometryGroup PTR
   SetLastError(m_pFactory->CreateGeometryGroup(fillMode, geometries, geometriesCount, geometryGroup))
   RETURN geometryGroup
END FUNCTION
' ========================================================================================

' ========================================================================================
' Transforms the specified geometry and stores the result as an ID2D1TransformedGeometry object.
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateTransformedGeometry (BYVAL sourceGeometry AS ID2D1Geometry PTR, _
BYREF transform AS D2D1_MATRIX_3X2_F) AS ID2D1TransformedGeometry PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM transformedGeometry AS ID2D1TransformedGeometry PTR
   SetLastError(m_pFactory->CreateTransformedGeometry(sourceGeometry, transform, transformedGeometry))
   RETURN transformedGeometry
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateTransformedGeometry (BYVAL sourceGeometry AS ID2D1Geometry PTR, _
BYVAL transform AS D2D1_MATRIX_3X2_F PTR) AS ID2D1TransformedGeometry PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM transformedGeometry AS ID2D1TransformedGeometry PTR
   SetLastError(m_pFactory->CreateTransformedGeometry(sourceGeometry, BYVAL transform, transformedGeometry))
   RETURN transformedGeometry
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an empty ID2D1PathGeometry.
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreatePathGeometry () AS ID2D1PathGeometry PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM pathGeometry AS ID2D1PathGeometry PTR
   SetLastError(m_pFactory->CreatePathGeometry(pathGeometry))
   RETURN pathGeometry
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an ID2D1StrokeStyle that describes start cap, dash pattern, and other features of a stroke.
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateStrokeStyle (BYREF strokeStyleProperties AS D2D1_STROKE_STYLE_PROPERTIES, _
BYREF dashes AS const FLOAT, BYVAL dashesCount AS UINT32) AS ID2D1StrokeStyle PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM strokeStyle AS ID2D1StrokeStyle PTR
   SetLastError(m_pFactory->CreateStrokeStyle(strokeStyleProperties, dashes, dashesCount, strokeStyle))
   RETURN strokeStyle
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateStrokeStyle (BYVAL strokeStyleProperties AS D2D1_STROKE_STYLE_PROPERTIES PTR, _
BYVAL dashes AS const FLOAT PTR, BYVAL dashesCount AS UINT32) AS ID2D1StrokeStyle PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM strokeStyle AS ID2D1StrokeStyle PTR
   SetLastError(m_pFactory->CreateStrokeStyle(BYVAL strokeStyleProperties, BYVAL dashes, dashesCount, strokeStyle))
   RETURN strokeStyle
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an ID2D1DrawingStateBlock that can be used with the SaveDrawingState and
' RestoreDrawingState methods of a render target.
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateDrawingStateBlock (BYREF drawingStateDescription AS const D2D1_DRAWING_STATE_DESCRIPTION, _
BYVAL textRenderingParams AS IDWriteRenderingParams PTR) AS ID2D1DrawingStateBlock PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM drawingStateBlock AS ID2D1DrawingStateBlock PTR
   SetLastError(m_pFactory->CreateDrawingStateBlock(drawingStateDescription, textRenderingParams, drawingStateBlock))
   RETURN drawingStateBlock
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateDrawingStateBlock (BYVAL drawingStateDescription AS const D2D1_DRAWING_STATE_DESCRIPTION PTR, _
BYVAL textRenderingParams AS IDWriteRenderingParams PTR) AS ID2D1DrawingStateBlock PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM drawingStateBlock AS ID2D1DrawingStateBlock PTR
   SetLastError(m_pFactory->CreateDrawingStateBlock(BYVAL drawingStateDescription, textRenderingParams, drawingStateBlock))
   RETURN drawingStateBlock
END FUNCTION
' ========================================================================================
  ' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateDrawingStateBlock (BYREF drawingStateDescription AS const D2D1_DRAWING_STATE_DESCRIPTION) AS ID2D1DrawingStateBlock PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM drawingStateBlock AS ID2D1DrawingStateBlock PTR
   SetLastError(m_pFactory->CreateDrawingStateBlock(drawingStateDescription, NULL, drawingStateBlock))
   RETURN drawingStateBlock
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateDrawingStateBlock (BYVAL drawingStateDescription AS const D2D1_DRAWING_STATE_DESCRIPTION PTR) AS ID2D1DrawingStateBlock PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM drawingStateBlock AS ID2D1DrawingStateBlock PTR
   SetLastError(m_pFactory->CreateDrawingStateBlock(BYVAL drawingStateDescription, NULL, drawingStateBlock))
   RETURN drawingStateBlock
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateDrawingStateBlock () AS ID2D1DrawingStateBlock PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM drawingStateBlock AS ID2D1DrawingStateBlock PTR
   SetLastError(m_pFactory->CreateDrawingStateBlock(BYVAL NULL, NULL, drawingStateBlock))
   RETURN drawingStateBlock
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a render target that renders to a Microsoft Windows Imaging Component (WIC) bitmap.
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateWicBitmapRenderTarget (BYVAL target AS IWICBitmap PTR, _
BYREF renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES) AS ID2D1RenderTarget PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM renderTarget AS ID2D1RenderTarget PTR
   SetLastError(m_pFactory->CreateWicBitmapRenderTarget(target, renderTargetProperties, renderTarget))
   RETURN renderTarget
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateWicBitmapRenderTarget (BYVAL target AS IWICBitmap PTR, _
BYVAL renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES PTR) AS ID2D1RenderTarget PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM renderTarget AS ID2D1RenderTarget PTR
   SetLastError(m_pFactory->CreateWicBitmapRenderTarget(target, BYVAL renderTargetProperties, renderTarget))
   RETURN renderTarget
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an ID2D1HwndRenderTarget, a render target that renders to a window.
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateHwndRenderTarget (BYREF renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES, _
BYREF hwndRenderTargetProperties AS D2D1_HWND_RENDER_TARGET_PROPERTIES) AS ID2D1HwndRenderTarget PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM hwndRenderTarget AS ID2D1HwndRenderTarget PTR
   SetLastError(m_pFactory->CreateHwndRenderTarget(renderTargetProperties, hwndRenderTargetProperties, hwndRenderTarget))
   RETURN hwndRenderTarget
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateHwndRenderTarget (BYVAL renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES PTR, _
BYVAL hwndRenderTargetProperties AS D2D1_HWND_RENDER_TARGET_PROPERTIES PTR) AS ID2D1HwndRenderTarget PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM hwndRenderTarget AS ID2D1HwndRenderTarget PTR
   SetLastError(m_pFactory->CreateHwndRenderTarget(BYVAL renderTargetProperties, BYVAL hwndRenderTargetProperties, hwndRenderTarget))
   RETURN hwndRenderTarget
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a render target that draws to a DirectX Graphics Infrastructure (DXGI) surface.
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateDxgiSurfaceRenderTarget (BYVAL dxgiSurface AS IDXGISurface PTR, _
BYREF renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES) AS ID2D1RenderTarget PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM renderTarget AS ID2D1RenderTarget PTR
   SetLastError(m_pFactory->CreateDxgiSurfaceRenderTarget(dxgiSurface, renderTargetProperties, renderTarget))
   RETURN renderTarget
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateDxgiSurfaceRenderTarget (BYVAL dxgiSurface AS IDXGISurface PTR, _
BYVAL renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES PTR) AS ID2D1RenderTarget PTR
   D2D1_DP("")
   IF m_pFactory = NULL THEN RETURN NULL
   DIM renderTarget AS ID2D1RenderTarget PTR
   SetLastError(m_pFactory->CreateDxgiSurfaceRenderTarget(dxgiSurface, BYVAL renderTargetProperties, renderTarget))
   RETURN renderTarget
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a render target that draws to a Windows Graphics Device Interface (GDI) device context.
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateDCRenderTarget (BYREF renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES) AS ID2D1DCRenderTarget PTR
   D2D1_DP("BYREF props: " & WSTR(VARPTR(renderTargetProperties)))
   IF m_pFactory = NULL THEN RETURN NULL
   DIM pRenderTarget AS ID2D1DCRenderTarget PTR
   SetLastError(m_pFactory->CreateDCRenderTarget(renderTargetProperties, pRenderTarget))
   D2D1_DP(WSTR(pRenderTarget))
   RETURN pRenderTarget
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateDCRenderTarget (BYVAL renderTargetProperties AS D2D1_RENDER_TARGET_PROPERTIES PTR) AS ID2D1DCRenderTarget PTR
   D2D1_DP("BYVAL props: " & WSTR(renderTargetProperties))
   IF m_pFactory = NULL THEN RETURN NULL
   DIM pRenderTarget AS ID2D1DCRenderTarget PTR
   SetLastError(m_pFactory->CreateDCRenderTarget(BYVAL renderTargetProperties, pRenderTarget))
   D2D1_DP(WSTR(pRenderTarget))
   RETURN pRenderTarget
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CD2D1Factory.CreateDCRenderTarget (BYVAL targetType AS D2D1_RENDER_TARGET_TYPE = D2D1_RENDER_TARGET_TYPE_DEFAULT,  _
BYVAL pixelFormat AS DXGI_FORMAT = DXGI_FORMAT_B8G8R8A8_UNORM, BYVAL alphaMode AS D2D1_ALPHA_MODE = D2D1_ALPHA_MODE_IGNORE,  _
BYVAL dpiX AS SINGLE = -1, BYVAL dpiY AS SINGLE = -1, BYVAL usage AS D2D1_RENDER_TARGET_USAGE = D2D1_RENDER_TARGET_USAGE_NONE, _
BYVAL minLevel AS D2D1_FEATURE_LEVEL = D2D1_FEATURE_LEVEL_DEFAULT) AS ID2D1DCRenderTarget PTR

   D2D1_DP("targetType: " & WSTR(targetType))
   IF m_pFactory = NULL THEN RETURN NULL
   ' // DPI values
   DIM AS SINGLE _dpiX, _dpiY
   m_pFactory->GetDesktopDpi(_dpiX, _dpiY)
   IF dpiX = -1 THEN dpiX = _dpiX
   IF dpiY = -1 THEN dpiY = _dpiY
   ' // Render target properties
   DIM props AS D2D1_RENDER_TARGET_PROPERTIES
   props.type = targetType
   props.pixelFormat = D2D1_PixelFormat(pixelFormat, alphaMode)
   props.dpiX = dpiX
   props.dpiY = dpiY
   props.usage = usage
   props.minLevel = minLevel
   ' // Create the render target
   DIM pRenderTarget AS ID2D1DCRenderTarget PTR
   SetLastError(m_pFactory->CreateDCRenderTarget(props, pRenderTarget))
   D2D1_DP(WSTR(pRenderTarget))
   RETURN pRenderTarget

END FUNCTION
' ========================================================================================


' ########################################################################################
' CID2D1SolidColorBrush class
' ########################################################################################

' ========================================================================================
TYPE CID2D1SolidColorBrush
Public:

   m_pSolidColorBrush AS ID2D1SolidColorBrush PTR

   DECLARE CONSTRUCTOR (BYVAL psolidColorBrush AS ID2D1SolidColorBrush PTR)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS ID2D1SolidColorBrush PTR PTR
   DECLARE OPERATOR CAST () AS ID2D1SolidColorBrush PTR
   DECLARE OPERATOR CAST () AS ID2D1Brush PTR

   ' // ID2D1Resource
   DECLARE FUNCTION GetFactory () AS ID2D1Factory PTR

   ' // ID2D1Brush
   DECLARE SUB SetOpacity (BYVAL opacity AS FLOAT)
   DECLARE FUNCTION GetOpacity () AS FLOAT
   DECLARE SUB SetTransform (BYREF transform AS D2D1_MATRIX_3X2_F)
   DECLARE SUB SetTransform (BYVAL transform AS D2D1_MATRIX_3X2_F PTR)
   DECLARE SUB GetTransform (BYREF transform AS D2D1_MATRIX_3X2_F)
   DECLARE SUB GetTransform (BYVAL transform AS D2D1_MATRIX_3X2_F PTR)
   DECLARE FUNCTION GetTransform () AS D2D1_MATRIX_3X2_F

   ' // ID2D1SolidColorBrush
   DECLARE SUB SetColor (BYREF clr AS D2D1_COLOR_F)
   DECLARE SUB SetColor (BYVAL clr AS D2D1_COLOR_F PTR)
   DECLARE FUNCTION GetColor () AS D2D1_COLOR_F
   DECLARE SUB GetColor (BYREF clr AS D2D1_COLOR_F)
   DECLARE SUB GetColor (BYVAL clr AS D2D1_COLOR_F PTR)

END TYPE
' ========================================================================================

' ========================================================================================
' Creates an instance of the CID2D1SolidColorBrush class.
' ========================================================================================
PRIVATE CONSTRUCTOR CID2D1SolidColorBrush (BYVAL psolidColorBrush AS ID2D1SolidColorBrush PTR)
   D2D1_DP("psolidColorBrush - " & WSTR(psolidColorBrush))
   m_pSolidColorBrush = psolidColorBrush
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CID2D1SolidColorBrush
   D2D1_DP("m_pSolidColorBrush: " & WSTR(m_pSolidColorBrush))
   IF m_pSolidColorBrush THEN m_pSolidColorBrush->Release
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the underlying ID2D1SolidColorBrush interface pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF solidColorBrush AS CID2D1SolidColorBrush) AS ID2D1SolidColorBrush PTR
   RETURN solidColorBrush.m_pSolidColorBrush
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of the underlying ID2D1SolidColorBrush interface pointer
' ========================================================================================
PRIVATE OPERATOR CID2D1SolidColorBrush.@ () AS ID2D1SolidColorBrush PTR PTR
   RETURN @m_pSolidColorBrush
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the underlying ID2D1SolidColorBrush interface pointer
' ========================================================================================
PRIVATE OPERATOR CID2D1SolidColorBrush.CAST () AS ID2D1SolidColorBrush PTR
   RETURN m_pSolidColorBrush
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the underlying ID2D1Brush interface pointer
' ========================================================================================
PRIVATE OPERATOR CID2D1SolidColorBrush.CAST () AS ID2D1Brush PTR
   RETURN cast(ID2D1Brush PTR, m_pSolidColorBrush)
END OPERATOR
' ========================================================================================

' ========================================================================================
' Inherited ID2D1Resource inteface method
' ========================================================================================

' ========================================================================================
' Retrieves the factory associated with this resource.
' ========================================================================================
PRIVATE FUNCTION CID2D1SolidColorBrush.GetFactory () AS ID2D1Factory PTR
   D2D1_DP("")
   DIM factory AS ID2D1Factory PTR
   DIM pBrush AS ID2D1Brush PTR = CAST(ID2D1Brush PTR, m_pSolidColorBrush)
   DIM pResource AS ID2D1Resource PTR = CAST(ID2D1Resource PTR, pBrush)
   pResource->GetFactory(factory)
   RETURN factory
END FUNCTION
' ========================================================================================
 
' ========================================================================================
' Inherited ID2D1Brush interface methods
' ========================================================================================

' ========================================================================================
' Sets the degree of opacity of this brush.
' ========================================================================================
PRIVATE SUB CID2D1SolidColorBrush.SetOpacity (BYVAL opacity AS FLOAT)
   D2D1_DP(WSTR(opacity))
   IF m_pSolidColorBrush THEN cast(ID2D1Brush PTR, m_pSolidColorBrush)->SetOpacity(opacity)
END SUB
' ========================================================================================

' ========================================================================================
' Gets the degree of opacity of this brush.
' ========================================================================================
PRIVATE FUNCTION CID2D1SolidColorBrush.GetOpacity () AS FLOAT
   D2D1_DP("")
   IF m_pSolidColorBrush THEN RETURN cast(ID2D1Brush PTR, m_pSolidColorBrush)->GetOpacity
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the transformation applied to the brush.
' ========================================================================================
PRIVATE SUB CID2D1SolidColorBrush.SetTransform (BYREF transform AS D2D1_MATRIX_3X2_F)
   D2D1_DP("")
   IF m_pSolidColorBrush THEN cast(ID2D1Brush PTR, m_pSolidColorBrush)->SetTransform(transform)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1SolidColorBrush.SetTransform (BYVAL transform AS D2D1_MATRIX_3X2_F PTR)
   D2D1_DP("")
   IF m_pSolidColorBrush THEN cast(ID2D1Brush PTR, m_pSolidColorBrush)->SetTransform(BYVAL transform)
END SUB
' ========================================================================================

' ========================================================================================
' Gets the transform applied to this brush.
' ========================================================================================
PRIVATE SUB CID2D1SolidColorBrush.GetTransform (BYREF transform AS D2D1_MATRIX_3X2_F)
   D2D1_DP("")
   IF m_pSolidColorBrush THEN cast(ID2D1Brush PTR, m_pSolidColorBrush)->GetTransform(transform)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1SolidColorBrush.GetTransform (BYVAL transform AS D2D1_MATRIX_3X2_F PTR)
   D2D1_DP("")
   IF m_pSolidColorBrush THEN cast(ID2D1Brush PTR, m_pSolidColorBrush)->GetTransform(BYVAL transform)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1SolidColorBrush.GetTransform () AS D2D1_MATRIX_3X2_F
   D2D1_DP("")
   DIM transform AS D2D1_MATRIX_3X2_F
   IF m_pSolidColorBrush THEN cast(ID2D1Brush PTR, m_pSolidColorBrush)->GetTransform(transform)
   RETURN transform
END FUNCTION
' ========================================================================================

' ========================================================================================
' ID2D1SolidColorBrush interface methods
' ========================================================================================

' ========================================================================================
' Specifies the color of this solid-color brush.
' ========================================================================================
PRIVATE SUB CID2D1SolidColorBrush.SetColor (BYREF clr AS D2D1_COLOR_F)
   D2D1_DP("")
   IF m_pSolidColorBrush THEN m_pSolidColorBrush->SetColor(clr)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1SolidColorBrush.SetColor (BYVAL clr AS D2D1_COLOR_F PTR)
   D2D1_DP("")
   IF m_pSolidColorBrush THEN m_pSolidColorBrush->SetColor(BYVAL clr)
END SUB
' ========================================================================================

' ========================================================================================
' Retrieves the color of the solid color brush.
' ========================================================================================
PRIVATE FUNCTION CID2D1SolidColorBrush.GetColor () AS D2D1_COLOR_F
   D2D1_DP("")
   DIM clr AS D2D1_COLOR_F
   IF m_pSolidColorBrush THEN m_pSolidColorBrush->GetColor(clr)
   RETURN clr
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1SolidColorBrush.GetColor (BYREF clr AS D2D1_COLOR_F)
   D2D1_DP("")
   IF m_pSolidColorBrush THEN m_pSolidColorBrush->GetColor(clr)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1SolidColorBrush.GetColor (BYVAL clr AS D2D1_COLOR_F PTR)
   D2D1_DP("")
   IF m_pSolidColorBrush THEN m_pSolidColorBrush->GetColor(BYVAL clr)
END SUB
' ========================================================================================


' ########################################################################################
' CID2D1RenderTarget class
' ########################################################################################

' ========================================================================================
TYPE CID2D1RenderTarget
Public:

   m_pRenderTarget AS ID2D1RenderTarget PTR

   DECLARE CONSTRUCTOR
   DECLARE CONSTRUCTOR (BYVAL pRenderTarget AS ID2D1RenderTarget PTR)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS ID2D1RenderTarget PTR PTR
   DECLARE OPERATOR CAST () AS ID2D1RenderTarget PTR
   DECLARE FUNCTION GetFactory () AS ID2D1Factory PTR
   DECLARE FUNCTION CreateBitmap (BYVAL size AS D2D1_SIZE_U, BYVAL srcData AS const ANY PTR, _
           BYVAL pitch AS UINT32, BYREF bitmapProperties AS D2D1_BITMAP_PROPERTIES) AS ID2D1Bitmap PTR
   DECLARE FUNCTION CreateBitmap (BYVAL size AS D2D1_SIZE_U, BYVAL srcData AS const ANY PTR, _
           BYVAL pitch AS UINT32, BYVAL bitmapProperties AS D2D1_BITMAP_PROPERTIES PTR) AS ID2D1Bitmap PTR
   DECLARE FUNCTION CreateBitmapFromWicBitmap (BYVAL wicBitmapSource AS IWICBitmapSource PTR, _
           BYREF bitmapProperties AS const D2D1_BITMAP_PROPERTIES) AS ID2D1Bitmap PTR
   DECLARE FUNCTION CreateBitmapFromWicBitmap (BYVAL wicBitmapSource AS IWICBitmapSource PTR, _
           BYVAL bitmapProperties AS const D2D1_BITMAP_PROPERTIES PTR) AS ID2D1Bitmap PTR
   DECLARE FUNCTION CreateBitmapFromWicBitmap (BYVAL wicBitmapSource AS IWICBitmapSource PTR) AS ID2D1Bitmap PTR
   DECLARE FUNCTION CreateSharedBitmap (BYREF iid AS IID, BYVAL pdata AS ANY PTR, _
           BYREF bitmapProperties AS const D2D1_BITMAP_PROPERTIES) AS ID2D1Bitmap PTR
   DECLARE FUNCTION CreateSharedBitmap (BYVAL iid AS IID PTR, BYVAL pdata AS ANY PTR, _
           BYVAL bitmapProperties AS const D2D1_BITMAP_PROPERTIES PTR) AS ID2D1Bitmap PTR
   DECLARE FUNCTION CreateBitmapBrush (BYVAL bitmap AS ID2D1Bitmap PTR, BYREF bitmapBrushProperties AS const D2D1_BITMAP_BRUSH_PROPERTIES, _
           BYREF brushProperties AS const D2D1_BRUSH_PROPERTIES) AS ID2D1BitmapBrush PTR
   DECLARE FUNCTION CreateBitmapBrush (BYVAL bitmap AS ID2D1Bitmap PTR, BYVAL bitmapBrushProperties AS const D2D1_BITMAP_BRUSH_PROPERTIES PTR, _
           BYVAL brushProperties AS const D2D1_BRUSH_PROPERTIES PTR) AS ID2D1BitmapBrush PTR
   DECLARE FUNCTION CreateBitmapBrush (BYVAL bitmap AS ID2D1Bitmap PTR, _
           BYREF bitmapBrushProperties AS const D2D1_BITMAP_BRUSH_PROPERTIES) AS ID2D1BitmapBrush PTR
   DECLARE FUNCTION CreateBitmapBrush (BYVAL bitmap AS ID2D1Bitmap PTR, _
           BYVAL bitmapBrushProperties AS const D2D1_BITMAP_BRUSH_PROPERTIES PTR) AS ID2D1BitmapBrush PTR
   DECLARE FUNCTION CreateSolidColorBrush (BYREF clr AS D2D1_COLOR_F, _
           BYREF brushProperties AS const D2D1_BRUSH_PROPERTIES) AS ID2D1SolidColorBrush PTR
   DECLARE FUNCTION CreateSolidColorBrush (BYVAL clr AS D2D1_COLOR_F PTR, _
           BYVAL brushProperties AS const D2D1_BRUSH_PROPERTIES PTR) AS ID2D1SolidColorBrush PTR
   DECLARE FUNCTION CreateGradientStopCollection (BYREF gradientStops AS D2D1_GRADIENT_STOP, _
           BYVAL gradientStopsCount AS UINT32, BYVAL colorInterpolationGamma AS D2D1_GAMMA, _
           BYVAL extendMode AS D2D1_EXTEND_MODE) AS ID2D1GradientStopCollection PTR
   DECLARE FUNCTION CreateGradientStopCollection (BYVAL gradientStops AS D2D1_GRADIENT_STOP PTR, _
           BYVAL gradientStopsCount AS UINT32, BYVAL colorInterpolationGamma AS D2D1_GAMMA, _
           BYVAL extendMode AS D2D1_EXTEND_MODE) AS ID2D1GradientStopCollection PTR
   DECLARE FUNCTION CreateGradientStopCollection (BYREF gradientStops AS D2D1_GRADIENT_STOP, _
           BYVAL gradientStopsCount AS UINT32) AS ID2D1GradientStopCollection PTR
   DECLARE FUNCTION CreateGradientStopCollection (BYVAL gradientStops AS D2D1_GRADIENT_STOP PTR, _
           BYVAL gradientStopsCount AS UINT32) AS ID2D1GradientStopCollection PTR
   DECLARE FUNCTION CreateLinearGradientBrush ( _
           BYREF linearGradientBrushProperties AS const D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, _
           BYREF brushProperties AS const D2D1_BRUSH_PROPERTIES, _
           BYVAL gradientStopCollection AS ID2D1GradientStopCollection PTR) AS ID2D1LinearGradientBrush PTR
   DECLARE FUNCTION CreateLinearGradientBrush ( _
           BYVAL linearGradientBrushProperties AS const D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES PTR, _
           BYVAL brushProperties AS const D2D1_BRUSH_PROPERTIES PTR, _
           BYVAL gradientStopCollection AS ID2D1GradientStopCollection PTR) AS ID2D1LinearGradientBrush PTR
   DECLARE FUNCTION CreateLinearGradientBrush ( _
           BYREF linearGradientBrushProperties AS const D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, _
           BYVAL gradientStopCollection AS ID2D1GradientStopCollection PTR) AS ID2D1LinearGradientBrush PTR
   DECLARE FUNCTION CreateRadialGradientBrush ( _
           BYREF radialGradientBrushProperties AS const D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, _
           BYREF brushProperties AS const D2D1_BRUSH_PROPERTIES, _
           BYVAL gradientStopCollection AS ID2D1GradientStopCollection PTR) AS ID2D1RadialGradientBrush PTR
   DECLARE FUNCTION CreateRadialGradientBrush ( _
           BYVAL radialGradientBrushProperties AS const D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES PTR, _
           BYVAL brushProperties AS const D2D1_BRUSH_PROPERTIES PTR, _
           BYVAL gradientStopCollection AS ID2D1GradientStopCollection PTR) AS ID2D1RadialGradientBrush PTR
   DECLARE FUNCTION CreateRadialGradientBrush ( _
           BYREF radialGradientBrushProperties AS const D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, _
           BYVAL gradientStopCollection AS ID2D1GradientStopCollection PTR) AS ID2D1RadialGradientBrush PTR
   DECLARE FUNCTION CreateCompatibleRenderTarget (BYREF desiredSize AS const D2D1_SIZE_F, _
           BYREF desiredPixelSize AS const D2D1_SIZE_U, BYREF desiredFormat AS const D2D1_PIXEL_FORMAT, _
           BYVAL options AS D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS) AS ID2D1BitmapRenderTarget PTR
   DECLARE FUNCTION CreateCompatibleRenderTarget (BYVAL desiredSize AS const D2D1_SIZE_F PTR, _
           BYVAL desiredPixelSize AS const D2D1_SIZE_U PTR, BYVAL desiredFormat AS const D2D1_PIXEL_FORMAT PTR, _
           BYVAL options AS D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS) AS ID2D1BitmapRenderTarget PTR
   DECLARE FUNCTION CreateCompatibleRenderTarget () AS ID2D1BitmapRenderTarget PTR  
   DECLARE FUNCTION CreateCompatibleRenderTarget (BYREF desiredSize AS const D2D1_SIZE_F) AS ID2D1BitmapRenderTarget PTR
   DECLARE FUNCTION CreateCompatibleRenderTarget (BYVAL desiredSize AS const D2D1_SIZE_F PTR) AS ID2D1BitmapRenderTarget PTR
   DECLARE FUNCTION CreateCompatibleRenderTarget (BYREF desiredSize AS const D2D1_SIZE_F, _
           BYREF desiredPixelSize AS const D2D1_SIZE_U) AS ID2D1BitmapRenderTarget PTR
   DECLARE FUNCTION CreateCompatibleRenderTarget (BYVAL desiredSize AS const D2D1_SIZE_F PTR, _
           BYVAL desiredPixelSize AS const D2D1_SIZE_U PTR) AS ID2D1BitmapRenderTarget PTR
   DECLARE FUNCTION CreateLayer (BYREF size AS const D2D1_SIZE_F) AS ID2D1Layer PTR
   DECLARE FUNCTION CreateLayer (BYVAL size AS const D2D1_SIZE_F PTR) AS ID2D1Layer PTR
   DECLARE FUNCTION CreateLayer () AS ID2D1Layer PTR
   DECLARE FUNCTION CreateMesh () AS ID2D1Mesh PTR
   DECLARE SUB DrawLine (BYVAL point0 AS D2D1_POINT_2F, BYVAL point1 AS D2D1_POINT_2F, BYVAL brush AS ID2D1Brush PTR, _
           BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL)
   DECLARE SUB DrawRectangle (BYREF rect AS const D2D1_RECT_F, BYVAL brush AS ID2D1Brush PTR, _
           BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL)
   DECLARE SUB DrawRectangle (BYVAL rect AS const D2D1_RECT_F PTR, BYVAL brush AS ID2D1Brush PTR, _
           BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL)
   DECLARE SUB FillRectangle (BYREF rect AS const D2D1_RECT_F, BYVAL brush AS ID2D1Brush PTR)
   DECLARE SUB FillRectangle (BYVAL rect AS const D2D1_RECT_F PTR, BYVAL brush AS ID2D1Brush PTR)
   DECLARE SUB DrawRoundedRectangle (BYREF roundedRect AS const D2D1_ROUNDED_RECT, BYVAL brush AS ID2D1Brush PTR, _
           BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL)
   DECLARE SUB DrawRoundedRectangle (BYVAL roundedRect AS const D2D1_ROUNDED_RECT PTR, BYVAL brush AS ID2D1Brush PTR, _
           BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL)
   DECLARE SUB FillRoundedRectangle (BYREF roundedRect AS const D2D1_ROUNDED_RECT, BYVAL brush AS ID2D1Brush PTR)
   DECLARE SUB FillRoundedRectangle (BYVAL roundedRect AS const D2D1_ROUNDED_RECT PTR, BYVAL brush AS ID2D1Brush PTR)
   DECLARE SUB DrawEllipse (BYREF tEllipse AS const D2D1_ELLIPSE, BYVAL brush AS ID2D1Brush PTR, _
           BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL) 
   DECLARE SUB DrawEllipse (BYVAL tEllipse AS const D2D1_ELLIPSE PTR, BYVAL brush AS ID2D1Brush PTR, _
           BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL) 
   DECLARE SUB FillEllipse (BYREF tEllipse AS const D2D1_ELLIPSE, BYVAL brush AS ID2D1Brush PTR)
   DECLARE SUB FillEllipse (BYVAL tEllipse AS const D2D1_ELLIPSE PTR, BYVAL brush AS ID2D1Brush PTR)
   DECLARE SUB DrawGeometry (BYVAL geometry AS ID2D1Geometry PTR, BYVAL brush AS ID2D1Brush PTR, _
           BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL) 
   DECLARE SUB FillGeometry (BYVAL geometry AS ID2D1Geometry PTR, BYVAL brush AS ID2D1Brush PTR)
   DECLARE SUB FillMesh (BYVAL mesh AS ID2D1Mesh PTR, BYVAL brush AS ID2D1Brush PTR)
   DECLARE SUB FillOpacityMask (BYVAL opacityMask AS ID2D1Bitmap PTR, BYVAL brush AS ID2D1Brush PTR, _
           BYVAL content AS D2D1_OPACITY_MASK_CONTENT, BYREF destinationRectangle AS const D2D1_RECT_F, _
           BYREF sourceRectangle AS const D2D1_RECT_F)
   DECLARE SUB FillOpacityMask (BYVAL opacityMask AS ID2D1Bitmap PTR, BYVAL brush AS ID2D1Brush PTR, _
           BYVAL content AS D2D1_OPACITY_MASK_CONTENT, BYVAL destinationRectangle AS const D2D1_RECT_F PTR, _
           BYVAL sourceRectangle AS const D2D1_RECT_F PTR)
   DECLARE SUB DrawBitmap (BYVAL bitmap AS ID2D1Bitmap PTR, BYREF destinationRectangle AS const D2D1_RECT_F, _
           BYVAL opacity AS FLOAT = 1.0!, BYVAL InterpolationMode AS D2D1_BITMAP_INTERPOLATION_MODE = D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, _
           BYREF sourceRectangle AS const D2D1_RECT_F)
   DECLARE SUB DrawBitmap (BYVAL bitmap AS ID2D1Bitmap PTR, BYVAL destinationRectangle AS const D2D1_RECT_F PTR, _
           BYVAL opacity AS FLOAT = 1.0!, BYVAL InterpolationMode AS D2D1_BITMAP_INTERPOLATION_MODE = D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, _
           BYVAL sourceRectangle AS const D2D1_RECT_F PTR)
   DECLARE SUB DrawText (BYREF wzSstr AS WSTRING, BYVAL stringLength AS UINT32, _
           BYVAL textFormat AS IDWriteTextFormat PTR, BYREF layoutRect AS D2D1_RECT_F, BYVAL defaultForegroundBrush AS ID2D1Brush PTR, _
           BYVAL options AS D2D1_DRAW_TEXT_OPTIONS = D2D1_DRAW_TEXT_OPTIONS_NONE, _
           BYVAL measuringMode AS DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE_NATURAL)
   DECLARE SUB DrawText (BYVAL pwzSstr AS WSTRING PTR, BYVAL stringLength AS UINT32, _
           BYVAL textFormat AS IDWriteTextFormat PTR, BYVAL layoutRect AS D2D1_RECT_F PTR, BYVAL defaultForegroundBrush AS ID2D1Brush PTR, _
           BYVAL options AS D2D1_DRAW_TEXT_OPTIONS = D2D1_DRAW_TEXT_OPTIONS_NONE, _
           BYVAL measuringMode AS DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE_NATURAL)
   DECLARE SUB DrawTextLayout (BYVAL origin AS D2D1_POINT_2F, BYVAL textLayout AS IDWriteTextLayout PTR, _
           BYVAL defaultFillBrush AS ID2D1Brush PTR, BYVAL options AS D2D1_DRAW_TEXT_OPTIONS = D2D1_DRAW_TEXT_OPTIONS_NONE) 
   DECLARE SUB DrawGlyphRun (BYVAL baselineOrigin AS D2D1_POINT_2F, BYREF glyphRun AS const DWRITE_GLYPH_RUN, _
           BYVAL foregroundBrush AS ID2D1Brush PTR, BYVAL measuringMode AS DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE_NATURAL)
   DECLARE SUB DrawGlyphRun (BYVAL baselineOrigin AS D2D1_POINT_2F, BYVAL glyphRun AS const DWRITE_GLYPH_RUN PTR, _
           BYVAL foregroundBrush AS ID2D1Brush PTR, BYVAL measuringMode AS DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE_NATURAL)
   DECLARE SUB SetTransform (BYREF transform AS D2D1_MATRIX_3X2_F)
   DECLARE SUB SetTransform (BYVAL transform AS D2D1_MATRIX_3X2_F PTR)
   DECLARE SUB GetTransform (BYREF transform AS D2D1_MATRIX_3X2_F)
   DECLARE SUB GetTransform (BYVAL transform AS D2D1_MATRIX_3X2_F PTR)
   DECLARE FUNCTION GetTransform () AS D2D1_MATRIX_3X2_F
   DECLARE SUB SetAntialiasMode (BYVAL antialiasMode AS D2D1_ANTIALIAS_MODE)
   DECLARE FUNCTION GetAntialiasMode () AS D2D1_ANTIALIAS_MODE
   DECLARE SUB SetTextAntialiasMode (BYVAL textAntialiasMode AS D2D1_TEXT_ANTIALIAS_MODE)
   DECLARE FUNCTION GetTextAntialiasMode () AS D2D1_TEXT_ANTIALIAS_MODE
   DECLARE SUB SetTextRenderingParams (BYVAL textRenderingParams AS IDWriteRenderingParams PTR)
   DECLARE SUB GetTextRenderingParams (BYREF textRenderingParams AS IDWriteRenderingParams PTR)
   DECLARE FUNCTION GetTextRenderingParams () AS IDWriteRenderingParams PTR
   DECLARE SUB SetTags (BYVAL tag1 AS D2D1_TAG, BYVAL tag2 AS D2D1_TAG)
   DECLARE SUB GetTags (BYREF tag1 AS D2D1_TAG = NULL, BYREF tag2 AS D2D1_TAG = NULL)
   DECLARE SUB GetTags (BYVAL tag1 AS D2D1_TAG PTR = NULL, BYVAL tag2 AS D2D1_TAG PTR = NULL)
   DECLARE SUB PushLayer (BYREF layerParameters AS D2D1_LAYER_PARAMETERS, BYVAL layer AS ID2D1Layer PTR)
   DECLARE SUB PushLayer (BYVAL layerParameters AS D2D1_LAYER_PARAMETERS PTR, BYVAL layer AS ID2D1Layer PTR)
   DECLARE SUB PopLayer
   DECLARE FUNCTION Flush (BYREF tag1 AS D2D1_TAG = NULL, BYREF tag2 AS D2D1_TAG = NULL) AS HRESULT
   DECLARE FUNCTION Flush (BYVAL tag1 AS D2D1_TAG PTR = NULL, BYVAL tag2 AS D2D1_TAG PTR = NULL) AS HRESULT
   DECLARE SUB SaveDrawingState (BYVAL drawingStateBlock AS ID2D1DrawingStateBlock PTR)
   DECLARE SUB RestoreDrawingState (BYVAL drawingStateBlock AS ID2D1DrawingStateBlock PTR)
   DECLARE SUB PushAxisAlignedClip (BYREF clipRect AS D2D1_RECT_F, BYVAL antialiasMode AS D2D1_ANTIALIAS_MODE)
   DECLARE SUB PushAxisAlignedClip (BYVAL clipRect AS D2D1_RECT_F PTR, BYVAL antialiasMode AS D2D1_ANTIALIAS_MODE)
   DECLARE SUB PopAxisAlignedClip
   DECLARE SUB Clear (BYREF clearColor AS const D2D1_COLOR_F)
   DECLARE SUB Clear (BYVAL clearColor AS const D2D1_COLOR_F PTR)
   DECLARE SUB BeginDraw
   DECLARE FUNCTION EndDraw (BYREF tag1 AS D2D1_TAG, BYREF tag2 AS D2D1_TAG) AS HRESULT
   DECLARE FUNCTION EndDraw (BYVAL tag1 AS D2D1_TAG PTR = NULL, BYVAL tag2 AS D2D1_TAG PTR = NULL) AS HRESULT
   DECLARE SUB GetPixelFormat (BYREF pixelFormat AS D2D1_PIXEL_FORMAT)
   DECLARE SUB GetPixelFormat (BYVAL pixelFormat AS D2D1_PIXEL_FORMAT PTR)
   DECLARE FUNCTION GetPixelFormat () AS D2D1_PIXEL_FORMAT
   DECLARE SUB SetDpi (BYVAL dpiX AS FLOAT, BYVAL dpiY AS FLOAT)
   DECLARE SUB GetDpi (BYREF dpiX AS FLOAT, BYREF dpiY AS FLOAT)
   DECLARE SUB GetSize (BYREF size AS D2D1_SIZE_F)
   DECLARE SUB GetSize (BYVAL size AS D2D1_SIZE_F PTR)
   DECLARE FUNCTION GetSize () AS D2D1_SIZE_F
   DECLARE SUB GetPixelSize (BYREF pixelSize AS D2D1_SIZE_U)
   DECLARE SUB GetPixelSize (BYVAL pixelSize AS D2D1_SIZE_U PTR)
   DECLARE FUNCTION GetPixelSize () AS D2D1_SIZE_U
   DECLARE FUNCTION GetMaximumBitmapSize () AS UINT32
   DECLARE FUNCTION IsSupported (BYREF renderTargetProperties AS const D2D1_RENDER_TARGET_PROPERTIES) AS BOOLEAN
   DECLARE FUNCTION IsSupported (BYVAL renderTargetProperties AS const D2D1_RENDER_TARGET_PROPERTIES PTR) AS BOOLEAN

 END TYPE
' ========================================================================================

' ========================================================================================
' Creates an instance of the CID2D1RenderTarget class.
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CID2D1RenderTarget
   D2D1_DP("Default")
END CONSTRUCTOR
' ========================================================================================
' ========================================================================================
PRIVATE CONSTRUCTOR CID2D1RenderTarget (BYVAL pRenderTarget AS ID2D1RenderTarget PTR)
   D2D1_DP("pRenderTarget - " & WSTR(pRenderTarget))
   m_pRenderTarget = pRenderTarget
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CID2D1RenderTarget
   D2D1_DP("m_pRenderTarget: " & WSTR(m_pRenderTarget))
   IF m_pRenderTarget THEN m_pRenderTarget->Release
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the underlying ID2D1RenderTarget interface pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF renderTarget AS CID2D1RenderTarget) AS ID2D1RenderTarget PTR
   RETURN renderTarget.m_pRenderTarget
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of the underlying ID2D1RenderTarget interface pointer
' ========================================================================================
PRIVATE OPERATOR CID2D1RenderTarget.@ () AS ID2D1RenderTarget PTR PTR
   RETURN @m_pRenderTarget
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the underlying ID2D1RenderTarget interface pointer
' ========================================================================================
PRIVATE OPERATOR CID2D1RenderTarget.CAST () AS ID2D1RenderTarget PTR
   RETURN m_pRenderTarget
END OPERATOR
' ========================================================================================

' ========================================================================================
' ID2D1Resource
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.GetFactory () AS ID2D1Factory PTR
   D2D1_DP("")
   DIM factory AS ID2D1Factory PTR
   DIM pResource AS ID2D1Resource PTR = CAST(ID2D1RenderTarget PTR, m_pRenderTarget)
   pResource->GetFactory(factory)
   RETURN factory
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an uninitialized Direct2D bitmap.
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateBitmap (BYVAL size AS D2D1_SIZE_U, BYVAL srcData AS const ANY PTR, _
BYVAL pitch AS UINT32, BYREF bitmapProperties AS D2D1_BITMAP_PROPERTIES) AS ID2D1Bitmap PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmap AS ID2D1Bitmap PTR
   m_pRenderTarget->CreateBitmap(size, srcData, pitch, bitmapProperties, bitmap)
   RETURN bitmap
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateBitmap (BYVAL size AS D2D1_SIZE_U, BYVAL srcData AS const ANY PTR, _
BYVAL pitch AS UINT32, BYVAL bitmapProperties AS D2D1_BITMAP_PROPERTIES PTR) AS ID2D1Bitmap PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmap AS ID2D1Bitmap PTR
   SetLastError(m_pRenderTarget->CreateBitmap(size, srcData, pitch, BYVAL bitmapProperties, bitmap))
   RETURN bitmap
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an ID2D1Bitmap by copying the specified Microsoft Windows Imaging Component (WIC) bitmap.
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateBitmapFromWicBitmap (BYVAL wicBitmapSource AS IWICBitmapSource PTR, _
BYREF bitmapProperties AS const D2D1_BITMAP_PROPERTIES) AS ID2D1Bitmap PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmap AS ID2D1Bitmap PTR
   SetLastError(m_pRenderTarget->CreateBitmapFromWicBitmap(wicBitmapSource, bitmapProperties, bitmap))
   RETURN bitmap
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateBitmapFromWicBitmap (BYVAL wicBitmapSource AS IWICBitmapSource PTR, _
BYVAL bitmapProperties AS const D2D1_BITMAP_PROPERTIES PTR) AS ID2D1Bitmap PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmap AS ID2D1Bitmap PTR
   SetLastError(m_pRenderTarget->CreateBitmapFromWicBitmap(wicBitmapSource, BYVAL bitmapProperties, bitmap))
   RETURN bitmap
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateBitmapFromWicBitmap (BYVAL wicBitmapSource AS IWICBitmapSource PTR) AS ID2D1Bitmap PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmap AS ID2D1Bitmap PTR
   SetLastError(m_pRenderTarget->CreateBitmapFromWicBitmap(wicBitmapSource, BYVAL NULL, bitmap))
   RETURN bitmap
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateSharedBitmap (BYREF iid AS IID, BYVAL pdata AS ANY PTR, _
BYREF bitmapProperties AS const D2D1_BITMAP_PROPERTIES) AS ID2D1Bitmap PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmap AS ID2D1Bitmap PTR
   SetLastError(m_pRenderTarget->CreateSharedBitmap(iid, pdata, bitmapProperties, bitmap))
   RETURN bitmap
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateSharedBitmap (BYVAL iid AS IID PTR, BYVAL pdata AS ANY PTR, _
BYVAL bitmapProperties AS const D2D1_BITMAP_PROPERTIES PTR) AS ID2D1Bitmap PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmap AS ID2D1Bitmap PTR
   SetLastError(m_pRenderTarget->CreateSharedBitmap(BYVAL iid, pdata, BYVAL bitmapProperties, bitmap))
   RETURN bitmap
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an ID2D1BitmapBrush from the specified bitmap.
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateBitmapBrush (BYVAL bitmap AS ID2D1Bitmap PTR, BYREF bitmapBrushProperties AS const D2D1_BITMAP_BRUSH_PROPERTIES, _
BYREF brushProperties AS const D2D1_BRUSH_PROPERTIES) AS ID2D1BitmapBrush PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmapBrush AS ID2D1BitmapBrush PTR
   SetLastError(m_pRenderTarget->CreateBitmapBrush(bitmap, bitmapBrushProperties, brushProperties, bitmapBrush))
   RETURN bitmapBrush
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateBitmapBrush (BYVAL bitmap AS ID2D1Bitmap PTR, BYVAL bitmapBrushProperties AS const D2D1_BITMAP_BRUSH_PROPERTIES PTR, _
BYVAL brushProperties AS const D2D1_BRUSH_PROPERTIES PTR) AS ID2D1BitmapBrush PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmapBrush AS ID2D1BitmapBrush PTR
   SetLastError(m_pRenderTarget->CreateBitmapBrush(bitmap, BYVAL bitmapBrushProperties, BYVAL brushProperties, bitmapBrush))
   RETURN bitmapBrush
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateBitmapBrush (BYVAL bitmap AS ID2D1Bitmap PTR, _
BYREF bitmapBrushProperties AS const D2D1_BITMAP_BRUSH_PROPERTIES) AS ID2D1BitmapBrush PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmapBrush AS ID2D1BitmapBrush PTR
   SetLastError(m_pRenderTarget->CreateBitmapBrush(bitmap, bitmapBrushProperties, BYVAL NULL, bitmapBrush))
   RETURN bitmapBrush
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateBitmapBrush (BYVAL bitmap AS ID2D1Bitmap PTR, _
BYVAL bitmapBrushProperties AS const D2D1_BITMAP_BRUSH_PROPERTIES PTR) AS ID2D1BitmapBrush PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmapBrush AS ID2D1BitmapBrush PTR
   SetLastError(m_pRenderTarget->CreateBitmapBrush(bitmap, BYVAL bitmapBrushProperties, BYVAL NULL, bitmapBrush))
   RETURN bitmapBrush
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateSolidColorBrush (BYREF clr AS D2D1_COLOR_F, _
BYREF brushProperties AS const D2D1_BRUSH_PROPERTIES) AS ID2D1SolidColorBrush PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM solidColorBrush AS ID2D1SolidColorBrush PTR
   SetLastError(m_pRenderTarget->CreateSolidColorBrush(clr, brushProperties, solidColorBrush))
   RETURN solidColorBrush
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateSolidColorBrush (BYVAL clr AS D2D1_COLOR_F PTR, _
BYVAL brushProperties AS const D2D1_BRUSH_PROPERTIES PTR) AS ID2D1SolidColorBrush PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM solidColorBrush AS ID2D1SolidColorBrush PTR
   SetLastError(m_pRenderTarget->CreateSolidColorBrush(BYVAL clr, BYVAL brushProperties, solidColorBrush))
   RETURN solidColorBrush
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an ID2D1GradientStopCollection from the specified array of D2D1_GRADIENT_STOP structures.
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateGradientStopCollection (BYREF gradientStops AS D2D1_GRADIENT_STOP, _
BYVAL gradientStopsCount AS UINT32, BYVAL colorInterpolationGamma AS D2D1_GAMMA, _
BYVAL extendMode AS D2D1_EXTEND_MODE) AS ID2D1GradientStopCollection PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM gradientStopCollection AS ID2D1GradientStopCollection PTR
   SetLastError(m_pRenderTarget->CreateGradientStopCollection(gradientStops, gradientStopsCount, colorInterpolationGamma, _
      extendMode, gradientStopCollection))
   RETURN gradientStopCollection
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateGradientStopCollection (BYVAL gradientStops AS D2D1_GRADIENT_STOP PTR, _
BYVAL gradientStopsCount AS UINT32, BYVAL colorInterpolationGamma AS D2D1_GAMMA, _
BYVAL extendMode AS D2D1_EXTEND_MODE) AS ID2D1GradientStopCollection PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM gradientStopCollection AS ID2D1GradientStopCollection PTR
   SetLastError(m_pRenderTarget->CreateGradientStopCollection(BYVAL gradientStops, gradientStopsCount, colorInterpolationGamma, _
      extendMode, gradientStopCollection))
   RETURN gradientStopCollection
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateGradientStopCollection (BYREF gradientStops AS D2D1_GRADIENT_STOP, _
BYVAL gradientStopsCount AS UINT32) AS ID2D1GradientStopCollection PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM gradientStopCollection AS ID2D1GradientStopCollection PTR
   SetLastError(m_pRenderTarget->CreateGradientStopCollection(gradientStops, gradientStopsCount, D2D1_GAMMA_2_2, _
      D2D1_EXTEND_MODE_CLAMP, gradientStopCollection))
   RETURN gradientStopCollection
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateGradientStopCollection (BYVAL gradientStops AS D2D1_GRADIENT_STOP PTR, _
BYVAL gradientStopsCount AS UINT32) AS ID2D1GradientStopCollection PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM gradientStopCollection AS ID2D1GradientStopCollection PTR
   SetLastError(m_pRenderTarget->CreateGradientStopCollection(BYVAL gradientStops, gradientStopsCount, D2D1_GAMMA_2_2, _
      D2D1_EXTEND_MODE_CLAMP, gradientStopCollection))
   RETURN gradientStopCollection
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an ID2D1LinearGradientBrush object for painting areas with a linear gradient.
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateLinearGradientBrush ( _
BYREF linearGradientBrushProperties AS const D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, _
BYREF brushProperties AS const D2D1_BRUSH_PROPERTIES, _
BYVAL gradientStopCollection AS ID2D1GradientStopCollection PTR) AS ID2D1LinearGradientBrush PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM linearGradientBrush AS ID2D1LinearGradientBrush PTR
   SetLastError(m_pRenderTarget->CreateLinearGradientBrush(linearGradientBrushProperties, brushProperties, _
      gradientStopCollection, linearGradientBrush))
   RETURN linearGradientBrush
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateLinearGradientBrush ( _
BYVAL linearGradientBrushProperties AS const D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES PTR, _
BYVAL brushProperties AS const D2D1_BRUSH_PROPERTIES PTR, _
BYVAL gradientStopCollection AS ID2D1GradientStopCollection PTR) AS ID2D1LinearGradientBrush PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM linearGradientBrush AS ID2D1LinearGradientBrush PTR
   SetLastError(m_pRenderTarget->CreateLinearGradientBrush(BYVAL linearGradientBrushProperties, BYVAL brushProperties, _
      gradientStopCollection, linearGradientBrush))
   RETURN linearGradientBrush
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateLinearGradientBrush ( _
BYREF linearGradientBrushProperties AS const D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES, _
BYVAL gradientStopCollection AS ID2D1GradientStopCollection PTR) AS ID2D1LinearGradientBrush PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM linearGradientBrush AS ID2D1LinearGradientBrush PTR
   SetLastError(m_pRenderTarget->CreateLinearGradientBrush(linearGradientBrushProperties, BYVAL NULL, _
      gradientStopCollection, linearGradientBrush))
   RETURN linearGradientBrush
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates an ID2D1RadialGradientBrush object that can be used to paint areas with a radial gradient.
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateRadialGradientBrush ( _
BYREF radialGradientBrushProperties AS const D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, _
BYREF brushProperties AS const D2D1_BRUSH_PROPERTIES, _
BYVAL gradientStopCollection AS ID2D1GradientStopCollection PTR) AS ID2D1RadialGradientBrush PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM radialGradientBrush AS ID2D1RadialGradientBrush PTR
   SetLastError(m_pRenderTarget->CreateRadialGradientBrush(radialGradientBrushProperties, brushProperties, _
      gradientStopCollection, radialGradientBrush))
   RETURN radialGradientBrush
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateRadialGradientBrush ( _
BYVAL radialGradientBrushProperties AS const D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES PTR, _
BYVAL brushProperties AS const D2D1_BRUSH_PROPERTIES PTR, _
BYVAL gradientStopCollection AS ID2D1GradientStopCollection PTR) AS ID2D1RadialGradientBrush PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM radialGradientBrush AS ID2D1RadialGradientBrush PTR
   SetLastError(m_pRenderTarget->CreateRadialGradientBrush(BYVAL radialGradientBrushProperties, BYVAL brushProperties, _
      gradientStopCollection, radialGradientBrush))
   RETURN radialGradientBrush
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateRadialGradientBrush ( _
BYREF radialGradientBrushProperties AS const D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES, _
BYVAL gradientStopCollection AS ID2D1GradientStopCollection PTR) AS ID2D1RadialGradientBrush PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM radialGradientBrush AS ID2D1RadialGradientBrush PTR
   SetLastError(m_pRenderTarget->CreateRadialGradientBrush(radialGradientBrushProperties, BYVAL NULL, _
      gradientStopCollection, radialGradientBrush))
   RETURN radialGradientBrush
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a new bitmap render target for use during intermediate offscreen drawing that is
' compatible with the current render target .
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateCompatibleRenderTarget (BYREF desiredSize AS const D2D1_SIZE_F, _
BYREF desiredPixelSize AS const D2D1_SIZE_U, BYREF desiredFormat AS const D2D1_PIXEL_FORMAT, _
BYVAL options AS D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS) AS ID2D1BitmapRenderTarget PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmapRenderTarget AS ID2D1BitmapRenderTarget PTR
   SetLastError(m_pRenderTarget->CreateCompatibleRenderTarget(desiredSize, desiredPixelSize, _
      desiredFormat, options, bitmapRenderTarget))
   RETURN bitmapRenderTarget
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateCompatibleRenderTarget (BYVAL desiredSize AS const D2D1_SIZE_F PTR, _
BYVAL desiredPixelSize AS const D2D1_SIZE_U PTR, BYVAL desiredFormat AS const D2D1_PIXEL_FORMAT PTR, _
BYVAL options AS D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS) AS ID2D1BitmapRenderTarget PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmapRenderTarget AS ID2D1BitmapRenderTarget PTR
   SetLastError(m_pRenderTarget->CreateCompatibleRenderTarget(BYVAL desiredSize, BYVAL desiredPixelSize, _
      BYVAL desiredFormat, options, bitmapRenderTarget))
   RETURN bitmapRenderTarget
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateCompatibleRenderTarget () AS ID2D1BitmapRenderTarget PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmapRenderTarget AS ID2D1BitmapRenderTarget PTR
   SetLastError(m_pRenderTarget->CreateCompatibleRenderTarget(BYVAL NULL, BYVAL NULL, _
      BYVAL NULL, D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS_NONE, bitmapRenderTarget))
   RETURN bitmapRenderTarget
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateCompatibleRenderTarget (BYREF desiredSize AS const D2D1_SIZE_F) AS ID2D1BitmapRenderTarget PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmapRenderTarget AS ID2D1BitmapRenderTarget PTR
   SetLastError(m_pRenderTarget->CreateCompatibleRenderTarget(desiredSize, BYVAL NULL, _
      BYVAL NULL, D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS_NONE, bitmapRenderTarget))
   RETURN bitmapRenderTarget
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateCompatibleRenderTarget (BYVAL desiredSize AS const D2D1_SIZE_F PTR) AS ID2D1BitmapRenderTarget PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmapRenderTarget AS ID2D1BitmapRenderTarget PTR
   SetLastError(m_pRenderTarget->CreateCompatibleRenderTarget(BYVAL desiredSize, BYVAL NULL, _
      BYVAL NULL, D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS_NONE, bitmapRenderTarget))
   RETURN bitmapRenderTarget
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateCompatibleRenderTarget (BYREF desiredSize AS const D2D1_SIZE_F, _
BYREF desiredPixelSize AS const D2D1_SIZE_U) AS ID2D1BitmapRenderTarget PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmapRenderTarget AS ID2D1BitmapRenderTarget PTR
   SetLastError(m_pRenderTarget->CreateCompatibleRenderTarget(desiredSize, desiredPixelSize, _
      BYVAL NULL, D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS_NONE, bitmapRenderTarget))
   RETURN bitmapRenderTarget
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateCompatibleRenderTarget (BYVAL desiredSize AS const D2D1_SIZE_F PTR, _
BYVAL desiredPixelSize AS const D2D1_SIZE_U PTR) AS ID2D1BitmapRenderTarget PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM bitmapRenderTarget AS ID2D1BitmapRenderTarget PTR
   SetLastError(m_pRenderTarget->CreateCompatibleRenderTarget(BYVAL desiredSize, BYVAL desiredPixelSize, _
      BYVAL NULL, D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS_NONE, bitmapRenderTarget))
   RETURN bitmapRenderTarget
END FUNCTION
' ========================================================================================

' ========================================================================================
' Creates a layer resource that can be used with this render target and its compatible render targets.
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateLayer (BYREF size AS const D2D1_SIZE_F) AS ID2D1Layer PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM layer AS ID2D1Layer PTR
   SetLastError(m_pRenderTarget->CreateLayer(size, layer))
   RETURN layer
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateLayer (BYVAL size AS const D2D1_SIZE_F PTR) AS ID2D1Layer PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM layer AS ID2D1Layer PTR
   SetLastError(m_pRenderTarget->CreateLayer(BYVAL size, layer))
   RETURN layer
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateLayer () AS ID2D1Layer PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM layer AS ID2D1Layer PTR
   SetLastError(m_pRenderTarget->CreateLayer(BYVAL NULL, layer))
   RETURN layer
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.CreateMesh () AS ID2D1Mesh PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN NULL
   DIM mesh AS ID2D1Mesh PTR
   SetLastError(m_pRenderTarget->CreateMesh(mesh))
   RETURN mesh
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawLine (BYVAL point0 AS D2D1_POINT_2F, BYVAL point1 AS D2D1_POINT_2F, BYVAL brush AS ID2D1Brush PTR, _
BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawLine(point0, point1, brush, strokeWidth, strokeStyle)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawRectangle (BYREF rect AS const D2D1_RECT_F, BYVAL brush AS ID2D1Brush PTR, _
BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawRectangle(rect, brush, strokeWidth, strokeStyle)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawRectangle (BYVAL rect AS const D2D1_RECT_F PTR, BYVAL brush AS ID2D1Brush PTR, _
BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawRectangle(BYVAL rect, brush, strokeWidth, strokeStyle)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.FillRectangle (BYREF rect AS const D2D1_RECT_F, BYVAL brush AS ID2D1Brush PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->FillRectangle(rect, brush)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.FillRectangle (BYVAL rect AS const D2D1_RECT_F PTR, BYVAL brush AS ID2D1Brush PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->FillRectangle(BYVAL rect, brush)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawRoundedRectangle (BYREF roundedRect AS const D2D1_ROUNDED_RECT, BYVAL brush AS ID2D1Brush PTR, _
BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawRoundedRectangle(roundedRect, brush, strokeWidth, strokeStyle)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawRoundedRectangle (BYVAL roundedRect AS const D2D1_ROUNDED_RECT PTR, BYVAL brush AS ID2D1Brush PTR, _
BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawRoundedRectangle(BYVAL roundedRect, brush, strokeWidth, strokeStyle)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.FillRoundedRectangle (BYREF roundedRect AS const D2D1_ROUNDED_RECT, BYVAL brush AS ID2D1Brush PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->FillRoundedRectangle(roundedRect, brush)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.FillRoundedRectangle (BYVAL roundedRect AS const D2D1_ROUNDED_RECT PTR, BYVAL brush AS ID2D1Brush PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->FillRoundedRectangle(BYVAL roundedRect, brush)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawEllipse (BYREF tEllipse AS const D2D1_ELLIPSE, BYVAL brush AS ID2D1Brush PTR, _
BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL) 
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawEllipse(tEllipse, brush, strokeWidth, strokeStyle)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawEllipse (BYVAL tEllipse AS const D2D1_ELLIPSE PTR, BYVAL brush AS ID2D1Brush PTR, _
BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL) 
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawEllipse(BYVAL tEllipse, brush, strokeWidth, strokeStyle)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.FillEllipse (BYREF tEllipse AS const D2D1_ELLIPSE, BYVAL brush AS ID2D1Brush PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->FillEllipse(tEllipse, brush)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.FillEllipse (BYVAL tEllipse AS const D2D1_ELLIPSE PTR, BYVAL brush AS ID2D1Brush PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->FillEllipse(BYVAL tEllipse, brush)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawGeometry (BYVAL geometry AS ID2D1Geometry PTR, BYVAL brush AS ID2D1Brush PTR, _
BYVAL strokeWidth AS FLOAT = 1.0!, BYVAL strokeStyle AS ID2D1StrokeStyle PTR = NULL) 
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawGeometry(geometry, brush, strokeWidth, strokeStyle)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.FillGeometry (BYVAL geometry AS ID2D1Geometry PTR, BYVAL brush AS ID2D1Brush PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->FillGeometry(geometry, brush)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.FillMesh (BYVAL mesh AS ID2D1Mesh PTR, BYVAL brush AS ID2D1Brush PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->FillMesh(mesh, brush)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.FillOpacityMask (BYVAL opacityMask AS ID2D1Bitmap PTR, BYVAL brush AS ID2D1Brush PTR, _
BYVAL content AS D2D1_OPACITY_MASK_CONTENT, BYREF destinationRectangle AS const D2D1_RECT_F, _
BYREF sourceRectangle AS const D2D1_RECT_F)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->FillOpacityMask(opacityMask, brush, content, destinationRectangle, sourceRectangle)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.FillOpacityMask (BYVAL opacityMask AS ID2D1Bitmap PTR, BYVAL brush AS ID2D1Brush PTR, _
BYVAL content AS D2D1_OPACITY_MASK_CONTENT, BYVAL destinationRectangle AS const D2D1_RECT_F PTR, _
BYVAL sourceRectangle AS const D2D1_RECT_F PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->FillOpacityMask(opacityMask, brush, content, BYVAL destinationRectangle, BYVAL sourceRectangle)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawBitmap (BYVAL bitmap AS ID2D1Bitmap PTR, BYREF destinationRectangle AS const D2D1_RECT_F, _
BYVAL opacity AS FLOAT = 1.0!, BYVAL InterpolationMode AS D2D1_BITMAP_INTERPOLATION_MODE = D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, _
BYREF sourceRectangle AS const D2D1_RECT_F)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawBitmap(bitmap, destinationRectangle, opacity, InterpolationMode, sourceRectangle)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawBitmap (BYVAL bitmap AS ID2D1Bitmap PTR, BYVAL destinationRectangle AS const D2D1_RECT_F PTR, _
BYVAL opacity AS FLOAT = 1.0!, BYVAL InterpolationMode AS D2D1_BITMAP_INTERPOLATION_MODE = D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, _
BYVAL sourceRectangle AS const D2D1_RECT_F PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawBitmap(bitmap, BYVAL destinationRectangle, opacity, InterpolationMode, BYVAL sourceRectangle)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawText (BYREF wzSstr AS WSTRING, BYVAL stringLength AS UINT32, _
BYVAL textFormat AS IDWriteTextFormat PTR, BYREF layoutRect AS D2D1_RECT_F, BYVAL defaultForegroundBrush AS ID2D1Brush PTR, _
BYVAL options AS D2D1_DRAW_TEXT_OPTIONS = D2D1_DRAW_TEXT_OPTIONS_NONE, _
BYVAL measuringMode AS DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE_NATURAL)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawText(wzSstr, stringLength, textFormat, layoutRect, defaultForegroundBrush, options, measuringMode)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawText (BYVAL pwzSstr AS WSTRING PTR, BYVAL stringLength AS UINT32, _
BYVAL textFormat AS IDWriteTextFormat PTR, BYVAL layoutRect AS D2D1_RECT_F PTR, BYVAL defaultForegroundBrush AS ID2D1Brush PTR, _
BYVAL options AS D2D1_DRAW_TEXT_OPTIONS = D2D1_DRAW_TEXT_OPTIONS_NONE, _
BYVAL measuringMode AS DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE_NATURAL)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawText(BYVAL pwzSstr, stringLength, textFormat, BYVAL layoutRect, defaultForegroundBrush, options, measuringMode)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawTextLayout (BYVAL origin AS D2D1_POINT_2F, BYVAL textLayout AS IDWriteTextLayout PTR, _
BYVAL defaultFillBrush AS ID2D1Brush PTR, BYVAL options AS D2D1_DRAW_TEXT_OPTIONS = D2D1_DRAW_TEXT_OPTIONS_NONE) 
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawTextLayout(origin, textLayout, defaultFillBrush, options)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawGlyphRun (BYVAL baselineOrigin AS D2D1_POINT_2F, BYREF glyphRun AS const DWRITE_GLYPH_RUN, _
BYVAL foregroundBrush AS ID2D1Brush PTR, BYVAL measuringMode AS DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE_NATURAL)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawGlyphRun(baselineOrigin, glyphRun, foregroundBrush, measuringMode)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.DrawGlyphRun (BYVAL baselineOrigin AS D2D1_POINT_2F, BYVAL glyphRun AS const DWRITE_GLYPH_RUN PTR, _
BYVAL foregroundBrush AS ID2D1Brush PTR, BYVAL measuringMode AS DWRITE_MEASURING_MODE = DWRITE_MEASURING_MODE_NATURAL)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->DrawGlyphRun(baselineOrigin, BYVAL glyphRun, foregroundBrush, measuringMode)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.SetTransform (BYREF transform AS D2D1_MATRIX_3X2_F)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->SetTransform(transform)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.SetTransform (BYVAL transform AS D2D1_MATRIX_3X2_F PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->SetTransform(BYVAL transform)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.GetTransform (BYREF transform AS D2D1_MATRIX_3X2_F)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->GetTransform(transform)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.GetTransform (BYVAL transform AS D2D1_MATRIX_3X2_F PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->GetTransform(BYVAL transform)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.GetTransform () AS D2D1_MATRIX_3X2_F
   D2D1_DP("")
   DIM transform AS D2D1_MATRIX_3X2_F
   IF m_pRenderTarget = NULL THEN RETURN transform
   m_pRenderTarget->GetTransform(transform)
   RETURN transform
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.SetAntialiasMode (BYVAL antialiasMode AS D2D1_ANTIALIAS_MODE)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->SetAntialiasMode(antialiasMode)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.GetAntialiasMode () AS D2D1_ANTIALIAS_MODE
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT FUNCTION
   RETURN m_pRenderTarget->GetAntialiasMode
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.SetTextAntialiasMode (BYVAL textAntialiasMode AS D2D1_TEXT_ANTIALIAS_MODE)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->SetTextAntialiasMode(textAntialiasMode)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.GetTextAntialiasMode () AS D2D1_TEXT_ANTIALIAS_MODE
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT FUNCTION
   RETURN m_pRenderTarget->GetTextAntialiasMode
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.SetTextRenderingParams (BYVAL textRenderingParams AS IDWriteRenderingParams PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->SetTextRenderingParams(textRenderingParams)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.GetTextRenderingParams (BYREF textRenderingParams AS IDWriteRenderingParams PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->GetTextRenderingParams(textRenderingParams)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.GetTextRenderingParams () AS IDWriteRenderingParams PTR
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT FUNCTION
   DIM textRenderingParams AS IDWriteRenderingParams PTR
   m_pRenderTarget->GetTextRenderingParams(textRenderingParams)
   RETURN textRenderingParams
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.SetTags (BYVAL tag1 AS D2D1_TAG, BYVAL tag2 AS D2D1_TAG)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->SetTags(tag1, tag2)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.GetTags (BYREF tag1 AS D2D1_TAG = NULL, BYREF tag2 AS D2D1_TAG = NULL)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->GetTags(tag1, tag2)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.GetTags (BYVAL tag1 AS D2D1_TAG PTR = NULL, BYVAL tag2 AS D2D1_TAG PTR = NULL)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->GetTags(BYVAL tag1, BYVAL tag2)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.PushLayer (BYREF layerParameters AS D2D1_LAYER_PARAMETERS, BYVAL layer AS ID2D1Layer PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->PushLayer(layerParameters, layer)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.PushLayer (BYVAL layerParameters AS D2D1_LAYER_PARAMETERS PTR, BYVAL layer AS ID2D1Layer PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->PushLayer(BYVAL layerParameters, layer)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.PopLayer
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->PopLayer
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.Flush (BYREF tag1 AS D2D1_TAG = NULL, BYREF tag2 AS D2D1_TAG = NULL) AS HRESULT
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT FUNCTION
   RETURN m_pRenderTarget->Flush(tag1, tag2)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.Flush (BYVAL tag1 AS D2D1_TAG PTR = NULL, BYVAL tag2 AS D2D1_TAG PTR = NULL) AS HRESULT
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT FUNCTION
   RETURN m_pRenderTarget->Flush(BYVAL tag1, BYVAL tag2)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.SaveDrawingState (BYVAL drawingStateBlock AS ID2D1DrawingStateBlock PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->SaveDrawingState(drawingStateBlock)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.RestoreDrawingState (BYVAL drawingStateBlock AS ID2D1DrawingStateBlock PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->RestoreDrawingState(drawingStateBlock)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.PushAxisAlignedClip (BYREF clipRect AS D2D1_RECT_F, BYVAL antialiasMode AS D2D1_ANTIALIAS_MODE)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->PushAxisAlignedClip(clipRect, antialiasMode)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.PushAxisAlignedClip (BYVAL clipRect AS D2D1_RECT_F PTR, BYVAL antialiasMode AS D2D1_ANTIALIAS_MODE)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->PushAxisAlignedClip(BYVAL clipRect, antialiasMode)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.PopAxisAlignedClip
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->PopAxisAlignedClip
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.Clear (BYREF clearColor AS const D2D1_COLOR_F)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->Clear(clearColor)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.Clear (BYVAL clearColor AS const D2D1_COLOR_F PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->Clear(BYVAL clearColor)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.BeginDraw
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->BeginDraw
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.EndDraw (BYREF tag1 AS D2D1_TAG, BYREF tag2 AS D2D1_TAG) AS HRESULT
   D2D1_DP("BYREF")
   IF m_pRenderTarget = NULL THEN RETURN E_POINTER
   RETURN m_pRenderTarget->EndDraw(tag1, tag2)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.EndDraw (BYVAL tag1 AS D2D1_TAG PTR = NULL, BYVAL tag2 AS D2D1_TAG PTR = NULL) AS HRESULT
   D2D1_DP("BYVAL PTR")
   IF m_pRenderTarget = NULL THEN RETURN E_POINTER
   RETURN m_pRenderTarget->EndDraw(BYVAL tag1, BYVAL tag2)
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.GetPixelFormat (BYREF pixelFormat AS D2D1_PIXEL_FORMAT)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->GetPixelFormat(pixelFormat)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.GetPixelFormat (BYVAL pixelFormat AS D2D1_PIXEL_FORMAT PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->GetPixelFormat(BYVAL pixelFormat)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.GetPixelFormat () AS D2D1_PIXEL_FORMAT
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT FUNCTION
   DIM pixelFormat AS D2D1_PIXEL_FORMAT
   m_pRenderTarget->GetPixelFormat(pixelFormat)
   RETURN pixelFormat
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.SetDpi (BYVAL dpiX AS FLOAT, BYVAL dpiY AS FLOAT)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->SetDpi(dpiX, dpiY)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.GetDpi (BYREF dpiX AS FLOAT, BYREF dpiY AS FLOAT)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->GetDpi(dpiX, dpiY)
END SUB
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.GetSize (BYREF size AS D2D1_SIZE_F)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->GetSize(size)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.GetSize (BYVAL size AS D2D1_SIZE_F PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->GetSize(BYVAL size)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.GetSize () AS D2D1_SIZE_F
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT FUNCTION
   DIM size AS D2D1_SIZE_F
   m_pRenderTarget->GetSize(size)
   RETURN size
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.GetPixelSize (BYREF pixelSize AS D2D1_SIZE_U)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->GetPixelSize(pixelSize)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE SUB CID2D1RenderTarget.GetPixelSize (BYVAL pixelSize AS D2D1_SIZE_U PTR)
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT SUB
   m_pRenderTarget->GetPixelSize(BYVAL pixelSize)
END SUB
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.GetPixelSize () AS D2D1_SIZE_U
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN EXIT FUNCTION
   DIM pixelSize AS D2D1_SIZE_U
   m_pRenderTarget->GetPixelSize(pixelSize)
   RETURN pixelSize
END FUNCTION
' ========================================================================================

' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.IsSupported (BYREF renderTargetProperties AS const D2D1_RENDER_TARGET_PROPERTIES) AS BOOLEAN
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN FALSE
   m_pRenderTarget->IsSupported(renderTargetProperties)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1RenderTarget.IsSupported (BYVAL renderTargetProperties AS const D2D1_RENDER_TARGET_PROPERTIES PTR) AS BOOLEAN
   D2D1_DP("")
   IF m_pRenderTarget = NULL THEN RETURN FALSE
   m_pRenderTarget->IsSupported(BYVAL renderTargetProperties)
END FUNCTION
' ========================================================================================


' ########################################################################################
' CID2D1DCRenderTarget class
' ########################################################################################

' ========================================================================================
TYPE CID2D1DCRenderTarget EXTENDS CID2D1RenderTarget
Public:

   m_pDCRenderTarget AS ID2D1DCRenderTarget PTR

   DECLARE CONSTRUCTOR (BYVAL pDCRenderTarget AS ID2D1DCRenderTarget PTR)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS ID2D1DCRenderTarget PTR PTR
   DECLARE OPERATOR CAST () AS ID2D1DCRenderTarget PTR
   DECLARE FUNCTION BindDC (BYVAL hDC AS HDC, BYREF subRect AS const RECT) AS HRESULT
   DECLARE FUNCTION BindDC (BYVAL hDC AS HDC, BYVAL subRect AS const RECT PTR) AS HRESULT
   DECLARE FUNCTION BindDC (BYVAL hDC AS HDC, BYVAL nLeft AS LONG, BYVAL nTop AS LONG, BYVAL nRight AS LONG, BYVAL nBottom AS LONG) AS HRESULT

END TYPE

' ========================================================================================
' Creates an instance of the CID2D1DCRenderTarget class.
' ========================================================================================
PRIVATE CONSTRUCTOR CID2D1DCRenderTarget (BYVAL pDCRenderTarget AS ID2D1DCRenderTarget PTR)
   D2D1_DP("pDCRenderTarget - " & WSTR(pDCRenderTarget))
   m_pDCRenderTarget = pDCRenderTarget
   IF m_pDCRenderTarget THEN m_pDCRenderTarget->AddRef
   m_pRenderTarget = CAST(ID2D1RenderTarget PTR, m_pDCRenderTarget)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CID2D1DCRenderTarget
   D2D1_DP("m_pDCRenderTarget: " & WSTR(m_pDCRenderTarget))
   IF m_pDCRenderTarget THEN m_pDCRenderTarget->Release
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the underlying ID2D1DCRenderTarget interface pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF renderTarget AS CID2D1DCRenderTarget) AS ID2D1RenderTarget PTR
   RETURN renderTarget.m_pDCRenderTarget
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of the underlying ID2D1DCRenderTarget interface pointer
' ========================================================================================
PRIVATE OPERATOR CID2D1DCRenderTarget.@ () AS ID2D1DCRenderTarget PTR PTR
   RETURN @m_pDCRenderTarget
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the underlying ID2D1DCRenderTarget interface pointer
' ========================================================================================
PRIVATE OPERATOR CID2D1DCRenderTarget.CAST () AS ID2D1DCRenderTarget PTR
   RETURN m_pDCRenderTarget
END OPERATOR
' ========================================================================================

' ========================================================================================
' Binds the render target to the device context to which it issues drawing commands.
' ========================================================================================
PRIVATE FUNCTION CID2D1DCRenderTarget.BindDC (BYVAL hDC AS HDC, BYREF subRect AS const RECT) AS HRESULT
   D2D1_DP("")
   IF m_pDCRenderTarget = NULL THEN RETURN E_POINTER
   RETURN m_pDCRenderTarget->BindDC(hdc, subRect)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1DCRenderTarget.BindDC (BYVAL hDC AS HDC, BYVAL subRect AS const RECT PTR) AS HRESULT
   D2D1_DP("")
   IF m_pDCRenderTarget = NULL THEN RETURN E_POINTER
   RETURN m_pDCRenderTarget->BindDC(hdc, BYVAL subRect)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1DCRenderTarget.BindDC (BYVAL hDC AS HDC, BYVAL nLeft AS LONG, BYVAL nTop AS LONG, BYVAL nRight AS LONG, BYVAL nBottom AS LONG) AS HRESULT
   D2D1_DP("")
   IF m_pDCRenderTarget = NULL THEN RETURN E_POINTER
   DIM rc AS RECT = (nLeft, nTop, nRight, nBottom)
   RETURN m_pDCRenderTarget->BindDC(hdc, rc)
END FUNCTION
' ========================================================================================


' ########################################################################################
' CID2D1HwndRenderTarget class
' Renders drawing instructions to a window.
' ########################################################################################

' ========================================================================================
TYPE CID2D1HwndRenderTarget EXTENDS CID2D1RenderTarget
Public:

   m_pHwndRenderTarget AS ID2D1HwndRenderTarget PTR

   DECLARE CONSTRUCTOR (BYVAL pHwndRenderTarget AS ID2D1HwndRenderTarget PTR)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS ID2D1HwndRenderTarget PTR PTR
   DECLARE OPERATOR CAST () AS ID2D1HwndRenderTarget PTR
   DECLARE FUNCTION CheckWindowState () AS D2D1_WINDOW_STATE
   DECLARE FUNCTION Resize (BYREF pixelSize AS D2D1_SIZE_U) AS HRESULT
   DECLARE FUNCTION Resize (BYVAL pixelSize AS D2D1_SIZE_U PTR) AS HRESULT
   DECLARE FUNCTION Resize () AS D2D1_SIZE_U
   DECLARE FUNCTION GetHwnd () AS HWND

END TYPE

' ========================================================================================
' Creates an instance of the CID2D1HwndRenderTarget class.
' ========================================================================================
PRIVATE CONSTRUCTOR CID2D1HwndRenderTarget (BYVAL pHwndRenderTarget AS ID2D1HwndRenderTarget PTR)
   D2D1_DP("HwndRenderTarget - " & WSTR(pHwndRenderTarget))
   m_pHwndRenderTarget = pHwndRenderTarget
   IF m_pHwndRenderTarget THEN m_pHwndRenderTarget->AddRef
   m_pRenderTarget = CAST(ID2D1RenderTarget PTR, m_pHwndRenderTarget)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CID2D1HwndRenderTarget
   D2D1_DP("m_pHwndRenderTarget: " & WSTR(m_pHwndRenderTarget))
   IF m_pHwndRenderTarget THEN m_pHwndRenderTarget->Release
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the underlying ID2D1HwndRenderTarget interface pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF renderTarget AS CID2D1HwndRenderTarget) AS ID2D1HwndRenderTarget PTR
   RETURN renderTarget.m_pHwndRenderTarget
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of the underlying ID2D1HwndRenderTarget interface pointer
' ========================================================================================
PRIVATE OPERATOR CID2D1HwndRenderTarget.@ () AS ID2D1HwndRenderTarget PTR PTR
   RETURN @m_pHwndRenderTarget
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the underlying ID2D1HwndRenderTarget interface pointer
' ========================================================================================
PRIVATE OPERATOR CID2D1HwndRenderTarget.CAST () AS ID2D1HwndRenderTarget PTR
   RETURN m_pHwndRenderTarget
END OPERATOR
' ========================================================================================

' ========================================================================================
' Indicates whether the HWND associated with this render target is occluded.
' ========================================================================================
PRIVATE FUNCTION CID2D1HwndRenderTarget.CheckWindowState () AS D2D1_WINDOW_STATE
   D2D1_DP("")
   IF m_pHwndRenderTarget = NULL THEN RETURN 0
   RETURN m_pHwndRenderTarget->CheckWindowState
END FUNCTION
' ========================================================================================

' ========================================================================================
' Changes the size of the render target to the specified pixel size.
' ========================================================================================
PRIVATE FUNCTION CID2D1HwndRenderTarget.Resize (BYREF pixelSize AS D2D1_SIZE_U) AS HRESULT
   D2D1_DP("")
   IF m_pHwndRenderTarget = NULL THEN RETURN 0
   RETURN m_pHwndRenderTarget->Resize(pixelSize)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1HwndRenderTarget.Resize (BYVAL pixelSize AS D2D1_SIZE_U PTR) AS HRESULT
   D2D1_DP("")
   IF m_pHwndRenderTarget = NULL THEN RETURN 0
   RETURN m_pHwndRenderTarget->Resize(BYVAL pixelSize)
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CID2D1HwndRenderTarget.Resize () AS D2D1_SIZE_U
   D2D1_DP("")
   IF m_pHwndRenderTarget = NULL THEN SetLastError(E_POINTER) : EXIT FUNCTION
   DIM pixelSize AS D2D1_SIZE_U
   SetLastError(m_pHwndRenderTarget->Resize(pixelSize))
   RETURN pixelSize
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns the HWND associated with this render target.
' ========================================================================================
PRIVATE FUNCTION CID2D1HwndRenderTarget.GetHwnd () AS HWND
   D2D1_DP("")
   IF m_pHwndRenderTarget = NULL THEN SetLastError(E_POINTER) : EXIT FUNCTION
   RETURN m_pHwndRenderTarget->GetHwnd
END FUNCTION
' ========================================================================================


' ########################################################################################
' CID2D1BitmapRenderTarget class
' Renders to an intermediate texture created by the CreateCompatibleRenderTarget method.
' ########################################################################################

' ========================================================================================
TYPE CID2D1BitmapRenderTarget EXTENDS CID2D1RenderTarget
Public:

   m_pBitmapRenderTarget AS ID2D1BitmapRenderTarget PTR

   DECLARE CONSTRUCTOR (BYVAL pBitmapRenderTarget AS ID2D1BitmapRenderTarget PTR)
   DECLARE DESTRUCTOR
   DECLARE OPERATOR @ () AS ID2D1BitmapRenderTarget PTR PTR
   DECLARE OPERATOR CAST () AS ID2D1BitmapRenderTarget PTR
   DECLARE FUNCTION GetBitmap () AS ID2D1Bitmap PTR

END TYPE

' ========================================================================================
' Creates an instance of the CID2D1HwndRenderTarget class.
' ========================================================================================
PRIVATE CONSTRUCTOR CID2D1BitmapRenderTarget (BYVAL pBitmapRenderTarget AS ID2D1BitmapRenderTarget PTR)
   D2D1_DP("BitmapRenderTarget - " & WSTR(pBitmapRenderTarget))
   m_pBitmapRenderTarget = pBitmapRenderTarget
   IF m_pBitmapRenderTarget THEN m_pBitmapRenderTarget->AddRef
   m_pRenderTarget = CAST(ID2D1RenderTarget PTR, m_pBitmapRenderTarget)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Cleanup
' ========================================================================================
PRIVATE DESTRUCTOR CID2D1BitmapRenderTarget
   D2D1_DP("m_pBitmapRenderTarget: " & WSTR(m_pBitmapRenderTarget))
   IF m_pBitmapRenderTarget THEN m_pBitmapRenderTarget->Release
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the underlying ID2D1BitmapRenderTarget interface pointer
' ========================================================================================
PRIVATE OPERATOR * (BYREF renderTarget AS CID2D1BitmapRenderTarget) AS ID2D1BitmapRenderTarget PTR
   RETURN renderTarget.m_pBitmapRenderTarget
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the address of the underlying ID2D1BitmapRenderTarget interface pointer
' ========================================================================================
PRIVATE OPERATOR CID2D1BitmapRenderTarget.@ () AS ID2D1BitmapRenderTarget PTR PTR
   RETURN @m_pBitmapRenderTarget
END OPERATOR
' ========================================================================================
' ========================================================================================
' Returns the underlying ID2D1BitmapRenderTarget interface pointer
' ========================================================================================
PRIVATE OPERATOR CID2D1BitmapRenderTarget.CAST () AS ID2D1BitmapRenderTarget PTR
   RETURN m_pBitmapRenderTarget
END OPERATOR
' ========================================================================================

' ========================================================================================
' Indicates whether the HWND associated with this render target is occluded.
' ========================================================================================
PRIVATE FUNCTION CID2D1BitmapRenderTarget.GetBitmap () AS ID2D1Bitmap PTR
   D2D1_DP("")
   IF m_pBitmapRenderTarget = NULL THEN RETURN NULL
   DIM bitmap AS ID2D1Bitmap PTR
   SetLasTError(m_pBitmapRenderTarget->GetBitmap(bitmap))
   RETURN bitmap
END FUNCTION
' ========================================================================================

END NAMESPACE
