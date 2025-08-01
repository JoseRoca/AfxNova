# Printer Functions



**Include file**: AfxNova/AfxPrinter.inc

### Methods

| Name       | Description |
| ---------- | ----------- |
| [AfxEnumLocalPrinterPorts](#afxenumlocalprinterports) | Returns a list with port names of the locally installed printers. |
| [AfxEnumPrinterNames](#afxenumprinternames) | Returns a list with the available printers, print servers, domains, or print providers. |
| [AfxEnumPrinterPorts](#afxenumprinterports) | Returns a list with the ports that available for printing on a specified server. |
| [AfxGetDefaultPrinter](#afxgetdefaultprinter) | Retrieves the name of the default printer. |
| [AfxGetDefaultPrinterDriver](#afxgetdefaultprinterdriver) | Retrieves the name of the default printer driver. |
| [AfxGetDefaultPrinterPort](#afxgetdefaultprinterport) | Retrieves the name of the default printer port. |
| [AfxGetPrinterCollate](#afxgetprintercollate) | If the printer supports collating, the return value is TRUE; otherwise, the return value is FALSE. |
| [AfxGetPrinterCollateStatus](#afxgetprintercollatestatus) | Returns the printer collate status. |
| [AfxGetPrinterColorMode](#afxgetprintercolormode) | If the printer supports color printing, the return value is TRUE; otherwise, the return value is FALSE. |
| [AfxGetPrinterCopies](#afxgetprintercopies) | Returns the number of copies printed if the device supports multiple-page copies. |
| [AfxGetPrinterDC](#afxgetprinterdc) | Returns the handle of the default printer device context. |
| [AfxGetPrinterDriverVersion](#afxgetprinterdriverversion) | Returns the version number of the Windows-based printer driver. |
| [AfxGetPrinterDuplex](#afxgetprinterduplex) | If the printer supports duplex printing, the return value is TRUE; otherwise, the return value is FALSE. |
| [AfxGetPrinterDuplexMode](#afxgetprinterduplexname) | If the printer supports duplex printing, returns the current duplex mode. |
| [AfxGetPrinterFromPort](#afxgetprinterfromport) | Returns the printer name for a given port name. |
| [AfxGetPrinterHorizontalResolution](#afxgetprinterhorizontalresolution) | Returns the width, in pixels, of the printable area of the page. |
| [AfxGetPrinterVerticalResolution](#afxgetprinterverticalresolution) | Returns the Height, in pixels, of the printable area of the page. |
| [AfxGetPrinterMaxCopies](#afxgetprintermaxcopies) | Returns the maximum number of copies the device can print. |
| [AfxGetPrinterMaxPaperHeight](#afxgetprintermaxpaperheight) | Returns the maximum paper height in tenths of a millimeter. |
| [AfxGetPrinterMaxPaperWidth](#afxgetprintermaxpaperwidth) | Returns the maximum paper width in tenths of a millimeter. |
| [AfxGetPrinterMediaReady](#afxgetprintermediaready) | Retrieves the names of the paper forms that are currently available for use. |
| [AfxGetPrinterMinPaperHeight](#afxgetprinterminpaperheight) | Returns the minimum paper height in tenths of a millimeter. |
| [AfxGetPrinterMinPaperWidth](#afxgetprinterminpaperWidth) | Returns the minimum paper width in tenths of a millimeter. |
| [AfxGetPrinterOrientation](#afxgetprinterorientation) | Returns the printer orientation. |
| [AfxGetPrinterOrientationDegrees](#afxgetprinterorientationdegrees) | Returns the relationship between portrait and landscape orientations for a device, in terms of the number of degrees that portrait orientation is rotated counterclockwise to produce landscape orientation. |
| [AfxGetPrinterPaperLength](#afxgetprinterpaperlength) | Returns the paper length. |
| [AfxGetPrinterPaperWidth](#afxgetprinterpaperwidth) | Returns the paper width. |
| [AfxGetPrinterPaperNames](#afxgetprinterpapernames) | Returns a list of supported paper names (for example, Letter or Legal). |
| [AfxGetPrinterPaperSize](#afxgetprinterpapersize) | Returns the paper size for which the printer is currently configured. |
| [AfxGetPrinterPaperSizes](#afxgetprinterpapersizes) | Returns a list of each supported paper sizes, in tenths of a millimeter. |
| [AfxGetPrinterPhysicalHeight](#Afxgetprinterphysicalheight) | Returns the height of the physical page, in device units. |
| [AfxGetPrinterPhysicalOffsetX](#Afxgetprinterphysicaloffsetx) | Returns the distance from the left edge of the physical page to the left edge of the printable area, in device units. |
| [AfxGetPrinterPhysicalOffsetY](#Afxgetprinterphysicaloffsety) | Returns the distance from the top edge of the physical page to the top edge of the printable area, in device units. |
| [AfxGetPrinterPhysicalWidth](#Afxgetprinterphysicalwidth) | Returns the width of the physical page, in device units. |
| [AfxGetPrinterPort](#afxgetprinterport) | Returns the port name for a given printer name. |
| [AfxGetPrinterPPIX](#afxgetprinterppix) | Retrieves the number of pixels per inch of the specified host printer page (horizontal resolution). |
| [AfxGetPrinterPPIY](#afxgetprinterppiy) | Retrieves the number of pixels per inch of the specified host printer page (vertical resolution). |
| [AfxGetPrinterQuality](#afxgetprinterquality) | Returns the printer print quality mode. |
| [AfxGetPrinterRate](#afxgetprinterrate) | Returns the printer's print rate. |
| [AfxGetPrinterRateUnit](#afxgetprinterrateunit) | Returns the printer's rate unit. |
| [AfxGetPrinterRatePPM](#afxgetprinterrateppm) | Returns the printer's print rate, in pages per minute. |
| [AfxGetPrinterScale](#afxgetprinterscale) | Returns the factor by which the printed output is to be scaled. |
| [AfxGetPrinterScalingFactorX](#afxgetprinterscalingfactorx) | Returns the scaling factor for the x-axis of the printer. |
| [AfxGetPrinterScalingFactorY](#afxgetprinterscalingfactory) | Returns the scaling factor for the y-axis of the printer. |
| [AfxGetPrinterTray](#afxgetprintertray) | Returns the paper source |
| [AfxGetPrinterTrayNames](#afxgetprintertraynames) | Returns a list with the names of the printer's paper bins. |
| [AfxGetPrinterTrueType](#afxgetprintertruetype) | Retrieves the abilities of the driver to use TrueType fonts. |
| [AfxOpenPrintersFolder](#afxopenprinterfolder) | Opens an instance of Explorer with the Printers and Faxes folder selected. |
| [AfxPrinterDialog](#afxprinterdialog) | Displays the printer dialog. Returns TRUE or FALSE. |
| [AfxSetPrinterCollageStatus](#afxsetprintercollagestatus) | Specifies whether collation should be used when printing multiple copies. |
| [AfxSetPrinterColorMode](#afxsetprintercolormode) | Switches between color and monochrome on color printers. |
| [AfxSetPrinterCopies](#afxsetprintercopies) | Selects the number of copies printed if the device supports multiple-page copies. |
| [AfxSetPrinterDuplexMode](#afxsetprinterduplexmode) | Sets the printer duplex mode. |
| [AfxSetPrinterInfo](#afxsetprinterinfo) | Sets data for a specified printer. |
| [AfxSetPrinterOrientation](#afxsetprinterorientation) | Sets the printer orientation. |
| [AfxSetPrinterPaperSize](#afxsetprinterpapersize) | Sets the printer paper size. |
| [AfxSetPrinterQuality](#afxsetprinterquality) | Specifies the print quality mode. |
| [AfxSetPrinterScale](#afxsetprinterscale) | Specifies the factor by which the printed output is to be scaled. |
| [AfxSetPrinterTray](#afxsetprintertray) | Sets the paper source. |

---

## AfxEnumLocalPrinterPorts

Returns a list with port names of the locally installed printers.

```
FUNCTION AfxEnumLocalPrinterPorts () AS DWSTRING
```

#### Return value

The port names. Names are separated with a carriage return and a line feed characters.

---

## AfxEnumPrinterNames

Returns a list with the available printers, print servers, domains, or print providers.

```
FUNCTION AfxEnumPrinterNames () AS DWSTRING
```

#### Return value

The printer names. Names are separated with a carriage return and a line feed characters.

---

## AfxEnumPrinterPorts

Returns a list with the ports that are available for printing on a specified server.

```
FUNCTION AfxEnumPrinterPorts (BYVAL pwszServerName AS WSTRING PTR = NULL) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *pwszServerName* | The name of the printer to attach (as shown in the Devices and Printers applet of the Control Panel). If *pwszServerName* is NULL, the function enumerates the local machine's printer ports. |

#### Return value

A list of port names. Names are separated with a carriage return and a line feed characters.

---

## AfxGetDefaultPrinter

Retrieves the name of the default printer.

```
FUNCTION AfxGetDefaultPrinter () AS DWSTRING
```

#### Return value

The name of the default printer.

---

## AfxGetDefaultPrinterDriver

Retrieves the name of the default printer driver.

```
FUNCTION AfxGetDefaultPrinterDriver () AS DWSTRING
```

#### Return value

The name of the default printer driver.

---

## AfxGetDefaultPrinterPort

Retrieves the name of the default printer port. Use **AfxGetPrinterPort** to retrieve the port name for a named printer.

```
FUNCTION AfxGetDefaultPrinterPort () AS DWSTRING
```

#### Return value

The name of the default printer port.

---

## AfxGetPrinterCollate

If the printer supports collating, the return value is TRUE; otherwise, the return value is FALSE.

```
FUNCTION AfxGetPrinterCollate (BYREF wszPrinterName AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

TRUE or FALSE.

#### Remarks

If TRUE, the pages that are printed should be collated. To collate is to print out the entire document before printing the next copy, as opposed to printing out each page of the document the required number of times.

---

## AfxGetPrinterCollateStatus

Returns the printer collate status.

```
FUNCTION AfxGetPrinterCollateStatus (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The printer collate status. It can be one of the following: DMCOLLATE_FALSE = Collate is turned off; DMCOLLATE_TRUE = Collate is turned on.

---

## AfxGetPrinterColorMode

Returns the printer color mode.

```
FUNCTION AfxGetPrinterColorMode (BYREF wszPrinterName AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

If the printer supports color printing, the return value is TRUE; otherwise, the return value is FALSE.

#### Remarks

Some color printers have the capability to print using true black instead of a combination of cyan, magenta, and yellow (CMY). This usually creates darker and sharper text for documents. This option is only useful for color printers that support true black printing.

---

## AfxGetPrinterCopies

Returns the number of copies to print if the device supports multiple-page copies.

```
FUNCTION AfxGetPrinterCopies (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The number of copies to print.

---

## AfxGetPrinterDC

Returns the handle of the default printer device context.

```
FUNCTION AfxGetPrinterDC () AS HDC
```

#### Return value

The handle of the default printer device context.

---

## AfxGetPrinterDriverVersion

Returns the version number of the Windows-based printer driver.

```
FUNCTION AfxGetPrinterDriverVersion (BYREF wszPrinterName AS WSTRING) AS LONG
```

#### Return value

The version number of the Windows-based printer driver. The version numbers are created and maintained by the driver manufacturer.

---

## AfxGetPrinterDuplex

Retrieves if the printer supports duplex printing.

```
FUNCTION AfxGetPrinterDuplex (BYREF wszPrinterName AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

If the printer supports duplex printing, the return value is TRUE; otherwise, the return value is FALSE.

---

## AfxGetPrinterDuplexMode

If the printer supports duplex printing, returns the current duplex mode

```
FUNCTION AfxGetPrinterDuplex (BYREF wszPrinterName AS WSTRING) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

* DMDUP_SIMPLEX = 1 (Single sided printing)
* DMDUP_VERTICAL = 2 (Page flipped on the vertical edge)
* DMDUP_HORIZONTAL = 3 (Page flipped on the horizontal edge)

---

## AfxGetPrinterFromPort

Returns the printer name for a given port name.

```
FUNCTION AfxGetPrinterFromPort (BYREF wszPortName AS WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPortName* | The port name. |

#### Return value

The printer name.

---

## AfxGetPrinterHorizontalResolution

Returns the width, in pixels, of the printable area of the page.

```
FUNCTION AfxGetPrinterHorizontalResolution (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The width, in pixels, of the printable area of the page.

---

## AfxGetPrinterVerticalResolution

Returns the height, in pixels, of the printable area of the page.

```
FUNCTION AfxGetPrinterVerticalResolution (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The height, in pixels, of the printable area of the page.

---

## AfxGetPrinterMaxCopies

Returns the maximum number of copies the device can print.

```
FUNCTION AfxGetPrinterMaxCopies (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The maximum number of copies the device can print.

---

## AfxGetPrinterMaxPaperHeight

Returns the maximum paper width in tenths of a millimeter.

```
FUNCTION AfxGetPrinterMaxPaperHeight (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The maximum paper height in tenths of a millimeter.

---

## AfxGetPrinterMaxPaperWidth

Returns the maximum paper width in tenths of a millimeter.

```
FUNCTION AfxGetPrinterMaxPaperWidth (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The maximum paper width in tenths of a millimeter.

---

## AfxGetPrinterMediaReady

Retrieves the names of the paper forms that are currently available for use.

```
FUNCTION AfxGetPrinterMediaReady (BYREF wszPrinterName AS WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The names of the paper forms that are currently available for use.

---

## AfxGetPrinterMinPaperHeight

Returns the minimum paper height in tenths of a millimeter.

```
FUNCTION AfxGetPrinterMinPaperHeight (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The minimum paper height in tenths of a millimeter. If the function returns -1, this may mean either that the capability is not supported or there was a general function failure.


---

## AfxGetPrinterMinPaperWidth

Returns the minimum paper width in tenths of a millimeter.

```
FUNCTION AfxGetPrinterMinPaperHeight (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The minimum paper width in tenths of a millimeter. If the function returns -1, this may mean either that the capability is not supported or there was a general function failure.

---

## AfxGetPrinterOrientation

Returns the printer orientation.

```
FUNCTION AfxGetPrinterOrientation (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The printer orientation:

* DMORIENT_PORTRAIT = Portrait
* DMORIENT_LANDSCAPE = Landscape

---

## AfxGetPrinterOrientationDegrees

Returns the relationship between portrait and landscape orientations for a device, in terms of the number of degrees that portrait orientation is rotated counterclockwise to produce landscape orientation.

```
FUNCTION AfxGetPrinterOrientationDegrees (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The printer orientation degrees:

*   0  No landscape orientation.
*  90  Portrait is rotated 90 degrees to produce landscape.
* 180  Portrait is rotated 180 degrees to produce landscape.
* 270  Portrait is rotated 270 degrees to produce landscape.

If the function returns -1, this may mean either that the capability is not supported or there was a general function failure.

---

## AfxGetPrinterPaperLength

Returns the paper length.

```
FUNCTION AfxGetPrinterPaperLength (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The paper length.

---

## AfxGetPrinterPaperWidth

Returns the paper width.

```
FUNCTION AfxGetPrinterPaperWidth (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The paper width.

---

## AfxGetPrinterPaperNames

Returns a list of supported paper names (for example, Letter or Legal).

```
FUNCTION AfxGetPrinterPaperNames (BYREF wszPrinterName AS WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

A list of papernames are separated by a carriage return and a line feed characters.

---

## AfxGetPrinterPaperSizes

Returns a list of supported paper sizes, in tenths of a millimeter.

```
FUNCTION AfxGetPrinterPaperSizes (BYREF wszPrinterName AS WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

A list of papernames. Each entry if formatted as "<width> x <height>" and separated by a carriage return and a line feed characters.

---

## AfxGetPrinterPhysicalHeight

Returns the height of the physical page, in device units.

```
FUNCTION AfxGetPrinterPhysicalHeight (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The height of the physical page, in device units. For example, a printer set to print at 600 dpi on 8.5-by-11-inch paper has a physical height value of 6600 device units. Note that the physical page is almost always greater than the printable area of the page, and never smaller.

---

## AfxGetPrinterPhysicalOffsetX

Returns the distance from the left edge of the physical page to the left edge of the printable area, in device units.

```
FUNCTION AfxGetPrinterPhysicalOffsetX (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The distance from the left edge of the physical page to the left edge of the printable area, in device units. For example, a printer set to print at 600 dpi on 8.5-by-11-inch paper, that cannot print on the leftmost 0.25-inch of paper, has a horizontal physical offset of 150 device units.

---

## AfxGetPrinterPhysicalOffsetY

Returns the distance from the left edge of the physical page to the left edge of the printable area, in device units.

```
FUNCTION AfxGetPrinterPhysicalOffsetY (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The distance from the top edge of the physical page to the top edge of the printable area, in device units. For example, a printer set to print at 600 dpi on 8.5-by-11-inch paper, that cannot print on the topmost 0.5-inch of paper, has a vertical physical offset of 300 device units.

---

## AfxGetPrinterPhysicalWidth

Returns the width of the physical page, in device units.

```
FUNCTION AfxGetPrinterPhysicalWidth (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The width of the physical page, in device units. For example, a printer set to print at 600 dpi on 8.5-x11-inch paper has a physical width value of 5100 device units. Note that the physical page is almost always greater than the printable area of the page, and never smaller.

---

## AfxGetPrinterPort

Returns the port name for a given printer name.

```
FUNCTION AfxGetPrinterPort (BYREF wszPrinterName AS WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The port name.

---

## AfxGetPrinterPPIX

Retrieves the number of pixels per inch of the specified host printer page (horizontal resolution).

```
FUNCTION AfxGetPrinterPPIX (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The number of pixels per inch of the specified host printer page (horizontal resolution).

---

## AfxGetPrinterPPIY

Retrieves the number of pixels per inch of the specified host printer page (vertical resolution).

```
FUNCTION AfxGetPrinterPPIY (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The number of pixels per inch of the specified host printer page (vertical resolution).

---

## AfxGetPrinterQuality

Returns the printer print quality mode.

```
FUNCTION AfxGetPrinterQuality (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The printer print quality mode.

There are four predefined device-independent values:

* DMRES_DRAFT  = Draft
* DMRES_LOW    = Low
* DMRES_MEDIUM = Medium
* DMRES_HIGH   = High

If a positive value is returned, it specifies the number of dots per inch (DPI) and is therefore device dependent.

---

## AfxGetPrinterRate

Returns the printer's print rate.

```
FUNCTION AfxGetPrinterRate (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The printer's print rate.

If the function returns -1, this may mean either that the capability is not supported or there was a general function failure.

---

---

## AfxGetPrinterRatePPM

Returns the printer's print rate, in pages per minute.

```
FUNCTION AfxGetPrinterRatePPM (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The printer's print rate, in pages per minute.

If the function returns -1, this may mean either that the capability is not supported or there was a general function failure.

---

## AfxGetPrinterRateUnit

Returns the printer's print rate.

```
FUNCTION AfxGetPrinterRateUnit (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The printer's print rate. Can be one of the following values:

* PRINTRATEUNIT_CPS   Characters per second.
* PRINTRATEUNIT_IPM   Inches per minute.
* PRINTRATEUNIT_LPM   Lines per minute.
* PRINTRATEUNIT_PPM   Pages per minute.

If the function returns -1, this may mean either that the capability is not supported or there was a general function failure.

---
