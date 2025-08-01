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
| [AfxGetPrinterRatePPM](#afxgetprinterrateppm) | Returns the printer's print rate, in pages per minute. |
| [AfxGetPrinterRateUnit](#afxgetprinterrateunit) | Returns the printer's rate unit. |
| [AfxGetPrinterScale](#afxgetprinterscale) | Returns the factor by which the printed output is to be scaled. |
| [AfxGetPrinterScalingFactorX](#afxgetprinterscalingfactorx) | Returns the scaling factor for the x-axis of the printer. |
| [AfxGetPrinterScalingFactorY](#afxgetprinterscalingfactory) | Returns the scaling factor for the y-axis of the printer. |
| [AfxGetPrinterTray](#afxgetprintertray) | Returns the paper source |
| [AfxGetPrinterTrayNames](#afxgetprintertraynames) | Returns a list with the names of the printer's paper bins. |
| [AfxGetPrinterTrueType](#afxgetprintertruetype) | Retrieves the abilities of the driver to use TrueType fonts. |
| [AfxOpenPrintersFolder](#afxopenprinterfolder) | Opens an instance of Explorer with the Printers and Faxes folder selected. |
| [AfxPrinterDialog](#afxprinterdialog) | Displays the printer dialog. Returns TRUE or FALSE. |
| [AfxSetPrinterCollateStatus](#afxsetprintercollatestatus) | Specifies whether collation should be used when printing multiple copies. |
| [AfxSetPrinterColorMode](#afxsetprintercolormode) | Switches between color and monochrome on color printers. |
| [AfxSetPrinterCopies](#afxsetprintercopies) | Selects the number of copies printed if the device supports multiple-page copies. |
| [AfxSetPrinterDuplexMode](#afxsetprinterduplexmode) | Sets the printer duplex mode. |
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

## AfxGetPrinterScale

Returns the factor by which the printed output is to be scaled.

```
FUNCTION AfxGetPrinterScale (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The factor by which the printed output is to be scaled.

The apparent page size is scaled from the physical page size by a factor of dmScale /100. For example, a letter-sized page with a dmScale value of 50 would contain as much data as a page of 17- by 22-inches because the output text and graphics would be half their original height and width.

---

## AfxGetPrinterScalingFactorX

Returns the scaling factor for the x-axis of the printer.

```
FUNCTION AfxGetPrinterScalingFactorX (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The scaling factor for the x-axis of the printer.

---

## AfxGetPrinterScalingFactorY

Returns the scaling factor for the y-axis of the printer.

```
FUNCTION AfxGetPrinterScalingFactorY (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The scaling factor for the y-axis of the printer.

---

## AfxGetPrinterTray

Returns the paper source.

```
FUNCTION AfxGetPrinterTray (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The paper saurce. Can be one of the following values:

* DMBIN_UPPER = 1
* DMBIN_LOWER = 2
* DMBIN_MIDDLE = 3
* DMBIN_MANUAL = 4
* DMBIN_ENVELOPE = 5
* DMBIN_ENVMANUAL = 6
* DMBIN_AUTO = 7
* DMBIN_TRACTOR = 8
* DMBIN_SMALLFMT = 9
* DMBIN_LARGEFMT = 10
* DMBIN_LARGECAPACITY = 11
* DMBIN_CASSETTE = 14
* DMBIN_FORMSOURCE = 15

---

## AfxGetPrinterTrayNames

Returns a list with the names of the printer's paper bins.

```
FUNCTION AfxGetPrinterTrayNames (BYREF wszPrinterName AS WSTRING) AS DWSTRING
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

Returns a list with the names of the printer's paper bins. The names are separated by a carriage return and a line feed characters.

---

## AfxGetPrinterTrueType

Retrieves the abilities of the driver to use TrueType fonts.

```
FUNCTION AfxGetPrinterTrueType (BYREF wszPrinterName AS WSTRING) AS LONG
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |

#### Return value

The return value can be one or more of the following:

* DCTT_BITMAP   Device can print TrueType fonts as graphics.
* DCTT_DOWNLOAD Device can download TrueType fonts.
* DCTT_SUBDEV   Device can substitute device fonts for TrueType fonts.

If the function returns -1, this may mean either that the capability is not supported or there was a general function failure.

---

## AfxOpenPrintersFolder

Opens an instance of Explorer with the Printers and Faxes folder selected.

```
FUNCTION AfxOpenPrintersFolder () AS HINSTANCE
```

#### Return value

If the function succeeds, it returns a value greater than 32. If the function fails, it returns an error value that indicates the cause of the failure. The return value is cast as an HINSTANCE for backward compatibility with 16-bit Windows applications. It is not a true HINSTANCE, however. It can be cast only to an INT_PTR and compared to either 32 or the following error codes: See [ShellExecuteW](https://learn.microsoft.com/en-us/windows/win32/api/shellapi/nf-shellapi-shellexecutew).

---

## AfxPrinterDialog

Displays the printer dialog. Returns TRUE or FALSE.

```
FUNCTION AfxPrinterDialog (BYVAL hwndOwner AS HWND, BYREF flags AS DWORD, _
   BYREF hdc AS HDC, BYREF nCopies AS WORD, BYREF nFromPage AS WORD, BYREF nToPage AS WORD, _
   BYREF nMinPage AS WORD, BYREF nMaxPage AS WORD) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *hwndOwner* | A handle to the window that owns the dialog box. This member can be any valid window handle, or it can be NULL if the dialog box has no owner. |
| *flags* | Initializes the Print dialog box. When the dialog box returns, it sets these flags to indicate the user's input. See [PRINTDLGW structure](https://learn.microsoft.com/en-us/windows/win32/api/commdlg/ns-commdlg-printdlgw). |
| *hdc* | A handle to a device context or an information context, depending on whether the Flags member specifies the PD_RETURNDC or PC_RETURNIC flag. If neither flag is specified, the value of this member is undefined. If both flags are specified, PD_RETURNDC has priority. |
| *nCopies* | The number of copies. |
| *nFromPage* | The initial page. |
| *nToPage* | The final page. |
| *nMinPage* | The minimum value for the page range specified in the From and To page edit controls. If nMinPage equals nMaxPage, the Pages radio button and the starting and ending page edit controls are disabled. |
| *nMaxPage* | The maximum value for the page range specified in the From and To page edit controls. |

#### Return value

Returns TRUE or FALSE. The caller is resposible of deleting the returned HDC handle with **DeleteDC**.

#### Usage example
```
DIM flags AS DWORD = PD_RETURNDC, hdc AS HDC, nCopies AS WORD
DIM nFromPage AS WORD, nToPage AS WORD, nMinPage AS WORD, nMaxPage AS WORD
IF AfxPrinterDialog(hwnd, flags, hdc, nCopies, nFromPage, nToPage, nMinPage, nMaxPage) THEN
  ' // do printing
  IF hdc THEN DeleteDC hdc
END IF
```
---

## AfxSetPrinterCollateStatus

Specifies whether collation should be used when printing multiple copies.

```
FUNCTION AfxSetPrinterCollateStatus (BYREF wszPrinterName AS WSTRING, BYVAL nMode AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |
| *nMode* | ' The following are the possible values: DMCOLLATE_TRUE, DMCOLLATE_FALSE. |

#### Return value

TRUE or FALSE.

---

## AfxSetPrinterColorMode

Switches between color and monochrome on color printers.

```
FUNCTION AfxSetPrinterColorMode (BYREF wszPrinterName AS WSTRING, BYVAL nMode AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |
| *nMode* | ' The following are the possible values: DMCOLOR_COLOR, DMCOLOR_MONOCHROME. |

#### Return value

TRUE or FALSE.

---

## AfxSetPrinterCopies

Selects the number of copies printed if the device supports multiple-page copies.

```
FUNCTION AfxSetPrinterCopies (BYREF wszPrinterName AS WSTRING, BYVAL nCopies AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |
| *nCopies* | The number of copies. |

#### Return value

TRUE or FALSE.

---

## AfxSetPrinterDuplexMode

Sets the printer duplex mode.

```
FUNCTION AfxSetPrinterDuplexMode (BYREF wszPrinterName AS WSTRING, _
   BYVAL nDuplexMode AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |
| *nDuplexMode* | One of these values:<br>DMDUP_SIMPLEX = Single sided printing.<br>DMDUP_VERTICAL = Page flipped on the vertical edge.<br>DMDUP_HORIZONTAL = Page flipped on the horizontal edge. |

#### Return value

TRUE or FALSE.

---

## AfxSetPrinterOrientation

Sets the printer orientation.

```
FUNCTION AfxSetPrinterOrientation (BYREF wszPrinterName AS WSTRING, BYVAL nOrientation AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |
| *nOrientation* | One of these values:<br>DMORIENT_PORTRAIT = Portrait.<br>DMORIENT_LANDSCAPE = Landscape. |

#### Return value

TRUE or FALSE.

---

## AfxSetPrinterPaperSize

Sets the printer paper size.

```
FUNCTION AfxSetPrinterPaperSize (BYREF wszPrinterName AS WSTRING, BYVAL nSize AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |
| *nSize* | The paper size. Can be one of these values: DMPAPER_LETTER = 1; DMPAPER_FIRST = DMPAPER_LETTER; DMPAPER_LETTERSMALL = 2; DMPAPER_TABLOID = 3; DMPAPER_LEDGER = 4; DMPAPER_LEGAL = 5; DMPAPER_STATEMENT = 6; DMPAPER_EXECUTIVE = 7; DMPAPER_A3 = 8; DMPAPER_A4 = 9; DMPAPER_A4SMALL = 10; DMPAPER_A5 = 11; DMPAPER_B4 = 12; DMPAPER_B5 = 13; DMPAPER_FOLIO = 14; DMPAPER_QUARTO = 15; DMPAPER_10X14 = 16; DMPAPER_11X17 = 17; DMPAPER_NOTE = 18; DMPAPER_ENV_9 = 19; DMPAPER_ENV_10 = 20; DMPAPER_ENV_11 = 21; DMPAPER_ENV_12 = 22; DMPAPER_ENV_14 = 23; DMPAPER_CSHEET = 24; DMPAPER_DSHEET = 25; DMPAPER_ESHEET = 26; DMPAPER_ENV_DL = 27; DMPAPER_ENV_C5 = 28; DMPAPER_ENV_C3 = 29; DMPAPER_ENV_C4 = 30; DMPAPER_ENV_C6 = 31; DMPAPER_ENV_C65 = 32; DMPAPER_ENV_B4 = 33; DMPAPER_ENV_B5 = 34; DMPAPER_ENV_B6 = 35; DMPAPER_ENV_ITALY = 36; DMPAPER_ENV_MONARCH = 37; DMPAPER_ENV_PERSONAL = 38; DMPAPER_FANFOLD_US = 39; DMPAPER_FANFOLD_STD_GERMAN = 40;  DMPAPER_FANFOLD_LGL_GERMAN = 41; DMPAPER_ISO_B4 = 42; DMPAPER_JAPANESE_POSTCARD = 43; DMPAPER_9X11 = 44;  DMPAPER_10X11 = 45; DMPAPER_15X11 = 46; DMPAPER_ENV_INVITE = 47; DMPAPER_RESERVED_48 = 48; DMPAPER_RESERVED_49 = 49; DMPAPER_LETTER_EXTRA = 50; DMPAPER_LEGAL_EXTRA = 51; DMPAPER_TABLOID_EXTRA = 52; DMPAPER_A4_EXTRA = 53; DMPAPER_LETTER_TRANSVERSE = 54; DMPAPER_A4_TRANSVERSE = 55; DMPAPER_LETTER_EXTRA_TRANSVERSE = 56; DMPAPER_A_PLUS = 57; DMPAPER_B_PLUS = 58; DMPAPER_LETTER_PLUS = 59; DMPAPER_A4_PLUS = 60; DMPAPER_A5_TRANSVERSE = 61; DMPAPER_B5_TRANSVERSE = 62; DMPAPER_A3_EXTRA = 63; DMPAPER_A5_EXTRA = 64; DMPAPER_B5_EXTRA = 65; DMPAPER_A2 = 66; DMPAPER_A3_TRANSVERSE = 67; DMPAPER_A3_EXTRA_TRANSVERSE = 68; DMPAPER_DBL_JAPANESE_POSTCARD = 69; DMPAPER_A6 = 70; DMPAPER_JENV_KAKU2 = 71; DMPAPER_JENV_KAKU3 = 72; DMPAPER_JENV_CHOU3 = 73; DMPAPER_JENV_CHOU4 = 74; DMPAPER_LETTER_ROTATED = 75; DMPAPER_A3_ROTATED = 76; DMPAPER_A4_ROTATED = 77; DMPAPER_A5_ROTATED = 78; DMPAPER_B4_JIS_ROTATED = 79; DMPAPER_B5_JIS_ROTATED = 80; DMPAPER_JAPANESE_POSTCARD_ROTATED = 81; DMPAPER_DBL_JAPANESE_POSTCARD_ROTATED = 82; DMPAPER_A6_ROTATED = 83; DMPAPER_JENV_KAKU2_ROTATED = 84; DMPAPER_JENV_KAKU3_ROTATED = 85; DMPAPER_JENV_CHOU3_ROTATED = 86; DMPAPER_JENV_CHOU4_ROTATED = 87; DMPAPER_B6_JIS = 88; DMPAPER_B6_JIS_ROTATED = 89; DMPAPER_12X11 = 90; DMPAPER_JENV_YOU4 = 91; DMPAPER_JENV_YOU4_ROTATED = 92; DMPAPER_P16K = 93; DMPAPER_P32K = 94; DMPAPER_P32KBIG = 95; DMPAPER_PENV_1 = 96; DMPAPER_PENV_2 = 97; DMPAPER_PENV_3 = 98; DMPAPER_PENV_4 = 99; DMPAPER_PENV_5 = 100; DMPAPER_PENV_6 = 101; DMPAPER_PENV_7 = 102; DMPAPER_PENV_8 = 103; DMPAPER_PENV_9 = 104; DMPAPER_PENV_10 = 105; DMPAPER_P16K_ROTATED = 106; DMPAPER_P32K_ROTATED = 107; DMPAPER_P32KBIG_ROTATED = 108; DMPAPER_PENV_1_ROTATED = 109; DMPAPER_PENV_2_ROTATED = 110; DMPAPER_PENV_3_ROTATED = 111; DMPAPER_PENV_4_ROTATED = 112; DMPAPER_PENV_5_ROTATED = 113; DMPAPER_PENV_6_ROTATED = 114; DMPAPER_PENV_7_ROTATED = 115; DMPAPER_PENV_8_ROTATED = 116; DMPAPER_PENV_9_ROTATED = 117; DMPAPER_PENV_10_ROTATED = 118; DMPAPER_LAST = DMPAPER_PENV_10_ROTATED; DMPAPER_USER = 256 |

#### Return value

TRUE or FALSE.

---

## AfxSetPrinterQuality

Specifies the print quality mode.

```
FUNCTION AfxSetPrinterQuality (BYREF wszPrinterName AS WSTRING, BYVAL nMode AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |
| *nMode* | The printer qauality. |

There are four predefined device-independent values:

* DMRES_DRAFT  = Draft
* DMRES_LOW    = Low
* DMRES_MEDIUM = Medium
* DMRES_HIGH   = High

#### Return value

TRUE or FALSE.

---

## AfxSetPrinterScale

Specifies the factor by which the printed output is to be scaled.

```
FUNCTION AfxSetPrinterScale (BYREF wszPrinterName AS WSTRING, BYVAL dmScale AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |
| *dmScale* | The scale value. |

#### Return value

TRUE or FALSE.

#### Remarks

The apparent page size is scaled from the physical page size by a factor of *dmScale /100*. For example, a letter-sized page with a *dmScale* value of 50 would contain as much data as a page of 17- by 22-inches because the output text and graphics would be half their original height and width.

---

## AfxSetPrinterTray

Sets the paper source. Can be one of the following values, or it can be a device-specific value greater than or equal to DMBIN_USER.

```
FUNCTION AfxSetPrinterTray (BYREF wszPrinterName AS WSTRING, BYVAL nTray AS LONG) AS BOOLEAN
```

| Parameter  | Description |
| ---------- | ----------- |
| *wszPrinterName* | The printer name. |
| *nTray* | The paper source. |

The paper source can be one of the following values, or it can be a device-specific value greater than or equal to DMBIN_USER.

* DMBIN_UPPER = 1
* DMBIN_LOWER = 2
* DMBIN_MIDDLE = 3
* DMBIN_MANUAL = 4
* DMBIN_ENVELOPE = 5
* DMBIN_ENVMANUAL = 6
* DMBIN_AUTO = 7
* DMBIN_TRACTOR = 8
* DMBIN_SMALLFMT = 9
* DMBIN_LARGEFMT = 10
* DMBIN_LARGECAPACITY = 11
* DMBIN_CASSETTE = 14
* DMBIN_FORMSOURCE = 15

#### Return value

TRUE or FALSE.

---
