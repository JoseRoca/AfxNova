# Printer Functions



**Include file**: AfxNova/AfxPrinter.inc

### Methods

| Name       | Description |
| ---------- | ----------- |
| [AfxEnumLocalPrinterPorts](#afxenumlocalprinterports) | Returns a list with port names of the locally installed printers. |
| [AfxEnumPrinterNames](#afxenumpriinternames) | Returns a list with the available printers, print servers, domains, or print providers. |
| [AfxEnumPrinterPorts](#afxenumprinterports) | Returns a list with the ports that available for printing on a specified server. |
| [AfxGetDefaultPrinter](#afxgetdefaultprinter) | Retrieves the name of the default printer. |
| [AfxGetDefaultPrinterDriver](#afxgetdefaultprinterdriver) | Retrieves the name of the default printer driver. |
| [AfxGetDefaultPrinterPort](#afxgetdefaultprinterport) | Retrieves the name of the default printer port. |
| [AfxGetDocumentProperties](#afxgetdocumentproperties) | Retrieves printer initialization information. |
| [AfxGetPrinterCollate](#afxgetprintercollate) | If the printer supports collating, the return value is TRUE; otherwise, the return value is FALSE. |
| [AfxGetPrinterCollateStatus](#afxgetprintercollatestatus) | Returns the printer collate status. |
| [AfxGetPrinterColorMode](#afxgetprintercolormode) | If the printer supports color printing, the return value is TRUE; otherwise, the return value is FALSE. |
| [AfxGetPrinterCopies](#afxgetprintercopies) | Returns the number of copies printed if the device supports multiple-page copies. |
| [AfxGetPrinterDC](#afxgetprinterdc) | Returns the handle of the default printer device context. |
| [AfxGetPrinterDriverVersion](#afxgetprinterdriverversion) | Returns the version number of the Windows-based printer driver. |
| [AfxGetPrinterDuplex](#afxgetprinterduplex) | If the printer supports duplex printing, the return value is TRUE; otherwise, the return value is FALSE. |
| [AfxGetPrinterDuplexMode](#afxgetprinterduplexname) | If the printer supports duplex printing, returns the current duplex mode. |
| [AfxGetPrinterFromPort](#afxgetprinterfromport) | Returns the printer name for a given port name. |
| [AfxGetPrinterFromPort](#afxgetprinterfromport) | Returns the printer name for a given port name. |
| [AfxGetPrinterHorizontalResolution](#afxgetprinterhorizontalresolution) | Returns the width, in pixels, of the printable area of the page. |
| [AfxGetPrinterVorizontalResolution](#afxgetprintervorizontalresolution) | Returns the Height, in pixels, of the printable area of the page. |
| [AfxGetPrinterMaxCopies](#afxgetprintermaxcopies) | Returns the maximum number of copies the device can print. |
| [AfxGetPrinterMaxPaperHeight](#afxgetprintermaxpaperheight) | Returns the maximum paper width in tenths of a millimeter. |
| [AfxGetPrinterMaxPaperWidth](#afxgetprintermaxpaperwidth) | Returns the maximum paper width in tenths of a millimeter. |
| [AfxGetPrinterMediaReady](#afxgetprintermediaready) | Retrieves the names of the paper forms that are currently available for use. |
| [AfxGetPrinterMinPaperHeight](#afxgetprinterminpaperheight) | Returns the minimum paper width in tenths of a millimeter. |
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
