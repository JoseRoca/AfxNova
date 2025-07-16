' ########################################################################################
' Platform: Microsoft Windows
' Filename: AfxSysInfo.inc
' Contents: System information wrapper functions
' Compiler: FreeBASIC 32 & 64 bit
' Copyright (c) 2025 José Roca
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
#INCLUDE ONCE "windows.bi"
#INCLUDE ONCE "win/uxtheme.bi"
#INCLUDE ONCE "AfxNova/AfxWin.inc"
#INCLUDE ONCE "AfxNova/DWSTRING.inc"
USING AfxNova

NAMESPACE AfxNova

' ========================================================================================
' The raw SMBIOS table provider ('RSMB') retrieves the contents of the raw SMBIOS firmware
' table. The pFirmwareTableBuffer parameter of the GetSystemFirmwareTable function 
' receives the following data:
' ========================================================================================

TYPE RawSMBIOSData
   Used20CallingMethod AS BYTE
   SMBIOSMajorVersion AS BYTE
   SMBIOSMinorVersion AS BYTE
   DmiRevision AS BYTE
   Length As DWORD
   SMBIOSTableData(0) AS BYTE
END TYPE

' ========================================================================================
' Retrieves the Bios serial number.
' ========================================================================================
FUNCTION AfxGetBiosSerialNumber() AS DWSTRING
   DIM size AS DWORD = GetSystemFirmwareTable(&H52534D42, 0, 0, 0) ' 'RSMB'
   IF size = 0 THEN RETURN ""
   DIM buffer AS ANY PTR = Allocate(size)
   IF buffer = 0 THEN RETURN ""
   DIM result AS DWORD = GetSystemFirmwareTable(&H52534D42, 0, buffer, size)
   IF result <> size THEN
      Deallocate(buffer)
      RETURN ""
   END IF
   DIM pRaw AS RawSMBIOSData PTR = buffer
   DIM pData AS UBYTE PTR = @pRaw->SMBIOSTableData(0)
   Dim i AS LONG
   WHILE i < pRaw->Length
      DIM structType AS BYTE = pData[i]
      DIM structLength AS BYTE = pData[i + 1]
      IF structType = 1 THEN ' Type 1 = System Information
         DIM serialIndex As Byte = pData[i + 7]
         DIM pStrings AS UByte Ptr = pData + i + structLength
         DIM strCount AS LONG = 1
         DIM serial AS DWSTRING = ""
         WHILE *pStrings <> 0
            DIM dws AS DWSTRING = ""
            WHILE *pStrings <> 0
               dws += WCHR(*pStrings)
               pStrings += 1
            WEND
            pStrings += 1
            IF strCount = serialIndex THEN
               serial = dws
               EXIT WHILE
            END IF
            strCount += 1
         WEND
         Deallocate(buffer)
         RETURN serial
      END IF
      ' Move to next structure
      DIM pNext AS UBYTE PTR = pData + i + structLength
      WHILE *(pNext) <> 0 OR *(pNext + 1) <> 0
         pNext += 1
      WEND
      pNext += 2
      i = pNext - pData
   WEND
   Deallocate(buffer)
   RETURN ""
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieves the baseboard (motherboard/system board) serial number using Windows API.
' ========================================================================================
PRIVATE FUNCTION AfxGetBaseBoardSerialNumber () AS DWSTRING
   DIM size AS DWORD = GetSystemFirmwareTable(&H52534D42, 0, 0, 0) ' RSMB
   IF size = 0 THEN RETURN ""
   DIM buffer AS ANY PTR = Allocate(size)
   IF buffer = 0 THEN RETURN ""
   DIM result AS DWORD = GetSystemFirmwareTable(&H52534D42, 0, buffer, size)
   IF result <> size THEN
      Deallocate(buffer)
      RETURN ""
   END IF
   DIM pRaw AS RawSMBIOSData PTR = buffer
   DIM pData AS UBYTE PTR = @pRaw->SMBIOSTableData(0)
   DIM i AS LONG
   WHILE i < pRaw->Length
      DIM structType AS BYTE = pData[i]
      DIM structLength AS BYTE = pData[i + 1]
      IF structType = 2 THEN ' Baseboard Information (Type 2)
         DIM serialIndex AS BYTE = pData[i + 7]
         DIM pStrings AS UBYTE PTR = pData + i + structLength
         DIM strCount AS LONG = 1
         DIM serial AS DWSTRING = ""
         WHILE *pStrings <> 0
            DIM dws AS DWSTRING = ""
            WHILE *pStrings <> 0
               dws += WCHR(*pStrings)
               pStrings += 1
            WEND
            pStrings += 1
            IF strCount = serialIndex THEN
               serial = dws
               EXIT WHILE
            END IF
            strCount += 1
         WEND
         Deallocate(buffer)
         RETURN serial
      END IF
      ' Move to next structure
      DIM pNext AS UBYTE PTR = pData + i + structLength
      WHILE *(pNext) <> 0 OR *(pNext + 1) <> 0
         pNext += 1
      WEND
      pNext += 2
      i = pNext - pData
   WEND
   Deallocate(buffer)
   RETURN ""
END FUNCTION
' ========================================================================================

END NAMESPACE

' Is there more useful information to be extracted using GetSystemFirmwareTable?
' Absolutely—and you’ve only scratched the surface by parsing Type 1 (System Info) and
' Type 2 (Baseboard Info) from the SMBIOS table. The real power of GetSystemFirmwareTable
' with provider 'RSMB' is that it exposes the entire SMBIOS structure, which contains
' dozens of types with rich hardware metadata.
' Here’s a curated list of notable SMBIOS structure types you can extract for further insight:
' Hardware Identity & Configuration
' Type	Description	Notable Fields
' 0	BIOS Information	Version, Vendor, Release Date
' 1	System Information	Product Name, Serial, UUID
' 2	Baseboard Information	Manufacturer, Product, Serial
' 3	System Enclosure (Chassis)	Type (Laptop, Tower), Serial, Lock
' 4	Processor Information	Socket, Type, Manufacturer, Core Count
' 17	Memory Device	Size, Speed, Type, Serial
' Capabilities & Features
' Type	Description	Notable Fields
' 5	Memory Controller	Error Correction, Supported Speeds
' 7	Cache Information	Level, Size, Associativity
' 32	System Boot Information	Boot Status
' 15	System Event Log	Logging methods, format
' Device Mapping
' Type	Description	Usefulness
' 8	Port Connector Info	I/O ports on board
' 9	System Slots	PCI, AGP, M.2 slot enumeration
' 10	On Board Devices Info (Old)	Legacy mapping (deprecated but useful)
' 41	On Board Devices Extended Info	Modern mapping, device type flags
' Security & Power
' Type	Description	Notable Fields
' 24	Hardware Security	Reset Switch, Power Locks
' 27	Cooling Device	Type (Fan, Liquid), Speed, Status
