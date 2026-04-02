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
#INCLUDE ONCE "win/winioctl.bi"
#INCLUDE ONCE "AfxNova/AfxWin.inc"
#INCLUDE ONCE "AfxNova/DWSTRING.inc"
USING AfxNova

NAMESPACE AfxNova

' ========================================================================================
' // Converts four characters into a little endian dword
' // Usage example:
' // DIM rsmb AS DWORD = MAKEFOURCCLE("RSMB")
' // Print HEX(rsmb, 8)  ' Output: &h52534D42
' ========================================================================================
#ifndef MAKEFOURCCLE
PRIVATE FUNCTION MAKEFOURCCLE (BYREF s AS STRING) AS DWORD
   IF LEN(s) <> 4 THEN RETURN 0   ' // input string must be exactly 4 characters
   ' Combine bytes in little-endian order
   RETURN ASC(MID(s, 4, 1)) OR _
          ASC(MID(s, 3, 1)) SHL 8 OR _
          ASC(MID(s, 2, 1)) SHL 16 OR _
          ASC(MID(s, 1, 1)) SHL 24
END FUNCTION
#endif
' ========================================================================================

' ========================================================================================
' // Converts four characters into a big endian dword
' // DIM rsmb AS DWORD = MAKEFOURCCBE("RSMB")
' // Print HEX(rsmb, 8)  ' Output: &h424D5352
' ========================================================================================
#ifndef MAKEFOURCCBE
PRIVATE FUNCTION MAKEFOURCCBE (BYREF s AS STRING) AS DWORD
   IF LEN(s) <> 4 THEN RETURN 0   ' // input string must be exactly 4 characters
   ' Combine bytes in big-endian order
   RETURN ASC(MID(s, 4, 1)) SHL 24 OR _
          ASC(MID(s, 3, 1)) SHL 16 OR _
          ASC(MID(s, 2, 1)) SHL 8 OR _
          ASC(MID(s, 1, 1))
END FUNCTION
#endif
' ========================================================================================

' ========================================================================================
' The raw SMBIOS table provider ('RSMB') retrieves the contents of the raw SMBIOS firmware
' table. The pFirmwareTableBuffer parameter of the GetSystemFirmwareTable function 
' receives the following data:
' ========================================================================================
#ifndef RawSMBIOSData
TYPE RawSMBIOSData
   Used20CallingMethod AS BYTE
   SMBIOSMajorVersion AS BYTE
   SMBIOSMinorVersion AS BYTE
   DmiRevision AS BYTE
   Length As DWORD
   SMBIOSTableData(0) AS BYTE
END TYPE
#endif

#ifndef RawSMBIOSTable
TYPE RawSMBIOSTable
   ' // Spec version 2.0+
   AS BYTE Type       ' // Offset 00h - Platform Firmware Information indicator
   AS BYTE Length     ' // Offset 01h - 12h + number of Firmware Characteristics Extension Bytes
   AS WORD Handle     ' // Offset 02h - Varies
END TYPE
#endif

' Type 0 - Firmware information
#ifndef RawSMBIOSBiosInfo
TYPE RawSMBIOSBiosInfo
   AS RawSMBIOSTable Header
   ' // Spec 2.0+
   AS BYTE Vendor               ' // Offset 04h - String number of the Firmware Vendor's Name
   AS BYTE Version              ' // Offset 05h - String number of the Firmware Version
   AS WORD StartingAddrSeg      ' // Offset 06h - Bios Starting Address Segment
   AS BYTE ReleaseDate          ' // Offset 08h - Firmware Release Date
   AS BYTE ROMSize              ' // Offset 09h - Firmware ROM Size
   AS ULONG64 Characteristics   ' // Offset 0Ah - Firmware Characteistics
   ' // Spec 2.4+
   AS BYTE Extension(1)         ' // Offset 12h - Firmware Characteristics Extension Bytes
   AS BYTE MajorRelease         ' // Offset 14h - Platform Firmare Major Release
   AS BYTE MinorRelease         ' // Offset 15h - Platform Firmare Minor Release
   AS BYTE ECFirmwareMajor      ' // Offset 16h - Embedded Controller Firmare Major Release
   AS BYTE ECFirmwareMinor      ' // Offset 17h - Embedded Controller Firmare Major Release
   ' // Spec 3.1+
   AS WORD EFROMSize            ' // Offset 18h - Extended Firmware ROM Size
END TYPE
#endif

' Type 1 - System information
#ifndef RawSMBIOSSystemInfo
TYPE RawSMBIOSSystemInfo
   ' // Spec 2.0+
   AS RawSMBIOSTable Header
   AS BYTE Manufacturer   ' // Offset 04h - Manufacturer
   AS BYTE ProductName    ' // Offset 05h - Product name
   AS BYTE Version        ' // Offset 06h - Version
   AS BYTE SerialNumber   ' // Offset 07h - Serial number
   ' // Spec 2.1+
   AS BYTE UUID(15)       ' // Offset 08h - 16 bytes - Universal unique ID number
   AS BYTE Wakeup_Type    ' // Offset 18h - Enum
   ' // Spec 2.4+
   AS BYTE SKUNumber      ' // Offset 19h - SKU Number
   AS BYTE Family         ' // Offset 1Ah - Family
END TYPE
#endif

' ########################################################################################
' CBiosTable class
' Loads and owns the SMBIOS memory block.
' ########################################################################################

' ========================================================================================
TYPE CBiosTable

   Private:
      m_pTable AS RawSMBIOSTable PTR   ' // Pointer to the raw firmware bios table
      m_pRaw AS RawSMBIOSData PTR      ' // Pointer to the persistent firmware bios data

   Public:
      DECLARE CONSTRUCTOR (BYVAL wantedType AS BYTE)
      DECLARE DESTRUCTOR ()
      DECLARE FUNCTION GetTable() AS RawSMBIOSTable PTR
      DECLARE FUNCTION FindTable (BYVAL wantedType AS BYTE) AS BOOLEAN
      DECLARE FUNCTION GetString (BYVAL index AS BYTE) AS DWSTRING

END TYPE
' ========================================================================================

' ========================================================================================
' Creates an instance of the CBiosTable class and retrieves the specified firmware table.
' ========================================================================================
PRIVATE CONSTRUCTOR CBiosTable (BYVAL wantedType AS BYTE)
   ' // Retrieve the system firmware table
   DIM size AS DWORD = GetSystemFirmwareTable(MAKEFOURCCLE("RSMB"), 0, 0, 0)
   IF size = 0 THEN EXIT CONSTRUCTOR
   ' // Allocate a buffer for the firmware table
   DIM pRoot AS ANY PTR = Allocate(size)
   IF pRoot = NULL THEN EXIT CONSTRUCTOR
   ' // Retrieves the specified firmware table from the firmware table provider
   IF GetSystemFirmwareTable(MAKEFOURCCLE("RSMB"), 0, pRoot, size) <> size THEN
      Deallocate(pRoot)
      EXIT CONSTRUCTOR
   END IF
   ' // Store the pointer to the persistent firmware bios data
   m_pRaw = CAST(RawSMBIOSData PTR, pRoot)
   ' // Find the address of the SMBIOS Table
   IF NOT this.FindTable(wantedType) THEN
      Deallocate(pRoot)
      EXIT CONSTRUCTOR
   END IF
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' Deallocates the persistent firmware table data.
' ========================================================================================
PRIVATE DESTRUCTOR CBiosTable
    IF m_pRaw <> NULL THEN Deallocate(m_pRaw)
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns a pointer to the raw SMBIOS table.
' ========================================================================================
PRIVATE FUNCTION CBiosTable.GetTable() AS RawSMBIOSTable PTR
    RETURN m_pTable
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieve the RawSMBIOSTable pointer
' ========================================================================================
PRIVATE FUNCTION CBiosTable.FindTable (BYVAL wantedType AS BYTE) AS BOOLEAN
   IF m_pRaw = NULL THEN RETURN FALSE
   DIM pData AS BYTE PTR = @m_pRaw->SMBIOSTableData(0)
   DIM i AS LONG
   WHILE i < m_pRaw->Length
      DIM structType AS BYTE = pData[i]
      DIM structLength AS BYTE = pData[i + 1]
      IF structLength = 0 THEN EXIT WHILE  ' Corrupted data
      ' It is the type we are looking for?
      IF structType = wantedType THEN
         m_pTable = CAST(RawSMBIOSTable PTR, pData + i)
         RETURN TRUE
      END IF
      ' Advance to the strings section
      DIM pNext AS BYTE PTR = pData + i + structLength
      ' Skip strings until a double null is found
      WHILE (pNext[0] <> 0) OR (pNext[1] <> 0)
         pNext += 1
      WEND
      pNext += 2  ' Skip the double null
      i = pNext - pData
   WEND
   RETURN FALSE   ' Not found
END FUNCTION
' ========================================================================================

' ========================================================================================
' Retrieve the SMBIOS string
' ========================================================================================
PRIVATE FUNCTION CBiosTable.GetString (BYVAL index AS BYTE) AS DWSTRING
   IF m_pTable = NULL THEN RETURN ""
   IF index = 0 THEN RETURN ""
   DIM pData AS BYTE PTR = cast(BYTE PTR, m_pTable)
   DIM structLength AS BYTE = m_pTable->Length
   ' First string immediately after the formatted part
   DIM pStr AS BYTE PTR = pData + structLength
   DIM current AS LONG = 1
   DIM result AS DWSTRING
   WHILE TRUE
      ' End of all strings? (double zero)
      IF pStr[0] = 0 AND pStr[1] = 0 THEN EXIT WHILE
      IF current = index THEN
         ' Extract string
         WHILE *pStr <> 0
            result += WCHR(*pStr)
            pStr += 1
         WEND
         RETURN result
      END IF
      ' Skip current string
      WHILE *pStr <> 0
         pStr += 1
      WEND
      ' Skip the terminating zero
      pStr += 1
      current += 1
   WEND
   RETURN ""
END FUNCTION
' ========================================================================================


' ########################################################################################

' ========================================================================================
' Generic SMBIOS string accessor: type + offset
' ========================================================================================
PRIVATE FUNCTION AfxGetBiosInfo (BYVAL wantedType AS BYTE, BYVAL offset AS LONG) AS DWSTRING
   DIM pBios AS CBiosTable = wantedType
   DIM pTable AS RawSMBIOSTable PTR = pBios.GetTable
   IF pTable = NULL THEN RETURN ""
   DIM idx AS BYTE = CAST(BYTE PTR, pTable)[offset]
   IF idx = 0 THEN RETURN ""
   RETURN pBios.GetString(idx)
END FUNCTION
' ========================================================================================

' // Macros:

' --- Type 0: BIOS Information ---
#define SMBIOS_BIOSVendor          AfxGetBiosInfo(0, &h04)
#define SMBIOS_BIOSVersion         AfxGetBiosInfo(0, &h05)
#define SMBIOS_BIOSReleaseDate     AfxGetBiosInfo(0, &h08)

' --- Type 1: System Information ---
#define SMBIOS_SystemManufacturer  AfxGetBiosInfo(1, &h04)
#define SMBIOS_SystemProductName   AfxGetBiosInfo(1, &h05)
#define SMBIOS_SystemVersion       AfxGetBiosInfo(1, &h06)
#define SMBIOS_SystemSerialNumber  AfxGetBiosInfo(1, &h07)
#define SMBIOS_SystemSKUNumber     AfxGetBiosInfo(1, &h19)
#define SMBIOS_SystemFamily        AfxGetBiosInfo(1, &h1A)

PRIVATE FUNCTION BIOS_SystemUUID () AS STRING
   DIM pBios AS CBiosTable = 1
   DIM pTable AS RawSMBIOSTable PTR = pBios.GetTable
   IF pTable = NULL THEN RETURN ""
   DIM p AS BYTE PTR = CAST(BYTE PTR, pTable) + &h08
   ' SMBIOS UUID is 16 bytes
   DIM uuid AS STRING
   FOR i AS LONG = 0 TO 15
      uuid &= HEX(p[i], 2)
      IF i = 3 OR i = 5 OR i = 7 OR i = 9 THEN uuid &= "-"
   NEXT
   RETURN uuid
END FUNCTION

' --- Type 2: Baseboard ---
#define SMBIOS_BaseBoardManufacturer  AfxGetBiosInfo(2, &h04)
#define SMBIOS_BaseBoardProduct       AfxGetBiosInfo(2, &h05)
#define SMBIOS_BaseBoardVersion       AfxGetBiosInfo(2, &h06)
#define SMBIOS_BaseBoardSerialNumber  AfxGetBiosInfo(2, &h07)
#define SMBIOS_BaseBoardAssetTag      AfxGetBiosInfo(2, &h08)

' --- Type 3: Chassis ---
#define SMBIOS_ChassisManufacturer    AfxGetBiosInfo(3, &h04)
#define SMBIOS_ChassisVersion         AfxGetBiosInfo(3, &h06)
#define SMBIOS_ChassisSerialNumber    AfxGetBiosInfo(3, &h07)
#define SMBIOS_ChassisAssetTag        AfxGetBiosInfo(3, &h08)

' --- Type 4: Processor ---
#define SMBIOS_ProcessorSocket        AfxGetBiosInfo(4, &h04)
#define SMBIOS_ProcessorManufacturer  AfxGetBiosInfo(4, &h07)
#define SMBIOS_ProcessorVersion       AfxGetBiosInfo(4, &h10)
#define SMBIOS_ProcessorSerialNumber  AfxGetBiosInfo(4, &h20)
#define SMBIOS_ProcessorAssetTag      AfxGetBiosInfo(4, &h21)
#define SMBIOS_ProcessorPartNumber    AfxGetBiosInfo(4, &h22)

' --- Type 17: Memory Device ---
#define SMBIOS_MemoryDeviceLocator       AfxGetBiosInfo(17, &h10)
#define SMBIOS_MemoryDeviceBankLocator   AfxGetBiosInfo(17, &h11)
#define SMBIOS_MemoryDeviceManufacturer  AfxGetBiosInfo(17, &h17)
#define SMBIOS_MemoryDeviceSerialNumber  AfxGetBiosInfo(17, &h18)
#define SMBIOS_MemoryDeviceAssetTag      AfxGetBiosInfo(17, &h19)
#define SMBIOS_MemoryDevicePartNumber    AfxGetBiosInfo(17, &h1A)

' ========================================================================================
' Retrieves the size of RAM in the first module (Type 17)
' The Size field is at offset 0Ch, but it has special rules:
' If it's FFFFh, use Extended Size (offset 1Ch).
' If it's 0, it's not installed.
' If it's 7FFFh, the size is unknown.
' ========================================================================================
PRIVATE FUNCTION SMBIOS_MemoryDeviceSize () AS ULONG
   DIM pBios AS CBiosTable = 17
   DIM pTable AS RawSMBIOSTable PTR = pBios.GetTable
   IF pTable = NULL THEN RETURN 0
   DIM p AS BYTE PTR = CAST(BYTE PTR, pTable)
   DIM size AS ULONG = CULNG(p[&h0C]) OR (CULNG(p[&h0D]) SHL 8)
   IF size = &hFFFF THEN
      ' Extended size (in MB)
      DIM ext AS ULONG = _
         p[&h1C] OR _
         (p[&h1D] SHL 8) OR _
         (p[&h1E] SHL 16) OR _
         (p[&h1F] SHL 24)
      RETURN ext
   END IF
   IF size = 0 OR size = &h7FFF THEN RETURN 0
   ' Size is in MB
   RETURN size
END FUNCTION
' ========================================================================================

' ========================================================================================
' Additional detail of the memory device type.
' ========================================================================================
PRIVATE FUNCTION SMBIOS_MemoryDeviceTypeDetail () AS STRING
   DIM pBios AS CBiosTable = 17
   DIM pTable AS RawSMBIOSTable PTR = pBios.GetTable
   IF pTable = NULL THEN RETURN ""
   DIM p AS BYTE PTR = CAST(BYTE PTR, pTable)
   DIM typeDetail AS USHORT = p[&h13] OR (p[&h14] SHL 8)
   IF (typeDetail AND &h0080) THEN RETURN "DDR"
   IF (typeDetail AND &h0200) THEN RETURN "DDR2"
   IF (typeDetail AND &h0400) THEN RETURN "DDR2 FB-DIMM"
   IF (typeDetail AND &h0800) THEN RETURN "DDR3"
   IF (typeDetail AND &h1000) THEN RETURN "FBD2"
   IF (typeDetail AND &h2000) THEN RETURN "DDR4"
   IF (typeDetail AND &h4000) THEN RETURN "LPDDR4"
   IF (typeDetail AND &h8000) THEN RETURN "DDR5"
   RETURN "Unknown"
END FUNCTION
' ========================================================================================

' ========================================================================================
' Identifies the maximum capable speed of the device, in megatransfers per second (MT/s). 
' ========================================================================================
PRIVATE FUNCTION SMBIOS_MemoryDeviceSpeed () AS USHORT
   DIM pBios AS CBiosTable = 17
   DIM pTable AS RawSMBIOSTable PTR = pBios.GetTable
   IF pTable = NULL THEN RETURN 0
   DIM p AS UBYTE PTR = CAST(UBYTE PTR, pTable)
   RETURN p[&h15] OR (p[&h16] SHL 8)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Number of enabled cores per processor socket. If the value is unknown, the field is set to 0.
' For core counts of 256 or greater, the Core Enabled field is set to FFh and the Core Enabled 2
' field is set to the number of enabled cores. 
' ========================================================================================
PRIVATE FUNCTION SMBIOS_CPUCoreCount () AS LONG
   DIM pBios AS CBiosTable = 4
   DIM p AS BYTE PTR = CAST(BYTE PTR, pBios.GetTable)
   IF p = NULL THEN RETURN 0
   RETURN p[&h23]
END FUNCTION

' ========================================================================================
' Number of threads per processor socket. If the value is unknown, the field is set to 0.
' For thread counts of 256 or greater, the Thread Count field is set to FFh and the Thread 
' Count 2 field is set to the number of threads. 
' ========================================================================================
PRIVATE FUNCTION SMBIOS_CPUThreadCount () AS LONG
   DIM pBios AS CBiosTable = 4
   DIM p AS BYTE PTR = CAST(BYTE PTR, pBios.GetTable)
   IF p = NULL THEN RETURN 0
   RETURN p[&h25]
END FUNCTION
' ========================================================================================

END NAMESPACE
