' ########################################################################################
' Microsoft Windows
' File: CDOMessage.inc
' Contents: CDO class to send emails.
' Compiler: FreeBasic 32 & 64-bit
' Copyright (c) 2025 José Roca. Freeware. Use at your own risk.
' THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
' EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
' MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
' ########################################################################################

#pragma once
#include once "windows.bi"
#include once "AfxNova/DWSTRING.inc"
#include once "AfxNova/AfxCDOSys.bi"
#include once "AfxNova/AfxCOM.inc"
USING AfxNova

NAMESPACE AfxNova

' ========================================================================================
' CDOMessage class
' ========================================================================================
TYPE CDOMessage

   Public:
      m_Result AS HRESULT                   ' // HRESULT
      m_pConfig AS Afx_IConfiguration PTR   ' // Pointer to the Afx_IConfiguration interface
      m_pMessage AS Afx_IMessage PTR        ' // Pointer to the Afx_IMessage interface
      m_pFields AS Afx_AdoFields PTR        ' // Pointer to the Afx_AdoFields interface

   Private:
      m_bUninitCOM AS BOOLEAN

   Public:
      DECLARE CONSTRUCTOR
      DECLARE DESTRUCTOR
      DECLARE FUNCTION GetLastResult () AS HRESULT
      DECLARE FUNCTION SetResult (BYVAL Result AS HRESULT) AS HRESULT
      DECLARE FUNCTION GetErrorInfo () AS DWSTRING
      DECLARE FUNCTION ConfigValue (BYREF wszField AS WSTRING, BYREF wszValue AS WSTRING) AS HRESULT
      DECLARE FUNCTION ConfigValue (BYREF wszField AS WSTRING, BYVAL value AS LONG) AS HRESULT
      DECLARE FUNCTION ConfigUpdate () AS HRESULT
      DECLARE FUNCTION BCC (BYREF wszBCC AS WSTRING) AS HRESULT
      DECLARE FUNCTION CC (BYREF wszCC AS WSTRING) AS HRESULT
      DECLARE FUNCTION From (BYREF wszFrom AS WSTRING) AS HRESULT
      DECLARE FUNCTION Sender (BYREF wszSender AS WSTRING) AS HRESULT
      DECLARE FUNCTION Recipients (BYREF wszTo AS WSTRING) AS HRESULT
      DECLARE FUNCTION ReplyTo (BYREF wszReplyTo AS WSTRING) AS HRESULT
      DECLARE FUNCTION Keywords (BYREF wszKeywords AS WSTRING) AS HRESULT
      DECLARE FUNCTION Organization (BYREF wszOrganization AS WSTRING) AS HRESULT
      DECLARE FUNCTION Subject (BYREF wszSubject AS WSTRING) AS HRESULT
      DECLARE FUNCTION TextBody (BYREF wszTextBody AS WSTRING) AS HRESULT
      DECLARE FUNCTION HTMLBody (BYREF wszTextBody AS WSTRING) AS HRESULT
      DECLARE FUNCTION DSNOptions (BYVAL pDSNOptions AS CdoDSNOptions) AS HRESULT
      DECLARE FUNCTION MimeFormatted (BYVAL pMimeFormatted AS BOOLEAN) AS HRESULT
      DECLARE FUNCTION MDNRequested (BYVAL pMDNRequested AS BOOLEAN) AS HRESULT
      DECLARE FUNCTION AddAttachment (BYREF wszURL AS WSTRING, BYREF wszUserName AS WSTRING = "", BYREF wszPassword AS WSTRING = "") AS HRESULT
      DECLARE FUNCTION Send () AS HRESULT
      DECLARE FUNCTION FollowUpTo (BYREF wszFollowUpTo AS WSTRING) AS HRESULT
      DECLARE FUNCTION Newsgroups (BYREF wszNewsgroups AS WSTRING) AS HRESULT
      DECLARE FUNCTION Post () AS HRESULT

END TYPE
' ========================================================================================

' ========================================================================================
' CDOMessage constructor
' ========================================================================================
PRIVATE CONSTRUCTOR CDOMessage
   ' // Initialize the COM library
   DIM hr AS HRESULT = CoInitialize(NULL)
   IF hr = S_OK OR hr = S_FALSE THEN m_bUninitCOM = TRUE
   ' // Create an instance of the IMessage interface
   m_pMessage = AfxNewCom("CDO.Message")
   ' // Create an instance of the IConfiguration interface
   m_pConfig = AfxNewCom("CDO.Configuration")
   ' // Get a reference to the Afx_AdoFields interface
   IF m_pConfig THEN m_pConfig->get_Fields(@m_pFields)
END CONSTRUCTOR
' ========================================================================================

' ========================================================================================
' CDOMessage destructor
' ========================================================================================
PRIVATE DESTRUCTOR CDOMessage
   ' // Release the interfaces
   IF m_pFields THEN m_pFields->Release
   IF m_pConfig THEN m_pConfig->Release
   IF m_pMessage THEN m_pMessage->Release
   ' // Free the COM library
   IF m_bUninitCOM THEN CoUninitialize
END DESTRUCTOR
' ========================================================================================

' ========================================================================================
' Returns the last result code.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.GetLastResult () AS HRESULT
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sets the last result code.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.SetResult (BYVAL Result AS HRESULT) AS HRESULT
   m_Result = Result
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Returns a description of the last result code.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.GetErrorInfo () AS DWSTRING
   DIM dws AS DWSTRING = "Error &h" & HEX(m_Result, 8)
   SELECT CASE UCASE(HEX(m_Result, 8))
      CASE "80040201" : dws += ": Uncaught exception"
      CASE "80040202" : dws += ": No data source has been opened for the object."
      CASE "80040203" : dws += ": The object does not support this type of data source"
      CASE "80040204" : dws += ": he object does not support the requested property name or namespace"
      CASE "80040205" : dws += ": The object does not support the requested property"
      CASE "80040206" : dws += ": The object is not active. It may have been deleted or it may not have been opened"
      CASE "80040207" : dws += ": The object does not support storing persistent state information for objects"
      CASE "80040208" : dws += ": The requested property or feature, while supported, is not available at this time or in this context"
      CASE "80040209" : dws += ": No default drop directory has been configured for this server"
      CASE "8004020A" : dws += ": The SMTP server name is required, and was not found in the configuration source"
      CASE "8004020B" : dws += ": The NNTP server name is required, and was not found in the configuration source"
      CASE "8004020C" : dws += ": At least one recipient is required, but none were found"
      CASE "8004020D" : dws += ": At least one of the From or Sender fields is required, and neither was found"
      CASE "8004020E" : dws += ": The server rejected the sender address"
      CASE "8004020F" : dws += ": The server rejected one or more recipient addresses"
      CASE "80040210" : dws += ": The message could not be posted to the NNTP server"
      CASE "80040211" : dws += ": The message could not be sent to the SMTP server"
      CASE "80040212" : dws += ": The transport lost its connection to the server"
      CASE "80040213" : dws += ": The transport failed to connect to the server"
      CASE "80040214" : dws += ": The Subject, From, and Newsgroup fields are all required, and one or more was not found"
      CASE "80040215" : dws += ": The server rejected the logon attempt due to authentication failure"
      CASE "80040216" : dws += ": The content type was not valid in this context"
      CASE "80040216" : dws += ": The content type was not valid in this context"
      CASE "80040217" : dws += ": The transport was unable to log on to the server"
      CASE "80040218" : dws += ": The requested resource could not be found"
      CASE "80040219" : dws += ": Access to the requested resource is denied"
      CASE "8004021A" : dws += ": The HTTP request failed"
      CASE "8004021B" : dws += ": This is a multipart body part. It has no content other than the body parts contained within it"
      CASE "8004021C" : dws += ": Multipart body parts must be encoded as 7bit, 8bit, or binary"
      CASE "8004021D" : dws += ": Unsafe operation"
      CASE "8004021E" : dws += ": The requested property was not found"
      CASE "80040220" : dws += ": The 'SendUsing' configuration value is invalid"
      CASE "80040221" : dws += ": The 'postUsing' configuration value is invalid"
      CASE "80040222" : dws += ": The pickup directory path is required and was not specified"
      CASE "80040223" : dws += ": One or more messages could not be deleted"
      CASE "80040224" : dws += ": The requested operation is not available on the underlying object"
      CASE "80040227" : dws += ": The property is read-only"
      CASE "80040228" : dws += ": The property cannot be deleted"
      CASE "80040229" : dws += ": Data  written to the object are inconsistent or invalid"
      CASE "8004022A" : dws += ": The requested property is not in the mail header namespace"
      CASE "8004022B" : dws += ": The requested character set is not installed on the computer"
      CASE "8004022C" : dws += ": The ADO stream has not been opened"
      CASE "8004022D" : dws += ": The content properties are missing"
      CASE "8004022E" : dws += ": Content properties XML must be encoded using UTF-8"
      CASE "8004022F" : dws += ": Failed to parse content properties XML"
      CASE "80040230" : dws += ": Failed to convert a property from XML to a requested type"
      CASE "80040231" : dws += ": No directories were specified for resolution"
      CASE "80040232" : dws += ": Failed to resolve against one or more of the specified directories"
      CASE "80040233" : dws += ": Could not find the Sender's mailbox"
      CASE "80040234" : dws += ": Binding to self is not allowed"
      CASE "80040235" : dws += ": Data of the Attendees in the object are inconsistent or invalid"
      CASE "80040236" : dws += ": No more roles of the desired type are available - maxed out"
      CASE "80040237" : dws += ": The task with its current task type does not permit a call on Assign. Please call CreateUnassignedCopy instead."
      CASE "80040238" : dws += ": The task is not assigned to the current user"
      CASE "80040239" : dws += ": The item being accessed or manipulated is out of date"
      CASE "80044000" : dws += ": The first argument is invalid"
      CASE "80044001" : dws += ": The second argument is invalid"
      CASE "80044002" : dws += ": The third argument is invalid"
      CASE "80044003" : dws += ": The fourth argument is invalid"
      CASE "80044004" : dws += ": The fifth argument is invalid"
      CASE "800CCE05" : dws += ": The requested body part was not found in this message"
      CASE "800CCE1D" : dws += ": The content encoding type is invalid"
      CASE ELSE
         DIM cbLen AS DWORD, pBuffer AS WSTRING PTR, dwsMsg AS DWSTRING
         cbLen = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER OR _
                 FORMAT_MESSAGE_FROM_SYSTEM OR FORMAT_MESSAGE_IGNORE_INSERTS, _
                 NULL, m_Result, BYVAL MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
                 cast(LPWSTR, @pBuffer), 0, NULL)
         IF cbLen THEN
            dwsMsg = *pBuffer
            LocalFree pBuffer
         END IF
         IF m_Result = 0 THEN
            dws = dwsMsg
         ELSE
            dws += " " & dwsMsg
         END IF
   END SELECT
   RETURN dws
END FUNCTION

' ========================================================================================
' Sets the value of the specified property.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.ConfigValue (BYREF wszField AS WSTRING, BYREF wszValue AS WSTRING) AS HRESULT
   IF m_pConfig = NULL OR m_pFields = NULL THEN RETURN E_NOINTERFACE
   DIM pField AS Afx_AdoField PTR
   DIM vField AS VARIANT: vField.vt = VT_BSTR : vField.bstrVal = SysAllocString(wszField)
   this.SetResult(m_pFields->get_Item(vField, @pField))
   VariantClear @vField
   IF m_Result <> S_OK THEN RETURN m_Result
   DIM vValue AS VARIANT: vValue.vt = VT_BSTR : vValue.bstrVal = SysAllocString(wszValue)
   this.SetResult(pField->put_Value(vValue))
   VariantClear @vValue
   pField->Release
   RETURN m_Result
END FUNCTION
' ========================================================================================
' ========================================================================================
PRIVATE FUNCTION CDOMessage.ConfigValue (BYREF wszField AS WSTRING, BYVAL value AS LONG) AS HRESULT
   IF m_pConfig = NULL OR m_pFields = NULL THEN RETURN E_NOINTERFACE
   DIM pField AS Afx_AdoField PTR
   DIM vField AS VARIANT: vField.vt = VT_BSTR : vField.bstrVal = SysAllocString(wszField)
   this.SetResult(m_pFields->get_Item(vField, @pField))
   VariantClear @vField
   IF m_Result <> S_OK THEN RETURN m_Result
   DIM vValue AS VARIANT: vValue.vt = VT_I4 : vValue.lVal = value
   this.SetResult(pField->put_Value(vValue))
   VariantClear @vValue
   pField->Release
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Saves changes made to the Fields collection.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.ConfigUpdate () AS HRESULT
   IF m_pFields = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pFields->Update)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds to the list of blind copy (Bcc) recipients.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.BCC (BYREF wszBCC AS WSTRING) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   DIM bstrBCC AS Afx_BSTR = SysAllocString(wszBCC)
   this.SetResult(m_pMessage->put_BCC(bstrBCC))
   SysFreeString bstrBCC
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds to the list of copy (Cc) recipients.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.CC (BYREF wszCC AS WSTRING) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   DIM bstrCC AS AFX_BSTR = SysAllocString(wszCC)
   this.SetResult(m_pMessage->put_CC(bstrCC))
   SysFreeString bstrCC
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The messaging address of the principal author of the message.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.From (BYREF wszFrom AS WSTRING) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   DIM bstrFrom AS AFX_BSTR = SysAllocString(wszFrom)
   this.SetResult(m_pMessage->put_From(bstrFrom))
   SysFreeString bstrFrom
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The messaging address of the message submitter.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.Sender (BYREF wszSender AS WSTRING) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   DIM bstrSender AS AFX_BSTR = SysAllocString(wszSender)
   this.SetResult(m_pMessage->put_Sender(bstrSender))
   SysFreeString bstrSender
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The principal (To) recipients for this message.
' Named as Recipients instead of To because To is a reserved word.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.Recipients (BYREF wszTo AS WSTRING) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   DIM bstrTo AS AFX_BSTR = SysAllocString(wszTo)
   this.SetResult(m_pMessage->put_To(bstrTo))
   SysFreeString bstrTo
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The messaging addresses to which replies to this message should be sent.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.ReplyTo (BYREF wszReplyTo AS WSTRING) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   DIM bstrReplyTo AS AFX_BSTR = SysAllocString(wszReplyTo)
   this.SetResult(m_pMessage->put_ReplyTo(bstrReplyTo))
   SysFreeString bstrReplyTo
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The list of keywords for this message.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.Keywords (BYREF wszKeywords AS WSTRING) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   DIM bstrKeywords AS AFX_BSTR = SysAllocString(wszKeywords)
   this.SetResult(m_pMessage->put_Keywords(bstrKeywords))
   SysFreeString bstrKeywords
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The organization of the sender.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.Organization (BYREF wszOrganization AS WSTRING) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   DIM bstrOrganization AS AFX_BSTR = SysAllocString(wszOrganization)
   this.SetResult(m_pMessage->put_Organization(bstrOrganization))
   SysFreeString bstrOrganization
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The message subject.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.Subject (BYREF wszSubject AS WSTRING) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   DIM bstrSubject AS AFX_BSTR = SysAllocString(wszSubject)
   this.SetResult(m_pMessage->put_Subject(bstrSubject))
   SysFreeString bstrSubject
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The plain text representation of the message.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.TextBody (BYREF wszTextBody AS WSTRING) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   DIM bstrTextBody AS AFX_BSTR = SysAllocString(wszTextBody)
   this.SetResult(m_pMessage->put_TextBody(bstrTextBody))
   SysFreeString bstrTextBody
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The plain text representation of the message.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.HTMLBody (BYREF wszHTMLBody AS WSTRING) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   DIM bstrHTMLBody AS AFX_BSTR = SysAllocString(wszHTMLBody)
   this.SetResult(m_pMessage->put_HTMLBody(bstrHTMLBody))
   SysFreeString bstrHTMLBody
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Includes a request for a return report on the delivery status of the message.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.DSNOptions (BYVAL pDSNOptions AS CdoDSNOptions) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_DSNOptions(pDSNOptions))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates whether a Message Disposition Notification is requested on a message.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.MimeFormatted (BYVAL pMimeFormatted AS BOOLEAN) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_MimeFormatted(pMimeFormatted))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Indicates whether a Message Disposition Notification is requested on a message.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.MDNRequested (BYVAL pMDNRequested AS BOOLEAN) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   RETURN this.SetResult(m_pMessage->put_MDNRequested(pMDNRequested))
END FUNCTION
' ========================================================================================

' ========================================================================================
' Adds an attachment to this message.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.AddAttachment (BYREF wszURL AS WSTRING, BYREF wszUserName AS WSTRING = "", BYREF wszPassword AS WsTRING = "") AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   DIM bstrURL AS AFX_BSTR = SysAllocString(wszURL)
   DIM bstrUserName AS AFX_BSTR = SysAllocString(wszUserName)
   DIM bstrPassword AS AFX_BSTR = SysAllocString(wszPassword)
   this.SetResult(m_pMessage->AddAttachment(bstrURL, bstrUserName, bstrPassword, NULL))
   SysFreeString bstrURL
   SysFreeString bstrUserName
   SysFreeString bstrPassword
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Sends the message.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.Send () AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   ' // Set the configuration
   m_pMessage->put_Configuration(m_pConfig)
   ' // Send the message
   RETURN this.SetResult(m_pMessage->Send)
END FUNCTION
' ========================================================================================

' ========================================================================================
' Newsgroups to which any responses to this message are posted.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.FollowUpTo (BYREF wszFollowUpTo AS WSTRING) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   DIM bstrFollowUpTo AS AFX_BSTR = SysAllocString(wszFollowUpTo)
   this.SetResult(m_pMessage->put_ReplyTo(bstrFollowUpTo))
   SysFreeString bstrFollowUpTo
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' The newsgroup recipients for the message.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.Newsgroups (BYREF wszNewsgroups AS WSTRING) AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   DIM bstrNewsgroups AS AFX_BSTR = SysAllocString(wszNewsgroups)
   this.SetResult(m_pMessage->put_To(bstrNewsgroups))
   SysFreeString bstrNewsgroups
   RETURN m_Result
END FUNCTION
' ========================================================================================

' ========================================================================================
' Posts the message.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.Post () AS HRESULT
   IF m_pMessage = NULL THEN RETURN E_NOINTERFACE
   ' // Set the configuration
   m_pMessage->put_Configuration(m_pConfig)
   ' // Post the message
   this.SetResult(m_pMessage->Post)
   RETURN m_Result
END FUNCTION
' ========================================================================================

END NAMESPACE
