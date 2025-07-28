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
' Returns the description of the most recent OLE error in the current logical thread and
' clears the error state for the thread. It should be called as soon as possible after
' calling a method of an Automation interface and will only succeed if the object supports
' the IErrorInfo interface.
' ========================================================================================
PRIVATE FUNCTION CDOMessage.GetErrorInfo () AS DWSTRING
   DIM pErrorInfo AS IErrorInfo PTR, dwsInfo AS DWSTRING
   DIM hr AS HRESULT = ..GetErrorInfo(0, @pErrorInfo)
   IF hr <> S_OK THEN RETURN ""
   DIM bstrDescription AS AFX_BSTR
   hr = pErrorInfo->lpvtbl->GetDescription(pErrorInfo, @bstrDescription)
   IF hr = S_OK THEN
      dwsInfo = *bstrDescription
      SysFreeString bstrDescription
      AfxSafeRelease(pErrorInfo)
   END IF
   RETURN dwsInfo
END FUNCTION
' ========================================================================================

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
