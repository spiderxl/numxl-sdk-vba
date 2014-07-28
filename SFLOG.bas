Attribute VB_Name = "SFLOG"
Option Explicit
Option Compare Text
'*  \addtogroup SFLOG
'*  @{
'*  \file SFLOG.bas
'*
'*  \copyright (c) 2007-2013 Spider Financial Corp.
'*  All rights reserved.
'*  \brief  VBA declaration for the public C API of SFLOG.DLL library
'*  \details  logging system support for NumXL SDK
'*  \copyright (c) 2007-2013 Spider Financial Corp.
'*             All rights reserved.
'*  \author Spider Financial Corp
'*  \version 1.62
'*  $Revision: 13918 $
'*  $Date: 2013-11-14 11:38:42 -0600 (Thu, 14 Nov 2013) $
Private Const msMODULE As String = "SFLOG"


#If Debugging = 0 Then
  #If VBA7 Then
    Public Declare PtrSafe Function SFLOG_Init Lib "SFLog.DLL" Alias "#100" (ByVal szAppName As String, _
                                                                                    ByVal szLogDir As String) As Integer
    Public Declare PtrSafe Function SFLOG_Shutdown Lib "SFLog.DLL" Alias "#105" () As Integer
    Public Declare PtrSafe Function SFLOG_LogMsg Lib "SFLog.DLL" Alias "#110" (ByVal nLevel As Integer, _
                                                                                    ByVal szFilename As String, _
                                                                                    ByVal szFuncName As String, _
                                                                                    ByVal szFuncSig As String, _
                                                                                    ByVal nLineNo As Integer, _
                                                                                    ByVal szMsg As String) As Integer
    Public Declare PtrSafe Function SFLOG_GETLEVEL Lib "SFLog.DLL" Alias "#115" (ByRef nLevel As Integer) As Integer
    Public Declare PtrSafe Function SFLOG_SETLEVEL Lib "SFLog.DLL" Alias "#120" (ByVal nLevel As Integer) As Integer
  #Else
    Public Declare Function SFLOG_Init Lib "SFLog.DLL" Alias "#100" (ByVal szAppName As String, ByVal szLogDir As String) As Integer
    Public Declare Function SFLOG_Shutdown Lib "SFLog.DLL" Alias "#105" () As Integer
    Public Declare Function SFLOG_LogMsg Lib "SFLog.DLL" Alias "#110" (ByVal nLevel As Integer, _
                                                                                    ByVal szFilename As String, _
                                                                                    ByVal szFuncName As String, _
                                                                                    ByVal szFuncSig As String, _
                                                                                    ByVal nLineNo As Integer, _
                                                                                    ByVal szMsg As String) As Integer
    Public Declare Function SFLOG_GETLEVEL Lib "SFLog.DLL" Alias "#115" (ByRef nLevel As Integer) As Integer
    Public Declare Function SFLOG_SETLEVEL Lib "SFLog.DLL" Alias "#120" (ByVal nLevel As Integer) As Integer
  #End If
#End If


'* @}


