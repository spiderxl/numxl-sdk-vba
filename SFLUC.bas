Attribute VB_Name = "SFLUC"
Option Explicit
Option Compare Text
'!  \addtogroup SFLUC
'!  @{
'!  \file SFLUC.bas
'!
'!  \copyright (c) 2007-2013 Spider Financial Corp.
'!  All rights reserved.
'!  \brief  header file for the public API of SFLUC library
'!  \details  License system support for NumXL SDK
'!  \copyright (c) 2007-2013 Spider Financial Corp.
'!             All rights reserved.
'!  \author Spider Financial Corp
'!  \version 1.62
'!  $Revision: 13918 $
'!  $Date: 2013-11-14 11:38:42 -0600 (Thu, 14 Nov 2013) $
Private Const msMODULE As String = "SFLUC"

' Are we running Excel 2010/2013?
#If Debugging = 0 Then
  #If VBA7 Then
    Public Declare PtrSafe Function SFLUC_Init Lib "SFLUC.DLL" Alias "#100" (ByVal szAppName As String, _
                                                                                    ByVal szLogDir As String, _
                                                                                    ByVal bEnableGracePeriod As Boolean) As Integer
    Public Declare PtrSafe Function SFLUC_Shutdown Lib "SFLUC.DLL" Alias "#105" () As Integer
    Public Declare PtrSafe Function SFLUC_CHECK_LICENSE Lib "SFLUC.DLL" Alias "#200" () As Integer
    Public Declare PtrSafe Function SFLUC_LICENSE_LEVEL Lib "SFLUC.DLL" Alias "#205" (ByRef nLevel As Integer) As Integer
    Public Declare PtrSafe Function SFLUC_LICENSE_STATUS Lib "SFLUC.DLL" Alias "#210" () As Integer
    Public Declare PtrSafe Function SFLUC_CHECK_KEYCODE Lib "SFLUC.DLL" Alias "#215" (ByVal PDKey As String, _
                                                                                            ByVal szKey As String, _
                                                                                            ByVal szActCode As String, _
                                                                                            ByRef ulExpiry As Long, _
                                                                                            ByRef nLevel As Integer) As Integer
    Public Declare PtrSafe Function SFLUC_SERVICEDATE Lib "SFLUC.DLL" Alias "#220" (ByVal szLicenseKey As String, ByRef serviceDate As Long) As Integer
    Public Declare PtrSafe Function SFLUC_UPDATEVERSION Lib "SFLUC.DLL" Alias "#225" (ByVal szLicenseKey As String, _
                                                                                            ByVal szFileVersion As String, _
                                                                                            ByVal newVersion As String, ByRef dwSize As Integer, _
                                                                                            ByVal downloadURL As String, ByRef dwSize2 As Integer) As Integer
    Public Declare PtrSafe Function SFLUC_MACHINEID Lib "SFLUC.DLL" Alias "#300" (ByVal szBuffer As String, ByRef nLevel As Integer) As Integer
    Public Declare PtrSafe Function SFLUC_LICENSE_KEY Lib "SFLUC.DLL" Alias "#305" (ByVal szBuffer As String, ByRef nLevel As Integer) As Integer
    Public Declare PtrSafe Function SFLUC_LICENSE_KEY_EXPIRY Lib "SFLUC.DLL" Alias "#310" (ByRef nExpiry As Long) As Integer
    Public Declare PtrSafe Function SFLUC_ACTIVATION_CODE Lib "SFLUC.DLL" Alias "#315" (ByVal szBuffer As String, ByRef nLevel As Integer) As Integer
    
  #Else
    Public Declare Function SFLUC_Init Lib "SFLUC.DLL" Alias "#100" (ByVal szAppName As String, _
                                                                                    ByVal szLogDir As String, _
                                                                                    ByVal bEnableGracePeriod As Boolean) As Integer
    Public Declare Function SFLUC_Shutdown Lib "SFLUC.DLL" Alias "#105" () As Integer
    Public Declare Function SFLUC_CHECK_LICENSE Lib "SFLUC.DLL" Alias "#200" () As Integer
    Public Declare Function SFLUC_LICENSE_LEVEL Lib "SFLUC.DLL" Alias "#205" (ByRef nLevel As Integer) As Integer
    Public Declare Function SFLUC_LICENSE_STATUS Lib "SFLUC.DLL" Alias "#210" () As Integer
    Public Declare Function SFLUC_CHECK_KEYCODE Lib "SFLUC.DLL" Alias "#215" (ByVal PDKey As String, _
                                                                                            ByVal szKey As String, _
                                                                                            ByVal szActCode As String, _
                                                                                            ByRef ulExpiry As Long, _
                                                                                            ByRef nLevel As Integer) As Integer
    Public Declare Function SFLUC_SERVICEDATE Lib "SFLUC.DLL" Alias "#220" (ByVal szLicenseKey As String, ByRef serviceDate As Long) As Integer
    Public Declare Function SFLUC_UPDATEVERSION Lib "SFLUC.DLL" Alias "#225" (ByVal szLicenseKey As String, _
                                                                                            ByVal szFileVersion As String, _
                                                                                            ByVal newVersion As String, ByRef dwSize As Integer, _
                                                                                            ByVal downloadURL As String, ByRef dwSize2 As Integer) As Integer
    Public Declare Function SFLUC_MACHINEID Lib "SFLUC.DLL" Alias "#300" (ByVal szBuffer As String, ByRef nLevel As Integer) As Integer
    Public Declare Function SFLUC_LICENSE_KEY Lib "SFLUC.DLL" Alias "#305" (ByVal szBuffer As String, ByRef nLevel As Integer) As Integer
    Public Declare Function SFLUC_LICENSE_KEY_EXPIRY Lib "SFLUC.DLL" Alias "#310" (ByRef nExpiry As Long) As Integer
    Public Declare Function SFLUC_ACTIVATION_CODE Lib "SFLUC.DLL" Alias "#315" (ByVal szBuffer As String, ByRef nLevel As Integer) As Integer
    
  #End If
#End If



'* @}
