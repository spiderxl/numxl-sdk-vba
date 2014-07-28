Attribute VB_Name = "SFDBM"
Option Explicit
Option Compare Text

'* \defgroup SFDBM SFDBM
'*   Public APIs to process date, holiday and calendar calculations.
'* @{
'*  \file SFDBM.bas
'*  \brief  header file for SFDBM public APIs
'*  \details  Internal database API calls; used for date, holiday and calendar calculations.
'*  \copyright (c) 2007-2013 Spider Financial Corp.
'*             All rights reserved.
'*  \author Spider Financial Corp
'*  \version 1.62
'*  $Revision: 13918 $
'*  $Date: 2013-11-14 11:38:42 -0600 (Thu, 14 Nov 2013) $
Private Const msMODULE As String = "SFDBM"


#If Debugging = 0 Then
  #If VBA7 Then
    Public Declare PtrSafe Function SFDB_Init Lib "SFDBM.DLL" Alias "#100" (ByVal szAppName As String, _
                                                                                   ByVal szKey As String, _
                                                                                   ByVal szActCode As String, _
                                                                                   ByVal szTmpPath As String) As Integer
    Public Declare PtrSafe Function SFDB_Shutdown Lib "SFDBM.DLL" Alias "#105" () As Integer
    Public Declare PtrSafe Function SFDB_EDATE Lib "SFDBM.DLL" Alias "#200" (ByVal argDate As Long, _
                                                                                   ByVal szPeriod As String, _
                                                                                   ByRef retVal As Long) As Integer
    Public Declare PtrSafe Function SFDB_NWKDAY Lib "SFDBM.DLL" Alias "#201" (ByVal weekdy As Integer, _
                                                                                   ByVal order As Integer, _
                                                                                   ByVal mnth As Integer, _
                                                                                   ByVal year As Integer, _
                                                                                   ByRef retVal As Long) As Integer
    Public Declare PtrSafe Function SFDB_WKDYOrder Lib "SFDBM.DLL" Alias "#202" (ByVal argDate As Long, _
                                                                                  ByRef retVal As Integer) As Integer
    Public Declare PtrSafe Function SFDB_WEEKDAY Lib "SFDBM.DLL" Alias "#203" (ByVal argDate As Long, _
                                                                                    ByVal argReturnType As Integer, _
                                                                                    ByRef retVal As Integer) As Integer
    Public Declare PtrSafe Function SFDB_DTADJUST Lib "SFDBM.DLL" Alias "#204" (ByVal argDate As Long, _
                                                                                      ByVal argNextPrev As Integer, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal wkndNo As Integer, _
                                                                                      ByRef retVal As Long) As Integer
    Public Declare PtrSafe Function SFDB_ISWRKDY Lib "SFDBM.DLL" Alias "#205" (ByVal argDate As Long, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal wkndNo As Integer) As Integer
    Public Declare PtrSafe Function SFDB_NETWRKDYS Lib "SFDBM.DLL" Alias "#206" (ByVal argStartDate As Long, _
                                                                                      ByVal argEndDate As Long, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal wkndNo As Integer, _
                                                                                      ByRef retVal As Long) As Integer
    Public Declare PtrSafe Function SFDB_WORKDAY Lib "SFDBM.DLL" Alias "#207" (ByVal argDate As Long, _
                                                                                      ByVal days As Integer, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal wkndNo As Integer, _
                                                                                      ByRef retVal As Long) As Integer
    ' Holidays
    Public Declare PtrSafe Function SFDB_HLDYS Lib "SFDBM.DLL" Alias "#220" (ByVal argHolidays As String, _
                                                                                      ByVal retVal As String, _
                                                                                      ByRef nLen As Long) As Integer
    Public Declare PtrSafe Function SFDB_FindHLDY Lib "SFDBM.DLL" Alias "#221" (ByVal argDate As Long, _
                                                                                      ByVal argHolidays As String, _
                                                                                      ByVal retVal As String, _
                                                                                      ByRef nLen As Long) As Integer
    Public Declare PtrSafe Function SFDB_HLDYName Lib "SFDBM.DLL" Alias "#222" (ByVal code As String, _
                                                                                      ByVal retVal As String, _
                                                                                      ByRef nLen As Long) As Integer
    Public Declare PtrSafe Function SFDB_HLDYDate Lib "SFDBM.DLL" Alias "#223" (ByVal argDate As Long, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal holidays As String, _
                                                                                      ByVal retType As Integer, _
                                                                                      ByRef retVal As Long) As Integer
    Public Declare PtrSafe Function SFDB_ISHLDY Lib "SFDBM.DLL" Alias "#224" (ByVal argDate As Long, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal holidays As String) As Integer
    Public Declare PtrSafe Function SFDB_HLDYDates Lib "SFDBM.DLL" Alias "#225" (ByVal argStartDate As Long, _
                                                                                      ByVal argEndDate As Long, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef retVal As Long, _
                                                                                      ByRef nSize As Long) As Integer
    Public Declare PtrSafe Function SFDB_ISVALIDHLDYCODE Lib "SFDBM.DLL" Alias "#226" (ByVal argCode As String) As Integer
    
    'Weekend
    Public Declare PtrSafe Function SFDB_WKNDCode Lib "SFDBM.DLL" Alias "#240" (ByVal argNumber As Integer, _
                                                                                      ByVal retVal As String, _
                                                                                      ByRef nSize As Long) As Integer
    Public Declare PtrSafe Function SFDB_WKNDNo Lib "SFDBM.DLL" Alias "#241" (ByVal argCode As String, _
                                                                                    ByRef nWkndNo As Integer) As Integer
    Public Declare PtrSafe Function SFDB_ISWKND Lib "SFDBM.DLL" Alias "#242" (ByVal argDate As Long, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal wkndNo As Integer, _
                                                                                      ByVal argOptions As Integer) As Integer
    Public Declare PtrSafe Function SFDB_WKNDur Lib "SFDBM.DLL" Alias "#243" (ByVal argDate As Long, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal wkndNo As Integer, _
                                                                                      ByRef retVal As Integer) As Integer
    Public Declare PtrSafe Function SFDB_WKNDate Lib "SFDBM.DLL" Alias "#244" (ByVal argDate As Long, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal wkndNo As Integer, _
                                                                                      ByVal wkndOptions As Integer, _
                                                                                      ByVal direction As Integer, _
                                                                                      ByRef retVal As Long) As Integer
    Public Declare PtrSafe Function SFDB_ISVALIDWKND Lib "SFDBM.DLL" Alias "#245" (ByVal szCode As String) As Integer
                                                                                      
    ' Calendar
    Public Declare PtrSafe Function SFDB_CALNAME Lib "SFDBM.DLL" Alias "#260" (ByVal szCode As String, _
                                                                                     ByVal szRetVal As String, _
                                                                                     ByRef nLen As Long) As Integer
    Public Declare PtrSafe Function SFDB_CALENDARS Lib "SFDBM.DLL" Alias "#261" (ByVal argName As String, _
                                                                                     ByVal szRetVal As String, _
                                                                                     ByRef nLen As Long, _
                                                                                     ByRef nNumber As Integer) As Integer
    Public Declare PtrSafe Function SFDB_CALHLDYS Lib "SFDBM.DLL" Alias "#262" (ByVal argCalCode As String, _
                                                                                     ByVal szRetVal As String, _
                                                                                     ByRef nLen As Long, _
                                                                                     ByRef nNumber As Integer) As Integer
    Public Declare PtrSafe Function SFDB_CALWKND Lib "SFDBM.DLL" Alias "#263" (ByVal argCalCode As String, _
                                                                                     ByRef nWkndNo As Integer) As Integer
    Public Declare PtrSafe Function SFDB_ISVALIDCALCODE Lib "SFDBM.DLL" Alias "#264" (ByVal argCalCode As String) As Integer
                                                                                      
    'Country Support
    Public Declare PtrSafe Function SFDB_ISVALIDCNTRYCODE Lib "SFDBM.DLL" Alias "#300" (ByVal argCalCode As String) As Integer
    Public Declare PtrSafe Function SFDB_GETWKNDFROMCNTRY Lib "SFDBM.DLL" Alias "#301" (ByVal argCalCode As String, _
                                                                                              ByVal retVal As String, _
                                                                                              ByRef nLen As Long) As Integer
    Public Declare PtrSafe Function SFDB_GETCALFROMCNTRY Lib "SFDBM.DLL" Alias "#302" (ByVal argCalCode As String, _
                                                                                              ByVal retVal As String, _
                                                                                              ByRef nLen As Long) As Integer
    'Currency Support
    Public Declare PtrSafe Function SFDB_ISVALIDCCYCODE Lib "SFDBM.DLL" Alias "#320" (ByVal szCCY As String) As Integer
    Public Declare PtrSafe Function SFDB_GETWKNDFROMCCY Lib "SFDBM.DLL" Alias "#322" (ByVal szCCY As String, _
                                                                                            ByVal retVal As String, _
                                                                                            ByRef nLen As Long) As Integer
    Public Declare PtrSafe Function SFDB_GETCALFROMCCY Lib "SFDBM.DLL" Alias "#323" (ByVal szCCY As String, _
                                                                                            ByVal retVal As String, _
                                                                                            ByRef nLen As Long) As Integer
    ' FX MKT Support
    Public Declare PtrSafe Function SFDB_GETVALIDCCYPAIR Lib "SFDBM.DLL" Alias "#340" (ByVal CCY1 As String, _
                                                                                              ByVal CCY2 As String, _
                                                                                              ByVal CCYPair As String, _
                                                                                              ByRef nLen As Long) As Integer
  #Else
    Public Declare Function SFDB_Init Lib "SFDBM.DLL" Alias "#100" (ByVal szAppName As String, _
                                                                                   ByVal szKey As String, _
                                                                                   ByVal szActCode As String, _
                                                                                   ByVal szTmpPath As String) As Integer
    Public Declare Function SFDB_Shutdown Lib "SFDBM.DLL" Alias "#105" () As Integer
    Public Declare Function SFDB_EDATE Lib "SFDBM.DLL" Alias "#200" (ByVal argDate As Long, _
                                                                                   ByVal szPeriod As String, _
                                                                                   ByRef retVal As Long) As Integer
    Public Declare Function SFDB_NWKDAY Lib "SFDBM.DLL" Alias "#201" (ByVal weekdy As Integer, _
                                                                                   ByVal order As Integer, _
                                                                                   ByVal mnth As Integer, _
                                                                                   ByVal year As Integer, _
                                                                                   ByRef retVal As Long) As Integer
    Public Declare Function SFDB_WKDYOrder Lib "SFDBM.DLL" Alias "#202" (ByVal argDate As Long, _
                                                                                  ByRef retVal As Integer) As Integer
    Public Declare Function SFDB_WEEKDAY Lib "SFDBM.DLL" Alias "#203" (ByVal argDate As Long, _
                                                                                    ByVal argReturnType As Integer, _
                                                                                    ByRef retVal As Integer) As Integer
    Public Declare Function SFDB_DTADJUST Lib "SFDBM.DLL" Alias "#204" (ByVal argDate As Long, _
                                                                                      ByVal argNextPrev As Integer, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal wkndNo As Integer, _
                                                                                      ByRef retVal As Long) As Integer
    Public Declare Function SFDB_ISWRKDY Lib "SFDBM.DLL" Alias "#205" (ByVal argDate As Long, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal wkndNo As Integer) As Integer
    Public Declare Function SFDB_NETWRKDYS Lib "SFDBM.DLL" Alias "#206" (ByVal argStartDate As Long, _
                                                                                      ByVal argEndDate As Long, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal wkndNo As Integer, _
                                                                                      ByRef retVal As Long) As Integer
    Public Declare Function SFDB_WORKDAY Lib "SFDBM.DLL" Alias "#207" (ByVal argDate As Long, _
                                                                                      ByVal days As Integer, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal wkndNo As Integer, _
                                                                                      ByRef retVal As Long) As Integer
    ' Holidays
    Public Declare Function SFDB_HLDYS Lib "SFDBM.DLL" Alias "#220" (ByVal argHolidays As String, _
                                                                                      ByVal retVal As String, _
                                                                                      ByRef nLen As Long) As Integer
    Public Declare Function SFDB_FindHLDY Lib "SFDBM.DLL" Alias "#221" (ByVal argDate As Long, _
                                                                                      ByVal argHolidays As String, _
                                                                                      ByVal retVal As String, _
                                                                                      ByRef nLen As Long) As Integer
    Public Declare Function SFDB_HLDYName Lib "SFDBM.DLL" Alias "#222" (ByVal code As String, _
                                                                                      ByVal retVal As String, _
                                                                                      ByRef nLen As Long) As Integer
    Public Declare Function SFDB_HLDYDate Lib "SFDBM.DLL" Alias "#223" (ByVal argDate As Long, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal holidays As String, _
                                                                                      ByVal retType As Integer, _
                                                                                      ByRef retVal As Long) As Integer
    Public Declare Function SFDB_ISHLDY Lib "SFDBM.DLL" Alias "#224" (ByVal argDate As Long, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal holidays As String) As Integer
    Public Declare Function SFDB_HLDYDates Lib "SFDBM.DLL" Alias "#225" (ByVal argStartDate As Long, _
                                                                                      ByVal argEndDate As Long, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef retVal As Long, _
                                                                                      ByRef nSize As Long) As Integer
    Public Declare Function SFDB_ISVALIDHLDYCODE Lib "SFDBM.DLL" Alias "#226" (ByVal argCode As String) As Integer
                                                                                      
    'Weekend
    Public Declare Function SFDB_WKNDCode Lib "SFDBM.DLL" Alias "#240" (ByVal argNumber As Integer, _
                                                                                      ByVal retVal As String, _
                                                                                      ByRef nSize As Long) As Integer
    Public Declare Function SFDB_WKNDNo Lib "SFDBM.DLL" Alias "#241" (ByVal argCode As String, _
                                                                                    ByRef nWkndNo As Integer) As Integer
    Public Declare Function SFDB_ISWKND Lib "SFDBM.DLL" Alias "#242" (ByVal argDate As Long, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal wkndNo As Integer, _
                                                                                      ByVal argOptions As Integer) As Integer
    Public Declare Function SFDB_WKNDur Lib "SFDBM.DLL" Alias "#243" (ByVal argDate As Long, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal wkndNo As Integer, _
                                                                                      ByRef retVal As Integer) As Integer
    Public Declare Function SFDB_WKNDate Lib "SFDBM.DLL" Alias "#244" (ByVal argDate As Long, _
                                                                                      ByVal holidays As String, _
                                                                                      ByRef zDates As Long, _
                                                                                      ByVal nSize As Long, _
                                                                                      ByVal wkndNo As Integer, _
                                                                                      ByVal wkndOptions As Integer, _
                                                                                      ByVal direction As Integer, _
                                                                                      ByRef retVal As Long) As Integer
    Public Declare Function SFDB_ISVALIDWKND Lib "SFDBM.DLL" Alias "#245" (ByVal szCode As String) As Integer
                                                                                      
    ' Calendar
    Public Declare Function SFDB_CALNAME Lib "SFDBM.DLL" Alias "#260" (ByVal szCode As String, _
                                                                                     ByVal szRetVal As String, _
                                                                                     ByRef nLen As Long) As Integer
    Public Declare Function SFDB_CALENDARS Lib "SFDBM.DLL" Alias "#261" (ByVal argName As String, _
                                                                                     ByVal szRetVal As String, _
                                                                                     ByRef nLen As Long, _
                                                                                     ByRef nNumber As Integer) As Integer
    Public Declare Function SFDB_CALHLDYS Lib "SFDBM.DLL" Alias "#262" (ByVal argCalCode As String, _
                                                                                     ByVal szRetVal As String, _
                                                                                     ByRef nLen As Long, _
                                                                                     ByRef nNumber As Integer) As Integer
    Public Declare Function SFDB_CALWKND Lib "SFDBM.DLL" Alias "#263" (ByVal argCalCode As String, _
                                                                                     ByRef nWkndNo As Integer) As Integer
    Public Declare Function SFDB_ISVALIDCALCODE Lib "SFDBM.DLL" Alias "#264" (ByVal argCalCode As String) As Integer
                                                                                      
    'Country Support
    Public Declare Function SFDB_ISVALIDCNTRYCODE Lib "SFDBM.DLL" Alias "#300" (ByVal argCalCode As String) As Integer
    Public Declare Function SFDB_GETWKNDFROMCNTRY Lib "SFDBM.DLL" Alias "#301" (ByVal argCalCode As String, _
                                                                                ByVal retVal As String, _
                                                                                ByRef nLen As Long) As Integer
    Public Declare Function SFDB_GETCALFROMCNTRY Lib "SFDBM.DLL" Alias "#302" (ByVal argCalCode As String, _
                                                                                ByVal retVal As String, _
                                                                                ByRef nLen As Long) As Integer
    'Currency Support
    Public Declare Function SFDB_ISVALIDCCYCODE Lib "SFDBM.DLL" Alias "#320" (ByVal szCCY As String) As Integer
    Public Declare Function SFDB_GETWKNDFROMCCY Lib "SFDBM.DLL" Alias "#322" (ByVal szCCY As String, _
                                                                              ByVal retVal As String, _
                                                                              ByRef nLen As Long) As Integer
    Public Declare Function SFDB_GETCALFROMCCY Lib "SFDBM.DLL" Alias "#323" (ByVal szCCY As String, _
                                                                              ByVal retVal As String, _
                                                                              ByRef nLen As Long) As Integer
    ' FX MKT Support
    Public Declare Function SFDB_GETVALIDCCYPAIR Lib "SFDBM.DLL" Alias "#340" (ByVal CCY1 As String, _
                                                                                ByVal CCY2 As String, _
                                                                                ByVal CCYPair As String, _
                                                                                ByRef nLen As Long) As Integer
  #End If
#End If



'* @}

