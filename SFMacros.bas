Attribute VB_Name = "SFMacros"
Option Explicit
Option Compare Text
'!  \file SFMacros.bas
'!  \brief  NumXL SDK error codes definitions
'!  The information on this page is intended to be used by programmers so that the software they write can better deal with errors.
'!
'!  \copyright (c) 2007-2013 Spider Financial Corp.
'!             All rights reserved.
'!  \author Spider Financial Corp
'!  \version 1.62
'!
'!  $Revision: 13865 $
'!  $Date: 2013-11-10 14:31:46 -0600 (Sun, 10 Nov 2013) $
Private Const msMODULE As String = "SFMacros"

'!  \enum  NDK_RETURN_TYPE
'!  \brief An enumeration of all return codes
'!
Public Enum NDK_RETURN_TYPE
  NDK_SUCCESS = 0                     '* SUCCESS return code.
  NDK_FAILED = -1                     '* FAILED: Internal error occured
  
  ' TRUE/FALSE aliases
  NDK_TRUE = 0                        '* SUCCESS: return value is TRUE.
  NDK_FALSE = 1                       '* SUCCESS: return value is FALSE
  
  NDK_SDK_UNINITIALIZED = -10         '* FAILED: The API NDK_INIT has not yet been called
  NDK_LOG_UNINITIALIZED = -11         '* FAILED: The API NDK_INIT has not yet been called
  NDK_LUC_UNINITIALIZED = -12         '* FAILED: The API NDK_INIT has not yet been called
  NDK_DBM_UNINITIALIZED = -13         '* FAILED: The API NDK_INIT has not yet been called
  
  NDK_LOG_INIT_FAILED = -20           '* FAILED: The logging system failed during initialization, check the configuration settings
  NDK_DB_INIT_FAILED = -21            '* FAILED: Missing or failed to open the database file
  NDK_LUC_INIT_FAILED = -22           '* FAILED: Missing or failed to open the database file
  
  ' Initialization error codes
  NDK_MISSING_CONF = -100             '* FAILED: The configuration file is missing
  NDK_BAD_CONF = -101                 '* FAILED: Access denied or corrupted file
  NDK_CONF_DATAPATH_INVALID = -102    '* FAILED: Invalid datapath value in the configuration file
  NDK_DATAPATH_INVALID = -103         '* FAILED: failed to retrieve/construct a temp path for logs and intermediate calculation
  NDK_CONF_PRODID_INVALID = -104      '* FAILED: Invalid value for [GLOBALS][PRODUCTID] entry in the conf file
  NDK_LOGFILE_INUSE = -105            '* FAILED: Failed to open the logfile for writing (permission error or file in use)
  NDK_MISSING_APP_ARG = -106          '* FAILED: invalid or Null argument (e.g. AppName for return value)
  NDK_MISSING_LICENSE_KEY = -107      '* FAILED: The product license ket is invalid
  NDK_INVALID_LICENSE_KEY = -108      '* FAILED: The product license ket is invalid
  NDK_INACTIVE_LICENSE_KEY = -109     '* FAILED: The license key has yet to be activated
  NDK_INVALID_KEY_CODE = -110         '* FAILED: The license key and code are not valid
  NDK_EXPIRED_LICENSE_KEY = -111      '* FAILED: The license key has expired
  NDK_LOW_LICENSE_LEVEL = -112        '* FAILED: The required license level is not met by current license
  
  ' Runtime error codes
  NDK_INVALID_ARG = -300              '* FAILED: an input argument with unexpected or invalid value.
  NDK_LENGTH_ERROR = -301             '* FAILED: The user's buffer is not big enough or Insufficient input data
  NDK_INVALID_VALUE = -302            '* FAILED: Invalid value of an argument
  NDK_EMPTY_TIME_SERIES = -303        '* FAILED: number of non-missing values is zero
  NDK_ZERO_INVALID_VARIANCE = -304    '* FAILED: number of non-missing values is zero
  NDK_CALIBRATION_ERROR = -305        '* FAILED: The optimizer failed to converge to a unique solution.
  NDK_INVALID_MODEL = -306            '* FAILED: The model's parameters values did not pass the stability test.
  
  ' Implementation status
  NDK_NOTSUPPORTED = -400             '* FAILED: The required operation is not currently implemented/supported
  
  ' Warnings codes
  NDK_RET_NAN = 100                   '* WARNING: The function returns an invalid (i.e. missing) value
  NDK_SKIP_INIT = 105                 '* WARNING: The DLL is already initialize, skipping !
  
  NDK_KEY_IN_GRACE_PEROD = 1000       '* INFORMATION: the trial license key is in the 7-day grace period
  NDK_KEY_IN_TRIAL_PEROD = 1005       '* INFORMATION: the trial license key is in the free trial period
  NDK_KEY_NOT_IN_TRIAL_PEROD = 1010   '* INFORMATION: the trial license key is not in the free trial period
  NDK_PERP_KEY_ACTIVE = 1015          '* INFORMATION: the perpetual license key is activated
  NDK_PERP_KEY_INACTIVE = 1020        '* INFORMATION: the perpetual license key is not activated
  NDK_SUB_KEY_ACTIVE = 1025           '* INFORMATION: the subscription license key is activated
  NDK_SUB_KEY_INACTIVE = 1030         '* INFORMATION: the subscription license key is not activated
End Enum

