Attribute VB_Name = "SFSDK"
Option Explicit
Option Compare Text
'*  \addtogroup SFSDK
'*  NumXL SDK Econometric and statistical APIs
'*  @{
'*  \file SFSDK.bas
'*
'*  \copyright (c) 2007-2013 Spider Financial Corp.
'*  All rights reserved.
'*  \brief  header file for the public API of SFSDK library
'*  \details function declaration for NumXL SDK Econometric and statistical APIs
'*  \copyright (c) 2007-2013 Spider Financial Corp.
'*             All rights reserved.
'*  \author Spider Financial Corp
'*  \version 1.62
'*  $Revision: 13920 $
'*  $Date: 2013-11-14 19:26:03 -0600 (Thu, 14 Nov 2013) $
Private Const msMODULE As String = "SFSDK"

#Const DllName = "SFSDK.DLL"




' Are we running Excel 2010/2013?
#If VBA7 Then
   Public Declare PtrSafe Function NDK_Init Lib "SFSDK.DLL" Alias "#100" (ByVal szAppName As String, _
                                                                                   ByVal szKey As String, _
                                                                                   ByVal szActivation As String, _
                                                                                   ByVal szLogDirectory As String) As Integer
   Public Declare PtrSafe Function NDK_Shutdown Lib "SFSDK.DLL" Alias "#105" () As Integer
   Public Declare PtrSafe Function NDK_INFO Lib "SFSDK.DLL" Alias "#110" (ByVal nRetType As Integer, _
                                                                                 ByVal szMsd As String, _
                                                                                 ByVal nlen As Long) As Integer
   Public Declare PtrSafe Function NDK_MSG Lib "SFSDK.DLL" Alias "#115" (ByVal nRetCode As Integer, _
                                                                               ByVal szMsd As String, _
                                                                               ByVal nlen As Long) As Integer
   'Time Series statistics
   Public Declare PtrSafe Function NDK_ACF Lib "SFSDK.DLL" Alias "#200" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_ACF_ERROR Lib "SFSDK.DLL" Alias "#205" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_ACFCI Lib "SFSDK.DLL" Alias "#210" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal alpha As Double, _
                                                                               ByRef UL As Double, _
                                                                               ByRef LL As Double) As Integer
   Public Declare PtrSafe Function NDK_PACF Lib "SFSDK.DLL" Alias "#215" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_PACF_ERROR Lib "SFSDK.DLL" Alias "#220" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_PACFCI Lib "SFSDK.DLL" Alias "#225" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal alpha As Double, _
                                                                               ByRef UL As Double, _
                                                                               ByRef LL As Double) As Integer
   'Statistical Tests
   Public Declare PtrSafe Function NDK_ACFTEST Lib "SFSDK.DLL" Alias "#300" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal target As Double, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_NORMALTEST Lib "SFSDK.DLL" Alias "#301" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_WNTEST Lib "SFSDK.DLL" Alias "#302" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_ARCHTEST Lib "SFSDK.DLL" Alias "#303" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
                                                                               
   Public Declare PtrSafe Function NDK_MEANTEST Lib "SFSDK.DLL" Alias "#304" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal target As Double, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
                                                                               
   Public Declare PtrSafe Function NDK_STDEVTEST Lib "SFSDK.DLL" Alias "#305" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal target As Double, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   
   Public Declare PtrSafe Function NDK_SKEWTEST Lib "SFSDK.DLL" Alias "#306" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
 
   Public Declare PtrSafe Function NDK_XKURTTEST Lib "SFSDK.DLL" Alias "#307" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_XCFTEST Lib "SFSDK.DLL" Alias "#308" (ByRef X As Double, _
                                                                                   ByRef Y As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal trger As Double, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_ADFTEST Lib "SFSDK.DLL" Alias "#309" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal options As Integer, _
                                                                               ByVal testdown As Boolean, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   ''' Perform Johansen cointegration test
   Public Declare PtrSafe Function NDK_JOHANSENTEST Lib "SFSDK.DLL" Alias "#311" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal M As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal PolyOrder As Integer, _
                                                                               ByVal tracetest As Boolean, _
                                                                               ByVal R As Integer, _
                                                                               ByVal alpha As Double, _
                                                                               ByRef retStat As Double, _
                                                                               ByRef retCV As Double) As Integer
 
   Public Declare PtrSafe Function NDK_COLNRTY_TEST Lib "SFSDK.DLL" Alias "#312" (ByRef XX As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal M As Long, _
                                                                               ByRef mask As Byte, _
                                                                               ByVal nMaskLen As Long, _
                                                                               ByVal method As Integer, _
                                                                               ByVal colIndex As Integer, _
                                                                               ByRef retVal As Double) As Integer
                                                                               
   Public Declare PtrSafe Function NDK_CHOWTEST Lib "SFSDK.DLL" Alias "#313" (ByRef XX1 As Double, _
                                                                               ByVal M As Long, _
                                                                               ByRef Y1 As Double, _
                                                                               ByVal N1 As Long, _
                                                                               ByRef XX2 As Double, _
                                                                               ByRef Y2 As Double, _
                                                                               ByVal N2 As Long, _
                                                                               ByRef mask As Byte, _
                                                                               ByVal nMaskLen As Long, _
                                                                               ByVal intercept As Double, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   ' General statistics
   Public Declare PtrSafe Function NDK_GINI Lib "SFSDK.DLL" Alias "#400" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef retVal As Double) As Integer
                                                                               
   Public Declare PtrSafe Function NDK_XCF Lib "SFSDK.DLL" Alias "#401" (ByRef X As Double, _
                                                                               ByRef Y As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
                                                                               
   Public Declare PtrSafe Function NDK_XKURT Lib "SFSDK.DLL" Alias "#403" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_SKEW Lib "SFSDK.DLL" Alias "#404" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_AVERAGE Lib "SFSDK.DLL" Alias "#405" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_VARIANCE Lib "SFSDK.DLL" Alias "#406" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_MIN Lib "SFSDK.DLL" Alias "#407" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_MAX Lib "SFSDK.DLL" Alias "#408" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_HURST_EXPONENT Lib "SFSDK.DLL" Alias "#409" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_QUANTILE Lib "SFSDK.DLL" Alias "#410" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal p As Double, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_IQR Lib "SFSDK.DLL" Alias "#411" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_RMS Lib "SFSDK.DLL" Alias "#412" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_MD Lib "SFSDK.DLL" Alias "#413" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_RMD Lib "SFSDK.DLL" Alias "#414" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_MAD Lib "SFSDK.DLL" Alias "#415" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_LRVAR Lib "SFSDK.DLL" Alias "#416" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal W As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_SAD Lib "SFSDK.DLL" Alias "#417" (ByRef Y As Double, _
                                                                               ByRef YHat As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_MAE Lib "SFSDK.DLL" Alias "#418" (ByRef Y As Double, _
                                                                               ByRef YHat As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_MAPE Lib "SFSDK.DLL" Alias "#419" (ByRef Y As Double, _
                                                                               ByRef YHat As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_RMSE Lib "SFSDK.DLL" Alias "#420" (ByRef Y As Double, _
                                                                               ByRef YHat As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_SSE Lib "SFSDK.DLL" Alias "#421" (ByRef Y As Double, _
                                                                               ByRef YHat As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_SORT_ASC Lib "SFSDK.DLL" Alias "#422" (ByRef Y As Double, _
                                                                               ByVal N As Long) As Integer
                                                                               
 
   ' Statistical distribution
   Public Declare PtrSafe Function NDK_GED_XKURT Lib "SFSDK.DLL" Alias "#500" (ByVal df As Double, _
                                                                                     ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_TDIST_XKURT Lib "SFSDK.DLL" Alias "#501" (ByVal df As Double, _
                                                                                     ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_EDF Lib "SFSDK.DLL" Alias "#502" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal target As Double, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_HIST_BINS Lib "SFSDK.DLL" Alias "#503" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal method As Integer, _
                                                                               ByRef retVal As Long) As Integer
   Public Declare PtrSafe Function NDK_HIST_BIN_LIMIT Lib "SFSDK.DLL" Alias "#504" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal nBins As Long, _
                                                                               ByVal index As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_HISTOGRAM Lib "SFSDK.DLL" Alias "#505" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal nBins As Long, _
                                                                               ByVal index As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_KERNEL_DENSITY_ESTIMATE Lib "SFSDK.DLL" Alias "#506" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal target As Double, _
                                                                               ByVal bandwidth As Double, _
                                                                               ByVal argKernelFunc As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_GAUSS_FORECI Lib "SFSDK.DLL" Alias "#507" (ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal upper As Boolean, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_TSTUDENT_FORECI Lib "SFSDK.DLL" Alias "#508" (ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal df As Double, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal upper As Boolean, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_GED_FORECI Lib "SFSDK.DLL" Alias "#509" (ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal df As Double, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal upper As Boolean, _
                                                                               ByRef retVal As Double) As Integer
   'ARMA Function
 Public Declare PtrSafe Function NDK_ARMA_GOF Lib "SFSDK.DLL" Alias "#600" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByRef phis As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef thetas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
                                                                               
   
   Public Declare PtrSafe Function NDK_ARMA_RESID Lib "SFSDK.DLL" Alias "#601" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByRef phis As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef thetas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal retType As Integer) As Integer
   Public Declare PtrSafe Function NDK_ARMA_FITTED Lib "SFSDK.DLL" Alias "#602" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByRef phis As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef thetas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal retType As Integer) As Integer
   Public Declare PtrSafe Function NDK_ARMA_FORE Lib "SFSDK.DLL" Alias "#603" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByRef phis As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef thetas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nSteps As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByVal alpha As Double, _
                                                                               ByRef retVal As Double) As Integer
  
   Public Declare PtrSafe Function NDK_ARMA_PARAM Lib "SFSDK.DLL" Alias "#605" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef avg As Double, _
                                                                               ByRef sigma As Double, _
                                                                               ByRef phis As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef thetas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByVal maxIter As Long) As Integer
   
   
   Public Declare PtrSafe Function NDK_ARMA_VALIDATE Lib "SFSDK.DLL" Alias "#606" (ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByRef phis As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef thetas As Double, _
                                                                               ByVal q As Long) As Integer
   'AirLine Function
   Public Declare PtrSafe Function NDK_AIRLINE_GOF Lib "SFSDK.DLL" Alias "#640" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal S As Integer, _
                                                                               ByVal theta As Double, _
                                                                               ByVal theta2 As Double, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_AIRLINE_RESID Lib "SFSDK.DLL" Alias "#641" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal S As Integer, _
                                                                               ByVal theta As Double, _
                                                                               ByVal theta2 As Double, _
                                                                               ByVal retType As Integer) As Integer
   Public Declare PtrSafe Function NDK_AIRLINE_FITTED Lib "SFSDK.DLL" Alias "#642" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal S As Integer, _
                                                                               ByVal theta As Double, _
                                                                               ByVal theta2 As Double, _
                                                                               ByVal retType As Integer) As Integer
   Public Declare PtrSafe Function NDK_AIRLINE_FORE Lib "SFSDK.DLL" Alias "#643" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal S As Integer, _
                                                                               ByVal theta As Double, _
                                                                               ByVal theta2 As Double, _
                                                                               ByVal nSteps As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByVal alpha As Double, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_AIRLINE_PARAM Lib "SFSDK.DLL" Alias "#645" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef mean As Double, _
                                                                               ByRef sigma As Double, _
                                                                               ByVal S As Integer, _
                                                                               ByRef theta As Double, _
                                                                               ByRef theta2 As Double, _
                                                                               ByVal retType As Integer, _
                                                                               ByVal maxIter As Long) As Integer
   Public Declare PtrSafe Function NDK_AIRLINE_VALIDATE Lib "SFSDK.DLL" Alias "#646" (ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal S As Integer, _
                                                                               ByVal theta As Double, _
                                                                               ByVal theta2 As Double) As Integer
   ' GARCH Function
   Public Declare PtrSafe Function NDK_GARCH_GOF Lib "SFSDK.DLL" Alias "#650" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mu As Double, _
                                                                               ByRef alphas As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef betas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nInnovationType As Integer, _
                                                                               ByVal nu As Double, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_GARCH_RESID Lib "SFSDK.DLL" Alias "#651" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mu As Double, _
                                                                               ByRef alphas As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef betas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nInnovationType As Integer, _
                                                                               ByVal nu As Double, _
                                                                               ByVal retType As Integer) As Integer
   Public Declare PtrSafe Function NDK_GARCH_FITTED Lib "SFSDK.DLL" Alias "#652" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mu As Double, _
                                                                               ByRef alphas As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef betas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nInnovationType As Integer, _
                                                                               ByVal nu As Double, _
                                                                               ByVal retType As Integer) As Integer
   Public Declare PtrSafe Function NDK_GARCH_FORE Lib "SFSDK.DLL" Alias "#653" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef sigmas As Double, _
                                                                               ByVal NSigmas As Long, _
                                                                               ByVal mu As Double, _
                                                                               ByRef alphas As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef betas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nInnovationType As Integer, _
                                                                               ByVal nu As Double, _
                                                                               ByVal nSteps As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByVal alpha As Double, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare PtrSafe Function NDK_GARCH_PARAM Lib "SFSDK.DLL" Alias "#655" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef mu As Double, _
                                                                               ByRef alphas As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef betas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nInnovationType As Integer, _
                                                                               ByRef nu As Double, _
                                                                               ByVal retType As Integer, _
                                                                               ByVal maxIter As Long) As Integer
   Public Declare PtrSafe Function NDK_GARCH_VALIDATE Lib "SFSDK.DLL" Alias "#656" (ByVal mu As Double, _
                                                                               ByRef alphas As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef betas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nInnovationType As Integer, _
                                                                               ByVal nu As Double) As Integer
   Public Declare PtrSafe Function NDK_GARCH_LRVAR Lib "SFSDK.DLL" Alias "#657" (ByVal mu As Double, _
                                                                               ByRef alphas As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef betas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nInnovationType As Integer, _
                                                                               ByVal nu As Double, _
                                                                               ByRef retVal As Double) As Integer
                                                                               
                                                                                   
                                                                               
                                                                               
                                                                               
                                                                               
   ' EGARCH Function
   ''' Computes the log-likelihood ((LLF), Akaike Information Criterion (AIC) or other goodness of fit function of the GARCH model. More...
   Public Declare PtrSafe Function NDK_EGARCH_GOF Lib "SFSDK.DLL" Alias "#660" (ByRef pData As Double, ByVal nSize As Long, ByVal mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                       ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, ByVal nInnovationType As Integer, _
                                                                                       ByVal nu As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
   
   
   ''' Returns an array of cells for the standardized residuals of a given GARCH model. More...
   Public Declare PtrSafe Function NDK_EGARCH_RESID Lib "SFSDK.DLL" Alias "#661" (ByRef pData As Double, ByVal nSize As Long, ByVal mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                       ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, ByVal nInnovationType As Integer, _
                                                                                       ByVal nu As Double, ByVal retType As Integer) As Long
   
   
   ''' Returns an array of cells for the initial (non-optimal), optimal or standard errors of the model's parameters. More...
   Public Declare PtrSafe Function NDK_EGARCH_PARAM Lib "SFSDK.DLL" Alias "#665" (ByRef pData As Double, ByVal nSize As Long, ByRef mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                       ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, ByVal nInnovationType As Integer, _
                                                                                       ByRef nu As Double, ByVal retType As Integer, ByVal maxIter As Long) As Long
   
    
   ''' Returns a simulated data series the underlying EGARCH process. More...
   Public Declare PtrSafe Function NDK_EGARCH_SIM Lib "SFSDK.DLL" Alias "#664" (ByRef pData As Double, ByVal nSize As Long, ByVal sigmas As Double, ByVal nSigmasSize As Long, _
                                                                                     ByVal mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                     ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                     ByVal nInnovationType As Integer, ByVal nu As Double, _
                                                                                     ByVal nSteps As Long, ByVal seed As Long, ByRef retVal As Double) As Long
   
    
   ''' Calculates the out-of-sample forecast statistics. More...
   Public Declare PtrSafe Function NDK_EGARCH_FORE Lib "SFSDK.DLL" Alias "#663" (ByRef pData As Double, ByVal nSize As Long, ByVal sigmas As Double, ByVal nSigmasSize As Long, _
                                                                                     ByVal mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                     ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                     ByVal nInnovationType As Integer, ByVal nu As Double, _
                                                                                     ByVal nSteps As Long, ByVal retType As Integer, ByVal alpha As Double, ByRef retVal As Double) As Long
   
    
   '''Returns an array of cells for the fitted values (i.e. mean, volatility and residuals) More...
   Public Declare PtrSafe Function NDK_EGARCH_FITTED Lib "SFSDK.DLL" Alias "#662" (ByRef pData As Double, ByVal nSize As Long, ByVal mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                     ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                     ByVal nInnovationType As Integer, ByVal nu As Double, ByVal retType As Integer) As Long
   
    
   '''   Calculates the long-run average volatility for a given E-GARCH model. More...
   Public Declare PtrSafe Function NDK_EGARCH_LRVAR Lib "SFSDK.DLL" Alias "#667" (ByVal mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                     ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                     ByVal nInnovationType As Integer, ByVal nu As Double, ByRef retVal As Double) As Long
   
    
   '''   Examines the model's parameters for stability constraints (e.g. stationary, positive variance, etc.). More...
   Public Declare PtrSafe Function NDK_EGARCH_VALIDATE Lib "SFSDK.DLL" Alias "#666" (ByVal mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                     ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                     ByVal nInnovationType As Integer, ByVal nu As Double) As Long
   
   ' GARCH M
   ''' Computes the log-likelihood ((LLF), Akaike Information Criterion (AIC) or other goodness of fit function of the GARCH model. More...
   Public Declare PtrSafe Function NDK_GARCHM_GOF Lib "SFSDK.DLL" Alias "#670" (ByRef pData As Double, ByVal nSize As Long, ByVal mu As Double, ByVal lambda As Double, _
                                                                                       ByRef alphas As Double, ByVal p As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                       ByVal nInnovationType As Integer, ByVal nu As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
   ''' Returns an array of cells for the standardized residuals of a given GARCH model. More...
   Public Declare PtrSafe Function NDK_GARCHM_RESID Lib "SFSDK.DLL" Alias "#671" (ByRef pData As Double, ByVal nSize As Long, ByVal mu As Double, ByVal lambda As Double, _
                                                                                       ByRef alphas As Double, ByVal p As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                       ByVal nInnovationType As Integer, ByVal nu As Double, ByVal retType As Integer) As Long
   ''' Returns an array of cells for the initial (non-optimal), optimal or standard errors of the model's parameters. More...
   Public Declare PtrSafe Function NDK_GARCHM_PARAM Lib "SFSDK.DLL" Alias "#675" (ByRef pData As Double, ByVal nSize As Long, ByRef mu As Double, ByRef lambda As Double, _
                                                                                       ByRef alphas As Double, ByVal p As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                       ByVal nInnovationType As Integer, ByRef nu As Double, ByVal retType As Integer, ByVal maxIter As Long) As Long
   ''' Returns a simulated data series the underlying EGARCH process. More...
   Public Declare PtrSafe Function NDK_GARCHM_SIM Lib "SFSDK.DLL" Alias "#674" (ByRef pData As Double, ByVal nSize As Long, ByVal sigmas As Double, ByVal nSigmasSize As Long, _
                                                                                     ByVal mu As Double, ByVal lambda As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                     ByRef betas As Double, ByVal q As Long, _
                                                                                     ByVal nInnovationType As Integer, ByVal nu As Double, _
                                                                                     ByVal nSteps As Long, ByVal seed As Long, ByRef retVal As Double) As Long
   ''' Calculates the out-of-sample forecast statistics. More...
   Public Declare PtrSafe Function NDK_GARCHM_FORE Lib "SFSDK.DLL" Alias "#673" (ByRef pData As Double, ByVal nSize As Long, ByVal sigmas As Double, ByVal nSigmasSize As Long, _
                                                                                     ByVal mu As Double, ByVal lambda As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                     ByRef betas As Double, ByVal q As Long, _
                                                                                     ByVal nInnovationType As Integer, ByVal nu As Double, _
                                                                                     ByVal nSteps As Long, ByVal retType As Integer, ByVal alpha As Double, ByRef retVal As Double) As Long
   ''' Returns an array of cells for the fitted values (i.e. mean, volatility and residuals) More...
   Public Declare PtrSafe Function NDK_GARCHM_FITTED Lib "SFSDK.DLL" Alias "#672" (ByRef pData As Double, ByVal nSize As Long, ByVal mu As Double, ByVal lambda As Double, _
                                                                                       ByRef alphas As Double, ByVal p As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                       ByVal nInnovationType As Integer, ByVal nu As Double, ByVal retType As Integer) As Long
   ''' Calculates the long-run average volatility for the given GARCH-M model. More..
   Public Declare PtrSafe Function NDK_GARCHM_LRVAR Lib "SFSDK.DLL" Alias "#677" (ByVal mu As Double, ByVal lambda As Double, _
                                                                                         ByRef alphas As Double, ByVal p As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                         ByVal nInnovationType As Integer, ByVal nu As Double, ByRef retVal As Double) As Long
   ''' Calculates the long-run average volatility for the given GARCH-M model. More..
   Public Declare PtrSafe Function NDK_GARCHM_VALIDATE Lib "SFSDK.DLL" Alias "#676" (ByVal mu As Double, ByVal lambda As Double, _
                                                                                         ByRef alphas As Double, ByVal p As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                         ByVal nInnovationType As Integer, ByVal nu As Double) As Long
    
   ''' Gneralized Linear Model Functions
   ''' Examines the model's parameters for constraints (e.g. positive variance, etc.) More...
   Public Declare PtrSafe Function NDK_GLM_VALIDATE Lib "SFSDK.DLL" Alias "#715" (ByRef betas As Double, ByVal nBetas As Long, ByVal phi As Double, ByVal Lvk As Integer) As Long
   ''' Computes the log-likelihood ((LLF), Akaike Information Criterion (AIC) or other goodness of fit function of the GLM model. More...
   Public Declare PtrSafe Function NDK_GLM_GOF Lib "SFSDK.DLL" Alias "#710" (ByRef Y As Double, ByVal nSize As Long, ByRef X As Double, ByVal nVars As Long, _
                                                                                   ByRef betas As Double, ByVal nBetas As Long, ByVal phi As Double, ByVal Lvk As Integer, _
                                                                                   ByVal retType As Integer, ByRef retVal As Double) As Long
   ''' Returns the standardized residuals/errors of a given GLM. More...
   Public Declare PtrSafe Function NDK_GLM_RESID Lib "SFSDK.DLL" Alias "#711" (ByRef Y As Double, ByVal nSize As Long, ByRef X As Double, ByVal nVars As Long, _
                                                                                   ByRef betas As Double, ByVal nBetas As Long, ByRef phi As Double, ByVal Lvk As Integer, _
                                                                                   ByVal retType As Integer) As Long
   ''' Returns an array of cells for the initial (non-optimal), optimal or standard errors of the model's parameters. More...
   Public Declare PtrSafe Function NDK_GLM_PARAM Lib "SFSDK.DLL" Alias "#714" (ByRef Y As Double, ByVal nSize As Long, ByRef X As Double, ByVal nVars As Long, _
                                                                                   ByRef betas As Double, ByVal nBetas As Long, ByVal phi As Double, ByVal Lvk As Integer, _
                                                                                   ByVal retType As Integer, ByVal maxIter As Long) As Long
   ''' calculates the expected response (i.e. mean) value; given the GLM model and the values of the explanatory variables.
   Public Declare PtrSafe Function NDK_GLM_FORE Lib "SFSDK.DLL" Alias "#713" (ByRef X As Double, ByVal nVars As Long, _
                                                                                   ByRef betas As Double, ByVal nBetas As Long, ByVal phi As Double, ByVal Lvk As Integer, _
                                                                                   ByVal retType As Integer, ByVal alpha As Double, ByRef retVal As Double) As Long
   ''' Returns the standardized residuals/errors of a given GLM. More...
   Public Declare PtrSafe Function NDK_GLM_FITTED Lib "SFSDK.DLL" Alias "#712" (ByRef Y As Double, ByVal nSize As Long, ByRef X As Double, ByVal nVars As Long, _
                                                                                   ByRef betas As Double, ByVal nBetas As Long, ByRef phi As Double, ByVal Lvk As Integer, _
                                                                                   ByVal retType As Integer) As Long
    
   
   
   
   ''' Multiple Linear Regression (MLR)
   
   ''' Returns the standardized residuals/errors of a given GLM. More...
   Public Declare PtrSafe Function NDK_SLR_PARAM Lib "SFSDK.DLL" Alias "#720" (ByRef X As Double, ByVal nXSize As Long, ByRef Y As Double, ByVal nYSize As Long, _
                                                                                     ByVal intercept As Double, ByVal alpha As Double, ByVal retType As Integer, _
                                                                                     ByVal ParamIndex As Integer, ByRef retVal As Double) As Long
   
   Public Declare PtrSafe Function NDK_SLR_FORE Lib "SFSDK.DLL" Alias "#721" (ByRef X As Double, ByVal nXSize As Long, ByRef Y As Double, ByVal nYSize As Long, _
                                                                                     ByVal intercept As Double, ByVal target As Double, ByVal alphas As Double, _
                                                                                     ByVal retType As Integer, ByRef retVal As Double) As Long
   
   Public Declare PtrSafe Function NDK_SLR_FITTED Lib "SFSDK.DLL" Alias "#722" (ByRef X As Double, ByVal nXSize As Long, ByRef Y As Double, ByVal nYSize As Long, _
                                                                                     ByVal intercept As Double, ByVal retType As Integer) As Long
   
   
   Public Declare PtrSafe Function NDK_SLR_ANOVA Lib "SFSDK.DLL" Alias "#723" (ByRef X As Double, ByVal nXSize As Long, ByRef Y As Double, ByVal nYSize As Long, _
                                                                                     ByVal intercept As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
   
   
   Public Declare PtrSafe Function NDK_SLR_GOF Lib "SFSDK.DLL" Alias "#724" (ByRef X As Double, ByVal nXSize As Long, ByRef Y As Double, ByVal nYSize As Long, _
                                                                                     ByVal intercept As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
    
   
   Public Declare PtrSafe Function NDK_MLR_PARAM Lib "SFSDK.DLL" Alias "#730" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByRef Y As Double, ByVal nYSize As Long, ByVal intercept As Double, _
                                                                                 ByVal alphas As Double, ByVal retType As Integer, ByVal nParamIndex As Integer, ByRef retVal As Double) As Long
    
   Public Declare PtrSafe Function NDK_MLR_FORE Lib "SFSDK.DLL" Alias "#731" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByRef Y As Double, ByVal nYSize As Long, ByVal intercept As Double, ByRef target As Double, _
                                                                                 ByVal alphas As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
   
    
   Public Declare PtrSafe Function NDK_MLR_FITTED Lib "SFSDK.DLL" Alias "#732" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByRef Y As Double, ByVal nYSize As Long, ByVal intercept As Double, ByVal retType As Integer) As Long
   
    
   Public Declare PtrSafe Function NDK_MLR_ANOVA Lib "SFSDK.DLL" Alias "#733" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByRef Y As Double, ByVal nYSize As Long, ByVal intercept As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
   
   Public Declare PtrSafe Function NDK_MLR_GOF Lib "SFSDK.DLL" Alias "#734" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByRef Y As Double, ByVal nYSize As Long, ByVal intercept As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
    
   Public Declare PtrSafe Function NDK_MLR_PRFTest Lib "SFSDK.DLL" Alias "#736" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef Y As Double, ByVal nYSize As Long, _
                                                                                    ByVal intercept As Double, ByRef mask1 As Byte, ByVal nMaskLen1 As Long, _
                                                                                    ByRef mask2 As Byte, ByVal nMaskLen2 As Long, _
                                                                                    ByVal alpha As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
    
   Public Declare PtrSafe Function NDK_MLR_STEPWISE Lib "SFSDK.DLL" Alias "#735" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByRef Y As Double, ByVal nYSize As Long, ByVal intercept As Double, ByVal alpha As Double, ByVal mode As Integer) As Long



   ''' PCA
   Public Declare PtrSafe Function NDK_PCA_COMP Lib "SFSDK.DLL" Alias "#740" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByVal standardize As Integer, ByVal nCompIndex As Integer, ByVal retType As Integer, _
                                                                                 ByRef retVal As Double, ByVal nOutSize As Long) As Long
   Public Declare PtrSafe Function NDK_PCA_VAR Lib "SFSDK.DLL" Alias "#741" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByVal standardize As Integer, ByVal nVarIndex As Integer, ByVal MaxPC As Integer, ByVal retType As Integer, _
                                                                                 ByRef retVal As Double, ByVal nOutSize As Long) As Long
   Public Declare PtrSafe Function NDK_PCR_PARAM Lib "SFSDK.DLL" Alias "#742" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                  ByRef pYData As Double, ByVal nYSize As Long, ByVal intercept As Double, ByVal alpha As Double, _
                                                                                  ByVal retType As Integer, ByVal nParamIndex, ByRef retVal As Double) As Long
   Public Declare PtrSafe Function NDK_PCR_FORE Lib "SFSDK.DLL" Alias "#743" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                  ByRef pYData As Double, ByVal nYSize As Long, ByVal intercept As Double, ByRef target As Double, _
                                                                                  ByVal alpha As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
   Public Declare PtrSafe Function NDK_PCR_FITTED Lib "SFSDK.DLL" Alias "#744" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                  ByRef pYData As Double, ByVal nYSize As Long, ByVal intercept As Double, _
                                                                                  ByVal retType As Integer) As Long
   Public Declare PtrSafe Function NDK_PCR_ANOVA Lib "SFSDK.DLL" Alias "#745" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                  ByRef pYData As Double, ByVal nYSize As Long, ByVal intercept As Double, _
                                                                                  ByVal retType As Integer, ByRef retVal As Double) As Long
   Public Declare PtrSafe Function NDK_PCR_GOF Lib "SFSDK.DLL" Alias "#746" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                  ByRef pYData As Double, ByVal nYSize As Long, ByVal intercept As Double, _
                                                                                  ByVal retType As Integer, ByRef retVal As Double) As Long
   Public Declare PtrSafe Function NDK_PCR_PRFTest Lib "SFSDK.DLL" Alias "#748" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef pYData As Double, ByVal nYSize As Long, _
                                                                                    ByVal intercept As Double, _
                                                                                    ByRef mask1 As Byte, ByVal nMaskLen1 As Long, ByRef mask2 As Byte, ByVal nMaskLen2 As Long, _
                                                                                    ByVal alpha As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
   Public Declare PtrSafe Function NDK_PCR_STEPWISE Lib "SFSDK.DLL" Alias "#747" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                  ByRef pYData As Double, ByVal nYSize As Long, ByVal intercept As Double, _
                                                                                  ByVal alpha As Double, ByVal nMode As Integer) As Long
    
    
   
   ''' Transform
   '''Returns an array of cells for the (backward shifted, backshifted or lagged time series. More...
   Public Declare PtrSafe Function NDK_LAG Lib "SFSDK.DLL" Alias "#1000" (ByRef X As Double, ByVal N As Long, ByVal K As Long) As Long
   
   '''Returns an array of cells for the differenced time series (i.e. (1-L^S)^D). . More..
   Public Declare PtrSafe Function NDK_DIFF Lib "SFSDK.DLL" Alias "#1005" (ByRef X As Double, ByVal N As Long, ByVal S As Long, ByVal D As Long) As Long
   
   '''Returns an array of cells for the integrated time series (inverse operator of NDK_DIFF). . More...
   Public Declare PtrSafe Function NDK_INTEG Lib "SFSDK.DLL" Alias "#1010" (ByRef X As Double, ByVal N As Long, ByVal S As Long, ByVal D As Long, ByRef X0 As Double, ByVal N0 As Long) As Long
   
   '''Returns an array of cells of a time series after removing all missing values. More...
   Public Declare PtrSafe Function NDK_RMNA Lib "SFSDK.DLL" Alias "#4001" (ByRef X As Double, ByVal N As Long) As Long
    
   '''Returns the time-reversed order time series (i.e. the first observation is swapped with the last observation, etc.): both missing and non-missing values. More.
   Public Declare PtrSafe Function NDK_REVERSE Lib "SFSDK.DLL" Alias "#1024" (ByRef X As Double, ByVal N As Long) As Long
    
   '''Returns an array of cells for the scaled time series. More...
   Public Declare PtrSafe Function NDK_SCALE Lib "SFSDK.DLL" Alias "#1023" (ByRef X As Double, ByVal N As Long, ByVal K As Double) As Long
    
   '''Returns an array of the difference between two time series. More...
   Public Declare PtrSafe Function NDK_SUB Lib "SFSDK.DLL" Alias "#1022" (ByRef X1 As Double, ByVal N1 As Long, ByVal X2 As Double, ByVal N2 As Long) As Long
    
   '''Returns an array of the difference between two time series. More...
   Public Declare PtrSafe Function NDK_ADD Lib "SFSDK.DLL" Alias "#1021" (ByRef X1 As Double, ByVal N1 As Long, ByVal X2 As Double, ByVal N2 As Long) As Long
    
   '''Computes the complementary log-log transformation, including its inverse. More...
   Public Declare PtrSafe Function NDK_CLOGLOG Lib "SFSDK.DLL" Alias "#4005" (ByRef X As Double, ByVal N As Long, ByVal retType As Integer) As Long
    
   '''Computes the probit transformation, including its inverse. More..
   Public Declare PtrSafe Function NDK_PROBIT Lib "SFSDK.DLL" Alias "#4004" (ByRef X As Double, ByVal N As Long, ByVal retType As Integer) As Long
    
   '''Computes the complementary log-log transformation, including its inverse. More...
   Public Declare PtrSafe Function NDK_LOGIT Lib "SFSDK.DLL" Alias "#4003" (ByRef X As Double, ByVal N As Long, ByVal retType As Integer) As Long
   
   '''Computes the complementary Box-Cox transformation, including its inverse. More...
   Public Declare PtrSafe Function NDK_BOXCOX Lib "SFSDK.DLL" Alias "#4002" (ByRef X As Double, ByVal N As Long, ByRef lambda As Double, ByRef alpha As Double, _
                                                                                   ByVal retType As Integer, ByRef retVal As Double) As Long
    
   '''Detrends a time series using a regression of y against a polynomial time trend of order p. More...
   Public Declare PtrSafe Function NDK_DETREND Lib "SFSDK.DLL" Alias "#4010" (ByRef X As Double, ByVal N As Long, ByVal PolyOrder As Integer) As Long
    
   '''Returns an array of the deseasonalized time series assuming a linear model. More...
   Public Declare PtrSafe Function NDK_RMSEASONAL Lib "SFSDK.DLL" Alias "#4017" (ByRef X As Double, ByVal N As Long, ByVal period As Long) As Long
   
   '''Returns an array of cells of a time series after substituting all missing values with the mean/median. More...
   Public Declare PtrSafe Function NDK_INTERP_NAN Lib "SFSDK.DLL" Alias "#4000" (ByRef X As Double, ByVal N As Long, ByVal nMethod As Integer, ByVal plug As Double) As Long
    
   '''Examine whether the given array has one/more missing values. More..
   Public Declare PtrSafe Function NDK_HASNA Lib "SFSDK.DLL" Alias "#4018" (ByRef X As Double, ByVal N As Long, ByVal intermediate As Boolean) As Long
   


   ''' Spectral Analysis
   ''' Returns an array of cells for the convolution operator of two time series. More...
   Public Declare PtrSafe Function NDK_CONVOLUTION Lib "SFSDK.DLL" Alias "#1032" (ByRef X1 As Double, ByVal N1 As Long, ByRef X2 As Double, ByVal N2 As Long, ByRef Z As Double, ByVal nZSize As Long) As Long
   
   
   ''' Calculates the inverse discrete fast Fourier transformation, recovering the time series. More...
   Public Declare PtrSafe Function NDK_IDFT Lib "SFSDK.DLL" Alias "#1031" (ByRef Amp As Double, ByRef Phase As Double, ByVal nSize As Long, ByRef X As Double, ByVal nXSize As Long) As Long
   
   
   ''' Calculates the discrete fast Fourier transformation for amplitude and phase. More...
   Public Declare PtrSafe Function NDK_DFT Lib "SFSDK.DLL" Alias "#1030" (ByRef X As Double, ByVal nXSize As Long, ByRef Amp As Double, ByRef Phase As Double, ByVal nSize As Long) As Long
   
   ''' computes cyclical component of given time series using the Hodrick?Prescott filter. More...
   Public Declare PtrSafe Function NDK_HodrickPrescotFilter Lib "SFSDK.DLL" Alias "#1033" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, ByVal lambda As Double) As Long
    
    
   ''' Computes trend and cyclical component of a macroeconomic time series using Baxter-King Fixed Length Symmetric Filter. More...
   Public Declare PtrSafe Function NDK_BaxterKingFilter Lib "SFSDK.DLL" Alias "#1034" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, _
                                                                                             ByVal period_min As Long, ByVal period_max As Long, ByVal K As Long, ByVal drift As Boolean, _
                                                                                             ByVal unitroot As Boolean, ByVal retType As Integer) As Long
   
   ''' Smoothing API functions calls
   '''Returns the weighted moving (rolling/running) average using the previous m data points. More...
   Public Declare PtrSafe Function NDK_WMA Lib "SFSDK.DLL" Alias "#2000" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, _
                                                                                ByRef weights As Double, ByVal nwSize As Long, ByVal nHorizon As Long, ByRef retVal As Double) As Long
   
   '''Returns the (Brown's) simple exponential (EMA) smoothing estimate of the value of X at time t+m (based on the raw data up to time t)..
   Public Declare PtrSafe Function NDK_SESMTH Lib "SFSDK.DLL" Alias "#2005" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, _
                                                                                ByRef alpha As Double, ByVal nHorizon As Long, ByVal optimize As Boolean, ByRef retVal As Double) As Long
   
   '''Returns the (Holt-Winter's) double exponential smoothing estimate of the value of X at time T+m.
   Public Declare PtrSafe Function NDK_DESMTH Lib "SFSDK.DLL" Alias "#2010" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, _
                                                                                 ByRef alpha As Double, ByRef beta As Double, ByVal nHorizon As Long, ByVal optimize As Boolean, _
                                                                                 ByRef retVal As Double) As Long
   
   '''Returns the (Brown's) Linear exponential smoothing estimate of the value of X at time T+m (based on the raw data up to time t).
   Public Declare PtrSafe Function NDK_LESMTH Lib "SFSDK.DLL" Alias "#2015" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, _
                                                                                ByRef alpha As Double, ByVal nHorizon As Long, ByVal optimize As Boolean, ByRef retVal As Double) As Long
    
   '''Returns the (Winters's) triple exponential smoothing estimate of the value of X at time T+m. More...
   Public Declare PtrSafe Function NDK_TESMTH Lib "SFSDK.DLL" Alias "#2020" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, _
                                                                                   ByRef alpha As Double, ByRef beta As Double, ByRef gamma As Double, ByVal S As Long, _
                                                                                   ByVal nHorizon As Long, ByVal optimize As Boolean, ByRef retVal As Double) As Long
   
   '''Returns values along a trend curve (e.g. linear, quadratic, exponential, etc.) at time T+m..
   Public Declare PtrSafe Function NDK_TREND Lib "SFSDK.DLL" Alias "#2021" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, _
                                                                                  ByRef trendType As Integer, ByVal PolyOrder As Integer, ByVal allowIntercept As Boolean, ByVal interecept As Double, _
                                                                                  ByVal nHorizon As Long, ByVal retType As Integer, ByVal alpha As Double, ByRef retVal As Double) As Long
    



   ''' Utilities
   '''estimate the value of the function represented by (x,y) data set at an intermediate x-value. More...
   Public Declare PtrSafe Function NDK_INTERPOLATE Lib "SFSDK.DLL" Alias "#3000" (ByRef X As Double, ByVal nX As Long, ByRef Y As Double, ByVal nY As Long, _
                                                                                        ByRef XT As Double, ByVal nXT As Long, ByVal uMethod As Integer, ByVal extrapolate As Boolean, _
                                                                                        ByRef YVal As Double, ByVal nYVals As Long) As Long
   
   
   '''Locate and return the full path of the default editor (e.g. notepad) in the system. More...
   Public Declare PtrSafe Function NDK_DEFAULT_EDITOR Lib "SFSDK.DLL" Alias "#3025" (ByVal szFullPath As String, ByRef nSize As Long) As Long
   
   
   '''Returns the n-th token/substring in a string after splitting it using a given delimiter.
   Public Declare PtrSafe Function NDK_TOKENIZE Lib "SFSDK.DLL" Alias "#3020" (ByVal szTxt As String, ByVal szDelim As String, ByVal nOrder As Integer, ByVal szRetVal As String, _
                                                                                      ByVal nSize As Long) As Long
    
   '''Returns TRUE if the string matches the regular expression expressed.
   Public Declare PtrSafe Function NDK_REGEX_MATCH Lib "SFSDK.DLL" Alias "#3010" (ByVal szLine As String, ByVal szPattern As String, ByVal ignoreCase As Boolean, _
                                                                                         ByVal partialOK As Boolean, ByRef bRetVal As Boolean) As Long
   
    
   ''' Returns TRUE if the string matches the regular expression expressed.
   Public Declare PtrSafe Function NDK_REGEX_REPLACE Lib "SFSDK.DLL" Alias "#3015" (ByVal szLine As String, ByVal szKey As String, ByVal szValue As String, ByVal ignoreCase As Boolean, ByVal bGlobal As Boolean, ByVal szRetVal As String, ByVal nSize As Long) As Long
    
   
   '''calculates the value of the regression function for an intermediate x-value.
   Public Declare PtrSafe Function NDK_REGRESSION Lib "SFSDK.DLL" Alias "#3005" (ByRef X As Double, ByVal nXSize As Long, ByRef Y As Double, ByVal nYSize As Long, _
                                                                                       ByVal nType As Integer, ByVal nPolyOrder As Integer, ByVal intercept As Double, _
                                                                                       ByVal target As Double, ByVal retType As Integer, ByVal alpha As Double, _
                                                                                       ByRef retVal As Double) As Long



   '''Seasonal ajustments using X12-ARIMA API functions calls
   
   ''' Prepare the X12-ARIMA scripting environment
   Public Declare PtrSafe Function NDK_X12_ENV_INIT Lib "SFSDK.DLL" Alias "#5000" (ByVal init As Boolean) As Long
   ''' Cleanup all files created by ARIMA program
   Public Declare PtrSafe Function NDK_X12_ENV_CLEANUP Lib "SFSDK.DLL" Alias "#5002" () As Long
   ''' Prepare the X12 Model
   Public Declare PtrSafe Function NDK_X12_SCEN_INIT Lib "SFSDK.DLL" Alias "#5005" (ByVal szScenarioName As String, ByRef X12Options As Any) As Long
   ''' cleanup all temp files
   Public Declare PtrSafe Function NDK_X12_SCEN_CLEAUP Lib "SFSDK.DLL" Alias "#5007" (ByVal szScenarioName As String) As Long
   ''' Write the data to the disk
   Public Declare PtrSafe Function NDK_X12_DATA_FILE Lib "SFSDK.DLL" Alias "#5010" (ByVal szScenarioName As String, ByRef data As Double, ByVal nlen As Long, _
                                                                                           ByVal monthly As Boolean, ByVal startDate As Long, ByVal fileType As Integer) As Long
   ''' Write teh SPC file to the disk
   Public Declare PtrSafe Function NDK_X12_SPC_FILE Lib "SFSDK.DLL" Alias "#5015" (ByVal szScenarioName As String, ByRef X12Options As Any) As Long
   ''' Run any batch program
   Public Declare PtrSafe Function NDK_X12_RUN_BATCH Lib "SFSDK.DLL" Alias "#5020" (ByVal szScenarioName As String, ByVal batchFilename As String, ByRef status As Integer) As Long
   ''' Run the scenario selected
   Public Declare PtrSafe Function NDK_X12_RUN_SCENARIO Lib "SFSDK.DLL" Alias "#5022" (ByVal szScenarioName As String, ByRef status As Integer) As Long
   ''' Examine the status of running X12a program
   Public Declare PtrSafe Function NDK_X12_RUN_STAT Lib "SFSDK.DLL" Alias "#5025" (ByVal szScenarioName As String, ByRef status As Integer, ByVal szMsg As String, ByRef nlen As Long) As Long
   ''' return the full name of the x12a output file
   Public Declare PtrSafe Function NDK_X12_OUT_FILE Lib "SFSDK.DLL" Alias "#5030" (ByVal szScenarioName As String, ByVal retType As Integer, ByVal szOutFile As String, _
                                                                                           ByRef nlen As Long, ByVal OpenFileFlag As Boolean) As Long
   ''' return the output time series (seasonally adjusted)
   Public Declare PtrSafe Function NDK_X12_OUT_SERIES Lib "SFSDK.DLL" Alias "#5035" (ByVal szScenarioName As String, ByVal nComponent As Integer, ByRef pData As Double, ByRef nSize As Long) As Long
   ''' return the forecast output time series
   Public Declare PtrSafe Function NDK_X12_FORE_SERIES Lib "SFSDK.DLL" Alias "#5040" (ByVal szScenarioName As String, ByVal nStep As Long, ByVal retType As Integer, ByRef retVal As Double) As Long
   
   
   
   '''Portfolio Analysis
   '''Compute the portfolio equivalent returns
   Public Declare PtrSafe Function NDK_PORTFOLIO_RET Lib "SFSDK.DLL" Alias "#5500" (ByRef weights As Double, ByVal nAssets As Long, ByRef returns As Double, ByRef retVal As Double) As Long
   '''Calculates the overall portfolio variance (volatility squared)
   Public Declare PtrSafe Function NDK_PORTFOLIO_VARIANCE Lib "SFSDK.DLL" Alias "#5502" (ByRef weights As Double, ByVal nAssets As Long, ByRef covar As Double, ByRef retVal As Double) As Long
   ''' Calculates the covariance between two portfolios
   Public Declare PtrSafe Function NDK_PORTFOLIO_COVARIANCE Lib "SFSDK.DLL" Alias "#5504" (ByRef weights1 As Double, ByRef weights2 As Double, ByVal nAssets As Long, ByRef covar As Double, ByRef retVal As Double) As Long
 #Else
   Public Declare Function NDK_Init Lib "DllName" Alias "#100" (ByVal szAppName As String, _
                                                                                   ByVal szKey As String, _
                                                                                   ByVal szActivation As String, _
                                                                                   ByVal szLogDirectory As String) As Integer
   Public Declare Function NDK_Shutdown Lib "SFSDK.DLL" Alias "#105" () As Integer
   Public Declare Function NDK_INFO Lib "SFSDK.DLL" Alias "#110" (ByVal nRetType As Integer, _
                                                                                 ByVal szMsd As String, _
                                                                                 ByVal nlen As Long) As Integer
   Public Declare Function NDK_MSG Lib "SFSDK.DLL" Alias "#115" (ByVal nRetCode As Integer, _
                                                                               ByVal szMsd As String, _
                                                                               ByVal nlen As Long) As Integer
   'Time Series statistics
   Public Declare Function NDK_ACF Lib "SFSDK.DLL" Alias "#200" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_ACF_ERROR Lib "SFSDK.DLL" Alias "#205" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_ACFCI Lib "SFSDK.DLL" Alias "#210" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal alpha As Double, _
                                                                               ByRef UL As Double, _
                                                                               ByRef LL As Double) As Integer
   Public Declare Function NDK_PACF Lib "SFSDK.DLL" Alias "#215" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_PACF_ERROR Lib "SFSDK.DLL" Alias "#220" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_PACFCI Lib "SFSDK.DLL" Alias "#225" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal alpha As Double, _
                                                                               ByRef UL As Double, _
                                                                               ByRef LL As Double) As Integer
   'Statistical Tests
   Public Declare Function NDK_ACFTEST Lib "SFSDK.DLL" Alias "#300" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal target As Double, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_NORMALTEST Lib "SFSDK.DLL" Alias "#301" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_WNTEST Lib "SFSDK.DLL" Alias "#302" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_ARCHTEST Lib "SFSDK.DLL" Alias "#303" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
                                                                               
   Public Declare Function NDK_MEANTEST Lib "SFSDK.DLL" Alias "#304" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal target As Double, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
                                                                               
   Public Declare Function NDK_STDEVTEST Lib "SFSDK.DLL" Alias "#305" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal target As Double, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   
   Public Declare Function NDK_SKEWTEST Lib "SFSDK.DLL" Alias "#306" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
 
   Public Declare Function NDK_XKURTTEST Lib "SFSDK.DLL" Alias "#307" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_XCFTEST Lib "SFSDK.DLL" Alias "#308" (ByRef X As Double, _
                                                                                   ByRef Y As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal trger As Double, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_ADFTEST Lib "SFSDK.DLL" Alias "#309" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal options As Integer, _
                                                                               ByVal testdown As Boolean, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal method As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   ''' Perform Johansen cointegration test
   Public Declare Function NDK_JOHANSENTEST Lib "SFSDK.DLL" Alias "#311" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal M As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal PolyOrder As Integer, _
                                                                               ByVal tracetest As Boolean, _
                                                                               ByVal R As Integer, _
                                                                               ByVal alpha As Double, _
                                                                               ByRef retStat As Double, _
                                                                               ByRef retCV As Double) As Integer
 
   Public Declare Function NDK_COLNRTY_TEST Lib "SFSDK.DLL" Alias "#312" (ByRef XX As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal M As Long, _
                                                                               ByRef mask As Byte, _
                                                                               ByVal nMaskLen As Long, _
                                                                               ByVal method As Integer, _
                                                                               ByVal colIndex As Integer, _
                                                                               ByRef retVal As Double) As Integer
                                                                               
   Public Declare Function NDK_CHOWTEST Lib "SFSDK.DLL" Alias "#313" (ByRef XX1 As Double, _
                                                                               ByVal M As Long, _
                                                                               ByRef Y1 As Double, _
                                                                               ByVal N1 As Long, _
                                                                               ByRef XX2 As Double, _
                                                                               ByRef Y2 As Double, _
                                                                               ByVal N2 As Long, _
                                                                               ByRef mask As Byte, _
                                                                               ByVal nMaskLen As Long, _
                                                                               ByVal intercept As Double, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   ' General statistics
   Public Declare Function NDK_GINI Lib "SFSDK.DLL" Alias "#400" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef retVal As Double) As Integer
                                                                               
   Public Declare Function NDK_XCF Lib "SFSDK.DLL" Alias "#401" (ByRef X As Double, _
                                                                               ByRef Y As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal K As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
                                                                               
   Public Declare Function NDK_XKURT Lib "SFSDK.DLL" Alias "#403" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_SKEW Lib "SFSDK.DLL" Alias "#404" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_AVERAGE Lib "SFSDK.DLL" Alias "#405" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_VARIANCE Lib "SFSDK.DLL" Alias "#406" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_MIN Lib "SFSDK.DLL" Alias "#407" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_MAX Lib "SFSDK.DLL" Alias "#408" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_HURST_EXPONENT Lib "SFSDK.DLL" Alias "#409" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_QUANTILE Lib "SFSDK.DLL" Alias "#410" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal p As Double, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_IQR Lib "SFSDK.DLL" Alias "#411" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_RMS Lib "SFSDK.DLL" Alias "#412" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_MD Lib "SFSDK.DLL" Alias "#413" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_RMD Lib "SFSDK.DLL" Alias "#414" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_MAD Lib "SFSDK.DLL" Alias "#415" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal reserved As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_LRVAR Lib "SFSDK.DLL" Alias "#416" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal W As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_SAD Lib "SFSDK.DLL" Alias "#417" (ByRef Y As Double, _
                                                                               ByRef YHat As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_MAE Lib "SFSDK.DLL" Alias "#418" (ByRef Y As Double, _
                                                                               ByRef YHat As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_MAPE Lib "SFSDK.DLL" Alias "#419" (ByRef Y As Double, _
                                                                               ByRef YHat As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_RMSE Lib "SFSDK.DLL" Alias "#420" (ByRef Y As Double, _
                                                                               ByRef YHat As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_SSE Lib "SFSDK.DLL" Alias "#421" (ByRef Y As Double, _
                                                                               ByRef YHat As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_SORT_ASC Lib "SFSDK.DLL" Alias "#422" (ByRef Y As Double, _
                                                                               ByVal N As Long) As Integer
                                                                               
 
   ' Statistical distribution
   Public Declare Function NDK_GED_XKURT Lib "SFSDK.DLL" Alias "#500" (ByVal df As Double, _
                                                                                     ByRef retVal As Double) As Integer
   Public Declare Function NDK_TDIST_XKURT Lib "SFSDK.DLL" Alias "#501" (ByVal df As Double, _
                                                                                     ByRef retVal As Double) As Integer
   Public Declare Function NDK_EDF Lib "SFSDK.DLL" Alias "#502" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal target As Double, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_HIST_BINS Lib "SFSDK.DLL" Alias "#503" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal method As Integer, _
                                                                               ByRef retVal As Long) As Integer
   Public Declare Function NDK_HIST_BIN_LIMIT Lib "SFSDK.DLL" Alias "#504" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal nBins As Long, _
                                                                               ByVal index As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_HISTOGRAM Lib "SFSDK.DLL" Alias "#505" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal nBins As Long, _
                                                                               ByVal index As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_KERNEL_DENSITY_ESTIMATE Lib "SFSDK.DLL" Alias "#506" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal target As Double, _
                                                                               ByVal bandwidth As Double, _
                                                                               ByVal argKernelFunc As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_GAUSS_FORECI Lib "SFSDK.DLL" Alias "#507" (ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal upper As Boolean, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_TSTUDENT_FORECI Lib "SFSDK.DLL" Alias "#508" (ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal df As Double, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal upper As Boolean, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_GED_FORECI Lib "SFSDK.DLL" Alias "#509" (ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal df As Double, _
                                                                               ByVal alpha As Double, _
                                                                               ByVal upper As Boolean, _
                                                                               ByRef retVal As Double) As Integer
   'ARMA Function
 Public Declare Function NDK_ARMA_GOF Lib "SFSDK.DLL" Alias "#600" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByRef phis As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef thetas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
                                                                               
   
   Public Declare Function NDK_ARMA_RESID Lib "SFSDK.DLL" Alias "#601" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByRef phis As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef thetas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal retType As Integer) As Integer
   Public Declare Function NDK_ARMA_FITTED Lib "SFSDK.DLL" Alias "#602" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByRef phis As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef thetas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal retType As Integer) As Integer
   Public Declare Function NDK_ARMA_FORE Lib "SFSDK.DLL" Alias "#603" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByRef phis As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef thetas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nSteps As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByVal alpha As Double, _
                                                                               ByRef retVal As Double) As Integer
  
   Public Declare Function NDK_ARMA_PARAM Lib "SFSDK.DLL" Alias "#605" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef avg As Double, _
                                                                               ByRef sigma As Double, _
                                                                               ByRef phis As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef thetas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByVal maxIter As Long) As Integer
   
   
   Public Declare Function NDK_ARMA_VALIDATE Lib "SFSDK.DLL" Alias "#606" (ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByRef phis As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef thetas As Double, _
                                                                               ByVal q As Long) As Integer
   'AirLine Function
   Public Declare Function NDK_AIRLINE_GOF Lib "SFSDK.DLL" Alias "#640" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal S As Integer, _
                                                                               ByVal theta As Double, _
                                                                               ByVal theta2 As Double, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_AIRLINE_RESID Lib "SFSDK.DLL" Alias "#641" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal S As Integer, _
                                                                               ByVal theta As Double, _
                                                                               ByVal theta2 As Double, _
                                                                               ByVal retType As Integer) As Integer
   Public Declare Function NDK_AIRLINE_FITTED Lib "SFSDK.DLL" Alias "#642" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal S As Integer, _
                                                                               ByVal theta As Double, _
                                                                               ByVal theta2 As Double, _
                                                                               ByVal retType As Integer) As Integer
   Public Declare Function NDK_AIRLINE_FORE Lib "SFSDK.DLL" Alias "#643" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal S As Integer, _
                                                                               ByVal theta As Double, _
                                                                               ByVal theta2 As Double, _
                                                                               ByVal nSteps As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByVal alpha As Double, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_AIRLINE_PARAM Lib "SFSDK.DLL" Alias "#645" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef mean As Double, _
                                                                               ByRef sigma As Double, _
                                                                               ByVal S As Integer, _
                                                                               ByRef theta As Double, _
                                                                               ByRef theta2 As Double, _
                                                                               ByVal retType As Integer, _
                                                                               ByVal maxIter As Long) As Integer
   Public Declare Function NDK_AIRLINE_VALIDATE Lib "SFSDK.DLL" Alias "#646" (ByVal mean As Double, _
                                                                               ByVal sigma As Double, _
                                                                               ByVal S As Integer, _
                                                                               ByVal theta As Double, _
                                                                               ByVal theta2 As Double) As Integer
   ' GARCH Function
   Public Declare Function NDK_GARCH_GOF Lib "SFSDK.DLL" Alias "#650" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mu As Double, _
                                                                               ByRef alphas As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef betas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nInnovationType As Integer, _
                                                                               ByVal nu As Double, _
                                                                               ByVal retType As Integer, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_GARCH_RESID Lib "SFSDK.DLL" Alias "#651" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mu As Double, _
                                                                               ByRef alphas As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef betas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nInnovationType As Integer, _
                                                                               ByVal nu As Double, _
                                                                               ByVal retType As Integer) As Integer
   Public Declare Function NDK_GARCH_FITTED Lib "SFSDK.DLL" Alias "#652" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByVal mu As Double, _
                                                                               ByRef alphas As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef betas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nInnovationType As Integer, _
                                                                               ByVal nu As Double, _
                                                                               ByVal retType As Integer) As Integer
   Public Declare Function NDK_GARCH_FORE Lib "SFSDK.DLL" Alias "#653" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef sigmas As Double, _
                                                                               ByVal NSigmas As Long, _
                                                                               ByVal mu As Double, _
                                                                               ByRef alphas As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef betas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nInnovationType As Integer, _
                                                                               ByVal nu As Double, _
                                                                               ByVal nSteps As Long, _
                                                                               ByVal retType As Integer, _
                                                                               ByVal alpha As Double, _
                                                                               ByRef retVal As Double) As Integer
   Public Declare Function NDK_GARCH_PARAM Lib "SFSDK.DLL" Alias "#655" (ByRef X As Double, _
                                                                               ByVal N As Long, _
                                                                               ByRef mu As Double, _
                                                                               ByRef alphas As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef betas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nInnovationType As Integer, _
                                                                               ByRef nu As Double, _
                                                                               ByVal retType As Integer, _
                                                                               ByVal maxIter As Long) As Integer
   Public Declare Function NDK_GARCH_VALIDATE Lib "SFSDK.DLL" Alias "#656" (ByVal mu As Double, _
                                                                               ByRef alphas As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef betas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nInnovationType As Integer, _
                                                                               ByVal nu As Double) As Integer
   Public Declare Function NDK_GARCH_LRVAR Lib "SFSDK.DLL" Alias "#657" (ByVal mu As Double, _
                                                                               ByRef alphas As Double, _
                                                                               ByVal p As Long, _
                                                                               ByRef betas As Double, _
                                                                               ByVal q As Long, _
                                                                               ByVal nInnovationType As Integer, _
                                                                               ByVal nu As Double, _
                                                                               ByRef retVal As Double) As Integer
                                                                               
                                                                                   
                                                                               
                                                                               
                                                                               
                                                                               
   ' EGARCH Function
   ''' Computes the log-likelihood ((LLF), Akaike Information Criterion (AIC) or other goodness of fit function of the GARCH model. More...
   Public Declare Function NDK_EGARCH_GOF Lib "SFSDK.DLL" Alias "#660" (ByRef pData As Double, ByVal nSize As Long, ByVal mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                       ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, ByVal nInnovationType As Integer, _
                                                                                       ByVal nu As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
   
   
   ''' Returns an array of cells for the standardized residuals of a given GARCH model. More...
   Public Declare Function NDK_EGARCH_RESID Lib "SFSDK.DLL" Alias "#661" (ByRef pData As Double, ByVal nSize As Long, ByVal mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                       ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, ByVal nInnovationType As Integer, _
                                                                                       ByVal nu As Double, ByVal retType As Integer) As Long
   
   
   ''' Returns an array of cells for the initial (non-optimal), optimal or standard errors of the model's parameters. More...
   Public Declare Function NDK_EGARCH_PARAM Lib "SFSDK.DLL" Alias "#665" (ByRef pData As Double, ByVal nSize As Long, ByRef mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                       ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, ByVal nInnovationType As Integer, _
                                                                                       ByRef nu As Double, ByVal retType As Integer, ByVal maxIter As Long) As Long
   
    
   ''' Returns a simulated data series the underlying EGARCH process. More...
   Public Declare Function NDK_EGARCH_SIM Lib "SFSDK.DLL" Alias "#664" (ByRef pData As Double, ByVal nSize As Long, ByVal sigmas As Double, ByVal nSigmasSize As Long, _
                                                                                     ByVal mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                     ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                     ByVal nInnovationType As Integer, ByVal nu As Double, _
                                                                                     ByVal nSteps As Long, ByVal seed As Long, ByRef retVal As Double) As Long
   
    
   ''' Calculates the out-of-sample forecast statistics. More...
   Public Declare Function NDK_EGARCH_FORE Lib "SFSDK.DLL" Alias "#663" (ByRef pData As Double, ByVal nSize As Long, ByVal sigmas As Double, ByVal nSigmasSize As Long, _
                                                                                     ByVal mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                     ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                     ByVal nInnovationType As Integer, ByVal nu As Double, _
                                                                                     ByVal nSteps As Long, ByVal retType As Integer, ByVal alpha As Double, ByRef retVal As Double) As Long
   
    
   '''Returns an array of cells for the fitted values (i.e. mean, volatility and residuals) More...
   Public Declare Function NDK_EGARCH_FITTED Lib "SFSDK.DLL" Alias "#662" (ByRef pData As Double, ByVal nSize As Long, ByVal mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                     ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                     ByVal nInnovationType As Integer, ByVal nu As Double, ByVal retType As Integer) As Long
   
    
   '''   Calculates the long-run average volatility for a given E-GARCH model. More...
   Public Declare Function NDK_EGARCH_LRVAR Lib "SFSDK.DLL" Alias "#667" (ByVal mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                     ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                     ByVal nInnovationType As Integer, ByVal nu As Double, ByRef retVal As IDouble) As Long
   
    
   '''   Examines the model's parameters for stability constraints (e.g. stationary, positive variance, etc.). More...
   Public Declare Function NDK_EGARCH_VALIDATE Lib "SFSDK.DLL" Alias "#666" (ByVal mu As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                     ByRef Gammas As Double, ByVal g As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                     ByVal nInnovationType As Integer, ByVal nu As Double) As Long
   
   ' GARCH M
   ''' Computes the log-likelihood ((LLF), Akaike Information Criterion (AIC) or other goodness of fit function of the GARCH model. More...
   Public Declare Function NDK_GARCHM_GOF Lib "SFSDK.DLL" Alias "#670" (ByRef pData As Double, ByVal nSize As Long, ByVal mu As Double, ByVal lambda As Double, _
                                                                                       ByRef alphas As Double, ByVal p As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                       ByVal nInnovationType As Integer, ByVal nu As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
   ''' Returns an array of cells for the standardized residuals of a given GARCH model. More...
   Public Declare Function NDK_GARCHM_RESID Lib "SFSDK.DLL" Alias "#671" (ByRef pData As Double, ByVal nSize As Long, ByVal mu As Double, ByVal lambda As Double, _
                                                                                       ByRef alphas As Double, ByVal p As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                       ByVal nInnovationType As Integer, ByVal nu As Double, ByVal retType As Integer) As Long
   ''' Returns an array of cells for the initial (non-optimal), optimal or standard errors of the model's parameters. More...
   Public Declare Function NDK_GARCHM_PARAM Lib "SFSDK.DLL" Alias "#675" (ByRef pData As Double, ByVal nSize As Long, ByRef mu As Double, ByRef lambda As Double, _
                                                                                       ByRef alphas As Double, ByVal p As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                       ByVal nInnovationType As Integer, ByRef nu As Double, ByVal retType As Integer, ByVal maxIter As Long) As Long
   ''' Returns a simulated data series the underlying EGARCH process. More...
   Public Declare Function NDK_GARCHM_SIM Lib "SFSDK.DLL" Alias "#674" (ByRef pData As Double, ByVal nSize As Long, ByVal sigmas As Double, ByVal nSigmasSize As Long, _
                                                                                     ByVal mu As Double, ByVal lambda As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                     ByRef betas As Double, ByVal q As Long, _
                                                                                     ByVal nInnovationType As Integer, ByVal nu As Double, _
                                                                                     ByVal nSteps As Long, ByVal seed As Long, ByRef retVal As Double) As Long
   ''' Calculates the out-of-sample forecast statistics. More...
   Public Declare Function NDK_GARCHM_FORE Lib "SFSDK.DLL" Alias "#673" (ByRef pData As Double, ByVal nSize As Long, ByVal sigmas As Double, ByVal nSigmasSize As Long, _
                                                                                     ByVal mu As Double, ByVal lambda As Double, ByRef alphas As Double, ByVal p As Long, _
                                                                                     ByRef betas As Double, ByVal q As Long, _
                                                                                     ByVal nInnovationType As Integer, ByVal nu As Double, _
                                                                                     ByVal nSteps As Long, ByVal retType As Integer, ByVal alpha As Double, ByRef retVal As Double) As Long
   ''' Returns an array of cells for the fitted values (i.e. mean, volatility and residuals) More...
   Public Declare Function NDK_GARCHM_FITTED Lib "SFSDK.DLL" Alias "#672" (ByRef pData As Double, ByVal nSize As Long, ByVal mu As Double, ByVal lambda As Double, _
                                                                                       ByRef alphas As Double, ByVal p As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                       ByVal nInnovationType As Integer, ByVal nu As Double, ByVal retType As Integer) As Long
   ''' Calculates the long-run average volatility for the given GARCH-M model. More..
   Public Declare Function NDK_GARCHM_LRVAR Lib "SFSDK.DLL" Alias "#677" (ByVal mu As Double, ByVal lambda As Double, _
                                                                                         ByRef alphas As Double, ByVal p As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                         ByVal nInnovationType As Integer, ByVal nu As Double, ByRef retVal As Double) As Long
   ''' Calculates the long-run average volatility for the given GARCH-M model. More..
   Public Declare Function NDK_GARCHM_VALIDATE Lib "SFSDK.DLL" Alias "#676" (ByVal mu As Double, ByVal lambda As Double, _
                                                                                         ByRef alphas As Double, ByVal p As Long, ByRef betas As Double, ByVal q As Long, _
                                                                                         ByVal nInnovationType As Integer, ByVal nu As Double) As Long
    
   ''' Gneralized Linear Model Functions
   ''' Examines the model's parameters for constraints (e.g. positive variance, etc.) More...
   Public Declare Function NDK_GLM_VALIDATE Lib "SFSDK.DLL" Alias "#715" (ByRef betas As Double, ByVal nBetas As Long, ByVal phi As Double, ByVal Lvk As Integer) As Long
   ''' Computes the log-likelihood ((LLF), Akaike Information Criterion (AIC) or other goodness of fit function of the GLM model. More...
   Public Declare Function NDK_GLM_GOF Lib "SFSDK.DLL" Alias "#710" (ByRef Y As Double, ByVal nSize As Long, ByRef X As Double, ByVal nVars As Long, _
                                                                                   ByRef betas As Double, ByVal nBetas As Long, ByVal phi As Double, ByVal Lvk As Integer, _
                                                                                   ByVal retType As Integer, ByRef retVal As Double) As Long
   ''' Returns the standardized residuals/errors of a given GLM. More...
   Public Declare Function NDK_GLM_RESID Lib "SFSDK.DLL" Alias "#711" (ByRef Y As Double, ByVal nSize As Long, ByRef X As Double, ByVal nVars As Long, _
                                                                                   ByRef betas As Double, ByVal nBetas As Long, ByRef phi As Double, ByVal Lvk As Integer, _
                                                                                   ByVal retType As Integer) As Long
   ''' Returns an array of cells for the initial (non-optimal), optimal or standard errors of the model's parameters. More...
   Public Declare Function NDK_GLM_PARAM Lib "SFSDK.DLL" Alias "#714" (ByRef Y As Double, ByVal nSize As Long, ByRef X As Double, ByVal nVars As Long, _
                                                                                   ByRef betas As Double, ByVal nBetas As Long, ByVal phi As Double, ByVal Lvk As Integer, _
                                                                                   ByVal retType As Integer, ByVal maxIter As Long) As Long
   ''' calculates the expected response (i.e. mean) value; given the GLM model and the values of the explanatory variables.
   Public Declare Function NDK_GLM_FORE Lib "SFSDK.DLL" Alias "#713" (ByRef X As Double, ByVal nVars As Long, _
                                                                                   ByRef betas As Double, ByVal nBetas As Long, ByVal phi As Double, ByVal Lvk As Integer, _
                                                                                   ByVal retType As Integer, ByVal alpha As Double, ByRef retVal As Double) As Long
   ''' Returns the standardized residuals/errors of a given GLM. More...
   Public Declare Function NDK_GLM_FITTED Lib "SFSDK.DLL" Alias "#712" (ByRef Y As Double, ByVal nSize As Long, ByRef X As Double, ByVal nVars As Long, _
                                                                                   ByRef betas As Double, ByVal nBetas As Long, ByRef phi As Double, ByVal Lvk As Integer, _
                                                                                   ByVal retType As Integer) As Long
    
   
   
   
   ''' Multiple Linear Regression (MLR)
   
   ''' Returns the standardized residuals/errors of a given GLM. More...
   Public Declare Function NDK_SLR_PARAM Lib "SFSDK.DLL" Alias "#720" (ByRef X As Double, ByVal nXSize As Long, ByRef Y As Double, ByVal nYSize As Long, _
                                                                                     ByVal intercept As Double, ByVal alpha As Double, ByVal retType As Integer, _
                                                                                     ByVal ParamIndex As Integer, ByRef retVal As Double) As Long
   
   Public Declare Function NDK_SLR_FORE Lib "SFSDK.DLL" Alias "#721" (ByRef X As Double, ByVal nXSize As Long, ByRef Y As Double, ByVal nYSize As Long, _
                                                                                     ByVal intercept As Double, ByVal target As Double, ByVal alphas As Double, _
                                                                                     ByVal retType As Integer, ByRef retVal As Double) As Long
   
   Public Declare Function NDK_SLR_FITTED Lib "SFSDK.DLL" Alias "#722" (ByRef X As Double, ByVal nXSize As Long, ByRef Y As Double, ByVal nYSize As Long, _
                                                                                     ByVal intercept As Double, ByVal retType As Integer) As Long
   
   
   Public Declare Function NDK_SLR_ANOVA Lib "SFSDK.DLL" Alias "#723" (ByRef X As Double, ByVal nXSize As Long, ByRef Y As Double, ByVal nYSize As Long, _
                                                                                     ByVal intercept As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
   
   
   Public Declare Function NDK_SLR_GOF Lib "SFSDK.DLL" Alias "#724" (ByRef X As Double, ByVal nXSize As Long, ByRef Y As Double, ByVal nYSize As Long, _
                                                                                     ByVal intercept As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
    
   
   Public Declare Function NDK_MLR_PARAM Lib "SFSDK.DLL" Alias "#730" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByRef Y As Double, ByVal nYSize As Long, ByVal intercept As Double, _
                                                                                 ByVal alphas As Double, ByVal retType As Integer, ByVal nParamIndex As Integer, ByRef retVal As Double) As Long
    
   Public Declare Function NDK_MLR_FORE Lib "SFSDK.DLL" Alias "#731" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByRef Y As Double, ByVal nYSize As Long, ByVal intercept As Double, ByRef target As Double, _
                                                                                 ByVal alphas As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
   
    
   Public Declare Function NDK_MLR_FITTED Lib "SFSDK.DLL" Alias "#732" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByRef Y As Double, ByVal nYSize As Long, ByVal intercept As Double, ByVal retType As Integer) As Long
   
    
   Public Declare Function NDK_MLR_ANOVA Lib "SFSDK.DLL" Alias "#733" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByRef Y As Double, ByVal nYSize As Long, ByVal intercept As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
   
   Public Declare Function NDK_MLR_GOF Lib "SFSDK.DLL" Alias "#734" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByRef Y As Double, ByVal nYSize As Long, ByVal intercept As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
    
   Public Declare Function NDK_MLR_PRFTest Lib "SFSDK.DLL" Alias "#736" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef Y As Double, ByVal nYSize As Long, _
                                                                                    ByVal intercept As Double, ByRef mask1 As Byte, ByVal nMaskLen1 As Long, _
                                                                                    ByRef mask2 As Byte, ByVal nMaskLen2 As Long, _
                                                                                    ByVal alpha As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
    
   Public Declare Function NDK_MLR_STEPWISE Lib "SFSDK.DLL" Alias "#735" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByRef Y As Double, ByVal nYSize As Long, ByVal intercept As Double, ByVal alpha As Double, ByVal mode As Integer) As Long



   ''' PCA
   Public Declare Function NDK_PCA_COMP Lib "SFSDK.DLL" Alias "#740" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByVal standardize As Integer, ByVal nCompIndex As Integer, ByVal retType As Integer, _
                                                                                 ByRef retVal As Double, ByVal nOutSize As Long) As Long
   Public Declare Function NDK_PCA_VAR Lib "SFSDK.DLL" Alias "#741" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                 ByVal standardize As Integer, ByVal nVarIndex As Integer, ByVal MaxPC As Integer, ByVal retType As Integer, _
                                                                                 ByRef retVal As Double, ByVal nOutSize As Long) As Long
   Public Declare Function NDK_PCR_PARAM Lib "SFSDK.DLL" Alias "#742" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                  ByRef pYData As Double, ByVal nYSize As Long, ByVal intercept As Double, ByVal alpha As Double, _
                                                                                  ByVal retType As Integer, ByVal nParamIndex, ByRef retVal As Double) As Long
   Public Declare Function NDK_PCR_FORE Lib "SFSDK.DLL" Alias "#743" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                  ByRef pYData As Double, ByVal nYSize As Long, ByVal intercept As Double, ByRef target As Double, _
                                                                                  ByVal alpha As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
   Public Declare Function NDK_PCR_FITTED Lib "SFSDK.DLL" Alias "#744" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                  ByRef pYData As Double, ByVal nYSize As Long, ByVal intercept As Double, _
                                                                                  ByVal retType As Integer) As Long
   Public Declare Function NDK_PCR_ANOVA Lib "SFSDK.DLL" Alias "#745" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                  ByRef pYData As Double, ByVal nYSize As Long, ByVal intercept As Double, _
                                                                                  ByVal retType As Integer, ByRef retVal As Double) As Long
   Public Declare Function NDK_PCR_GOF Lib "SFSDK.DLL" Alias "#746" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                  ByRef pYData As Double, ByVal nYSize As Long, ByVal intercept As Double, _
                                                                                  ByVal retType As Integer, ByRef retVal As Double) As Long
   Public Declare Function NDK_PCR_PRFTest Lib "SFSDK.DLL" Alias "#748" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef pYData As Double, ByVal nYSize As Long, _
                                                                                    ByVal intercept As Double, _
                                                                                    ByRef mask1 As Byte, ByVal nMaskLen1 As Long, ByRef mask2 As Byte, ByVal nMaskLen2 As Long, _
                                                                                    ByVal alpha As Double, ByVal retType As Integer, ByRef retVal As Double) As Long
   Public Declare Function NDK_PCR_STEPWISE Lib "SFSDK.DLL" Alias "#747" (ByRef X As Double, ByVal nXSize As Long, ByVal nxVars As Long, ByRef mask As Byte, ByVal nMaskLen As Long, _
                                                                                  ByRef pYData As Double, ByVal nYSize As Long, ByVal intercept As Double, _
                                                                                  ByVal alpha As Double, ByVal nMode As Integer) As Long
    
    
   
   ''' Transform
   '''Returns an array of cells for the (backward shifted, backshifted or lagged time series. More...
   Public Declare Function NDK_LAG Lib "SFSDK.DLL" Alias "#1000" (ByRef X As Double, ByVal N As Long, ByVal K As Long) As Long
   
   '''Returns an array of cells for the differenced time series (i.e. (1-L^S)^D). . More..
   Public Declare Function NDK_DIFF Lib "SFSDK.DLL" Alias "#1005" (ByRef X As Double, ByVal N As Long, ByVal S As Long, ByVal D As Long) As Long
   
   '''Returns an array of cells for the integrated time series (inverse operator of NDK_DIFF). . More...
   Public Declare Function NDK_INTEG Lib "SFSDK.DLL" Alias "#1010" (ByRef X As Double, ByVal N As Long, ByVal S As Long, ByVal D As Long, ByRef X0 As Double, ByVal N0 As Long) As Long
   
   '''Returns an array of cells of a time series after removing all missing values. More...
   Public Declare Function NDK_RMNA Lib "SFSDK.DLL" Alias "#4001" (ByRef X As Double, ByVal N As Long) As Long
    
   '''Returns the time-reversed order time series (i.e. the first observation is swapped with the last observation, etc.): both missing and non-missing values. More.
   Public Declare Function NDK_REVERSE Lib "SFSDK.DLL" Alias "#1024" (ByRef X As Double, ByVal N As Long) As Long
    
   '''Returns an array of cells for the scaled time series. More...
   Public Declare Function NDK_SCALE Lib "SFSDK.DLL" Alias "#1023" (ByRef X As Double, ByVal N As Long, ByVal K As Double) As Long
    
   '''Returns an array of the difference between two time series. More...
   Public Declare Function NDK_SUB Lib "SFSDK.DLL" Alias "#1022" (ByRef X1 As Double, ByVal N1 As Long, ByVal X2 As Double, ByVal N2 As Long) As Long
    
   '''Returns an array of the difference between two time series. More...
   Public Declare Function NDK_ADD Lib "SFSDK.DLL" Alias "#1021" (ByRef X1 As Double, ByVal N1 As Long, ByVal X2 As Double, ByVal N2 As Long) As Long
    
   '''Computes the complementary log-log transformation, including its inverse. More...
   Public Declare Function NDK_CLOGLOG Lib "SFSDK.DLL" Alias "#4005" (ByRef X As Double, ByVal N As Long, ByVal retType As Integer) As Long
    
   '''Computes the probit transformation, including its inverse. More..
   Public Declare Function NDK_PROBIT Lib "SFSDK.DLL" Alias "#4004" (ByRef X As Double, ByVal N As Long, ByVal retType As Integer) As Long
    
   '''Computes the complementary log-log transformation, including its inverse. More...
   Public Declare Function NDK_LOGIT Lib "SFSDK.DLL" Alias "#4003" (ByRef X As Double, ByVal N As Long, ByVal retType As Integer) As Long
   
   '''Computes the complementary Box-Cox transformation, including its inverse. More...
   Public Declare Function NDK_BOXCOX Lib "SFSDK.DLL" Alias "#4002" (ByRef X As Double, ByVal N As Long, ByRef lambda As Double, ByRef alpha As Double, _
                                                                                   ByVal retType As Integer, ByRef retVal As Double) As Long
    
   '''Detrends a time series using a regression of y against a polynomial time trend of order p. More...
   Public Declare Function NDK_DETREND Lib "SFSDK.DLL" Alias "#4010" (ByRef X As Double, ByVal N As Long, ByVal PolyOrder As Integer) As Long
    
   '''Returns an array of the deseasonalized time series assuming a linear model. More...
   Public Declare Function NDK_RMSEASONAL Lib "SFSDK.DLL" Alias "#4017" (ByRef X As Double, ByVal N As Long, ByVal period As Long) As Long
   
   '''Returns an array of cells of a time series after substituting all missing values with the mean/median. More...
   Public Declare Function NDK_INTERP_NAN Lib "SFSDK.DLL" Alias "#4000" (ByRef X As Double, ByVal N As Long, ByVal nMethod As Integer, ByVal plug As Double) As Long
    
   '''Examine whether the given array has one/more missing values. More..
   Public Declare Function NDK_HASNA Lib "SFSDK.DLL" Alias "#4018" (ByRef X As Double, ByVal N As Long, ByVal intermediate As Boolean) As Long
   


   ''' Spectral Analysis
   ''' Returns an array of cells for the convolution operator of two time series. More...
   Public Declare Function NDK_CONVOLUTION Lib "SFSDK.DLL" Alias "#1032" (ByRef X1 As Double, ByVal N1 As Long, ByRef X2 As Double, ByVal N2 As Long, ByRef Z As Double, ByVal nZSize As Long) As Long
   
   
   ''' Calculates the inverse discrete fast Fourier transformation, recovering the time series. More...
   Public Declare Function NDK_IDFT Lib "SFSDK.DLL" Alias "#1031" (ByRef Amp As Double, ByRef Phase As Double, ByVal nSize As Long, ByRef X As Double, ByVal nXSize As Long) As Long
   
   
   ''' Calculates the discrete fast Fourier transformation for amplitude and phase. More...
   Public Declare Function NDK_DFT Lib "SFSDK.DLL" Alias "#1030" (ByRef X As Double, ByVal nXSize As Long, ByRef Amp As Double, ByRef Phase As Double, ByVal nSize As Long) As Long
   
   ''' computes cyclical component of given time series using the Hodrick?Prescott filter. More...
   Public Declare Function NDK_HodrickPrescotFilter Lib "SFSDK.DLL" Alias "#1033" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, ByVal lambda As Double) As Long
    
    
   ''' Computes trend and cyclical component of a macroeconomic time series using Baxter-King Fixed Length Symmetric Filter. More...
   Public Declare Function NDK_BaxterKingFilter Lib "SFSDK.DLL" Alias "#1034" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, _
                                                                                             ByVal period_min As Long, ByVal period_max As Long, ByVal K As Long, ByVal drift As Boolean, _
                                                                                             ByVal unitroot As Boolean, ByVal retType As Integer) As Long
   
   ''' Smoothing API functions calls
   '''Returns the weighted moving (rolling/running) average using the previous m data points. More...
   Public Declare Function NDK_WMA Lib "SFSDK.DLL" Alias "#2000" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, _
                                                                                ByRef weights As Double, ByVal nwSize As Long, ByVal nHorizon As Long, ByRef retVal As Double) As Long
   
   '''Returns the (Brown's) simple exponential (EMA) smoothing estimate of the value of X at time t+m (based on the raw data up to time t)..
   Public Declare Function NDK_SESMTH Lib "SFSDK.DLL" Alias "#2005" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, _
                                                                                ByRef alpha As Double, ByVal nHorizon As Long, ByVal optimize As Boolean, ByRef retVal As Double) As Long
   
   '''Returns the (Holt-Winter's) double exponential smoothing estimate of the value of X at time T+m.
   Public Declare Function NDK_DESMTH Lib "SFSDK.DLL" Alias "#2010" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, _
                                                                                 ByRef alpha As Double, ByRef beta As Double, ByVal nHorizon As Long, ByVal optimize As Boolean, _
                                                                                 ByRef retVal As Double) As Long
   
   '''Returns the (Brown's) Linear exponential smoothing estimate of the value of X at time T+m (based on the raw data up to time t).
   Public Declare Function NDK_LESMTH Lib "SFSDK.DLL" Alias "#2015" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, _
                                                                                ByRef alpha As Double, ByVal nHorizon As Long, ByVal optimize As Boolean, ByRef retVal As Double) As Long
    
   '''Returns the (Winters's) triple exponential smoothing estimate of the value of X at time T+m. More...
   Public Declare Function NDK_TESMTH Lib "SFSDK.DLL" Alias "#2020" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, _
                                                                                   ByRef alpha As Double, ByRef beta As Double, ByRef gamma As Double, ByVal S As Long, _
                                                                                   ByVal nHorizon As Long, ByVal optimize As Boolean, ByRef retVal As Double) As Long
   
   '''Returns values along a trend curve (e.g. linear, quadratic, exponential, etc.) at time T+m..
   Public Declare Function NDK_TREND Lib "SFSDK.DLL" Alias "#2021" (ByRef X As Double, ByVal N As Long, ByVal Ascending As Boolean, _
                                                                                  ByRef trendType As Integer, ByVal PolyOrder As Integer, ByVal allowIntercept As Boolean, ByVal interecept As Double, _
                                                                                  ByVal nHorizon As Long, ByVal retType As Integer, ByVal alpha As Double, ByRef retVal As Double) As Long
    



   ''' Utilities
   '''estimate the value of the function represented by (x,y) data set at an intermediate x-value. More...
   Public Declare Function NDK_INTERPOLATE Lib "SFSDK.DLL" Alias "#3000" (ByRef X As Double, ByVal nX As Long, ByRef Y As Double, ByVal nY As Long, _
                                                                                        ByRef XT As Double, ByVal nXT As Long, ByVal uMethod As Integer, ByVal extrapolate As Boolean, _
                                                                                        ByRef YVal As Double, ByVal nYVals As Long) As Long
   
   
   '''Locate and return the full path of the default editor (e.g. notepad) in the system. More...
   Public Declare Function NDK_DEFAULT_EDITOR Lib "SFSDK.DLL" Alias "#3025" (ByVal szFullPath As String, ByRef nSize As Long) As Long
   
   
   '''Returns the n-th token/substring in a string after splitting it using a given delimiter.
   Public Declare Function NDK_TOKENIZE Lib "SFSDK.DLL" Alias "#3020" (ByVal szTxt As String, ByVal szDelim As String, ByVal nOrder As Integer, ByVal szRetVal As String, _
                                                                                      ByVal nSize As Long) As Long
    
   '''Returns TRUE if the string matches the regular expression expressed.
   Public Declare Function NDK_REGEX_MATCH Lib "SFSDK.DLL" Alias "#3010" (ByVal szLine As String, ByVal szPattern As String, ByVal ignoreCase As Boolean, _
                                                                                         ByVal partialOK As Boolean, ByRef bRetVal As Boolean) As Long
   
    
   ''' Returns TRUE if the string matches the regular expression expressed.
   Public Declare Function NDK_REGEX_REPLACE Lib "SFSDK.DLL" Alias "#3015" (ByVal szLine As String, ByVal szKey As String, ByVal szValue As String, ByVal ignoreCase As Boolean, ByVal bGlobal As Boolean, ByVal szRetVal As String, ByVal nSize As Long) As Long
    
   
   '''calculates the value of the regression function for an intermediate x-value.
   Public Declare Function NDK_REGRESSION Lib "SFSDK.DLL" Alias "#3005" (ByRef X As Double, ByVal nXSize As Long, ByRef Y As Double, ByVal nYSize As Long, _
                                                                                       ByVal nType As Integer, ByVal nPolyOrder As Integer, ByVal intercept As Double, _
                                                                                       ByVal target As Double, ByVal retType As Integer, ByVal alpha As Double, _
                                                                                       ByRef retVal As Double) As Long



   '''Seasonal ajustments using X12-ARIMA API functions calls
   
   ''' Prepare the X12-ARIMA scripting environment
   Public Declare Function NDK_X12_ENV_INIT Lib "SFSDK.DLL" Alias "#5000" (ByVal init As Boolean) As Long
   ''' Cleanup all files created by ARIMA program
   Public Declare Function NDK_X12_ENV_CLEANUP Lib "SFSDK.DLL" Alias "#5002" () As Long
   ''' Prepare the X12 Model
   Public Declare Function NDK_X12_SCEN_INIT Lib "SFSDK.DLL" Alias "#5005" (ByVal szScenarioName As String, ByRef X12Options As Any) As Long
   ''' cleanup all temp files
   Public Declare Function NDK_X12_SCEN_CLEAUP Lib "SFSDK.DLL" Alias "#5007" (ByVal szScenarioName As String) As Long
   ''' Write the data to the disk
   Public Declare Function NDK_X12_DATA_FILE Lib "SFSDK.DLL" Alias "#5010" (ByVal szScenarioName As String, ByRef data As Double, ByVal nlen As Long, _
                                                                                           ByVal monthly As Boolean, ByVal startDate As Long, ByVal fileType As Integer) As Long
   ''' Write teh SPC file to the disk
   Public Declare Function NDK_X12_SPC_FILE Lib "SFSDK.DLL" Alias "#5015" (ByVal szScenarioName As String, ByRef X12Options As Any) As Long
   ''' Run the X12a program
   Public Declare Function NDK_X12_RUN_BATCH Lib "SFSDK.DLL" Alias "#5020" (ByVal szScenarioName As String, ByVal batchFilename As String, ByRef status As Integer) As Long
   ''' Run the scenario selected
   Public Declare Function NDK_X12_RUN_SCENARIO Lib "SFSDK.DLL" Alias "#5022" (ByVal szScenarioName As String, ByRef status As Integer) As Long
   ''' Examine the status of running X12a program
   Public Declare Function NDK_X12_RUN_STAT Lib "SFSDK.DLL" Alias "#5025" (ByVal szScenarioName As String, ByRef status As Integer, ByVal szMsg As String, ByRef nlen As Long) As Long
   ''' return the full name of the x12a output file
   Public Declare Function NDK_X12_OUT_FILE Lib "SFSDK.DLL" Alias "#5030" (ByVal szScenarioName As String, ByVal retType As Integer, ByVal szOutFile As String, _
                                                                                           ByRef nlen As Long, ByVal OpenFileFlag As Boolean) As Long
   ''' return the output time series (seasonally adjusted)
   Public Declare Function NDK_X12_OUT_SERIES Lib "SFSDK.DLL" Alias "#5035" (ByVal szScenarioName As String, ByVal nComponent As Integer, ByRef pData As Double, ByRef nSize As Long) As Long
   ''' return the forecast output time series
   Public Declare Function NDK_X12_FORE_SERIES Lib "SFSDK.DLL" Alias "#5040" (ByVal szScenarioName As String, ByVal nStep As Long, ByVal retType As Integer, ByRef retVal As Double) As Long
   
   
   
   '''Portfolio Analysis
   '''Compute the portfolio equivalent returns
   Public Declare Function NDK_PORTFOLIO_RET Lib "SFSDK.DLL" Alias "#5500" (ByRef weights As Double, ByVal nAssets As Long, ByRef returns As Double, ByRef retVal As Double) As Long
   '''Calculates the overall portfolio variance (volatility squared)
   Public Declare Function NDK_PORTFOLIO_VARIANCE Lib "SFSDK.DLL" Alias "#5502" (ByRef weights As Double, ByVal nAssets As Long, ByRef covar As Double, ByRef retVal As Double) As Long
   ''' Calculates the covariance between two portfolios
   Public Declare Function NDK_PORTFOLIO_COVARIANCE Lib "SFSDK.DLL" Alias "#5504" (ByRef weights1 As Double, ByRef weights2 As Double, ByVal nAssets As Long, ByRef covar As Double, ByRef retVal As Double) As Long
#End If

'* @}


' Change the current directory
Public Sub ChgCurrentDirectory()
  ChDir Application.Workbooks("NumXLAPI.xla").Path
End Sub
