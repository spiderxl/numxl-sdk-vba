Attribute VB_Name = "STDEVTEST"
Sub Button1_Click()

Dim nRetCode As NDK_RETURN_TYPE
Dim data As Range
Dim target As Double
Dim alpha As Double
Dim retVal As Double
Dim inputArray(1 To 100) As Double

Call SFSDK.ChgCurrentDirectory

nRetCode = NDK_Init("TestApp", vbNullChar, vbNullChar, vbNullChar)

If nRetCode >= NDK_SUCCESS Then
  Set data = Range("$A$1:$A$101")
  
  For i = 1 To 100
    inputArray(i) = data(i, 1).Value
  Next i
  
  target = 2#
  alpha = 0.05
  retVal = -1
  

  nRetCode = NDK_STDEVTEST(inputArray(1), 100, target, alpha, 1, 1, retVal)
  

  nRetCode = NDK_Shutdown()
End If


End Sub
