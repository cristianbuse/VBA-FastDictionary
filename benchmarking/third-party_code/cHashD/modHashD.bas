Attribute VB_Name = "modHashD"
Option Explicit

Public Type SAFEARRAY1D
  cDims As Integer
  fFeatures As Integer
  cbElements As Long
  cLocks As Long
  pvData As Long
  cElements1D As Long
  lLbound1D As Long
End Type

Public Declare Sub BindArray Lib "kernel32" Alias "RtlMoveMemory" (PArr() As Any, pSrc&, Optional ByVal CB& = 4)
Public Declare Function VariantCopy Lib "oleaut32" (Dst As Any, Src As Any) As Long
Public Declare Function VariantCopyInd Lib "oleaut32" (Dst As Any, Src As Any) As Long
Private Declare Function CharLowerBuffW& Lib "user32" (lpsz As Any, ByVal cchLength&)

Public LWC(-32768 To 32767) As Integer

Public Sub InitLWC()
  Dim i As Long
  For i = -32768 To 32767: LWC(i) = i: Next 'init the Lookup-Array to the full WChar-range
  CharLowerBuffW LWC(-32768), 65536 '<-- and convert its whole content to LowerCase-WChars
End Sub