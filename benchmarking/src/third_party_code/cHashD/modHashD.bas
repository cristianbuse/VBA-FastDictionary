Attribute VB_Name = "modHashD"
Option Explicit

Public Type SAFEARRAY1D
  cDims As Integer
  fFeatures As Integer
  cbElements As Long
  cLocks As Long
#If VBA7 Then
  pvData As LongPtr
#Else
  pvData As Long
#End If
  cElements1D As Long
  lLbound1D As Long
End Type

#If Win64 Then
    Const PTR_SIZE As Long = 8
#Else
    Const PTR_SIZE As Long = 4
#End If

#If Mac Then
    #If VBA7 Then
        Public Declare PtrSafe Function BindArray Lib "/usr/lib/libc.dylib" Alias "memmove" (PArr() As Any, pSrc As LongPtr, Optional ByVal CB As LongPtr = PTR_SIZE) As LongPtr
    #Else
        Public Declare Function BindArray Lib "/usr/lib/libc.dylib" Alias "memmove" (PArr() As Any, pSrc As Long, Optional ByVal CB As Long = 4) As Long
    #End If
#Else
    #If VBA7 Then
        Public Declare PtrSafe Sub BindArray Lib "kernel32" Alias "RtlMoveMemory" (PArr() As Any, pSrc As LongPtr, Optional ByVal CB As LongPtr = PTR_SIZE)
        Private Declare PtrSafe Function CharLowerBuffW& Lib "user32" (lpsz As Any, ByVal cchLength&)
    #Else
        Public Declare Sub BindArray Lib "kernel32" Alias "RtlMoveMemory" (PArr() As Any, pSrc&, Optional ByVal CB& = 4)
        Public Declare Function VariantCopy Lib "oleaut32" (Dst As Any, src As Any) As Long
        Public Declare Function VariantCopyInd Lib "oleaut32" (Dst As Any, src As Any) As Long
        Private Declare Function CharLowerBuffW& Lib "user32" (lpsz As Any, ByVal cchLength&)
    #End If
#End If

Public LWC(-32768 To 32767) As Integer

Public Sub InitLWC()
  Dim i As Long
#If Mac Then
  For i = -32768 To 32767: LWC(i) = AscW(LCase$(ChrW$(i))): Next
#Else
  For i = -32768 To 32767: LWC(i) = i: Next 'init the Lookup-Array to the full WChar-range
  CharLowerBuffW LWC(-32768), 65536 '<-- and convert its whole content to LowerCase-WChars
#End If
End Sub
