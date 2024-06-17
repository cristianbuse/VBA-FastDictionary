VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHashD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'cHashD, Olaf Schmidt in August 2016 (enhanced about mixed, variable KeyTypes and Remove-functions in 2020)

Private Const NoEntry As Currency = -98765432123456@
Private Const DynTakeOver As Long = 4, HMul As Long = 3727 'fixed HashSlot-amount and Hash-Multiplikator

Private Type HashTableEntry
  Count As Long
  DataIdxsStat(0 To DynTakeOver - 1) As Long
  DataIdxsDyn() As Long
End Type

Private W() As Integer, saW As SAFEARRAY1D

Private mLastExpectedMaxCount As Long, mEnsureUniqueKeys As Boolean, mLastH As Long
Private mCount As Long, mDTUB As Long, mHashTableSize As Long, mCompareMode As VbCompareMethod
Private HashTable() As HashTableEntry, mKeys(), mValues()  'both pairing-arrays are of type Variant
Private mFIndexes() As Long, mFCount As Long, mFUB As Long 'three vars, to handle the deletes

Private Sub Class_Initialize()
  saW.cDims = 1:  saW.cbElements = 2
  saW.cLocks = 1: saW.fFeatures = &H11 'FADF_AUTO=&H1 || FADF_FIXEDSIZE=&H10
  BindArray W, VarPtr(saW)

  mCompareMode = vbBinaryCompare
  ReInit 16384 'at startup we set it up, to behave ideally for up to 16384 Items

  If LWC(97) = 0 Then InitLWC
End Sub

Public Sub ReInit(Optional ByVal ExpectedMaxCount As Long, Optional ByVal EnsureUniqueKeys As Boolean = True)
  If ExpectedMaxCount <= 0 Then ExpectedMaxCount = mLastExpectedMaxCount
  If ExpectedMaxCount < 128 Then ExpectedMaxCount = 128
  mLastExpectedMaxCount = ExpectedMaxCount
  mEnsureUniqueKeys = EnsureUniqueKeys

  mHashTableSize = 16
  Do Until mHashTableSize * 2 > ExpectedMaxCount: mHashTableSize = mHashTableSize * 2: Loop
  If mHashTableSize > 524288 Then mHashTableSize = 524288
  ReDim HashTable(0 To mHashTableSize - 1)

  mDTUB = mLastExpectedMaxCount
  ReDim mKeys(0 To mDTUB)
  ReDim mValues(0 To mDTUB)
  mCount = 0
  mFCount = 0: mFUB = 16: ReDim mFIndexes(0 To mFUB)
End Sub

Public Sub Clear()
  ReInit
End Sub

Public Property Get Count() As Long
  Count = mCount
End Property

Public Property Get HasUniqueKeys() As Boolean
  HasUniqueKeys = mEnsureUniqueKeys
End Property

Public Property Get StringCompareMode() As VbCompareMethod
  StringCompareMode = mCompareMode
End Property
Public Property Let StringCompareMode(ByVal RHS As VbCompareMethod)
  mCompareMode = RHS
End Property

Public Function Keys() As Variant()
  If mCount = 0 Then Keys = Array(): Exit Function

  Dim i As Long, j As Long, V(): ReDim V(0 To mCount - 1)
  For i = 0 To mCount - 1
    Do While mValues(j) = NoEntry: j = j + 1: Loop
    #If VBA7 Or Mac Then
      If IsObject(mKeys(j)) Then Set V(i) = mKeys(j) Else V(i) = mKeys(j)
      j = j + 1
    #Else
      VariantCopy V(i), ByVal VarPtr(mKeys(j)): j = j + 1
    #End If
  Next

  Keys = V
End Function
Public Function Items() As Variant()
  If mCount = 0 Then Items = Array(): Exit Function

  Dim i As Long, j As Long, V(): ReDim V(0 To mCount - 1)
  For i = 0 To mCount - 1
    Do While mValues(j) = NoEntry: j = j + 1: Loop
    #If VBA7 Or Mac Then
      If IsObject(mValues(j)) Then Set V(i) = mValues(j) Else V(i) = mValues(j)
      j = j + 1
    #Else
      VariantCopy V(i), ByVal VarPtr(mValues(j)): j = j + 1
    #End If
  Next

  Items = V
End Function
Public Function Pairs() As Variant() 'hand-out the Key-Values in a Pairs-Array
  If mCount = 0 Then Pairs = Array(): Exit Function

  Dim i As Long, j As Long, V(): ReDim V(0 To mCount - 1)
  For i = 0 To mCount - 1
    Do While mValues(j) = NoEntry: j = j + 1: Loop
    V(i) = Array(mKeys(j), mValues(j)): j = j + 1
  Next

  Pairs = V
End Function

Public Function Exists(Key) As Boolean
  Exists = FindIndex(Key) >= 0
End Function

Public Sub Add(Key, Item)
Dim H As Long, UB As Long, i As Long
  If mLastH Then
    H = mLastH: mLastH = 0
  ElseIf mEnsureUniqueKeys Then
    If FindIndex(Key, H) >= 0 Then Err.Raise 457
  Else
    H = -1: FindIndex Key, H 'a Value of -1 for H will skip the Index-Search, returning only H
  End If

  'add the new Pair, prolonging the Keys- and Values-arrays
  If mDTUB < mCount Then
     mDTUB = (mDTUB + 16) * 1.3
     ReDim Preserve mKeys(0 To mDTUB)
     ReDim Preserve mValues(0 To mDTUB)
  End If

  If mFCount > 0 Then
     mFCount = mFCount - 1: i = mFIndexes(mFCount)
  Else
     i = mCount
  End If
#If VBA7 Or Mac Then
  If IsObject(Item) Then Set mValues(i) = Item Else mValues(i) = Item
  If IsObject(Key) Then Set mKeys(i) = Key Else mKeys(i) = Key
#Else
  VariantCopyInd ByVal VarPtr(mValues(i)), ByVal VarPtr(Item)
  VariantCopyInd ByVal VarPtr(mKeys(i)), ByVal VarPtr(Key)
#End If

  'add the new DataIndex to the proper Hash-Buckets
  Select Case HashTable(H).Count
    Case Is < DynTakeOver
      HashTable(H).DataIdxsStat(HashTable(H).Count) = i
    Case DynTakeOver
      ReDim Preserve HashTable(H).DataIdxsDyn(DynTakeOver To DynTakeOver + 3)
      HashTable(H).DataIdxsDyn(DynTakeOver) = i
    Case Else
      UB = UBound(HashTable(H).DataIdxsDyn)
      If UB < HashTable(H).Count Then UB = (UB + 3) * 1.3: ReDim Preserve HashTable(H).DataIdxsDyn(DynTakeOver To UB)
      HashTable(H).DataIdxsDyn(HashTable(H).Count) = i
  End Select
  HashTable(H).Count = HashTable(H).Count + 1

  mCount = mCount + 1
End Sub

Public Sub Remove(Key)
If mCount = 0 Then Exit Sub
Dim H As Long, Idx As Long, i As Long, j As Long
    Idx = FindIndex(Key, H)
    If Idx < 0 Then Err.Raise 5

    For i = 0 To HashTable(H).Count - 2
      If i < DynTakeOver Then
         If j = 0 Then If HashTable(H).DataIdxsStat(i) = Idx Then j = i + 1
         If j Then
            If j < DynTakeOver Then
               HashTable(H).DataIdxsStat(i) = HashTable(H).DataIdxsStat(j): j = j + 1
            Else
               HashTable(H).DataIdxsStat(i) = HashTable(H).DataIdxsDyn(j): j = j + 1
            End If
         End If
      Else
         If j = 0 Then If HashTable(H).DataIdxsDyn(i) = Idx Then j = i + 1
         If j Then HashTable(H).DataIdxsDyn(i) = HashTable(H).DataIdxsDyn(j): j = j + 1
      End If
    Next
    HashTable(H).Count = i

    If mFUB < mFCount Then mFUB = mFUB + mFUB: ReDim Preserve mFIndexes(0 To mFUB)
    mFIndexes(mFCount) = Idx: mFCount = mFCount + 1

    mKeys(Idx) = NoEntry
    mValues(Idx) = NoEntry
    mCount = mCount - 1
End Sub

Public Sub RemoveByIndex(ByVal IndexZeroBased As Long)
  Remove KeyByIndex(IndexZeroBased)
End Sub

Public Function IndexByKey(Key) As Long
  Dim Idx As Long
      Idx = FindIndex(Key)
   If Idx >= 0 And mFCount > 0 Then AdjustIndex Idx, True
  IndexByKey = Idx
End Function

Public Property Get KeyByIndex(ByVal IndexZeroBased As Long)
  If IndexZeroBased < 0 Or IndexZeroBased >= mCount Then Err.Raise 9
  If mFCount Then AdjustIndex IndexZeroBased
#If VBA7 Or Mac Then
  If IsObject(mKeys(IndexZeroBased)) Then Set KeyByIndex = mKeys(IndexZeroBased) _
                                     Else KeyByIndex = mKeys(IndexZeroBased)
#Else
  VariantCopy KeyByIndex, ByVal VarPtr(mKeys(IndexZeroBased))
#End If
End Property

Public Property Get ItemByIndex(ByVal IndexZeroBased As Long)
  If IndexZeroBased < 0 Or IndexZeroBased >= mCount Then Err.Raise 9
  If mFCount Then AdjustIndex IndexZeroBased
#If VBA7 Or Mac Then
  If IsObject(mValues(IndexZeroBased)) Then Set ItemByIndex = mValues(IndexZeroBased) _
                                       Else ItemByIndex = mValues(IndexZeroBased)
#Else
  VariantCopy ItemByIndex, ByVal VarPtr(mValues(IndexZeroBased))
#End If
End Property
Public Property Let ItemByIndex(ByVal IndexZeroBased As Long, RHS)
  If IndexZeroBased < 0 Or IndexZeroBased >= mCount Then Err.Raise 9
  If mFCount Then AdjustIndex IndexZeroBased
#If VBA7 Or Mac Then
  If IsObject(RHS) Then Set mValues(IndexZeroBased) = RHS Else mValues(IndexZeroBased) = RHS
#Else
  VariantCopyInd ByVal VarPtr(mValues(IndexZeroBased)), ByVal VarPtr(RHS)
#End If
End Property
Public Property Set ItemByIndex(ByVal IndexZeroBased As Long, RHS)
  If IndexZeroBased < 0 Or IndexZeroBased >= mCount Then Err.Raise 9
  If mFCount Then AdjustIndex IndexZeroBased
#If VBA7 Or Mac Then
  If IsObject(RHS) Then Set mValues(IndexZeroBased) = RHS Else mValues(IndexZeroBased) = RHS
#Else
  VariantCopyInd ByVal VarPtr(mValues(IndexZeroBased)), ByVal VarPtr(RHS)
#End If
End Property

Public Property Get Item(Key)
Attribute Item.VB_UserMemId = 0
Dim Index As Long:  Index = FindIndex(Key)
 If Index >= 0 Then
#If VBA7 Or Mac Then
  If IsObject(mValues(Index)) Then Set Item = mValues(Index) Else Item = mValues(Index)
#Else
    VariantCopy Item, ByVal VarPtr(mValues(Index))
#End If
 End If
End Property
Public Property Let Item(Key, RHS)
Dim Index As Long:  Index = FindIndex(Key, mLastH)
 If Index = -1 Then
    Add Key, RHS
 Else
#If VBA7 Or Mac Then
  If IsObject(RHS) Then Set mValues(Index) = RHS Else mValues(Index) = RHS
#Else
    VariantCopyInd ByVal VarPtr(mValues(Index)), ByVal VarPtr(RHS)
#End If
 End If
 mLastH = 0
End Property
Public Property Set Item(Key, RHS)
Dim Index As Long:  Index = FindIndex(Key, mLastH)
 If Index = -1 Then
    Add Key, RHS
 Else
#If VBA7 Or Mac Then
  If IsObject(RHS) Then Set mValues(Index) = RHS Else mValues(Index) = RHS
#Else
    VariantCopyInd ByVal VarPtr(mValues(Index)), ByVal VarPtr(RHS)
#End If
 End If
 mLastH = 0
End Property

Private Function FindIndex(Key, Optional H As Long) As Long  'return -1, when no Key can be found
  Dim i As Long, C@, D#, L As LongPtr, F!, VT As VbVarType, HTUB As Long
  HTUB = mHashTableSize - 1
  FindIndex = H
  H = HTUB 'init the HashValue (all bits to 1)

  VT = VarType(Key)
  Select Case VT
    Case vbString
      saW.cElements1D = Len(Key): saW.pvData = StrPtr(Key)
      If mCompareMode = 0 Then
        For i = 0 To saW.cElements1D - 1: H = (H + W(i)) * HMul And HTUB: Next
        If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

        For i = 0 To HashTable(H).Count - 1
          If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
          If VarType(mKeys(FindIndex)) = VT Then If Key = mKeys(FindIndex) Then Exit Function
        Next
      Else
        For i = 0 To saW.cElements1D - 1: H = (H + LWC(W(i))) * HMul And HTUB: Next
        If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

        For i = 0 To HashTable(H).Count - 1
          If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
          If VarType(mKeys(FindIndex)) = VT Then If StrComp(Key, mKeys(FindIndex), mCompareMode) = 0 Then Exit Function
        Next
      End If

    Case vbObject
      L = ObjPtr(Key): saW.cElements1D = 2: saW.pvData = VarPtr(L)
        H = (H + W(0)) * HMul And HTUB 'loop-unrolling (we have only 2 16Bit integers in the array)
        H = (H + W(1)) * HMul And HTUB
      If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

      For i = 0 To HashTable(H).Count - 1
        If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
        If VarType(mKeys(FindIndex)) = VT Then If Key Is mKeys(FindIndex) Then Exit Function
      Next

    Case vbCurrency
      C = Key: saW.cElements1D = 4: saW.pvData = VarPtr(C)
        H = (H + W(0)) * HMul And HTUB 'loop-unrolling (we have only 4 16Bit integers in the array)
        H = (H + W(1)) * HMul And HTUB
        H = (H + W(2)) * HMul And HTUB
        H = (H + W(3)) * HMul And HTUB
      If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

      For i = 0 To HashTable(H).Count - 1
        If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
        If VarType(mKeys(FindIndex)) = VT Then If C = mKeys(FindIndex) Then Exit Function
      Next

    Case vbLong, vbInteger, vbByte
      L = Key: saW.cElements1D = 2: saW.pvData = VarPtr(L)
        H = (H + W(0)) * HMul And HTUB 'loop-unrolling (we have only 2 16Bit integers in the array)
        H = (H + W(1)) * HMul And HTUB
      If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

      For i = 0 To HashTable(H).Count - 1
        If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
        Select Case VarType(mKeys(FindIndex))
          Case vbLong, vbInteger, vbByte: If L = mKeys(FindIndex) Then Exit Function
        End Select
      Next

    Case vbDouble
      D = Key: saW.cElements1D = 4: saW.pvData = VarPtr(D)
        H = (H + W(0)) * HMul And HTUB 'loop-unrolling (we have only 4 16Bit integers in the array)
        H = (H + W(1)) * HMul And HTUB
        H = (H + W(2)) * HMul And HTUB
        H = (H + W(3)) * HMul And HTUB
      If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

      For i = 0 To HashTable(H).Count - 1
        If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
        If VarType(mKeys(FindIndex)) = VT Then If D = mKeys(FindIndex) Then Exit Function
      Next

    Case vbDate
      D = Key: saW.cElements1D = 4: saW.pvData = VarPtr(D)
        H = (H + W(0)) * HMul And HTUB 'loop-unrolling (we have only 4 16Bit integers in the array)
        H = (H + W(1)) * HMul And HTUB
        H = (H + W(2)) * HMul And HTUB
        H = (H + W(3)) * HMul And HTUB
      If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

      For i = 0 To HashTable(H).Count - 1
        If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
        If VarType(mKeys(FindIndex)) = VT Then If D = mKeys(FindIndex) Then Exit Function
      Next

    Case vbSingle
      F = Key: saW.cElements1D = 2: saW.pvData = VarPtr(F)
        H = (H + W(0)) * HMul And HTUB 'loop-unrolling (we have only 2 16Bit integers in the array)
        H = (H + W(1)) * HMul And HTUB
      If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

      For i = 0 To HashTable(H).Count - 1
        If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
        If VarType(mKeys(FindIndex)) = VT Then If F = mKeys(FindIndex) Then Exit Function
      Next
  End Select

  FindIndex = -1
End Function

Private Sub AdjustIndex(Idx As Long, Optional ByVal Inverse As Boolean)
  Dim i As Long, Incr As Long
  If Inverse Then Incr = -1 Else Incr = 1
  For i = 0 To mFCount - 1
     If mFIndexes(i) <= Idx Then Idx = Idx + Incr
  Next
End Sub

Friend Sub CheckHashDistribution()
Dim i As Long, Count As Long, cc As Long, Min As Long, Max As Long
  Min = &H7FFFFFFF
  For i = 0 To UBound(HashTable)
    Count = HashTable(i).Count
    If Count Then
      If Min > Count Then Min = Count
      If Max < Count Then Max = Count
      cc = cc + 1
    End If
  Next
  Debug.Print "Distribution of"; mCount; "entries over a HashTable with"; UBound(HashTable) + 1; "slots:"
  Debug.Print "Used-HashSlots:"; cc
  Debug.Print "Min-Entries:"; Min
  Debug.Print "Max-Entries:"; Max
End Sub

Public Property Get LoadFactor() As Double
    Dim i As Long
    Dim C As Long
    
    For i = 0 To UBound(HashTable)
        If HashTable(i).Count > DynTakeOver Then C = C + UBound(HashTable(i).DataIdxsDyn) - DynTakeOver + 1
    Next i
    LoadFactor = mCount / (C + DynTakeOver * (UBound(HashTable) + 1))
End Property
