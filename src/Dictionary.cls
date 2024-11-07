VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''=============================================================================
''' VBA Fast Dictionary
''' --------------------------------------------------
''' https://github.com/cristianbuse/VBA-FastDictionary
''' --------------------------------------------------
''' MIT License
'''
''' Copyright (c) 2024 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

''==============================================================================
'' Description:
''    * Scripting.Dictionary replacement for Windows
''    * Mac OS compatible
''    * Performant when dealing with many Items, large String or Object Keys
'' Methods:
''    * Add
''         - Adds a new Key-Item pair
''         - Keys are any data type except Arrays and User-Defined Types (UDTs)
''    * Exists
''         - Checks if a specified Key exists
''    * Factory
''         - Returns a new Dictionary instance
''    * Items
''         - Returns a 1D array of all the Items
''    * Keys
''         - Returns a 1D array of all the Keys
''    * PredictCount
''         - If the number of Key-Item pairs is known before adding or a good
''           guess is possible then a call to 'PredictCount' with the expected
''           count will prepare the internal size of the hash map so that there
''           are no calls made to 'Rehash' thus resulting in better performance
''    * Remove
''         - Removes an item by Key
''    * RemoveAll
''         - Removes all Key-Item pairs
''    * Self
''         - Self instance - useful in 'With New Dictionary' code blocks
'' Properties:
''    * CompareMode <Get/Let>
''         - Can only be changed if there are no stored items
''         - Can be: vbBinaryCompare (Default), vbTextCompare or a locale ID
''    * Count <Get>
''         - Returns the number of Key-Item pairs
''    * Item <Get>
''         - Returns an Item by Key
''         - Default Member. Can be omitted: d.Item(Key) can be called as d(Key)
''    * Item <Let><Set>
''         - Changes the value of an Item identified by the specified Key
''         - Default Member. d.Item(Key) = n can be called as d(Key) = n
''         - If Key does not exist then the pair is added via 'Add'
''    * Key <Let>
''         - Allows a Key value to be changed while preserving the Item
''    * LoadFactor <Get>
''         - Returns the current % load for the hash map containing indexes
''==============================================================================

'@PredeclaredId
Option Explicit
Option Compare Binary

#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
    #Else
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
    #End If
#Else 'Windows
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    #End If
#End If
#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

Const MaxLoadFactor As Single = 0.5
    
Private Enum InternalConstants 'Hides values from Locals window
#If Win64 Then
    isItemObjBit = &H80000000
    keyMask = &H7FFFFFFF
    ptrSize = 8
#Else
    ptrSize = 4
#End If
    notFound = -1
    intSize = 2
    longSize = 4
    intsPerDouble = 4
    groupSize = ptrSize
    nextItemOffset = ptrSize * 2
    initialGroupCount = 16
End Enum

Private Type Group
    Count As Long
    Index(0 To groupSize - 1) As Long
    Control As LongPtr
    WasEverFull As Boolean
End Type

Private Type HashMap
    Groups() As Group
    GroupCount As Long
    MaxLoad As Long
    GroupMask As Long
    ControlMask As Long
End Type

#If Win64 = 0 Then
Private Type EnumerableVariant
    Value As Variant
    Meta As Long
    IsItemObj As Boolean
    IsKeyObj As Boolean
    NextPtr As Long
End Type
#End If

Private Enum HashMeta
    hmRemoved = 2
    hmError = &H10000000
    hmNumber = &H20000000
    hmText = &H30000000
    hmObject = &H40000000
    [_modHM] = hmError
    [_maskHM] = [_modHM] - 1
End Enum

#Const Windows = (Mac = 0)

#If Windows Then
Private Type ScrDictLayout 'Scripting.Dictionary memory layout
    vTables(0 To 3) As LongPtr
    unkPtr1 As LongPtr
    refCount As Long
    firstItemPtr As LongPtr
    lastItemPtr As LongPtr
#If Win64 = 0 Then
    Dummy As Long
#End If
    hashTablePtr As LongPtr
    hashModulo As Long
    compMode As Long
    localeID As Long
    unkPtrs(0 To 2) As LongPtr
End Type
#End If

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

'Minimal version of MEMORY_ACCESSOR:
'https://github.com/cristianbuse/VBA-MemoryTools
Private Type IntegerAccessor
    ints() As Integer
    sa As SAFEARRAY_1D
End Type
Private Type PointerAccessor
    ptrs() As LongPtr
    sa As SAFEARRAY_1D
    isSet As Boolean
End Type

Private Enum CLPosition
    posComp = 0
    posLcid = 1
End Enum

'Both 'textHasher' and 'compLcid' will link into the default (Predeclared)
'   instance of this class. This is necessary to avoid speed issues on the heap
Private Type Hasher
#If Windows Then
    textHasher() As Collection
    compLcid() As Long 'Will use CLPosition Enum to distinguish index
    lcid As Long       'System localeID
    cachedLcid As Long 'To avoid 'If' statement for mode >= vbTextCompare
    isScrAvailable As Boolean
#End If
    ia As IntegerAccessor
    d As Double 'For hashing numbers
    dPtr As LongPtr
#If Win64 Then
    la As PointerAccessor
#End If
End Type

Private Type Lookups
    ByteShiftL(0 To groupSize - 1) As LongPtr
    ByteMask(0 To groupSize - 1) As LongPtr
    ByteOff(0 To groupSize - 1) As LongPtr
    PositionMask(0 To groupSize - 2) As LongPtr
    CountMask(0 To groupSize) As LongPtr
#If Win64 Then
    ModBytePosition(0 To 15) As Long
    ExpPos(0 To 63) As LongLong 'Double Float Exponent Position
#Else
    ModBytePosition(0 To 4) As Long
#End If
End Type

Private Type EnumProvider
    hasEnum As Boolean
    emptyColl As New Collection
    enumsColl As New Collection
    pa As PointerAccessor
End Type

Private Type DictionaryVariables
    Look() As Lookups
    Hash As HashMap
    Hasher As Hasher
    Enums As EnumProvider
    Compare As VbCompareMethod
    Items() As Variant
#If Win64 Then
    Keys() As Variant
    Meta() As Long
#Else
    Keys() As EnumerableVariant
#End If
    Count As Long
    UBound As Long
    UsedCount As Long
    DefInstance As Dictionary 'Avoids deallocation of default (Predeclared) dict
    IsInit As Boolean
End Type

'Class members
Private Vars As DictionaryVariables

'The 'NewEnum' method must always be the first method in the class so that it's
'   virtual table position is known. This is because we want to fix the x64 bug
'   when using a For Each.. loop:
'https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes
'
'Unfortunately we cannot use ITypeInfo::AddressOfMember because it overrides
'   some assembly bytes:
'https://stackoverflow.com/questions/65507735/address-of-class-method-crash-on-x64

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = DictEnum
End Function

Private Sub Init(Optional ByVal newGroupCount As Long = initialGroupCount)
    InitHashMap newGroupCount
    With Vars
        .Count = 0
        .UsedCount = 0
        .UBound = newGroupCount * groupSize / 2 - 1
        ReDim .Items(0 To .UBound)
    #If Win64 Then
        ReDim .Keys(0 To .UBound + 1) 'Extra member for safe IEnumVariant
        ReDim .Meta(0 To .UBound)
    #Else
        ReDim .Keys(0 To .UBound)
    #End If
    End With
    If Not Vars.IsInit Then InitLocalStructs
    Set Vars.Enums.emptyColl = Nothing 'Clear NextPtr in live IEnumVariant's
    Set Vars.Enums.enumsColl = Nothing
    Vars.Enums.hasEnum = False
End Sub

'Raises Error 457 if a duplicated key was specified
Public Sub Add(ByRef Key As Variant, ByRef Item As Variant)
    Dim hVal As Long
    Dim groupSlot As Long
    Dim i As Long
    Dim controlByte As Long
    '
    i = GetIndex(Key, hVal, groupSlot, controlByte)
    If i > notFound Then Err.Raise 457, TypeName(Me) & ".Add"
    '
    If Vars.Count > Vars.UBound Then
        Vars.UBound = Vars.UBound * 2 + 1
        ReDim Preserve Vars.Items(0 To Vars.UBound)
        '
        If Vars.Enums.hasEnum Then
            Dim ptr As LongPtr: ptr = VarPtr(Vars.Keys(0))
            RemoveUnusedEnums
        End If
        #If Win64 Then 'Extra member for safe IEnumVariant
            ReDim Preserve Vars.Keys(0 To Vars.UBound + 1)
            ReDim Preserve Vars.Meta(0 To Vars.UBound)
        #Else
            ReDim Preserve Vars.Keys(0 To Vars.UBound)
        #End If
        If Vars.Enums.hasEnum Then
            ShiftEnumPointers VarPtr(Vars.Keys(0)) - ptr
        End If
    End If
    i = Vars.Count
    Vars.Count = Vars.Count + 1
    Vars.UsedCount = Vars.UsedCount + 1
    If Vars.UsedCount > Vars.Hash.MaxLoad Then
        Rehash Vars.Hash.GroupCount * 2
        groupSlot = hVal Mod Vars.Hash.GroupCount
        Do While Vars.Hash.Groups(groupSlot).WasEverFull 'Unlikely after rehash
             groupSlot = (groupSlot + 1) Mod Vars.Hash.GroupCount
        Loop
        controlByte = (hVal And Vars.Hash.ControlMask) \ Vars.Hash.GroupCount
    ElseIf Vars.Hash.Groups(groupSlot).Count = groupSize Then
        'This can only happen if ClearMapIndex was previously called
        groupSlot = hVal Mod Vars.Hash.GroupCount
        Do While Vars.Hash.Groups(groupSlot).Count = groupSize
             groupSlot = (groupSlot + 1) Mod Vars.Hash.GroupCount
        Loop
    End If
    '
    With Vars.Hash.Groups(groupSlot)
        .Index(.Count) = i
        .Control = .Control Or (controlByte * Vars.Look(0).ByteShiftL(.Count))
        .Count = .Count + 1
        .WasEverFull = .WasEverFull Or (.Count = groupSize)
    End With
    '
#If Win64 Then
    If hVal And hmObject Then Set Vars.Keys(i) = Key Else Vars.Keys(i) = Key
    If Vars.Enums.hasEnum Then
        RemoveUnusedEnums
        If i > 0 Then
            ptr = VarPtr(Vars.Keys(i))
            Vars.Enums.pa.sa.pvData = ptr
            Vars.Enums.pa.sa.rgsabound0.cElements = 3
            Vars.Enums.pa.ptrs(2) = ptr
        End If
    End If
    If IsObject(Item) Then
        Vars.Meta(i) = hVal Or isItemObjBit
        Set Vars.Items(i) = Item
    Else
        Vars.Meta(i) = hVal
        Vars.Items(i) = Item
    End If
#Else
    With Vars.Keys(i)
        .IsKeyObj = CBool(hVal And hmObject)
        If .IsKeyObj Then Set .Value = Key Else .Value = Key
        If Vars.Enums.hasEnum Then
            RemoveUnusedEnums
            If i > 0 Then Vars.Keys(i - 1).NextPtr = VarPtr(.Value)
        End If
        .IsItemObj = IsObject(Item)
        If .IsItemObj Then
            Set Vars.Items(i) = Item
        Else
            Vars.Items(i) = Item
        End If
        .Meta = hVal
    End With
#End If
End Sub

'Rebuilds hash table using the stored hashes with metadata
Private Sub Rehash(ByVal newGroupCount As Long)
    Dim hVal As Long
    Dim ctrlByte As Long
    Dim groupSlot As Long
    Dim i As Long
    '
    InitHashMap newGroupCount
    For i = 0 To Vars.Count - 1
        #If Win64 Then
            hVal = Vars.Meta(i)
        #Else
            hVal = Vars.Keys(i).Meta
        #End If
        If hVal <> hmRemoved Then
            groupSlot = hVal And Vars.Hash.GroupMask
            Do While Vars.Hash.Groups(groupSlot).WasEverFull 'Unlikely after resize
                groupSlot = (groupSlot + 1) Mod Vars.Hash.GroupCount
            Loop
            With Vars.Hash.Groups(groupSlot)
                .Index(.Count) = i
                ctrlByte = (hVal And Vars.Hash.ControlMask) \ Vars.Hash.GroupCount
                .Control = .Control Or (ctrlByte * Vars.Look(0).ByteShiftL(.Count))
                .Count = .Count + 1
                .WasEverFull = .WasEverFull Or (.Count = groupSize)
            End With
        End If
    Next i
End Sub
Private Sub InitHashMap(ByVal newGroupCount As Long)
    Vars.Hash.GroupCount = newGroupCount
    ReDim Vars.Hash.Groups(0 To Vars.Hash.GroupCount - 1)
    Vars.Hash.MaxLoad = Vars.Hash.GroupCount * groupSize * MaxLoadFactor
    Vars.Hash.GroupMask = Vars.Hash.GroupCount - 1
    '&H80& bit cannot be multiplied with byteBroadcast without causing overflow
    Vars.Hash.ControlMask = Vars.Hash.GroupCount * &H7F&
End Sub

'Returns non-zero data index if key is found
'Returns (ByRef) a Long Integer / DWORD - 32 bits little-endian (LE) hash value
'Sign bit not used so that the following operations are straightforward:
'   - right bit shifting using \ operator and powers of 2
'   - using Mod operator for fast indexing into arrays with positive bounds
'Hash layout:
'|0|1|2|3|...|26|27|28|29|30|31| - Bit Position
'|h|h|h|h|...| h| h| m| m| m| 0| - 28 bits hash and 3 bits metadata about type
'                           | 0| - Not used - sign bit (see comments above)
'                        | 1|    - Bit mask - Input is Obj (HashMeta.hmObject)
'                  | 1| 1|       - Input is a String (HashMeta.hmText)
'                  | 1| 0|       - Input is a Number (HashMeta.hmNumber)
'                  | 0| 1|       - Input is an Error (HashMeta.hmError)
'|0|               | 0| 0|       - Input is Empty (all bits are 0)
'|1|               | 0| 0|       - Input is Null (bit 0 is 1 - vbNull)
'|2|               | 0| 0|       - Key was removed (HashMeta.hmRemoved)
'|?|?|?|?|...| ?| ?| 0| 0|       - Input is +Inf or -Inf or QNaN or SNaN
Private Function GetIndex(ByRef Key As Variant _
                        , Optional ByRef hVal As Long _
                        , Optional ByRef groupSlot As Long _
                        , Optional ByRef controlByte As Long _
                        , Optional ByRef posInGroup As Long _
                        , Optional ByVal hashOnly As Boolean = False) As Long
    #If (Win64 = 0) Or Mac Then
        Const vbLongLong = 20
    #End If
    If Vars.Hash.GroupCount = 0 Then Init
    Dim vt As VbVarType
    Dim i As Long
    '
    'Hash value
    If IsObject(Key) Then vt = vbObject Else vt = VarType(Key)
    If vt = vbString Then
        Const tPrime As Long = 131&
        Const tOverflowMask As Long = &H7FFFFF
#If Windows Then
        If Vars.Hasher.isScrAvailable Then
            Const minLenBinary As Long = 6
            If Vars.Compare = vbBinaryCompare Then
                Vars.Hasher.ia.sa.rgsabound0.cElements = Len(Key)
                If minLenBinary < Vars.Hasher.ia.sa.rgsabound0.cElements Then
                    Vars.Hasher.compLcid(posComp) = vbBinaryCompare
                    hVal = Vars.Hasher.textHasher(0).Item(Key) Or hmText 'Early-binded
                Else 'Faster to loop integers
                    Vars.Hasher.ia.sa.pvData = StrPtr(Key)
                    With Vars.Hasher.ia
                        For i = 0 To Vars.Hasher.ia.sa.rgsabound0.cElements - 1
                            hVal = (hVal And tOverflowMask) * tPrime + .ints(i)
                        Next i
                    End With
                    hVal = hVal And [_maskHM] Or hmText
                    Vars.Hasher.ia.sa.rgsabound0.cElements = intsPerDouble
                    Vars.Hasher.ia.sa.pvData = Vars.Hasher.dPtr
                End If
            Else
                Vars.Hasher.compLcid(posComp) = vbTextCompare
                Vars.Hasher.compLcid(posLcid) = Vars.Hasher.cachedLcid
                hVal = Vars.Hasher.textHasher(0).Item(Key) Or hmText 'Early-binded
            End If
        Else
#End If
            Vars.Hasher.ia.sa.rgsabound0.cElements = Len(Key)
            If Vars.Compare = vbBinaryCompare Then
                Vars.Hasher.ia.sa.pvData = StrPtr(Key)
            Else
                Dim s As String: s = LCase$(Key)
                Vars.Hasher.ia.sa.pvData = StrPtr(s)
            End If
            With Vars.Hasher.ia
                For i = 0 To Vars.Hasher.ia.sa.rgsabound0.cElements - 1
                    hVal = (hVal And tOverflowMask) * tPrime + .ints(i)
                Next i
            End With
            hVal = hVal And [_maskHM] Or hmText
            Vars.Hasher.ia.sa.rgsabound0.cElements = intsPerDouble
            Vars.Hasher.ia.sa.pvData = Vars.Hasher.dPtr
#If Windows Then
        End If
#End If
    ElseIf vt = vbObject Or vt = vbDataObject Then
        Const oPrime As Long = 2701&
        Static iUnk As stdole.IUnknown 'Dim is slower
        '
        vt = vbObject 'Replace vbDataObject if needed
        Set iUnk = Key
        #If Win64 Then
            Const oPreMask As LongLong = &H6FFFFFFFFFFFFFFF^
            Dim ll As LongLong
            ll = ObjPtr(iUnk) And oPreMask
            ll = ll + ll \ &H1000 + ll Mod oPrime
            hVal = CLng(ll And [_maskHM]) Or hmObject
        #Else
            Const oPreMask As Long = &H6FFFFFFF
            hVal = ObjPtr(iUnk) And oPreMask
            hVal = (hVal + hVal Mod oPrime) And [_maskHM] Or hmObject
        #End If
        Set iUnk = Nothing 'Must call because of Static but still faster
    ElseIf vt > vbLongLong Then
        Err.Raise 5, , "Cannot hash an Array or User Defined Type"
    ElseIf vt > vbNull Then
        Dim m As HashMeta
        If vt = vbError Then
            m = hmError
            Vars.Hasher.d = CDbl(Key)
        Else
            m = hmNumber
            Vars.Hasher.d = Key
        End If
        Const n1& = 1201, n2& = 2701, n3& = 131, n4& = 28571
        Const infOrNaN As Integer = &H7FF0 'Exponent bits for Double Float
        #If Win64 Then
            Const maskLL As LongLong = HashMeta.[_maskHM]
            Const iSignBit As Integer = &H8000
            Const factionalBits As LongLong = &HFFFFFFFFFFFFF^
            Const normalBit As LongLong = factionalBits + 1
            Dim h As LongLong, bitFinder As LongLong
            Dim highBits As Integer: highBits = Vars.Hasher.ia.ints(3)
            '
            If (highBits And infOrNaN) = infOrNaN Then
                hVal = (highBits * n1 + Vars.Hasher.ia.ints(0) * n4) And [_maskHM]
            Else
                h = Vars.Hasher.la.ptrs(0) And factionalBits Or normalBit
                bitFinder = h Or Vars.Look(0).ExpPos((highBits And &H3F0) \ &H10)
                h = h \ (bitFinder And -bitFinder)
                If highBits And iSignBit Then h = Not h - 1
                hVal = CLng(h And maskLL) Or m
            End If
        #Else
            With Vars.Hasher.ia
                If (.ints(3) And infOrNaN) = infOrNaN Then m = 0
                hVal = (.ints(0) * n4 + .ints(1) * n3 _
                      + .ints(2) * n2 + .ints(3) * n1) And [_maskHM] Or m
            End With
        #End If
    Else
        hVal = vt 'vbEmpty (0) or vbNull (1)
    End If
    If hashOnly Then Exit Function
    '
    'Calculate sub-hashes
    groupSlot = hVal Mod Vars.Hash.GroupCount
    controlByte = (hVal And Vars.Hash.ControlMask) \ Vars.Hash.GroupCount
    '
    #If Win64 Then
        Const signBit As LongLong = &H8000000000000000^
        Const highBitOff As LongLong = &H7F7F7F7F7F7F7F7F^ 'Bit Off Per Each Byte
        Const nHighBitOff As LongLong = highBitOff Or signBit
        Const byteBroadcast As LongLong = &H101010101010101^
        Const positionPrime As LongLong = 19
    #Else
        Const signBit As Long = &H80000000
        Const highBitOff As Long = &H7F7F7F7F
        Const nHighBitOff As Long = highBitOff Or signBit
        Const byteBroadcast As Long = &H1010101
        Const positionPrime As Long = 7
    #End If
    Dim matches As LongPtr
    Dim cMask As LongPtr
    Dim startSlot As Long: startSlot = groupSlot
    Do
        With Vars.Hash.Groups(groupSlot)
            cMask = Vars.Look(0).CountMask(.Count)
            'Match bytes adapted from:
            'https://graphics.stanford.edu/~seander/bithacks.html#ZeroInWord
            matches = .Control Xor (controlByte * byteBroadcast)
            matches = ((((matches And highBitOff) + nHighBitOff) _
                    Xor signBit) Or matches) And cMask Xor cMask
            Do While matches
                If matches = signBit Then 'Matched last element only
                    posInGroup = groupSize - 1
                    matches = 0
                Else 'Find next match within group
                    posInGroup = Vars.Look(0).ModBytePosition(CLng( _
                                 (-matches And matches) Mod positionPrime))
                    matches = matches Xor Vars.Look(0).PositionMask(posInGroup)
                End If
                GetIndex = .Index(posInGroup)
#If Win64 Then
                If (Vars.Meta(GetIndex) And keyMask) = hVal Then
                    If vt = vbString And (Vars.Compare > vbBinaryCompare) Then
                        If StrComp(Vars.Keys(GetIndex) _
                                 , Key, Vars.Compare) = 0 Then Exit Function
                    ElseIf vt = vbObject Then 'vbDataObject was replaced
                        If Vars.Keys(GetIndex) Is Key Then Exit Function
                    ElseIf hVal < HashMeta.[_modHM] Then 'Already matched
                        Exit Function
                    Else 'Number, Error or case-sensitive Text
                        If Vars.Keys(GetIndex) = Key Then Exit Function
                    End If
                End If
#Else
                If Vars.Keys(GetIndex).Meta = hVal Then
                    If vt = vbString And (Vars.Compare > vbBinaryCompare) Then
                        If StrComp(Vars.Keys(GetIndex).Value _
                                 , Key, Vars.Compare) = 0 Then Exit Function
                    ElseIf vt = vbObject Then 'vbDataObject was replaced
                        If Vars.Keys(GetIndex).Value Is Key Then Exit Function
                    ElseIf hVal < HashMeta.[_modHM] Then 'Already matched
                        Exit Function
                    Else 'Number, Error or case-sensitive Text
                        If Vars.Keys(GetIndex).Value = Key Then Exit Function
                    End If
                End If
#End If
            Loop
            If Not .WasEverFull Then Exit Do
        End With
        groupSlot = (groupSlot + 1) Mod Vars.Hash.GroupCount
    Loop Until groupSlot = startSlot
    GetIndex = notFound
End Function

'Values greater than 2 can be used to compare using specific Locale IDs (LCID)
'https://learn.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/comparemode-property
'Can only be changed if there are no stored items
Public Property Get CompareMode() As VbCompareMethod
    CompareMode = Vars.Compare
End Property
Public Property Let CompareMode(ByVal compMode As VbCompareMethod)
    Const minMode As Long = 0
    Const maxMode As Long = 31890 'Same as StrComp limits
    '
    If Vars.UsedCount > 0 And compMode <> Vars.Compare Then
        Err.Raise 5, TypeName(Me) & ".CompareMode", "Dict already contains data"
    ElseIf compMode < minMode Or compMode > maxMode Then
        Err.Raise 5, TypeName(Me) & ".CompareMode", "Invalid compare method"
    End If
    Vars.Compare = compMode
#If Windows Then
    CacheLcid
#End If
End Property
#If Windows Then
Private Sub CacheLcid()
    If Vars.Compare > vbTextCompare Then
        Vars.Hasher.cachedLcid = Vars.Compare
    Else
        Vars.Hasher.cachedLcid = Vars.Hasher.lcid
    End If
End Sub
#End If

Public Property Get Count() As Long
    Count = Vars.UsedCount
End Property

Public Function Exists(ByRef Key As Variant) As Boolean
    Exists = (GetIndex(Key) > notFound)
End Function

'Raises Error:
'   -    9: invalid key (key is not associated with any element)
'   -  450: 'Set' is missing when assigning an object
'@DefaultMember
Public Property Get Item(ByRef Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_UserMemId = 0
    Dim i As Long: i = GetIndex(Key)
    '
    'A check like i = notFound is avoided for speed purposes
    '   because error 9 is raised anyway when accessing the data arrays
#If Win64 Then
    If Vars.Meta(i) And isItemObjBit Then
#Else
    If Vars.Keys(i).IsItemObj Then
#End If
        Set Item = Vars.Items(i)
    Else
        Item = Vars.Items(i)
    End If
End Property
Public Property Let Item(ByRef Key As Variant, ByRef Item As Variant)
    If IsObject(Item) Or (VarType(Item) = vbDataObject) Then
        Err.Raise 450, TypeName(Me) & ".Item", "'Set' is required"
    End If
    Dim i As Long: i = GetIndex(Key)
    If i > notFound Then
        #If Win64 Then
            Vars.Meta(i) = Vars.Meta(i) And keyMask
        #Else
            Vars.Keys(i).IsItemObj = False
        #End If
        Vars.Items(i) = Item
    Else
        Add Key, Item
    End If
End Property
Public Property Set Item(ByRef Key As Variant, ByRef Item As Object)
    Dim i As Long: i = GetIndex(Key)
    If i > notFound Then
        #If Win64 Then
            Vars.Meta(i) = Vars.Meta(i) Or isItemObjBit
        #Else
            Vars.Keys(i).IsItemObj = True
        #End If
        Set Vars.Items(i) = Item
    Else
        Add Key, Item
    End If
End Property

Public Property Get Items() As Variant()
    If Vars.UsedCount = 0 Then
        Items = Array()
        Exit Property
    End If
    If Vars.UsedCount = Vars.Count Then
        Items = Vars.Items
        ReDim Preserve Items(0 To Vars.Count - 1)
    Else
        Dim res() As Variant
        ReDim res(0 To Vars.UsedCount - 1)
        Dim i As Long
        Dim j As Long
        '
        For i = 0 To Vars.Count - 1
#If Win64 Then
            If Vars.Meta(i) <> hmRemoved Then
                If Vars.Meta(i) And isItemObjBit Then
#Else
            If Vars.Keys(i).Meta <> hmRemoved Then
                If Vars.Keys(i).IsItemObj Then
#End If
                    Set res(j) = Vars.Items(i)
                Else
                    res(j) = Vars.Items(i)
                End If
                j = j + 1
            End If
        Next i
        Items = res
    End If
End Property

'Change an existing key value
'Raises Error:
'   -   9: invalid OldKey (key is not associated with any element)
'   - 457: invalid NewKey (key is already associated with an element)
Public Property Let Key(ByRef OldKey As Variant, ByRef NewKey As Variant)
    Dim hVal As Long
    Dim groupSlot As Long
    Dim posInGroup As Long
    Dim controlByte As Long
    Dim i As Long
    '
    If GetIndex(NewKey, hVal, , controlByte) > notFound Then Err.Raise 457
    i = GetIndex(OldKey, , groupSlot, , posInGroup)
    '
    'A check like i = notFound is avoided for speed purposes
    '   because error 9 is raised anyway when accessing the Keys array
    '
    'Replace Key and Meta while preserving Item
#If Win64 Then
    If Vars.Enums.hasEnum Then
        RemoveUnusedEnums
        Vars.Enums.pa.sa.pvData = VarPtr(Vars.Keys(i))
        Vars.Enums.pa.sa.rgsabound0.cElements = 3
        Dim ptr As LongPtr: ptr = Vars.Enums.pa.ptrs(2)
    End If
    If hVal And HashMeta.hmObject Then
        Set Vars.Keys(i) = NewKey
    Else
        Vars.Keys(i) = NewKey
    End If
    If Vars.Enums.hasEnum Then Vars.Enums.pa.ptrs(2) = ptr
    If Vars.Meta(i) And isItemObjBit Then hVal = hVal Or isItemObjBit
    Vars.Meta(i) = hVal
#Else
    With Vars.Keys(i)
        .IsKeyObj = CBool(hVal And HashMeta.hmObject)
        If .IsKeyObj Then Set .Value = NewKey Else .Value = NewKey
        .Meta = hVal
    End With
#End If
    '
    ClearMapIndex groupSlot, posInGroup
    '
    'Update hash map with new key
    groupSlot = hVal And Vars.Hash.GroupMask
    Do While Vars.Hash.Groups(groupSlot).Count = groupSize
         groupSlot = (groupSlot + 1) Mod Vars.Hash.GroupCount
    Loop
    With Vars.Hash.Groups(groupSlot)
        .Index(.Count) = i
        .Control = .Control Or (controlByte * Vars.Look(0).ByteShiftL(.Count))
        .Count = .Count + 1
        .WasEverFull = .WasEverFull Or (.Count = groupSize)
    End With
End Property

Private Sub ClearMapIndex(ByRef groupSlot As Long, ByRef posInGroup As Long)
    Dim lastPos As Long
    '
    With Vars.Hash.Groups(groupSlot)
        lastPos = .Count - 1
        If posInGroup < lastPos Then 'Swap
            .Index(posInGroup) = .Index(lastPos)
            .Control = .Control And Vars.Look(0).ByteOff(posInGroup)
            .Control = .Control Or ((.Control And Vars.Look(0).ByteMask(lastPos)) _
                                   \ Vars.Look(0).ByteShiftL(lastPos - posInGroup))
        End If
        .Index(lastPos) = 0
        .Control = .Control And Vars.Look(0).ByteOff(lastPos)
        .Count = .Count - 1
    End With
End Sub

Public Property Get Keys() As Variant()
    If Vars.UsedCount = 0 Then
        Keys = Array()
        Exit Property
    End If
    '
    Dim res() As Variant
    Dim i As Long
    Dim j As Long
    '
#If Win64 Then
    If Vars.UsedCount = Vars.Count Then
        Keys = Vars.Keys
        ReDim Preserve Keys(0 To Vars.Count - 1)
    Else
        ReDim res(0 To Vars.UsedCount - 1)
        For i = 0 To Vars.Count - 1
            If Vars.Meta(i) <> hmRemoved Then
                If Vars.Meta(i) And hmObject Then
                    Set res(j) = Vars.Keys(i)
                Else
                    res(j) = Vars.Keys(i)
                End If
                j = j + 1
            End If
        Next i
        Keys = res
    End If
#Else
    ReDim res(0 To Vars.UsedCount - 1)
    For i = 0 To Vars.Count - 1
        With Vars.Keys(i)
            If .Meta <> hmRemoved Then
                If .IsKeyObj Then Set res(j) = .Value Else res(j) = .Value
                j = j + 1
            End If
        End With
    Next i
    Keys = res
#End If
End Property

Public Sub PredictCount(ByVal expectedCount As Long)
    Const maxGroups As Long = HashMeta.[_modHM] / groupSize
    Dim expectedGroups As Long
    '
    If expectedCount <= 0 Then Exit Sub
    expectedGroups = 2 ^ -Int(-(Log(expectedCount / groupSize) _
                              + Log(1 / MaxLoadFactor)) / Log(2))
    If expectedGroups > maxGroups Then expectedGroups = maxGroups
    '
    If Vars.Hash.GroupCount = 0 Then
        If expectedGroups > initialGroupCount Then Init expectedGroups
    ElseIf expectedGroups > Vars.Hash.GroupCount * 2 Then
        Rehash expectedGroups
    End If
End Sub

'Remove a single item
'Raises Error:
'   -   9: invalid key (key is not associated with any element)
Public Sub Remove(ByRef Key As Variant)
    Dim hVal As Long
    Dim groupSlot As Long
    Dim posInGroup As Long
    Dim lastPos As Long
    Dim i As Long
    Dim j As Long
    Dim ptr As LongPtr
    '
    i = GetIndex(Key, , groupSlot, , posInGroup)
    If i = notFound Then Err.Raise 9, TypeName(Me) & ".Remove"
    If Vars.UsedCount = 1 Then
        Init
        Exit Sub
    End If
    '
    ClearMapIndex groupSlot, posInGroup
    '
    lastPos = Vars.Count - 1
    Vars.UsedCount = Vars.UsedCount - 1
    If i < lastPos Then
        'Clear data and register removal
        Vars.Items(i) = Empty
        #If Win64 Then
            Vars.Keys(i) = Empty
            Vars.Meta(i) = hmRemoved
        #Else
            Vars.Keys(i).Value = Empty
            Vars.Keys(i).Meta = hmRemoved
        #End If
        If Vars.Enums.hasEnum Then 'Link previous to next
            j = i + 1
            #If Win64 Then
                Do While Vars.Meta(j) = hmRemoved: j = j + 1: Loop
            #Else
                Do While Vars.Keys(j).Meta = hmRemoved: j = j + 1: Loop
            #End If
            ptr = VarPtr(Vars.Keys(j))
            RemoveUnusedEnums VarPtr(Vars.Keys(i)), ptr
            j = i - 1
            Do While j >= 0
                #If Win64 Then
                    If Vars.Meta(j) <> hmRemoved Then Exit Do
                #Else
                    If Vars.Keys(j).Meta <> hmRemoved Then Exit Do
                #End If
                j = j - 1
            Loop
            If j >= 0 Then
                #If Win64 Then
                    Vars.Enums.pa.sa.pvData = VarPtr(Vars.Keys(j + 1))
                    Vars.Enums.pa.sa.rgsabound0.cElements = 3
                    Vars.Enums.pa.ptrs(2) = ptr
                #Else
                    Vars.Keys(j).NextPtr = ptr
                #End If
            End If
        End If
        Exit Sub
    End If
    If Vars.Count > Vars.UsedCount Then
        i = lastPos - 1
        #If Win64 Then
            Do While Vars.Meta(i) = hmRemoved: i = i - 1: Loop
        #Else
            Do While Vars.Keys(i).Meta = hmRemoved: i = i - 1: Loop
        #End If
        Vars.Count = Vars.Count - lastPos + i + 1
    End If
    If Vars.Enums.hasEnum Then
        RemoveUnusedEnums VarPtr(Vars.Keys(lastPos))
        #If Win64 Then
            Vars.Keys(i + 1) = Empty 'Just to clear the pointer
        #Else
            Vars.Keys(i).NextPtr = 0&
        #End If
    End If
    '
    'Clear Variants in case deallocation is needed e.g. String/Object
    Vars.Items(lastPos) = Empty
    #If Win64 Then
        Vars.Keys(lastPos) = Empty
    #Else
        Vars.Keys(lastPos).Value = Empty
    #End If
    Vars.Count = Vars.Count - 1
End Sub

Public Sub RemoveAll()
    Init
End Sub

'Get self instance. Useful in 'With New Dictionary' blocks
Public Function Self() As Dictionary
    Set Self = Me
End Function

'Returns the current % load for the hash map containing indexes
Public Property Get LoadFactor() As Single
    If Vars.Hash.GroupCount = 0 Then Exit Property
    LoadFactor = Vars.UsedCount / (Vars.Hash.GroupCount * groupSize)
End Property

Private Sub InitLookups(ByRef vlook As Lookups)
    #If Win64 Then
        Const highBitOnly As LongLong = &H8080808080808080^
    #Else
        Const highBitOnly As Long = &H80808080
    #End If
    Dim i As Long
    '
    With vlook
        .ByteShiftL(0) = 1
        .ByteMask(0) = &H7F
        .ByteOff(0) = Not .ByteMask(0)
        For i = 1 To groupSize - 1
            .PositionMask(i - 1) = .ByteShiftL(i - 1) * &H80
            .ByteShiftL(i) = .ByteShiftL(i - 1) * &H100
            .CountMask(i) = .PositionMask(i - 1) + .CountMask(i - 1)
            .ByteMask(i) = .ByteMask(i - 1) * &H100
            .ByteOff(i) = Not .ByteMask(i)
        Next i
        .CountMask(groupSize) = highBitOnly
        '
        'These are the relevant values if applying Modulo positionPrime
    #If Win64 Then
        .ModBytePosition(12) = 1
        .ModBytePosition(13) = 2
        .ModBytePosition(3) = 3
        .ModBytePosition(8) = 4
        .ModBytePosition(15) = 5
        .ModBytePosition(2) = 6
        .ExpPos(51) = 1
        For i = 50 To 0 Step -1
            .ExpPos(i) = .ExpPos(i + 1) * 2
        Next i
    #Else
        .ModBytePosition(1) = 1
        .ModBytePosition(4) = 2
    #End If
    End With
End Sub

'Initializes an Integer Array Accessor on Win and Mac
'Initializes a Collection that can call ScriptingDictionary.HashVal with early
'   binding speed (Win only)
Private Sub InitLocalStructs()
    InitSafeArray Vars.Hasher.ia.sa, intSize
    #If Win64 Then
        InitSafeArray Vars.Hasher.la.sa, ptrSize
    #End If
    '
    Dictionary.InitStructs Vars
    #If Windows Then
        CacheLcid
    #End If
    '
    With Vars.Hasher
        .dPtr = VarPtr(.d)
        .ia.sa.pvData = .dPtr
        .ia.sa.rgsabound0.cElements = intsPerDouble
        #If Win64 Then
            .la.sa.pvData = Vars.Hasher.dPtr
            .la.sa.rgsabound0.cElements = 1
        #End If
    End With
    Vars.IsInit = True
End Sub
Private Sub InitSafeArray(ByRef sa As SAFEARRAY_1D, ByVal elemSize As Long)
    Const FADF_AUTO As Long = &H1
    Const FADF_FIXEDSIZE As Long = &H10
    Const FADF_COMBINED As Long = FADF_AUTO Or FADF_FIXEDSIZE
    With sa
        .cDims = 1
        .fFeatures = FADF_COMBINED
        .cbElements = elemSize
        .cLocks = 1
    End With
End Sub

'To avoid API calls overhead in Win VBA7, a Pointer Accessor is cached in the
'   default instance (Attribute VB_PredeclaredId = True) thus allowing faster
'   memory manipulation
'The best alternative was to use a standard .bas module like 'LibMemory' at:
'   https://github.com/cristianbuse/VBA-MemoryTools
Friend Sub InitStructs(ByRef v As DictionaryVariables)
    Static vlook As Lookups
    Static saL As SAFEARRAY_1D
    Static isLookSet As Boolean
    '
    If Not Me Is Dictionary Then Exit Sub
    Set v.DefInstance = Me 'Avoid deallocation of this instance
    '
    If Not isLookSet Then
        InitLookups vlook
        InitSafeArray saL, LenB(vlook)
        saL.pvData = VarPtr(vlook)
        saL.rgsabound0.cElements = 1
        isLookSet = True
    End If
    '
#If Mac Then
    CopyMemory ByVal VarPtr(v.Hasher.ia), VarPtr(v.Hasher.ia.sa), ptrSize
    CopyMemory ByVal VarPtr(v), VarPtr(saL), ptrSize
    #If Win64 Then
        CopyMemory ByVal VarPtr(v.Hasher.la), VarPtr(v.Hasher.la.sa), ptrSize
    #End If
#Else
    Const dictVTables As Long = 4
    Const dictMainVTableSize As Long = 22
    Const opNumDictHashVal As Long = 21
    Const opNumCollItem As Long = 7
    Static fakeDict As Collection
    Static mainVTable(0 To dictMainVTableSize - 1) As LongPtr
    Static sdl As ScrDictLayout
    Static lcid As Long
    Static saCL As SAFEARRAY_1D
    Static saF As SAFEARRAY_1D
    Static pa As PointerAccessor
    Static fakePtr As LongPtr
    Static isScrAvailable As Boolean
    '
    If Not pa.isSet Then
        InitPointerAccessor pa
        '
        'Early bind a Collection interface to a fake Scripting.Dictionary
        '   where Collection.Item is mapped to Dictionary.HashVal. This allows
        '   calls to HashVal with early binding speed without a dll reference!
        '
        Dim temp As Object
        Dim i As Long
        '
        fakePtr = VarPtr(fakeDict)
        '
        On Error Resume Next 'In case scrun.dll not available
        Set temp = CreateObject("Scripting.Dictionary")
        On Error GoTo 0
        '
        isScrAvailable = Not (temp Is Nothing)
        If isScrAvailable Then
            'Copy Scripting.Dictionary virtual table addresses
            pa.sa.pvData = ObjPtr(temp)
            pa.sa.rgsabound0.cElements = dictMainVTableSize
            For i = 0 To dictVTables - 1
                sdl.vTables(i) = pa.ptrs(i)
            Next i
            '
            'Copy locale ID
            #If Win64 Then
                lcid = CLng(pa.ptrs(10))
            #Else
                lcid = pa.ptrs(12)
            #End If
            Set temp = Nothing 'Actual dictionary not needed anymore
            sdl.localeID = lcid
            '
            'Copy entire main virtual function table to our own
            pa.sa.pvData = sdl.vTables(0)
            For i = 0 To dictMainVTableSize - 1
                mainVTable(i) = pa.ptrs(i)
            Next i
            sdl.vTables(0) = VarPtr(mainVTable(0)) 'Replace main vTable
            '
            'Map Collection.Item to Dictionary.HashVal
            mainVTable(opNumCollItem) = mainVTable(opNumDictHashVal)
            '
            'Set up fake instance
            sdl.hashModulo = HashMeta.[_modHM]
            sdl.refCount = 2 'To avoid deallocation
            pa.sa.rgsabound0.cElements = 1
            pa.sa.pvData = fakePtr
            pa.ptrs(0) = VarPtr(sdl)
            '
            'Init safe array for comp/lcid
            InitSafeArray saCL, longSize
            saCL.pvData = VarPtr(sdl.compMode)
            saCL.rgsabound0.cElements = 2
            '
            'Init safe array for collection
            InitSafeArray saF, ptrSize
            saF.pvData = fakePtr
            saF.rgsabound0.cElements = 1
        End If
        '
        'Increase stack frame size for 'NewEnum' to avoid For Each.. x64 bug
        #If Win64 Then
            pa.sa.pvData = ObjPtr(Me) 'vtbl
            pa.sa.rgsabound0.cElements = 1
            '
            pa.sa.pvData = pa.ptrs(0) + ptrSize * 7 'NewEnum address
            pa.sa.pvData = pa.ptrs(0) + 21 'SUB R12, 0x10
            '
            If (pa.ptrs(0) And &HFFFF^) = &H8149^ Then
                'SUB found. Replace 0x10 (16) with 0x800 (2048) bytes
                pa.ptrs(0) = pa.ptrs(0) And &HFF00000000FFFFFF^ Or &H800000000^
            End If
        #End If
        pa.sa.pvData = fakePtr
        #If Win64 = 0 Then
            pa.sa.rgsabound0.cElements = 1
        #End If
    End If
    If isScrAvailable Then
        v.Hasher.lcid = lcid
        v.Hasher.isScrAvailable = True
        pa.sa.pvData = VarPtr(v.Hasher)
        pa.ptrs(0) = VarPtr(saF)
        '
        'Init compMode and lcid accesor
        pa.sa.pvData = pa.sa.pvData + ptrSize
        pa.ptrs(0) = VarPtr(saCL)
    End If
    '
    'Init integer accesor for hashing wide characters
    pa.sa.pvData = VarPtr(v.Hasher.ia)
    pa.ptrs(0) = VarPtr(v.Hasher.ia.sa)
    '
    'Init LongLong accessor for hashing numbers on x64
    #If Win64 Then
        pa.sa.pvData = VarPtr(v.Hasher.la)
        pa.ptrs(0) = VarPtr(v.Hasher.la.sa)
    #End If
    '
    'Init lookups accessor
    pa.sa.pvData = VarPtr(v)
    pa.ptrs(0) = VarPtr(saL)
    pa.sa.pvData = fakePtr
#End If
End Sub

Friend Sub InitPointerAccessor(ByRef paToInit As PointerAccessor)
    Static pa As PointerAccessor
    '
    If Not pa.isSet Then
        InitSafeArray pa.sa, ptrSize
        CopyMemory ByVal VarPtr(pa), VarPtr(pa.sa), ptrSize 'The only API call
        pa.sa.pvData = VarPtr(pa.sa) 'Some valid address
        pa.sa.rgsabound0.cElements = 1
        pa.isSet = True
    End If
    '
    InitSafeArray paToInit.sa, ptrSize
    pa.sa.pvData = VarPtr(paToInit)
    pa.ptrs(0) = VarPtr(paToInit.sa)
    paToInit.isSet = True
End Sub
    
Public Function Factory() As Dictionary
    Set Factory = New Dictionary
End Function

Public Property Get HashVal(ByRef Key As Variant) As Long
    GetIndex Key, HashVal, hashOnly:=True
End Property

Private Function DictEnum() As IUnknown
    Dim hadEnum As Boolean
    '
    With Vars.Enums
        hadEnum = .hasEnum
        If .hasEnum Then RemoveUnusedEnums
        Set DictEnum = .emptyColl.[_NewEnum]
        .hasEnum = True
        .enumsColl.Add DictEnum, CStr(ObjPtr(DictEnum))
        If Not .pa.isSet Then Dictionary.InitPointerAccessor .pa
    End With
    If Vars.UsedCount = 0 Then Exit Function
    '
    Dim i As Long
    Dim hasGaps As Boolean: hasGaps = (Vars.UsedCount < Vars.Count)
    '
    If hasGaps Then
        #If Win64 Then
            Do While Vars.Meta(i) = hmRemoved: i = i + 1: Loop
        #Else
            Do While Vars.Keys(i).Meta = hmRemoved: i = i + 1: Loop
        #End If
    End If
    With Vars.Enums.pa
        .sa.pvData = ObjPtr(DictEnum) + nextItemOffset
        .sa.rgsabound0.cElements = 1
        .ptrs(0) = VarPtr(Vars.Keys(i))
    End With
    If hadEnum Then Exit Function
    '
    Dim ptr As LongPtr: ptr = VarPtr(Vars.Keys(0))
    Dim j As Long
    '
#If Win64 Then
    Const variantSize As Long = 24
    With Vars.Enums.pa
        .sa.pvData = ptr
        .sa.rgsabound0.cElements = (Vars.Count + 1) * 3
        If hasGaps Then
            j = i * 3 + 5
            For i = i + 1 To Vars.Count - 1
                If Vars.Meta(i) <> hmRemoved Then
                    .ptrs(j) = ptr + variantSize * i
                    j = i * 3 + 5
                End If
            Next i
            .ptrs(j) = 0^
        Else
            For i = 5 To .sa.rgsabound0.cElements - 4 Step 3
                ptr = ptr + variantSize
                .ptrs(i) = ptr
            Next i
            .ptrs(i) = 0^
        End If
    End With
#Else
    Const enumVarSize As Long = 28
    If hasGaps Then
        j = i
        For i = i + 1 To Vars.Count - 1
            If Vars.Keys(i).Meta <> hmRemoved Then
                Vars.Keys(j).NextPtr = ptr + enumVarSize * i
                j = i
            End If
        Next i
        Vars.Keys(j).NextPtr = 0&
    Else
        ptr = ptr + enumVarSize
        For i = 0 To Vars.Count - 2
            Vars.Keys(i).NextPtr = ptr + enumVarSize * i
        Next i
        Vars.Keys(i).NextPtr = 0&
    End If
#End If
    Vars.Enums.pa.sa.rgsabound0.cElements = 0
    Vars.Enums.pa.sa.pvData = 0
End Function

Private Sub RemoveUnusedEnums(Optional ByVal searchPtr As LongPtr _
                            , Optional ByVal replacePtr As LongPtr)
    #If Win64 Then
        Const nullPtr As LongLong = 0^
    #Else
        Const nullPtr As Long = 0&
    #End If
    Dim e As Variant 'IEnumVARIANT does not work with For Each
    With Vars.Enums
        For Each e In .enumsColl
            Dim ptr As LongPtr: ptr = ObjPtr(e)
            .pa.sa.pvData = ptr + nextItemOffset
            .pa.sa.rgsabound0.cElements = 1
            If .pa.ptrs(0) = nullPtr Then
                .enumsColl.Remove CStr(ptr)
            ElseIf .pa.ptrs(0) = searchPtr Then
                If replacePtr = nullPtr Then
                    .enumsColl.Remove CStr(ptr)
                Else
                    .pa.ptrs(0) = replacePtr
                End If
            End If
        Next e
        .hasEnum = (.enumsColl.Count > 0)
        .pa.sa.rgsabound0.cElements = 0
        .pa.sa.pvData = nullPtr
    End With
End Sub

Private Sub ShiftEnumPointers(ByVal addrShift As LongPtr)
    Dim e As Variant 'IEnumVARIANT does not work with For Each
    With Vars.Enums
        For Each e In .enumsColl
            .pa.sa.pvData = ObjPtr(e) + nextItemOffset
            .pa.sa.rgsabound0.cElements = 1
            .pa.ptrs(0) = .pa.ptrs(0) + addrShift
        Next e
    End With
    '
    Dim hasGaps As Boolean: hasGaps = (Vars.UsedCount < Vars.Count)
    Dim i As Long
    '
#If Win64 Then
    With Vars.Enums.pa
        .sa.pvData = VarPtr(Vars.Keys(0))
        .sa.rgsabound0.cElements = Vars.Count * 3
        If hasGaps Then
            For i = 5 To .sa.rgsabound0.cElements - 1 Step 3
                If .ptrs(i) <> 0^ Then .ptrs(i) = .ptrs(i) + addrShift
            Next i
        Else
            For i = 5 To .sa.rgsabound0.cElements - 1 Step 3
                .ptrs(i) = .ptrs(i) + addrShift
            Next i
        End If
    End With
#Else
    If hasGaps Then
        For i = 0 To Vars.Count - 2
            With Vars.Keys(i)
                If .NextPtr <> 0& Then .NextPtr = .NextPtr + addrShift
            End With
        Next i
    Else
        For i = 0 To Vars.Count - 2
            Vars.Keys(i).NextPtr = Vars.Keys(i).NextPtr + addrShift
        Next i
    End If
#End If
    Vars.Enums.pa.sa.rgsabound0.cElements = 0
    Vars.Enums.pa.sa.pvData = 0
End Sub
