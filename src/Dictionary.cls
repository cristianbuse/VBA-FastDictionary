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
''' VBA Mini Dictionary
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
''         - Preserves insertion order (default) unless 'FastUnorderedRemove' is
''           set to 'True'
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
''    * FastUnorderedRemove <Get/Let>
''         - Can only be changed if there are no stored items
''         - If set to 'False' (Default) then calling 'Remove' preserves the
''           order in which the Key-Item pairs were added but calling 'Items'
''           and 'Keys' will be slower if there are any gaps in the storage
''           arrays. If 'Remove' is not called then this setting has no effect
''         - If set to 'True' then calling 'Remove' ruins the order in which the
''           Key-Item pairs were added but calling 'Items' and 'Keys' will be
''           as fast as copying the entire array in one instruction
''    * IsOrdered <Get>
''         - Returns 'False' only if 'FastUnorderedRemove' is set to 'True' and
''           order was ruined by calling 'Remove'. Otherwise returns 'True'
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
        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" _
                Alias "memmove" (Destination As Any _
                               , Source As Any _
                               , ByVal Length As LongPtr) As LongPtr
    #Else
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" _
                Alias "memmove" (Destination As Any _
                               , Source As Any _
                               , ByVal Length As Long) As Long
    #End If
#Else 'Windows
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any _
                                     , Source As Any _
                                     , ByVal Length As LongPtr)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any _
                                     , Source As Any _
                                     , ByVal Length As Long)
    #End If
#End If
#If VBA7 = 0 Then
    Private Enum LongPtr
        [_]
    End Enum
#End If

Const MAX_LOAD_FACTOR As Single = 0.5
Const INITIAL_GROUP_COUNT As Long = 16

#If Win64 Then
    Const NULL_PTR As LongLong = 0^
    Const PTR_SIZE As Long = 8
    Const SIGN_BIT As LongLong = &H8000000000000000^
    Const HIGH_BIT_OFF As LongLong = &H7F7F7F7F7F7F7F7F^  'Bit Off Per Each Byte
    Const HIGH_BIT_ONLY As LongLong = &H8080808080808080^
    Const NHIGH_BIT_OFF As LongLong = HIGH_BIT_OFF Or SIGN_BIT
    Const BYTE_BROADCAST As LongLong = &H101010101010101^
    Const POSITION_PRIME As LongLong = 19
    Const VARIANT_SIZE As Long = 24
#Else
    Const NULL_PTR As Long = 0&
    Const PTR_SIZE As Long = 4
    Const SIGN_BIT As Long = &H80000000
    Const HIGH_BIT_OFF As Long = &H7F7F7F7F
    Const NHIGH_BIT_OFF As Long = HIGH_BIT_OFF Or SIGN_BIT
    Const HIGH_BIT_ONLY As Long = &H80808080
    Const BYTE_BROADCAST As Long = &H1010101
    Const POSITION_PRIME As Long = 7
    Const ENUM_VAR_SIZE As Long = 28
#End If
#If (Win64 = 0) Or Mac Then
    Const vbLongLong = 20
#End If

Const GROUP_SIZE As Long = PTR_SIZE
Const INTS_IN_DOUBLE As Long = 4
Const NOT_FOUND As Long = -1
Const NEXT_ITEM_OFFSET As Long = PTR_SIZE * 2
#If Win64 Then
    Const IS_ITEM_OBJ As Long = &H80000000
    Const KEY_MASK As Long = &H7FFFFFFF
#End If

Private Type Group
    Count As Long
    Index(0 To GROUP_SIZE - 1) As Long
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

Private Type DataStorage
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
End Type

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
    defInstance As Dictionary 'Avoids deallocation of default (Predeclared) dict
    isScrAvailable As Boolean
#End If
    isSet As Boolean
    ia As IntegerAccessor
    d As Double 'For hashing numbers
    dPtr As LongPtr
End Type

Private Type Lookups
    ByteShiftL(0 To GROUP_SIZE - 1) As LongPtr
    ByteMask(0 To GROUP_SIZE - 1) As LongPtr
    ByteOff(0 To GROUP_SIZE - 1) As LongPtr
    PositionMask(0 To GROUP_SIZE - 2) As LongPtr
    CountMask(0 To GROUP_SIZE) As LongPtr
#If Win64 Then
    ModBytePosition(0 To 15) As Long
#Else
    ModBytePosition(0 To 4) As Long
#End If
End Type

Private Type EnumProvider
    hasEnum As Boolean
    coll As New Collection
    pa As PointerAccessor
End Type

'Class members
Private m_compare As VbCompareMethod
Private m_data As DataStorage
Private m_enumProvider As EnumProvider
Private m_fastUnorderedRemove As Boolean
Private m_hash As HashMap
Private m_hasher As Hasher
Private m_isUnordered As Boolean
Private m_lookups As Lookups

Private Sub Init(Optional ByVal newGroupCount As Long = INITIAL_GROUP_COUNT)
    InitHashMap newGroupCount
    m_data.Count = 0
    m_data.UBound = 0
    m_data.UsedCount = 0
    m_isUnordered = False
    ReDim m_data.Items(0 To 0)
    ReDim m_data.Keys(0 To 0)
#If Win64 Then
    ReDim m_data.Meta(0 To 0)
#End If
    If Not m_hasher.isSet Then
        InitLookups
        InitLocalHasher
    End If
    Set m_enumProvider.coll = Nothing 'Clears NextPtr in all live IEnumVariant's
    m_enumProvider.hasEnum = False
End Sub

'Raises Error 457 if a duplicated key was specified
Public Sub Add(ByRef Key As Variant, ByRef Item As Variant)
    Dim hVal As Long
    Dim groupSlot As Long
    Dim i As Long
    Dim controlByte As Long
    '
    i = GetIndex(Key, hVal, groupSlot, controlByte)
    If i > NOT_FOUND Then Err.Raise 457, TypeName(Me) & ".Add"
    '
    If m_data.Count > m_data.UBound Then
        m_data.UBound = m_data.UBound * 2 + 1
        ReDim Preserve m_data.Items(0 To m_data.UBound)
        '
        If m_enumProvider.hasEnum Then
            Dim ptr As LongPtr: ptr = VarPtr(m_data.Keys(0))
            RemoveUnusedEnums
        End If
        #If Win64 Then 'Extra member for safe IEnumVariant
            ReDim Preserve m_data.Keys(0 To m_data.UBound + 1)
            ReDim Preserve m_data.Meta(0 To m_data.UBound)
        #Else
            ReDim Preserve m_data.Keys(0 To m_data.UBound)
        #End If
        If m_enumProvider.hasEnum Then
            ShiftEnumPointers VarPtr(m_data.Keys(0)) - ptr
        End If
    End If
    i = m_data.Count
    m_data.Count = m_data.Count + 1
    m_data.UsedCount = m_data.UsedCount + 1
    If m_data.UsedCount > m_hash.MaxLoad Then
        Rehash m_hash.GroupCount * 2
        groupSlot = hVal Mod m_hash.GroupCount
        Do While m_hash.Groups(groupSlot).WasEverFull 'Unlikely after rehash
             groupSlot = (groupSlot + 1) Mod m_hash.GroupCount
        Loop
        controlByte = (hVal And m_hash.ControlMask) \ m_hash.GroupCount
    ElseIf m_hash.Groups(groupSlot).Count = GROUP_SIZE Then
        'This can only happen if ClearMapIndex was previously called
        groupSlot = hVal Mod m_hash.GroupCount
        Do While m_hash.Groups(groupSlot).Count = GROUP_SIZE
             groupSlot = (groupSlot + 1) Mod m_hash.GroupCount
        Loop
    End If
    '
    With m_hash.Groups(groupSlot)
        .Index(.Count) = i
        .Control = .Control Or (controlByte * m_lookups.ByteShiftL(.Count))
        .Count = .Count + 1
        .WasEverFull = .WasEverFull Or (.Count = GROUP_SIZE)
    End With
    '
#If Win64 Then
    If hVal And hmObject Then Set m_data.Keys(i) = Key Else m_data.Keys(i) = Key
    If m_enumProvider.hasEnum Then
        RemoveUnusedEnums
        If i > 0 Then
            ptr = VarPtr(m_data.Keys(i))
            m_enumProvider.pa.sa.pvData = ptr
            m_enumProvider.pa.sa.rgsabound0.cElements = 3
            m_enumProvider.pa.ptrs(2) = ptr
        End If
    End If
    If IsObject(Item) Then
        m_data.Meta(i) = hVal Or IS_ITEM_OBJ
        Set m_data.Items(i) = Item
    Else
        m_data.Meta(i) = hVal
        m_data.Items(i) = Item
    End If
#Else
    With m_data.Keys(i)
        .IsKeyObj = CBool(hVal And hmObject)
        If .IsKeyObj Then Set .Value = Key Else .Value = Key
        If m_enumProvider.hasEnum Then
            RemoveUnusedEnums
            If i > 0 Then m_data.Keys(i - 1).NextPtr = VarPtr(.Value)
        End If
        .IsItemObj = IsObject(Item)
        If .IsItemObj Then
            Set m_data.Items(i) = Item
        Else
            m_data.Items(i) = Item
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
    For i = 0 To m_data.Count - 1
        #If Win64 Then
            hVal = m_data.Meta(i)
        #Else
            hVal = m_data.Keys(i).Meta
        #End If
        If hVal <> hmRemoved Then
            groupSlot = hVal And m_hash.GroupMask
            Do While m_hash.Groups(groupSlot).WasEverFull 'Unlikely after resize
                groupSlot = (groupSlot + 1) Mod m_hash.GroupCount
            Loop
            With m_hash.Groups(groupSlot)
                .Index(.Count) = i
                ctrlByte = (hVal And m_hash.ControlMask) \ m_hash.GroupCount
                .Control = .Control Or (ctrlByte * m_lookups.ByteShiftL(.Count))
                .Count = .Count + 1
                .WasEverFull = .WasEverFull Or (.Count = GROUP_SIZE)
            End With
        End If
    Next i
End Sub
Private Sub InitHashMap(ByVal newGroupCount As Long)
    m_hash.GroupCount = newGroupCount
    ReDim m_hash.Groups(0 To m_hash.GroupCount - 1)
    m_hash.MaxLoad = m_hash.GroupCount * GROUP_SIZE * MAX_LOAD_FACTOR
    m_hash.GroupMask = m_hash.GroupCount - 1
    '&H80& bit cannot be multiplied with BYTE_BROADCAST without causing overflow
    m_hash.ControlMask = m_hash.GroupCount * &H7F&
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
Private Function GetIndex(ByRef Key As Variant _
                        , Optional ByRef hVal As Long _
                        , Optional ByRef groupSlot As Long _
                        , Optional ByRef controlByte As Long _
                        , Optional ByRef posInGroup As Long _
                        , Optional ByVal hashOnly As Boolean = False) As Long
    If m_hash.GroupCount = 0 Then Init
    Dim vt As VbVarType
    Dim i As Long
    '
    'Hash value
    If IsObject(Key) Then vt = vbObject Else vt = VarType(Key)
    If vt = vbString Then
        Const tPrime As Long = 131&
        Const tOverflowMask As Long = &H7FFFFF
#If Windows Then
        If m_hasher.isScrAvailable Then
            Const minLenBinary As Long = 6
            If m_compare = vbBinaryCompare Then
                m_hasher.ia.sa.rgsabound0.cElements = Len(Key)
                If minLenBinary < m_hasher.ia.sa.rgsabound0.cElements Then
                    m_hasher.compLcid(posComp) = vbBinaryCompare
                    hVal = m_hasher.textHasher(0).Item(Key) Or hmText 'Early-binded
                Else 'Faster to loop integers
                    m_hasher.ia.sa.pvData = StrPtr(Key)
                    With m_hasher.ia
                        For i = 0 To m_hasher.ia.sa.rgsabound0.cElements - 1
                            hVal = (hVal And tOverflowMask) * tPrime + .ints(i)
                        Next i
                    End With
                    hVal = hVal And [_maskHM] Or hmText
                    m_hasher.ia.sa.rgsabound0.cElements = INTS_IN_DOUBLE
                    m_hasher.ia.sa.pvData = m_hasher.dPtr
                End If
            Else
                m_hasher.compLcid(posComp) = vbTextCompare
                m_hasher.compLcid(posLcid) = m_hasher.cachedLcid
                hVal = m_hasher.textHasher(0).Item(Key) Or hmText 'Early-binded
            End If
        Else
#End If
            m_hasher.ia.sa.rgsabound0.cElements = Len(Key)
            If m_compare = vbBinaryCompare Then
                m_hasher.ia.sa.pvData = StrPtr(Key)
            Else
                Dim s As String: s = LCase$(Key)
                m_hasher.ia.sa.pvData = StrPtr(s)
            End If
            With m_hasher.ia
                For i = 0 To m_hasher.ia.sa.rgsabound0.cElements - 1
                    hVal = (hVal And tOverflowMask) * tPrime + .ints(i)
                Next i
            End With
            hVal = hVal And [_maskHM] Or hmText
            m_hasher.ia.sa.rgsabound0.cElements = INTS_IN_DOUBLE
            m_hasher.ia.sa.pvData = m_hasher.dPtr
#If Windows Then
        End If
#End If
    ElseIf vt = vbObject Or vt = vbDataObject Then
        Const oPrime As Long = 2701&
        Const oPreMask As Long = &H6FFFFFFF
        Static iUnk As stdole.IUnknown 'Dim is slower
        '
        vt = vbObject 'Replace vbDataObject if needed
        Set iUnk = Key
        hVal = CLng(ObjPtr(iUnk) And oPreMask) 'Ignores high bits on x64
        hVal = (hVal + hVal Mod oPrime) And [_maskHM] Or hmObject
        Set iUnk = Nothing 'Must call because of Static but still faster
    ElseIf vt > vbLongLong Then
        Err.Raise 5, , "Cannot hash an Array or User Defined Type"
    ElseIf vt > vbNull Then
        Dim m As HashMeta
        If vt = vbError Then
            m = hmError
            m_hasher.d = CDbl(Key)
        Else
            m = hmNumber
            m_hasher.d = Key
        End If
        Const n1& = 1201, n2& = 2701, n3& = 131, n4& = 28571
        With m_hasher.ia
            hVal = (.ints(0) * n4 + .ints(1) * n3 _
                  + .ints(2) * n2 + .ints(3) * n1) And HashMeta.[_maskHM] Or m
        End With
    Else
        hVal = vt 'vbEmpty (0) or vbNull (1)
    End If
    If hashOnly Then Exit Function
    '
    'Calculate sub-hashes
    groupSlot = hVal Mod m_hash.GroupCount
    controlByte = (hVal And m_hash.ControlMask) \ m_hash.GroupCount
    '
    Dim matches As LongPtr
    Dim cMask As LongPtr
    Dim startSlot As Long: startSlot = groupSlot
    Do
        With m_hash.Groups(groupSlot)
            cMask = m_lookups.CountMask(.Count)
            'Match bytes adapted from:
            'https://graphics.stanford.edu/~seander/bithacks.html#ZeroInWord
            matches = .Control Xor (controlByte * BYTE_BROADCAST)
            matches = ((((matches And HIGH_BIT_OFF) + NHIGH_BIT_OFF) _
                    Xor SIGN_BIT) Or matches) And cMask Xor cMask
            Do While matches
                If matches = SIGN_BIT Then 'Matched last element only
                    posInGroup = GROUP_SIZE - 1
                    matches = 0
                Else 'Find next match within group
                    posInGroup = m_lookups.ModBytePosition(CLng( _
                                 (-matches And matches) Mod POSITION_PRIME))
                    matches = matches Xor m_lookups.PositionMask(posInGroup)
                End If
                GetIndex = .Index(posInGroup)
#If Win64 Then
                If (m_data.Meta(GetIndex) And KEY_MASK) = hVal Then
                    If vt = vbString And (m_compare > vbBinaryCompare) Then
                        If StrComp(m_data.Keys(GetIndex) _
                                 , Key, m_compare) = 0 Then Exit Function
                    ElseIf vt = vbObject Then 'vbDataObject was replaced
                        If m_data.Keys(GetIndex) Is Key Then Exit Function
                    ElseIf hVal <= vbNull Then 'Already matched
                        Exit Function
                    Else 'Number, Error or case-sensitive Text
                        If m_data.Keys(GetIndex) = Key Then Exit Function
                    End If
                End If
#Else
                If m_data.Keys(GetIndex).Meta = hVal Then
                    If vt = vbString And (m_compare > vbBinaryCompare) Then
                        If StrComp(m_data.Keys(GetIndex).Value _
                                 , Key, m_compare) = 0 Then Exit Function
                    ElseIf vt = vbObject Then 'vbDataObject was replaced
                        If m_data.Keys(GetIndex).Value Is Key Then Exit Function
                    ElseIf hVal <= vbNull Then 'Already matched
                        Exit Function
                    Else 'Number, Error or case-sensitive Text
                        If m_data.Keys(GetIndex).Value = Key Then Exit Function
                    End If
                End If
#End If
            Loop
            If Not .WasEverFull Then Exit Do
        End With
        groupSlot = (groupSlot + 1) Mod m_hash.GroupCount
        If groupSlot = startSlot Then Exit Do 'Avoid infinite loop
    Loop
    GetIndex = NOT_FOUND
End Function

'Values greater than 2 can be used to compare using specific Locale IDs (LCID)
'https://learn.microsoft.com/en-us/office/vba/Language/Reference/User-Interface-Help/comparemode-property
'Can only be changed if there are no stored items
Public Property Get CompareMode() As VbCompareMethod
    CompareMode = m_compare
End Property
Public Property Let CompareMode(ByVal compMode As VbCompareMethod)
    Const minMode As Long = 0
    Const maxMode As Long = 31890 'Same as StrComp limits
    '
    If m_data.UsedCount > 0 And compMode <> m_compare Then
        Err.Raise 5, TypeName(Me) & ".CompareMode", "Dict already contains data"
    ElseIf compMode < minMode Or compMode > maxMode Then
        Err.Raise 5, TypeName(Me) & ".CompareMode", "Invalid compare method"
    End If
    m_compare = compMode
#If Windows Then
    If m_compare > vbTextCompare Then
        m_hasher.cachedLcid = m_compare
    Else
        m_hasher.cachedLcid = m_hasher.lcid
    End If
#End If
End Property

Public Property Get Count() As Long
    Count = m_data.UsedCount
End Property

'Can only be changed if there are no stored items
'When set to True, the order in which the Items/Keys were added via 'Add' is no
'   longer preserved to facilitate faster 'Items' and 'Keys' calls
Public Property Get FastUnorderedRemove() As Boolean
    FastUnorderedRemove = m_fastUnorderedRemove
End Property
Public Property Let FastUnorderedRemove(ByVal fastOn As Boolean)
    If m_data.UsedCount > 0 And fastOn <> m_fastUnorderedRemove Then
        Err.Raise 5, TypeName(Me) & ".CompareMode", "Dict already contains data"
    End If
    m_fastUnorderedRemove = fastOn
End Property

Public Function Exists(ByRef Key As Variant) As Boolean
    Exists = (GetIndex(Key) > NOT_FOUND)
End Function

Public Property Get IsOrdered() As Boolean
    IsOrdered = Not m_isUnordered
End Property

'Raises Error:
'   -    9: invalid key (key is not associated with any element)
'   -  450: 'Set' is missing when assigning an object
'@DefaultMember
Public Property Get Item(ByRef Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_UserMemId = 0
    Dim i As Long: i = GetIndex(Key)
    '
    'A check like i = NOT_FOUND is avoided for speed purposes
    '   because error 9 is raised anyway when accessing the data arrays
#If Win64 Then
    If m_data.Meta(i) And IS_ITEM_OBJ Then
#Else
    If m_data.Keys(i).IsItemObj Then
#End If
        Set Item = m_data.Items(i)
    Else
        Item = m_data.Items(i)
    End If
End Property
Public Property Let Item(ByRef Key As Variant, ByRef Item As Variant)
    If IsObject(Item) Or (VarType(Item) = vbDataObject) Then
        Err.Raise 450, TypeName(Me) & ".Item", "'Set' is required"
    End If
    Dim i As Long: i = GetIndex(Key)
    If i > NOT_FOUND Then
        #If Win64 Then
            m_data.Meta(i) = m_data.Meta(i) And KEY_MASK
        #Else
            m_data.Keys(i).IsItemObj = False
        #End If
        m_data.Items(i) = Item
    Else
        Add Key, Item
    End If
End Property
Public Property Set Item(ByRef Key As Variant, ByRef Item As Object)
    Dim i As Long: i = GetIndex(Key)
    If i > NOT_FOUND Then
        #If Win64 Then
            m_data.Meta(i) = m_data.Meta(i) Or IS_ITEM_OBJ
        #Else
            m_data.Keys(i).IsItemObj = True
        #End If
        Set m_data.Items(i) = Item
    Else
        Add Key, Item
    End If
End Property

Public Function Items() As Variant()
    If m_data.UsedCount = 0 Then
        Items = Array()
        Exit Function
    End If
    If m_data.UsedCount = m_data.Count Then
        Items = m_data.Items
        ReDim Preserve Items(0 To m_data.Count - 1)
    Else
        Dim res() As Variant
        ReDim res(0 To m_data.UsedCount - 1)
        Dim i As Long
        Dim j As Long
        '
        For i = 0 To m_data.Count - 1
#If Win64 Then
            If m_data.Meta(i) <> hmRemoved Then
                If m_data.Meta(i) And IS_ITEM_OBJ Then
#Else
            If m_data.Keys(i).Meta <> hmRemoved Then
                If m_data.Keys(i).IsItemObj Then
#End If
                    Set res(j) = m_data.Items(i)
                Else
                    res(j) = m_data.Items(i)
                End If
                j = j + 1
            End If
        Next i
        Items = res
    End If
End Function

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
    If GetIndex(NewKey, hVal, , controlByte) > NOT_FOUND Then Err.Raise 457
    i = GetIndex(OldKey, , groupSlot, , posInGroup)
    '
    'A check like i = NOT_FOUND is avoided for speed purposes
    '   because error 9 is raised anyway when accessing the Keys array
    '
    'Replace Key and Meta while preserving Item
#If Win64 Then
    If m_enumProvider.hasEnum Then
        RemoveUnusedEnums
        m_enumProvider.pa.sa.pvData = VarPtr(m_data.Keys(i))
        m_enumProvider.pa.sa.rgsabound0.cElements = 3
        Dim ptr As LongPtr: ptr = m_enumProvider.pa.ptrs(2)
    End If
    If hVal And HashMeta.hmObject Then
        Set m_data.Keys(i) = NewKey
    Else
        m_data.Keys(i) = NewKey
    End If
    If m_enumProvider.hasEnum Then m_enumProvider.pa.ptrs(2) = ptr
    If m_data.Meta(i) And IS_ITEM_OBJ Then hVal = hVal Or IS_ITEM_OBJ
    m_data.Meta(i) = hVal
#Else
    With m_data.Keys(i)
        .IsKeyObj = CBool(hVal And HashMeta.hmObject)
        If .IsKeyObj Then Set .Value = NewKey Else .Value = NewKey
        .Meta = hVal
    End With
#End If
    '
    ClearMapIndex groupSlot, posInGroup
    '
    'Update hash map with new key
    groupSlot = hVal Mod m_hash.GroupCount
    Do While m_hash.Groups(groupSlot).Count = GROUP_SIZE
         groupSlot = (groupSlot + 1) Mod m_hash.GroupCount
    Loop
    With m_hash.Groups(groupSlot)
        .Index(.Count) = i
        .Control = .Control Or (controlByte * m_lookups.ByteShiftL(.Count))
        .Count = .Count + 1
        .WasEverFull = .WasEverFull Or (.Count = GROUP_SIZE)
    End With
End Property

Private Sub ClearMapIndex(ByRef groupSlot As Long, ByRef posInGroup As Long)
    Dim lastPos As Long
    '
    With m_hash.Groups(groupSlot)
        lastPos = .Count - 1
        If posInGroup < lastPos Then 'Swap
            .Index(posInGroup) = .Index(lastPos)
            .Control = .Control And m_lookups.ByteOff(posInGroup)
            .Control = .Control Or ((.Control And m_lookups.ByteMask(lastPos)) _
                                   \ m_lookups.ByteShiftL(lastPos - posInGroup))
        End If
        .Index(lastPos) = 0
        .Control = .Control And m_lookups.ByteOff(lastPos)
        .Count = .Count - 1
    End With
End Sub

Public Function Keys() As Variant()
    If m_data.UsedCount = 0 Then
        Keys = Array()
        Exit Function
    End If
    '
    Dim res() As Variant
    Dim i As Long
    Dim j As Long
    '
#If Win64 Then
    If m_data.UsedCount = m_data.Count Then
        Keys = m_data.Keys
        ReDim Preserve Keys(0 To m_data.Count - 1)
    Else
        ReDim res(0 To m_data.UsedCount - 1)
        For i = 0 To m_data.Count - 1
            If m_data.Meta(i) <> hmRemoved Then
                If m_data.Meta(i) And hmObject Then
                    Set res(j) = m_data.Keys(i)
                Else
                    res(j) = m_data.Keys(i)
                End If
                j = j + 1
            End If
        Next i
        Keys = res
    End If
#Else
    ReDim res(0 To m_data.UsedCount - 1)
    For i = 0 To m_data.Count - 1
        With m_data.Keys(i)
            If .Meta <> hmRemoved Then
                If .IsKeyObj Then Set res(j) = .Value Else res(j) = .Value
                j = j + 1
            End If
        End With
    Next i
    Keys = res
#End If
End Function

Public Sub PredictCount(ByVal expectedCount As Long)
    Const maxGroups As Long = HashMeta.[_modHM] / GROUP_SIZE
    Dim expectedGroups As Long
    '
    If expectedCount <= 0 Then Exit Sub
    expectedGroups = 2 ^ -Int(-(Log(expectedCount / GROUP_SIZE) _
                              + Log(1 / MAX_LOAD_FACTOR)) / Log(2))
    If expectedGroups > maxGroups Then expectedGroups = maxGroups
    '
    If m_hash.GroupCount = 0 Then
        If expectedGroups > INITIAL_GROUP_COUNT Then Init expectedGroups
    ElseIf expectedGroups > m_hash.GroupCount * 2 Then
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
    If i = NOT_FOUND Then Err.Raise 9, TypeName(Me) & ".Remove"
    If m_data.UsedCount = 1 Then
        Init
        Exit Sub
    End If
    '
    ClearMapIndex groupSlot, posInGroup
    '
    lastPos = m_data.Count - 1
    m_data.UsedCount = m_data.UsedCount - 1
    If i < lastPos Then
        If m_fastUnorderedRemove Then
            m_isUnordered = True
            'Swap data
            #If Win64 Then
                hVal = m_data.Meta(lastPos)
                If hVal And IS_ITEM_OBJ Then
                    Set m_data.Items(i) = m_data.Items(lastPos)
                Else
                    m_data.Items(i) = m_data.Items(lastPos)
                End If
                If m_enumProvider.hasEnum Then
                    RemoveUnusedEnums VarPtr(m_data.Keys(lastPos)), NULL_PTR
                    m_enumProvider.pa.sa.pvData = VarPtr(m_data.Keys(i))
                    m_enumProvider.pa.sa.rgsabound0.cElements = 3
                    ptr = m_enumProvider.pa.ptrs(2)
                End If
                If hVal And HashMeta.hmObject Then
                    Set m_data.Keys(i) = m_data.Keys(lastPos)
                Else
                    m_data.Keys(i) = m_data.Keys(lastPos)
                End If
                If m_enumProvider.hasEnum Then m_enumProvider.pa.ptrs(2) = ptr
                m_data.Meta(i) = hVal
            #Else
                ptr = m_data.Keys(i).NextPtr
                m_data.Keys(i) = m_data.Keys(lastPos)
                m_data.Keys(i).NextPtr = ptr
                If m_enumProvider.hasEnum Then
                    RemoveUnusedEnums VarPtr(m_data.Keys(lastPos)), NULL_PTR
                    m_data.Keys(lastPos - 1).NextPtr = NULL_PTR
                End If
                With m_data.Keys(i)
                    If .IsItemObj Then
                        Set m_data.Items(i) = m_data.Items(lastPos)
                    Else
                        m_data.Items(i) = m_data.Items(lastPos)
                    End If
                    hVal = .Meta
                End With
            #End If
            '
            'Update index inside hash map
            groupSlot = hVal And m_hash.GroupMask
            Do
                With m_hash.Groups(groupSlot)
                    For j = 0 To .Count - 1
                        If .Index(j) = lastPos Then
                            .Index(j) = i
                            Exit Do
                        End If
                    Next j
                End With
                groupSlot = (groupSlot + 1) Mod m_hash.GroupCount
            Loop
        Else 'Clear data and register removal
            m_data.Items(i) = Empty
            #If Win64 Then
                m_data.Keys(i) = Empty
                m_data.Meta(i) = hmRemoved
            #Else
                m_data.Keys(i).Value = Empty
                m_data.Keys(i).Meta = hmRemoved
            #End If
            If m_enumProvider.hasEnum Then 'Link previous to next
                j = i + 1
                #If Win64 Then
                    Do While m_data.Meta(j) = hmRemoved: j = j + 1: Loop
                #Else
                    Do While m_data.Keys(j).Meta = hmRemoved: j = j + 1: Loop
                #End If
                ptr = VarPtr(m_data.Keys(j))
                RemoveUnusedEnums VarPtr(m_data.Keys(i)), ptr
                j = i - 1
                Do While j >= 0
                    #If Win64 Then
                        If m_data.Meta(j) <> hmRemoved Then Exit Do
                    #Else
                        If m_data.Keys(j).Meta <> hmRemoved Then Exit Do
                    #End If
                    j = j - 1
                Loop
                If j >= 0 Then
                    #If Win64 Then
                        m_enumProvider.pa.sa.pvData = VarPtr(m_data.Keys(j + 1))
                        m_enumProvider.pa.sa.rgsabound0.cElements = 3
                        m_enumProvider.pa.ptrs(2) = ptr
                    #Else
                        m_data.Keys(j).NextPtr = ptr
                    #End If
                End If
            End If
            Exit Sub
        End If
    Else
        If m_data.Count > m_data.UsedCount Then
            i = lastPos - 1
            #If Win64 Then
                Do While m_data.Meta(i) = hmRemoved: i = i - 1: Loop
            #Else
                Do While m_data.Keys(i).Meta = hmRemoved: i = i - 1: Loop
            #End If
            m_data.Count = m_data.Count - lastPos + i + 1
        End If
        If m_enumProvider.hasEnum Then
            RemoveUnusedEnums VarPtr(m_data.Keys(lastPos)), NULL_PTR
            #If Win64 Then
                m_data.Keys(i + 1) = Empty 'Just to clear the pointer
            #Else
                m_data.Keys(i).NextPtr = NULL_PTR
            #End If
        End If
    End If
    'Clear Variants in case deallocation is needed e.g. String/Object
    m_data.Items(lastPos) = Empty
    #If Win64 Then
        m_data.Keys(lastPos) = Empty
    #Else
        m_data.Keys(lastPos).Value = Empty
    #End If
    m_data.Count = m_data.Count - 1
    m_isUnordered = m_isUnordered And (m_data.UsedCount > 1)
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
    If m_hash.GroupCount = 0 Then Exit Property
    LoadFactor = m_data.UsedCount / (m_hash.GroupCount * GROUP_SIZE)
End Property

Private Sub InitLookups()
    Dim i As Long
    '
    With m_lookups
        .ByteShiftL(0) = 1
        .ByteMask(0) = &H7F
        .ByteOff(0) = Not .ByteMask(0)
        For i = 1 To GROUP_SIZE - 1
            .PositionMask(i - 1) = .ByteShiftL(i - 1) * &H80
            .ByteShiftL(i) = .ByteShiftL(i - 1) * &H100
            .CountMask(i) = .PositionMask(i - 1) + .CountMask(i - 1)
            .ByteMask(i) = .ByteMask(i - 1) * &H100
            .ByteOff(i) = Not .ByteMask(i)
        Next i
        .CountMask(GROUP_SIZE) = HIGH_BIT_ONLY
        '
        'These are the relevant values if applying Modulo POSITION_PRIME
    #If Win64 Then
        .ModBytePosition(12) = 1
        .ModBytePosition(13) = 2
        .ModBytePosition(3) = 3
        .ModBytePosition(8) = 4
        .ModBytePosition(15) = 5
        .ModBytePosition(2) = 6
    #Else
        .ModBytePosition(1) = 1
        .ModBytePosition(4) = 2
    #End If
    End With
End Sub

'Initializes an Integer Array Accessor on Win and Mac
'Initializes a Collection that can call ScriptingDictionary.HashVal with early
'   binding speed (Win only)
Private Sub InitLocalHasher()
    Const INT_SIZE As Long = 2
    InitSafeArray m_hasher.ia.sa, INT_SIZE
#If Mac Then
    CopyMemory ByVal VarPtr(m_hasher.ia), VarPtr(m_hasher.ia.sa), PTR_SIZE
#Else
    Dictionary.InitHasher m_hasher
    CompareMode = m_compare
#End If
    With m_hasher
        .dPtr = VarPtr(.d)
        .ia.sa.pvData = .dPtr
        .ia.sa.rgsabound0.cElements = INTS_IN_DOUBLE
        .isSet = True
    End With
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

'To avoid API calls overhead in VBA7, a Pointer Accessor is cached in the
'   default instance (Attribute VB_PredeclaredId = True) thus allowing faster
'   memory manipulation
'The best alternative was to use a standard .bas module like 'LibMemory' at:
'   https://github.com/cristianbuse/VBA-MemoryTools
#If Windows Then
Friend Sub InitHasher(ByRef h As Hasher)
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
    '
    If Not Me Is Dictionary Then Exit Sub
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
        h.isScrAvailable = Not (temp Is Nothing)
        If h.isScrAvailable Then
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
            Const LONG_SIZE As Long = 4
            InitSafeArray saCL, LONG_SIZE
            saCL.pvData = VarPtr(sdl.compMode)
            saCL.rgsabound0.cElements = 2
            '
            'Init safe array for collection
            InitSafeArray saF, PTR_SIZE
            saF.pvData = fakePtr
            saF.rgsabound0.cElements = 1
        Else
            pa.sa.pvData = fakePtr
            pa.sa.rgsabound0.cElements = 1
        End If
    End If
    If h.isScrAvailable Then
        h.lcid = lcid
        Set h.defInstance = Me 'Avoid deallocation of this instance
        pa.sa.pvData = VarPtr(h)
        pa.ptrs(0) = VarPtr(saF)
        '
        'Init compMode and lcid accesor
        pa.sa.pvData = pa.sa.pvData + PTR_SIZE
        pa.ptrs(0) = VarPtr(saCL)
    End If
    '
    'Init integer accesor
    pa.sa.pvData = VarPtr(h.ia)
    pa.ptrs(0) = VarPtr(h.ia.sa)
    '
    pa.sa.pvData = fakePtr
End Sub
#End If

Friend Sub InitPointerAccessor(ByRef paToInit As PointerAccessor)
    Static pa As PointerAccessor
    '
    If Not pa.isSet Then
        InitSafeArray pa.sa, PTR_SIZE
        CopyMemory ByVal VarPtr(pa), VarPtr(pa.sa), PTR_SIZE 'The only API call
        pa.sa.pvData = VarPtr(pa.sa) 'Some valid address
        pa.sa.rgsabound0.cElements = 1
        pa.isSet = True
    End If
    '
    InitSafeArray paToInit.sa, PTR_SIZE
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

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_UserMemId = -4
    Dim hadEnum As Boolean
    '
    With m_enumProvider
        hadEnum = .hasEnum
        If .hasEnum Then RemoveUnusedEnums
        Set NewEnum = .coll.[_NewEnum]
        .hasEnum = True
        .coll.Add NewEnum, CStr(ObjPtr(NewEnum))
        If Not .pa.isSet Then Dictionary.InitPointerAccessor .pa
    End With
    If m_data.UsedCount = 0 Then Exit Function
    '
    Dim i As Long
    Dim hasGaps As Boolean: hasGaps = (m_data.UsedCount < m_data.Count)
    '
    If hasGaps Then
        #If Win64 Then
            Do While m_data.Meta(i) = hmRemoved: i = i + 1: Loop
        #Else
            Do While m_data.Keys(i).Meta = hmRemoved: i = i + 1: Loop
        #End If
    End If
    With m_enumProvider.pa
        .sa.pvData = ObjPtr(NewEnum) + NEXT_ITEM_OFFSET
        .sa.rgsabound0.cElements = 1
        .ptrs(0) = VarPtr(m_data.Keys(i))
    End With
    If hadEnum Then Exit Function
    '
    Dim ptr As LongPtr: ptr = VarPtr(m_data.Keys(0))
    Dim j As Long
    '
#If Win64 Then
    With m_enumProvider.pa
        .sa.pvData = ptr
        .sa.rgsabound0.cElements = (m_data.Count + 1) * 3
        If hasGaps Then
            j = i * 3 + 5
            For i = i + 1 To m_data.Count - 1
                If m_data.Meta(i) <> hmRemoved Then
                    .ptrs(j) = ptr + VARIANT_SIZE * i
                    j = i * 3 + 5
                End If
            Next i
            .ptrs(j) = NULL_PTR
        Else
            For i = 5 To .sa.rgsabound0.cElements - 4 Step 3
                ptr = ptr + VARIANT_SIZE
                .ptrs(i) = ptr
            Next i
            .ptrs(i) = NULL_PTR
        End If
    End With
#Else
    If hasGaps Then
        j = i
        For i = i + 1 To m_data.Count - 1
            If m_data.Keys(i).Meta <> hmRemoved Then
                m_data.Keys(j).NextPtr = ptr + ENUM_VAR_SIZE * i
                j = i
            End If
        Next i
        m_data.Keys(j).NextPtr = NULL_PTR
    Else
        ptr = ptr + ENUM_VAR_SIZE
        For i = 0 To m_data.Count - 2
            m_data.Keys(i).NextPtr = ptr + ENUM_VAR_SIZE * i
        Next i
        m_data.Keys(i).NextPtr = NULL_PTR
    End If
#End If
    m_enumProvider.pa.sa.rgsabound0.cElements = 0
    m_enumProvider.pa.sa.pvData = NULL_PTR
End Function

Private Sub RemoveUnusedEnums(Optional ByVal searchPtr As LongPtr _
                            , Optional ByVal replacePtr As LongPtr)
    Dim e As Variant 'IEnumVARIANT does not work with For Each
    With m_enumProvider
        For Each e In .coll
            Dim ptr As LongPtr: ptr = ObjPtr(e)
            .pa.sa.pvData = ptr + NEXT_ITEM_OFFSET
            .pa.sa.rgsabound0.cElements = 1
            If .pa.ptrs(0) = NULL_PTR Then
                .coll.Remove CStr(ptr)
            ElseIf .pa.ptrs(0) = searchPtr Then
                If replacePtr = NULL_PTR Then
                    .coll.Remove CStr(ptr)
                Else
                    .pa.ptrs(0) = replacePtr
                End If
            End If
        Next e
        .hasEnum = (.coll.Count > 0)
        .pa.sa.rgsabound0.cElements = 0
        .pa.sa.pvData = NULL_PTR
    End With
End Sub

Private Sub ShiftEnumPointers(ByVal addrShift As LongPtr)
    Dim e As Variant 'IEnumVARIANT does not work with For Each
    With m_enumProvider
        For Each e In .coll
            .pa.sa.pvData = ObjPtr(e) + NEXT_ITEM_OFFSET
            .pa.sa.rgsabound0.cElements = 1
            .pa.ptrs(0) = .pa.ptrs(0) + addrShift
        Next e
    End With
    '
    Dim hasGaps As Boolean: hasGaps = (m_data.UsedCount < m_data.Count)
    Dim i As Long
    '
#If Win64 Then
    With m_enumProvider.pa
        .sa.pvData = VarPtr(m_data.Keys(0))
        .sa.rgsabound0.cElements = m_data.Count * 3
        If hasGaps Then
            For i = 5 To .sa.rgsabound0.cElements - 1 Step 3
                If .ptrs(i) <> NULL_PTR Then .ptrs(i) = .ptrs(i) + addrShift
            Next i
        Else
            For i = 5 To .sa.rgsabound0.cElements - 1 Step 3
                .ptrs(i) = .ptrs(i) + addrShift
            Next i
        End If
    End With
#Else
    If hasGaps Then
        For i = 0 To m_data.Count - 2
            With m_data.Keys(i)
                If .NextPtr <> NULL_PTR Then .NextPtr = .NextPtr + addrShift
            End With
        Next i
    Else
        For i = 0 To m_data.Count - 2
            m_data.Keys(i).NextPtr = m_data.Keys(i).NextPtr + addrShift
        Next i
    End If
#End If
    m_enumProvider.pa.sa.rgsabound0.cElements = 0
    m_enumProvider.pa.sa.pvData = NULL_PTR
End Sub
