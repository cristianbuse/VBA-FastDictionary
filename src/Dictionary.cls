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
''    * Windows:
''         - Scripting.Dictionary replacement
''         - Performant when dealing with many Items
''         - Performant when dealing with large String Keys or Object Keys
''    * Mac:
''         - Supported
'' Methods:
''    * Add
''         - Adds a new Key-Item pair
''         - Keys are any data type except Arrays and User-Defined Types (UDTs)
''    * Exists
''         - Checks if a specified Key exists
''    * Items
''         - Returns a 1D array of all the Items
''    * Keys
''         - Returns a 1D array of all the Keys
''    * Remove
''         - Removes an item by Key
''    * RemoveAll
''         - Removes all Items
''    * Self
''         - Self instance - useful in 'With New Dictionary' blocks
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
''           and 'Keys' will be slower. If 'Remove' is not called then this
''           setting has no effect
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
''    * Key <Let>
''         - Allows a Key value to be changed while preserving the Item
''         - If the Key to be replaced does not exist then an error is raised
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
        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal length As Long) As Long
    #End If
#Else 'Windows
    #If VBA7 Then
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    #Else
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
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
    Const PTR_SIZE As Long = 8
    Const SIGN_BIT As LongLong = &H8000000000000000^
    Const HIGH_BIT_OFF As LongLong = &H7F7F7F7F7F7F7F7F^  'Per Byte
    Const HIGH_BIT_ONLY As LongLong = &H8080808080808080^ 'Per Byte
    Const BYTE_BROADCAST As LongLong = &H101010101010101^
    Const POSITION_PRIME As LongLong = 19
#Else
    Const PTR_SIZE As Long = 4
    Const SIGN_BIT As Long = &H80000000
    Const HIGH_BIT_OFF As Long = &H7F7F7F7F
    Const HIGH_BIT_ONLY As Long = &H80808080
    Const BYTE_BROADCAST As Long = &H1010101
    Const POSITION_PRIME As Long = 7
#End If
#If Win64 = 0 Or Mac Then
    Const vbLongLong = 20
#End If

Const GROUP_SIZE As Long = PTR_SIZE
Const NHIGH_BIT_OFF = HIGH_BIT_OFF Or SIGN_BIT
Const INTS_IN_DOUBLE As Long = 4
Const NOT_FOUND As Long = -1
Const IS_ITEM_OBJ As Long = &H80000000
Const KEY_MASK As Long = &H7FFFFFFF

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

Private Type DataStorage
    Keys() As Variant
    Items() As Variant
    Meta() As Long
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

#If Mac = 0 Then
Private Type ScrDictLayout 'Scripting.Dictionary memory layout
    vTables(0 To 3) As LongPtr
    unkPtr1 As LongPtr
    refCount As Long
    firstItemPtr As LongPtr
    lastItemPtr As LongPtr
#If Win64 = 0 Then
    dummy As Long
#End If
    hashTablePtr As LongPtr
    hashModulo As Long
    compMode As Long
    localeID As Long
    unkPtr2 As LongPtr
    unkPtr3 As LongPtr
End Type
#End If

Private Type IntegerAccesor
    ints() As Integer
End Type
#If Mac = 0 Then
    Private Type PointerAccesor
        ptrs() As LongPtr
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

'Minimal version of MEMORY_ACCESOR:
'https://github.com/cristianbuse/VBA-MemoryTools
Private Type MemoryAccessor
    ic As IntegerAccesor
    sa As SAFEARRAY_1D
End Type

Private Type Hasher
#If Mac = 0 Then
    dictInstance As ScrDictLayout
    textHasher As Collection
    lcid As Long
#End If
    ma As MemoryAccessor
    isSet As Boolean
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

'Class members
Private m_compare As VbCompareMethod
Private m_data As DataStorage
Private m_fastUnorderedRemove As Boolean
Private m_hash As HashMap
Private m_hasher As Hasher
Private m_isOrdered As Boolean
Private m_lookups As Lookups

Private Sub Init()
    InitHashMap INITIAL_GROUP_COUNT
    m_data.Count = 0
    m_data.UBound = 0
    m_data.UsedCount = 0
    m_isOrdered = True
    ReDim m_data.Keys(0 To 0)
    ReDim m_data.Items(0 To 0)
    ReDim m_data.Meta(0 To 0)
    If Not m_hasher.isSet Then
        InitLookups
        InitLocalHasher
    End If
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
        ReDim Preserve m_data.Keys(0 To m_data.UBound)
        ReDim Preserve m_data.Items(0 To m_data.UBound)
        ReDim Preserve m_data.Meta(0 To m_data.UBound)
    End If
    i = m_data.Count
    m_data.Count = m_data.Count + 1
    m_data.UsedCount = m_data.UsedCount + 1
    If m_data.UsedCount > m_hash.MaxLoad Then
        Rehash
        groupSlot = hVal Mod m_hash.GroupCount
        Do While m_hash.Groups(groupSlot).WasEverFull 'Unlikely after rehash
             groupSlot = (groupSlot + 1) Mod m_hash.GroupCount
        Loop
        controlByte = (hVal And m_hash.ControlMask) \ m_hash.GroupCount
    End If
    '
    With m_hash.Groups(groupSlot)
        .Index(.Count) = i
        .Control = .Control Or (controlByte * m_lookups.ByteShiftL(.Count))
        .Count = .Count + 1
        .WasEverFull = .WasEverFull Or (.Count = GROUP_SIZE)
    End With
    '
    If hVal And hmObject Then Set m_data.Keys(i) = Key Else m_data.Keys(i) = Key
    If IsObject(Item) Then
        m_data.Meta(i) = hVal Or IS_ITEM_OBJ
        Set m_data.Items(i) = Item
    Else
        m_data.Meta(i) = hVal
        m_data.Items(i) = Item
    End If
End Sub

'Rebuilds hash table using the stored hashes with metadata
Private Sub Rehash()
    Dim hVal As Long
    Dim ctrlByte As Long
    Dim groupSlot As Long
    Dim i As Long
    '
    InitHashMap m_hash.GroupCount * 2
    For i = 0 To m_data.Count - 1
        hVal = m_data.Meta(i)
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
                        , Optional ByRef posInGroup As Long) As Long
    If m_hash.GroupCount = 0 Then Init
    Dim vt As VbVarType
    '
    'Hash value
    If IsObject(Key) Then vt = vbObject Else vt = VarType(Key)
    If vt = vbString Then
        Const tPrime As Long = 131&
        Const tOverflowMask As Long = &H7FFFFF
#If Mac = 0 Then
        Const minLenBinary As Long = 6
        If m_compare = vbBinaryCompare Then
            m_hasher.ma.sa.rgsabound0.cElements = Len(Key)
            If minLenBinary < m_hasher.ma.sa.rgsabound0.cElements Then
                hVal = m_hasher.textHasher(Key) Or hmText 'Early-bind speed
            Else 'Faster to loop integers
                m_hasher.ma.sa.pvData = StrPtr(Key)
#Else
                m_hasher.ma.sa.rgsabound0.cElements = Len(Key)
                If m_compare = vbBinaryCompare Then
                    m_hasher.ma.sa.pvData = StrPtr(Key)
                Else
                    Dim s As String: s = LCase$(Key)
                    m_hasher.ma.sa.pvData = StrPtr(s)
                End If
#End If
                With m_hasher.ma.ic
                    Dim i As Long
                    For i = 0 To m_hasher.ma.sa.rgsabound0.cElements - 1
                        hVal = (hVal And tOverflowMask) * tPrime + .ints(i)
                    Next i
                End With
                hVal = hVal And [_maskHM] Or hmText
                m_hasher.ma.sa.rgsabound0.cElements = INTS_IN_DOUBLE
                m_hasher.ma.sa.pvData = m_hasher.dPtr
#If Mac = 0 Then
            End If
        Else
            hVal = m_hasher.textHasher(Key) Or hmText 'Early-binding speed
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
        'Prime 13 is low enough that there is no need to mask for overflow
        Const n1& = 13 ^ 1, n2& = 13 ^ 2, n3& = 13 ^ 3, n4& = 13 ^ 4
        With m_hasher.ma.ic
            hVal = (.ints(0) * n4 + .ints(1) * n3 _
                  + .ints(2) * n2 + .ints(3) * n1) And HashMeta.[_maskHM] Or m
        End With
    Else
        hVal = vt 'vbEmpty (0) or vbNull (1)
    End If
    '
    'Calculate sub-hashes
    groupSlot = hVal Mod m_hash.GroupCount
    controlByte = (hVal And m_hash.ControlMask) \ m_hash.GroupCount
    '
    Dim matches As LongPtr
    Dim cMask As LongPtr
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
                If (m_data.Meta(GetIndex) And KEY_MASK) = hVal Then
                    If vt = vbString And (m_compare <> vbBinaryCompare) Then
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
            Loop
            If Not .WasEverFull Then Exit Do
        End With
        groupSlot = (groupSlot + 1) Mod m_hash.GroupCount
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
    '
#If Mac = 0 Then
    With m_hasher.dictInstance
        If m_compare > vbTextCompare Then
            .compMode = vbTextCompare
            .localeID = m_compare
        Else
            .compMode = m_compare
            .localeID = m_hasher.lcid
        End If
    End With
#End If
End Property

Public Property Get Count() As Long
    Count = m_data.UsedCount
End Property

'Can only be changed if there are no stored items
'When set to True, the order in which the Items/Keys were added via 'Add' is no
'   longer preserved to facilitate faster 'Remove' calls
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
    IsOrdered = m_isOrdered
End Property

'Raises Error:
'   -    5: invalid key (key is not associated with any element)
'   -  450: 'Set' is missing when assigning an object
'@DefaultMember
Public Property Get Item(ByRef Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_UserMemId = 0
    Dim i As Long: i = GetIndex(Key)
    If m_data.Meta(i) And IS_ITEM_OBJ Then
        Set Item = m_data.Items(i)
    Else
        Item = m_data.Items(i)
    End If
End Property
Public Property Let Item(ByRef Key As Variant, ByRef Item As Variant)
    If IsObject(Item) Or (VarType(Item) = vbDataObject) Then
        Err.Raise 450, TypeName(Me) & ".Item", "Set is required"
    End If
    Dim i As Long: i = GetIndex(Key)
    If i > NOT_FOUND Then
        m_data.Meta(i) = m_data.Meta(i) And KEY_MASK
        m_data.Items(i) = Item
    Else
        Add Key, Item
    End If
End Property
Public Property Set Item(ByRef Key As Variant, ByRef Item As Object)
    Dim i As Long: i = GetIndex(Key)
    If i > NOT_FOUND Then
        m_data.Meta(i) = m_data.Meta(i) Or IS_ITEM_OBJ
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
        ReDim Preserve Items(0 To m_data.UBound)
    Else
        ReDim res(0 To m_data.UsedCount - 1)
        Dim i As Long
        Dim j As Long
        '
        For i = 0 To m_data.Count - 1
            If m_data.Meta(i) <> hmRemoved Then
                If m_data.Meta(i) And IS_ITEM_OBJ Then
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
'   -   5: invalid OldKey (key is not associated with any element)
'   - 457: invalid NewKey (key is already associated with an element)
Public Property Let Key(ByRef OldKey As Variant, ByRef NewKey As Variant)
    Dim hVal As Long
    Dim groupSlot As Long
    Dim posInGroup As Long
    Dim controlByte As Long
    Dim lastPos As Long
    Dim i As Long
    '
    If GetIndex(NewKey, hVal, , controlByte) > NOT_FOUND Then Err.Raise 457
    i = GetIndex(OldKey, , groupSlot, , posInGroup)
    If i = NOT_FOUND Then Err.Raise 5, TypeName(Me) & ".Key", "Key not found"
    '
    'Replace Key and Meta while preserving Item
    If hVal And HashMeta.hmObject Then
        Set m_data.Keys(i) = NewKey
    Else
        m_data.Keys(i) = NewKey
    End If
    If m_data.Meta(i) And IS_ITEM_OBJ Then hVal = hVal Or IS_ITEM_OBJ
    m_data.Meta(i) = hVal
    '
    ClearMapIndex groupSlot, posInGroup
    '
    'Update hash map with new key
    groupSlot = hVal Mod m_hash.GroupCount
    Do While m_hash.Groups(groupSlot).WasEverFull
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
    If m_data.UsedCount = m_data.Count Then
        Keys = m_data.Keys
        ReDim Preserve Keys(0 To m_data.UBound)
    Else
        ReDim res(0 To m_data.UsedCount - 1)
        Dim i As Long
        Dim j As Long
        '
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
End Function

'Remove a single item
'Raises Error:
'   -   5: invalid key (key is not associated with any element)
Public Sub Remove(ByRef Key As Variant)
    Dim hVal As Long
    Dim groupSlot As Long
    Dim posInGroup As Long
    Dim lastPos As Long
    Dim i As Long
    Dim j As Long
    '
    i = GetIndex(Key, , groupSlot, , posInGroup)
    If i = NOT_FOUND Then Err.Raise 5, TypeName(Me) & ".Remove"
    If m_data.UsedCount = 1 Then Init
    '
    ClearMapIndex groupSlot, posInGroup
    '
    lastPos = m_data.Count - 1
    m_data.Count = lastPos
    m_data.UsedCount = m_data.UsedCount - 1
    If i < lastPos Then
        If m_fastUnorderedRemove Then
            m_isOrdered = False
            'Swap data
            hVal = m_data.Meta(lastPos)
            If hVal And IS_ITEM_OBJ Then
                Set m_data.Items(i) = m_data.Items(lastPos)
            Else
                m_data.Items(i) = m_data.Items(lastPos)
            End If
            If hVal And HashMeta.hmObject Then
                Set m_data.Keys(i) = m_data.Keys(lastPos)
            Else
                m_data.Keys(i) = m_data.Keys(lastPos)
            End If
            m_data.Meta(i) = hVal
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
            m_data.Keys(i) = Empty
            m_data.Meta(i) = hmRemoved
            m_data.UsedCount = m_data.UsedCount - 1
        End If
    ElseIf m_data.Count > m_data.UsedCount Then
        i = lastPos - 1
        Do While m_data.Meta(i) = hmRemoved
            m_data.Count = m_data.Count - 1
            i = i - 1 'Cannot go out of bounds due to previous check
        Loop
    End If
    'Clear Variants in case deallocation is needed e.g. String/Object
    m_data.Items(lastPos) = Empty
    m_data.Keys(lastPos) = Empty
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
    InitSafeArray m_hasher.ma.sa, INT_SIZE
#If Mac Then
    CopyMemory ByVal VarPtr(m_hasher.ma.ic), VarPtr(m_hasher.ma.sa), PTR_SIZE
#Else
    With m_hasher
        .dictInstance.hashModulo = HashMeta.[_modHM]
        .dictInstance.refCount = 10 'Larger than 2 to avoid deallocation
        Dictionary.InitHasher .dictInstance.vTables _
                            , .textHasher, .dictInstance.localeID _
                            , VarPtr(.ma.ic), VarPtr(.ma.sa)
        .lcid = .dictInstance.localeID
    End With
    CompareMode = m_compare
#End If
    With m_hasher
        .dPtr = VarPtr(.d)
        .ma.sa.pvData = .dPtr
        .ma.sa.rgsabound0.cElements = INTS_IN_DOUBLE
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
'Cannot have public defined types in a class so we cannot have type arguments
'   in a Public method. Smaller components and pointers are passed instead
#If Mac = 0 Then
Public Sub InitHasher(ByRef allVTables() As LongPtr _
                    , ByRef fakeDict As Collection _
                    , ByRef lcid As Long _
                    , ByRef intsPtr As LongPtr _
                    , ByRef saPtr As LongPtr)
    Const dictVTableSize As Long = 22
    Const opNumDictHashVal As Long = 21
    Const opNumCollItem As Long = 7
    Static vTable(0 To dictVTableSize - 1) As LongPtr
    Static fakeVTables(0 To 3) As LongPtr
    Static dictLcid As Long
    Static isSet As Boolean
    Static sa As SAFEARRAY_1D
    Static pa As PointerAccesor
    '
    If Not Me Is Dictionary Then Exit Sub
    If Not isSet Then
        'Init pointer accessor
        InitSafeArray sa, PTR_SIZE
        CopyMemory ByVal VarPtr(pa), VarPtr(sa), PTR_SIZE 'The only API call
        '
        'Early bind a Collection interface to a fake Scripting.Dictionary
        '   where Collection.Item is mapped to Dictionary.HashVal. This allows
        '   calls to HashVal with early binding speed without a dll reference!
        '
        Dim temp As Object: Set temp = CreateObject("Scripting.Dictionary")
        Dim i As Long
        '
        'Copy Scripting.Dictionary virtual table addresses
        sa.pvData = ObjPtr(temp)
        sa.rgsabound0.cElements = dictVTableSize
        For i = 0 To 3
            fakeVTables(i) = pa.ptrs(i)
        Next i
        '
        'Copy locale ID
        dictLcid = CLng(pa.ptrs(10))
        Set temp = Nothing 'Actual dictionary not needed anymore
        '
        'Copy entire main virtual function table to our own
        sa.pvData = fakeVTables(0)
        For i = 0 To dictVTableSize - 1
            vTable(i) = pa.ptrs(i)
        Next i
        fakeVTables(0) = VarPtr(vTable(0)) 'Replace main vTable
        '
        vTable(opNumCollItem) = vTable(opNumDictHashVal) 'Map Item to HashVal
        sa.rgsabound0.cElements = 1
        isSet = True
    End If
    '
    'Create fake instance while avoiding calls to CopyMemory (slow in Win VBA7)
    'We need all 4 virtual tables to prevent crashes if debugging
    allVTables(0) = fakeVTables(0)
    allVTables(1) = fakeVTables(1)
    allVTables(2) = fakeVTables(2)
    allVTables(3) = fakeVTables(3)
    lcid = dictLcid
    sa.pvData = VarPtr(fakeDict)
    pa.ptrs(0) = VarPtr(allVTables(0))
    '
    'Init integer accesor
    sa.pvData = intsPtr
    pa.ptrs(0) = saPtr
End Sub
#End If
