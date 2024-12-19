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
        Private Declare PtrSafe Function LocalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
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
#Const Windows = (Mac = 0)
#Const x64 = Win64
#Const x32 = (x64 = 0) 'We don't care about x16
    
Private Enum InternalConstants 'Hides constants from Locals window
#If x64 Then
    isItemObjBit = &H80000000
    keyBitMask = &H7FFFFFFF
    ptrSize = 8
#Else
    ptrSize = 4
#End If
    prevOffset = ptrSize 'Previous instance pointer immediately after vTable
    notFound = -1
    intSize = 2
    intsPerDouble = 4
    groupSize = ptrSize
    nextItemOffset = ptrSize * 2
    initialGroupCount = 16
End Enum

Private Enum HashMeta
    hmRemoved = 2
    hmError = &H10000000
    hmNumber = &H20000000
    hmText = &H30000000
    hmObject = &H40000000
    [_modHM] = hmError
    [_maskHM] = [_modHM] - 1
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

#If x32 Then
Private Type EnumerableVariant
    Value As Variant
    Meta As Long
    IsItemObj As Boolean
    IsKeyObj As Boolean
    NextPtr As Long
End Type
#End If

#If Windows Then
Private Type ScrDictLayout 'Scripting.Dictionary memory layout
    vTables(0 To 3) As LongPtr
    unkPtr1 As LongPtr
    refCount As Long
    firstItemPtr As LongPtr
    lastItemPtr As LongPtr
#If x32 Then
    Dummy As Long
#End If
    hashTablePtr As LongPtr
    hashModulo As Long
    compMode As Long
    lcid As Long
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
             
'Data shared across all class instances
Private Type Globals
#If Windows Then
    sdl As ScrDictLayout
    lcid As Long 'System localeID
    isScrAvailable As Boolean
#End If
    saI As SAFEARRAY_1D
    saP As SAFEARRAY_1D
    d As Double 'For hashing numbers
    dPtr As LongPtr
    lastInstancePtr As LongPtr
    lastTerminatedPtr As LongPtr
    lastTerminatedNextPtr As LongPtr
    nextPtrOffset As LongPtr
    defPtrOffset As LongPtr
End Type

Private Type Lookups
    ByteShiftL(0 To groupSize - 1) As LongPtr
    ByteMask(0 To groupSize - 1) As LongPtr
    ByteOff(0 To groupSize - 1) As LongPtr
    PositionMask(0 To groupSize - 2) As LongPtr
    CountMask(0 To groupSize) As LongPtr
#If x64 Then
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
End Type

'These will link into the default (Predeclared) instance of this class
'This is necessary to avoid speed issues on the heap with fake Scripting Dict
'Faster deallocation of memory by orders of magnitude when having many instances
Private Type MemoryAccessors
    Common() As Globals
    Look() As Lookups
    RInt() As Integer
    RPtr() As LongPtr
#If Windows Then
    TextHash() As Collection
#End If
End Type

Private Type DeallocVars
    ThisPtr As LongPtr
    NextDictPtr As LongPtr
End Type

Private Type DictionaryVariables
    Hash As HashMap
    Enums As EnumProvider
    Compare As VbCompareMethod
    LocaleID As Long
    Items() As Variant
#If x64 Then
    Keys() As Variant
    Meta() As Long
#Else
    Keys() As EnumerableVariant
#End If
    Count As Long
    UBound As Long
    UsedCount As Long
    DefInstance As Dictionary 'Avoids deallocation of default (Predeclared) dict
    Dealloc As DeallocVars
    IsInitialized As Boolean
End Type

'Class members
Private Vars As DictionaryVariables
Private Mem As MemoryAccessors

'NewEnum must be the first method so that it's virtual table position is known.
'This is because we want to fix the x64 bug when using a For Each.. loop:
'https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes

'Cannot use ITypeInfo::AddressOfMember because it overrides some assembly bytes:
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
    #If x64 Then
        ReDim .Keys(0 To .UBound + 1) 'Extra member for safe IEnumVariant
        ReDim .Meta(0 To .UBound)
    #Else
        ReDim .Keys(0 To .UBound)
    #End If
    End With
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
        #If x64 Then 'Extra member for safe IEnumVariant
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
        .Control = .Control Or (controlByte * Mem.Look(0).ByteShiftL(.Count))
        .Count = .Count + 1
        .WasEverFull = .WasEverFull Or (.Count = groupSize)
    End With
    '
#If x64 Then
    If hVal And hmObject Then Set Vars.Keys(i) = Key Else Vars.Keys(i) = Key
    If Vars.Enums.hasEnum Then
        RemoveUnusedEnums
        If i > 0 Then
            ptr = VarPtr(Vars.Keys(i))
            With Mem.Common(0)
                .saP.pvData = ptr + nextItemOffset
                Mem.RPtr(0) = ptr
                .saP.pvData = .dPtr
            End With
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
        #If x64 Then
            hVal = Vars.Meta(i)
        #Else
            hVal = Vars.Keys(i).Meta
        #End If
        If hVal <> hmRemoved Then
            groupSlot = hVal And Vars.Hash.GroupMask
            Do While Vars.Hash.Groups(groupSlot).WasEverFull 'Unlikely (resized)
                groupSlot = (groupSlot + 1) Mod Vars.Hash.GroupCount
            Loop
            With Vars.Hash.Groups(groupSlot)
                .Index(.Count) = i
                ctrlByte = (hVal And Vars.Hash.ControlMask) \ Vars.Hash.GroupCount
                .Control = .Control Or (ctrlByte * Mem.Look(0).ByteShiftL(.Count))
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
    #If (x32) Or Mac Then
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
        With Mem.Common(0)
#If Windows Then
            If .isScrAvailable Then
                Const minLenBinary As Long = 6
                If Vars.Compare = vbBinaryCompare Then
                    .saI.rgsabound0.cElements = Len(Key)
                    If minLenBinary < .saI.rgsabound0.cElements Then
                        .sdl.compMode = vbBinaryCompare
                        hVal = Mem.TextHash(0)(Key) Or hmText 'Early-binded
                    Else 'Faster to loop integers
                        .saI.pvData = StrPtr(Key)
                        For i = 0 To .saI.rgsabound0.cElements - 1
                            hVal = (hVal And tOverflowMask) * tPrime + Mem.RInt(i)
                        Next i
                        hVal = hVal And [_maskHM] Or hmText
                        .saI.rgsabound0.cElements = intsPerDouble
                        .saI.pvData = .dPtr
                    End If
                Else
                    .sdl.compMode = vbTextCompare
                    .sdl.lcid = Vars.LocaleID
                    hVal = Mem.TextHash(0)(Key) Or hmText 'Early-binded
                End If
            Else
#End If
                .saI.rgsabound0.cElements = Len(Key)
                If Vars.Compare = vbBinaryCompare Then
                    .saI.pvData = StrPtr(Key)
                Else
                    Dim s As String: s = LCase$(Key)
                    .saI.pvData = StrPtr(s)
                End If
                For i = 0 To .saI.rgsabound0.cElements - 1
                    hVal = (hVal And tOverflowMask) * tPrime + Mem.RInt(i)
                Next i
                hVal = hVal And [_maskHM] Or hmText
                .saI.rgsabound0.cElements = intsPerDouble
                .saI.pvData = .dPtr
#If Windows Then
            End If
#End If
        End With
    ElseIf vt = vbObject Or vt = vbDataObject Then
        Const oPrime As Long = 2701&
        Static iUnk As stdole.IUnknown 'Dim is slower
        '
        vt = vbObject 'Replace vbDataObject if needed
        Set iUnk = Key
        #If x64 Then
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
        Set iUnk = Nothing 'Because of Static but still faster than just Dim
    ElseIf vt > vbLongLong Then
        Err.Raise 5, , "Cannot hash an Array or User Defined Type"
    ElseIf vt > vbNull Then
        Dim m As HashMeta
        If vt = vbError Then
            m = hmError
            Mem.Common(0).d = CDbl(Key)
        Else
            m = hmNumber
            Mem.Common(0).d = Key
        End If
        Const n1& = 1201, n2& = 2701, n3& = 131, n4& = 28571
        Const infOrNaN As Integer = &H7FF0 'Exponent bits for Double Float
        #If x64 Then
            Const maskLL As LongLong = HashMeta.[_maskHM]
            Const iSignBit As Integer = &H8000
            Const factionalBits As LongLong = &HFFFFFFFFFFFFF^
            Const normalBit As LongLong = factionalBits + 1
            Dim h As LongLong, bitFinder As LongLong
            Dim highBits As Integer: highBits = Mem.RInt(3)
            '
            If (highBits And infOrNaN) = infOrNaN Then
                hVal = (highBits * n1 + Mem.RInt(0) * n4) And [_maskHM]
            Else
                h = Mem.RPtr(0) And factionalBits Or normalBit
                bitFinder = h Or Mem.Look(0).ExpPos((highBits And &H3F0) \ &H10)
                h = h \ (bitFinder And -bitFinder)
                If highBits And iSignBit Then h = Not h - 1
                hVal = CLng(h And maskLL) Or m
            End If
        #Else
            If (Mem.RInt(3) And infOrNaN) = infOrNaN Then m = 0
            hVal = (Mem.RInt(0) * n4 + Mem.RInt(1) * n3 _
                  + Mem.RInt(2) * n2 + Mem.RInt(3) * n1) And [_maskHM] Or m
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
    #If x64 Then
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
            cMask = Mem.Look(0).CountMask(.Count)
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
                    posInGroup = Mem.Look(0).ModBytePosition(CLng( _
                                 (-matches And matches) Mod positionPrime))
                    matches = matches Xor Mem.Look(0).PositionMask(posInGroup)
                End If
                GetIndex = .Index(posInGroup)
#If x64 Then
                If (Vars.Meta(GetIndex) And keyBitMask) = hVal Then
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
    If Vars.Compare > vbTextCompare Then
        Vars.LocaleID = Vars.Compare
    Else
        Vars.LocaleID = Mem.Common(0).lcid
    End If
#End If
End Property

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
#If x64 Then
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
        #If x64 Then
            Vars.Meta(i) = Vars.Meta(i) And keyBitMask
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
        #If x64 Then
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
#If x64 Then
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
#If x64 Then
    Dim isObj As Boolean: isObj = CBool(hVal And HashMeta.hmObject)
    If Vars.Enums.hasEnum Then
        RemoveUnusedEnums
        With Mem.Common(0)
            .saP.pvData = VarPtr(Vars.Keys(i)) + nextItemOffset
            Dim ptr As LongPtr: ptr = Mem.RPtr(0)
            If isObj Then Set Vars.Keys(i) = NewKey Else Vars.Keys(i) = NewKey
            Mem.RPtr(0) = ptr
            .saP.pvData = .dPtr
        End With
    Else
        If isObj Then Set Vars.Keys(i) = NewKey Else Vars.Keys(i) = NewKey
    End If
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
        .Control = .Control Or (controlByte * Mem.Look(0).ByteShiftL(.Count))
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
            .Control = .Control And Mem.Look(0).ByteOff(posInGroup)
            .Control = .Control Or ((.Control And Mem.Look(0).ByteMask(lastPos)) _
                                   \ Mem.Look(0).ByteShiftL(lastPos - posInGroup))
        End If
        .Index(lastPos) = 0
        .Control = .Control And Mem.Look(0).ByteOff(lastPos)
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
#If x64 Then
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
        #If x64 Then
            Vars.Keys(i) = Empty
            Vars.Meta(i) = hmRemoved
        #Else
            Vars.Keys(i).Value = Empty
            Vars.Keys(i).Meta = hmRemoved
        #End If
        If Vars.Enums.hasEnum Then 'Link previous to next
            j = i + 1
            #If x64 Then
                Do While Vars.Meta(j) = hmRemoved: j = j + 1: Loop
            #Else
                Do While Vars.Keys(j).Meta = hmRemoved: j = j + 1: Loop
            #End If
            ptr = VarPtr(Vars.Keys(j))
            RemoveUnusedEnums VarPtr(Vars.Keys(i)), ptr
            j = i - 1
            Do While j >= 0
                #If x64 Then
                    If Vars.Meta(j) <> hmRemoved Then Exit Do
                #Else
                    If Vars.Keys(j).Meta <> hmRemoved Then Exit Do
                #End If
                j = j - 1
            Loop
            If j >= 0 Then
                #If x64 Then
                    With Mem.Common(0)
                        .saP.pvData = VarPtr(Vars.Keys(j + 1)) + nextItemOffset
                        Mem.RPtr(0) = ptr
                        .saP.pvData = .dPtr
                    End With
                #Else
                    Vars.Keys(j).NextPtr = ptr
                #End If
            End If
        End If
        Exit Sub
    End If
    If Vars.Count > Vars.UsedCount Then
        i = lastPos - 1
        #If x64 Then
            Do While Vars.Meta(i) = hmRemoved: i = i - 1: Loop
        #Else
            Do While Vars.Keys(i).Meta = hmRemoved: i = i - 1: Loop
        #End If
        Vars.Count = Vars.Count - lastPos + i + 1
    End If
    If Vars.Enums.hasEnum Then
        RemoveUnusedEnums VarPtr(Vars.Keys(lastPos))
        #If x64 Then
            Vars.Keys(i + 1) = Empty 'Just to clear the pointer
        #Else
            Vars.Keys(i).NextPtr = 0&
        #End If
    End If
    '
    'Clear Variants in case deallocation is needed e.g. String/Object
    Vars.Items(lastPos) = Empty
    #If x64 Then
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
    #If x64 Then
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
    #If x64 Then
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

'To avoid API calls overhead memory accessors are cached in the default instance
'   (Attribute VB_PredeclaredId = True)
'On Windows it initializes a Collection to call ScriptingDictionary.HashVal with
'   early binding speed
Friend Sub InitStructs(ByRef v As DictionaryVariables _
                     , ByRef m As MemoryAccessors)
    #If x64 Then
        Const nullPtr As LongLong = 0^
    #Else
        Const nullPtr As Long = 0&
    #End If
    Static l As Lookups
    Static h As Globals
    Static saL As SAFEARRAY_1D
    Static saH As SAFEARRAY_1D
    Static saPtrs(0 To 4) As LongPtr
    Dim temp As Object
    '
    If Not Vars.DefInstance Is Nothing Then
        Vars.DefInstance.InitStructs v, m
        Exit Sub
    End If
    '
    If saL.cDims = 0 Then
        h.nextPtrOffset = VarPtr(Vars.Dealloc.NextDictPtr) - Vars.Dealloc.ThisPtr
        h.defPtrOffset = VarPtr(Vars.DefInstance) - Vars.Dealloc.ThisPtr
        '
        saPtrs(0) = VarPtr(saH)
        saPtrs(1) = VarPtr(saL)
        saPtrs(2) = VarPtr(h.saI)
        saPtrs(3) = VarPtr(h.saP)
        '
        InitLookups l
        InitSafeArray saH, LenB(h)
        InitSafeArray saL, LenB(l)
        InitSafeArray h.saI, intSize
        InitSafeArray h.saP, ptrSize
        '
        saH.pvData = VarPtr(h)
        saL.pvData = VarPtr(l)
        h.dPtr = VarPtr(h.d)
        h.saI.pvData = h.dPtr
        h.saP.pvData = h.dPtr
        '
        saH.rgsabound0.cElements = 1
        saL.rgsabound0.cElements = 1
        h.saI.rgsabound0.cElements = intsPerDouble
        h.saP.rgsabound0.cElements = 1
        '
        'The only API call on Windows
        CopyMemory ByVal VarPtr(Mem) + ptrSize * 3, saPtrs(3), ptrSize
#If Windows Then
        Const dictVTables As Long = 4
        Const dictMainVTableSize As Long = 22
        Const opNumDictHashVal As Long = 21
        Const opNumCollItem As Long = 7
        Static mainVTable(0 To dictMainVTableSize - 1) As LongPtr
        Static c As Collection
        Static saC As SAFEARRAY_1D
        Dim i As Long
        '
        'Early bind a Collection interface to a fake Scripting.Dictionary
        '   where Collection.Item is mapped to Dictionary.HashVal. This allows
        '   calls to HashVal with early binding speed without a dll reference
        '
        InitSafeArray saC, ptrSize
        saC.pvData = VarPtr(c)
        saC.rgsabound0.cElements = 1
        saPtrs(4) = VarPtr(saC)
        '
        If InIde() Then
            On Error Resume Next 'In case scrun.dll not available
            Set temp = CreateObject("Scripting.Dictionary")
            On Error GoTo 0
        End If
        '
        h.isScrAvailable = Not (temp Is Nothing)
        If h.isScrAvailable Then
            'Copy Scripting.Dictionary virtual table addresses
            h.saP.pvData = ObjPtr(temp)
            h.saP.rgsabound0.cElements = dictMainVTableSize
            For i = 0 To dictVTables - 1
                h.sdl.vTables(i) = Mem.RPtr(i)
            Next i
            '
            'Copy locale ID
            #If x64 Then
                h.lcid = CLng(Mem.RPtr(10))
            #Else
                h.lcid = Mem.RPtr(12)
            #End If
            Set temp = Nothing 'Actual dictionary not needed anymore
            h.sdl.lcid = h.lcid
            '
            'Copy entire main virtual function table to our own
            h.saP.pvData = h.sdl.vTables(0)
            For i = 0 To dictMainVTableSize - 1
                mainVTable(i) = Mem.RPtr(i)
            Next i
            h.sdl.vTables(0) = VarPtr(mainVTable(0)) 'Replace main vTable
            '
            'Map Collection.Item to Dictionary.HashVal
            mainVTable(opNumCollItem) = mainVTable(opNumDictHashVal)
            '
            'Set up fake instance
            h.sdl.hashModulo = HashMeta.[_modHM]
            h.sdl.refCount = 2 'To avoid deallocation
            h.saP.rgsabound0.cElements = 1
            h.saP.pvData = VarPtr(c)
            Mem.RPtr(0) = VarPtr(h.sdl)
        End If
        '
        #If x64 Then
            'Increase stack frame size for 'NewEnum' to avoid For Each.. x64 bug
            'https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes
            h.saP.pvData = Vars.Dealloc.ThisPtr 'vtbl
            h.saP.pvData = Mem.RPtr(0) + ptrSize * 7 'NewEnum address
            h.saP.pvData = Mem.RPtr(0) + 21 'SUB R12, 0x10
            '
            If (Mem.RPtr(0) And &HFFFF^) = &H8149^ Then
                'SUB found. Replace 0x10 (16) with 0x800 (2048) bytes
                Mem.RPtr(0) = Mem.RPtr(0) And &HFF00000000FFFFFF^ Or &H800000000^
            End If
            '
            'Increase stack frame size for 'Class_Terminate' to avoid x64 bugs
            'https://github.com/cristianbuse/VBA-FastDictionary/issues/10
            'https://stackoverflow.com/questions/65041832/vba-takes-wrong-branch-at-if-statement-severe-compiler-bug
            h.saP.pvData = Vars.Dealloc.ThisPtr + LocalSize(Vars.Dealloc.ThisPtr) - ptrSize * 3 'IClassModuleEvt
            h.saP.pvData = Mem.RPtr(0) + ptrSize * 4 '_Terminate address
            h.saP.pvData = Mem.RPtr(0) + 21 'SUB R12, 0x08
            '
            If (Mem.RPtr(0) And &HFFFF^) = &H8149^ Then
                'SUB found. Replace 0x8 (8) with 0x80 (128) bytes
                Mem.RPtr(0) = Mem.RPtr(0) And &HFF00000000FFFFFF^ Or &H80000000^
            End If
        #End If
#End If
    End If
    '
    'Avoid deallocation of Global Instance
    If v.Dealloc.ThisPtr <> Vars.Dealloc.ThisPtr Then Set v.DefInstance = Me
    '
#If Mac Then
    CopyMemory ByVal VarPtr(m), saPtrs(0), ptrSize * 4
#Else
    If h.isScrAvailable Then
        If v.Compare <= vbTextCompare Then v.LocaleID = h.lcid
    End If
    '
    h.saP.rgsabound0.cElements = 5
    h.saP.pvData = VarPtr(m)
    For i = 0 To 4
        Mem.RPtr(i) = saPtrs(i)
    Next i
    h.saP.rgsabound0.cElements = 1
#End If
    h.saP.pvData = v.Dealloc.ThisPtr + prevOffset
    '
    Dim prevPtr As LongPtr: prevPtr = Mem.RPtr(0)
    Dim tempPtr As LongPtr
    '
    If prevPtr = nullPtr Then 'Can only be def instance
        h.lastInstancePtr = v.Dealloc.ThisPtr
        h.saP.pvData = h.dPtr
        v.IsInitialized = True
        Exit Sub
    End If
    '
    'In case user modified the global instance e.g. Set Dictionary = Nothing
    If v.DefInstance Is Nothing Then
        Do
            h.saP.pvData = prevPtr + h.defPtrOffset
            tempPtr = Mem.RPtr(0)
            If tempPtr Then Exit Do
            '
            h.saP.pvData = prevPtr + prevOffset
            If Mem.RPtr(0) = nullPtr Then
                h.saP.pvData = prevPtr + h.nextPtrOffset
                If Mem.RPtr(0) Then tempPtr = prevPtr
                Exit Do
            End If
            prevPtr = Mem.RPtr(0)
        Loop
        If (tempPtr <> nullPtr) And (tempPtr <> Vars.Dealloc.ThisPtr) Then
            'Link to the 'real' def instance
            h.saP.pvData = VarPtr(temp)
            Mem.RPtr(0) = tempPtr 'Unmanaged - ref count not increased
            Set Vars.DefInstance = temp
            Mem.RPtr(0) = nullPtr 'Ref count not decreased
            '
            #If Windows Then
                Set c = Nothing
            #End If
            Vars.DefInstance.InitStructs Vars, Mem
            With Mem.Common(0)
                .saP.pvData = .lastTerminatedPtr + prevOffset
                If Mem.RPtr(0) = v.Dealloc.ThisPtr Then
                    'Previous instance is currently initializing and it is
                    '   definitely reusing previously terminated memory
                    v.Dealloc.NextDictPtr = .lastTerminatedPtr
                    .saP.pvData = .lastTerminatedPtr + .nextPtrOffset
                    Mem.RPtr(0) = .lastTerminatedNextPtr
                    .lastTerminatedNextPtr = v.Dealloc.ThisPtr
                    .saP.pvData = v.Dealloc.ThisPtr + prevOffset
                    .lastTerminatedPtr = Mem.RPtr(0)
                End If
                .saP.pvData = .dPtr
            End With
            Exit Sub
        End If
    End If
    '
    If v.Dealloc.ThisPtr = h.lastTerminatedPtr Then
        'Reusing previously terminated memory
        v.Dealloc.NextDictPtr = h.lastTerminatedNextPtr
        h.lastTerminatedPtr = prevPtr
        h.lastTerminatedNextPtr = v.Dealloc.ThisPtr
    Else
        If prevPtr = h.lastInstancePtr Then
            h.lastInstancePtr = v.Dealloc.ThisPtr
        Else 'The previous instance could be initializing
            h.saP.pvData = prevPtr + prevOffset
            If Mem.RPtr(0) = h.lastInstancePtr Then
                h.lastInstancePtr = v.Dealloc.ThisPtr
            End If
        End If
    End If
    h.saP.pvData = prevPtr + h.nextPtrOffset
    Mem.RPtr(0) = v.Dealloc.ThisPtr
    '
    h.saP.pvData = h.dPtr
    v.IsInitialized = True
End Sub

'https://stackoverflow.com/a/9516565/8488913
Public Property Get InIde() As Boolean
    Debug.Assert SetTrue(InIde)
End Property
Private Function SetTrue(ByRef b As Boolean) As Boolean
    b = True
    SetTrue = True
End Function
    
Public Function Factory() As Dictionary
    Set Factory = New Dictionary
End Function

Public Property Get HashVal(ByRef Key As Variant) As Long
    GetIndex Key, HashVal, hashOnly:=True
End Property

Private Function DictEnum() As IUnknown
    Dim hadEnum As Boolean
    '
    If Not Vars.IsInitialized Then Exit Function 'In case called from IDE before Init
    With Vars.Enums
        hadEnum = .hasEnum
        If .hasEnum Then RemoveUnusedEnums
        Set DictEnum = .emptyColl.[_NewEnum]
        .hasEnum = True
        .enumsColl.Add DictEnum, CStr(ObjPtr(DictEnum))
    End With
    If Vars.UsedCount = 0 Then Exit Function
    '
    Dim i As Long
    Dim hasGaps As Boolean: hasGaps = (Vars.UsedCount < Vars.Count)
    '
    If hasGaps Then
        #If x64 Then
            Do While Vars.Meta(i) = hmRemoved: i = i + 1: Loop
        #Else
            Do While Vars.Keys(i).Meta = hmRemoved: i = i + 1: Loop
        #End If
    End If
    With Mem.Common(0)
        .saP.pvData = ObjPtr(DictEnum) + nextItemOffset
        Mem.RPtr(0) = VarPtr(Vars.Keys(i))
        .saP.pvData = .dPtr
    End With
    If hadEnum Then Exit Function
    '
    Dim ptr As LongPtr: ptr = VarPtr(Vars.Keys(0))
    Dim j As Long
    '
#If x64 Then
    Const variantSize As Long = 24
    With Mem.Common(0)
        .saP.pvData = ptr
        .saP.rgsabound0.cElements = (Vars.Count + 1) * 3
        If hasGaps Then
            j = i * 3 + 5
            For i = i + 1 To Vars.Count - 1
                If Vars.Meta(i) <> hmRemoved Then
                    Mem.RPtr(j) = ptr + variantSize * i
                    j = i * 3 + 5
                End If
            Next i
            Mem.RPtr(j) = 0^
        Else
            For i = 5 To .saP.rgsabound0.cElements - 4 Step 3
                ptr = ptr + variantSize
                Mem.RPtr(i) = ptr
            Next i
            Mem.RPtr(i) = 0^
        End If
        .saP.rgsabound0.cElements = 1
        .saP.pvData = .dPtr
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
End Function

Private Sub RemoveUnusedEnums(Optional ByVal searchPtr As LongPtr _
                            , Optional ByVal replacePtr As LongPtr)
    #If x64 Then
        Const nullPtr As LongLong = 0^
    #Else
        Const nullPtr As Long = 0&
    #End If
    Dim e As Variant 'IEnumVARIANT does not work with For Each
    With Mem.Common(0)
        For Each e In Vars.Enums.enumsColl
            Dim ptr As LongPtr: ptr = ObjPtr(e)
            .saP.pvData = ptr + nextItemOffset
            If Mem.RPtr(0) = nullPtr Then
                Vars.Enums.enumsColl.Remove CStr(ptr)
            ElseIf Mem.RPtr(0) = searchPtr Then
                If replacePtr = nullPtr Then
                    Vars.Enums.enumsColl.Remove CStr(ptr)
                Else
                    Mem.RPtr(0) = replacePtr
                End If
            End If
        Next e
        Vars.Enums.hasEnum = (Vars.Enums.enumsColl.Count > 0)
        .saP.pvData = .dPtr
    End With
End Sub

Private Sub ShiftEnumPointers(ByVal addrShift As LongPtr)
    Dim e As Variant 'IEnumVARIANT does not work with For Each
    For Each e In Vars.Enums.enumsColl
        Mem.Common(0).saP.pvData = ObjPtr(e) + nextItemOffset
        Mem.RPtr(0) = Mem.RPtr(0) + addrShift
    Next e
    '
    Dim hasGaps As Boolean: hasGaps = (Vars.UsedCount < Vars.Count)
    Dim i As Long
    '
#If x64 Then
    With Mem.Common(0)
        .saP.pvData = VarPtr(Vars.Keys(0))
        .saP.rgsabound0.cElements = Vars.Count * 3
        If hasGaps Then
            For i = 5 To .saP.rgsabound0.cElements - 1 Step 3
                If Mem.RPtr(i) <> 0^ Then Mem.RPtr(i) = Mem.RPtr(i) + addrShift
            Next i
        Else
            For i = 5 To .saP.rgsabound0.cElements - 1 Step 3
                Mem.RPtr(i) = Mem.RPtr(i) + addrShift
            Next i
        End If
        .saP.rgsabound0.cElements = 1
        .saP.pvData = .dPtr
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
End Sub

'Only initialize memory manipulation structs
Private Sub Class_Initialize()
    Vars.Dealloc.ThisPtr = ObjPtr(Me)
    Dictionary.InitStructs Vars, Mem
End Sub

'Postpones termination to a later stage where we have full control over how VBA
'   traverses the linked list of all dictionary instances
Private Sub Class_Terminate()
    'Deallocate enumerators just in case they are pointing to Keys
    Set Vars.Enums.emptyColl = Nothing
    Set Vars.Enums.enumsColl = Nothing
    '
    'Nested Dictionaries must be terminated before we deallocate this instance
    Erase Vars.Keys
    Erase Vars.Items
    '
    'Cache this instance inside the global instance and destroy later
    Vars.DefInstance.DelayTermination Me, Vars
    Set Vars.DefInstance = Nothing
End Sub

'When VB* terminates a class instance, it traverses all instances starting from
'   the last created instance all the way to the first. On itself this would be
'   fast but unfortunately VB also makes checks and can reclaim memory that is
'   unrelated to the instance being terminated. So, this traversal becomes
'   exponentially slower the more instances there are - O(n^2)
'This method 'tricks' VB into traversing only a handful of instances thus making
'   the whole termination process linear - O(n)
Friend Sub DelayTermination(ByRef dictToDelay As Dictionary _
                          , ByRef v As DictionaryVariables)
    #If x64 Then
        Const nullPtr As LongLong = 0^
    #Else
        Const nullPtr As Long = 0&
    #End If
    Static pendingDict As Dictionary
    Static pendingPtr As LongPtr
    Static lastDict As Dictionary
    Static lastDictPtr As LongPtr
    Dim prevPtr As LongPtr
    Dim followPtr As LongPtr
    Dim secondLastPtr As LongPtr
    Dim tempDict As Dictionary
    '
    If pendingDict Is Nothing Then
        Set pendingDict = dictToDelay
        pendingPtr = v.Dealloc.ThisPtr
        Exit Sub
    End If
    With Mem.Common(0)
        If pendingPtr = .lastInstancePtr Then
            'We force keep the last instance active to avoid extra logic
            If lastDict Is Nothing Then
                Set lastDict = pendingDict
                lastDictPtr = pendingPtr
                Set pendingDict = dictToDelay
                pendingPtr = v.Dealloc.ThisPtr
                Exit Sub
            End If
            '
            Set tempDict = lastDict
            Set lastDict = pendingDict
            Set pendingDict = tempDict
            Set tempDict = Nothing
            pendingPtr = lastDictPtr
            lastDictPtr = .lastInstancePtr
        End If
        '
        If .lastTerminatedPtr = nullPtr Then
            .lastTerminatedPtr = Vars.Dealloc.ThisPtr 'Use Def instance
        End If
        '
        .saP.pvData = pendingPtr + prevOffset
        prevPtr = Mem.RPtr(0)
        '
        If prevPtr = .lastTerminatedPtr Then
            .saP.pvData = pendingPtr + .nextPtrOffset
            followPtr = Mem.RPtr(0)
        Else 'Insert after last terminated
            Mem.RPtr(0) = .lastTerminatedPtr
            .saP.pvData = .lastTerminatedPtr + .nextPtrOffset
            '
            Dim tempPtr As LongPtr: tempPtr = Mem.RPtr(0)
            Mem.RPtr(0) = pendingPtr
            '
            .saP.pvData = tempPtr + prevOffset
            Mem.RPtr(0) = pendingPtr
            '
            .saP.pvData = pendingPtr + .nextPtrOffset
            followPtr = Mem.RPtr(0)
            Mem.RPtr(0) = tempPtr
            '
            If prevPtr <> nullPtr Then 'Not first ever instance
                .saP.pvData = prevPtr + .nextPtrOffset
                Mem.RPtr(0) = followPtr
            End If
            '
            .saP.pvData = followPtr + prevOffset
            Mem.RPtr(0) = prevPtr
            followPtr = tempPtr
        End If
        '
        'Make VB 'believe' that trailing instance is the last instance
        '   so that a shorter list is traversed when we terminate 'pending'
        .saP.pvData = .lastInstancePtr + prevOffset
        secondLastPtr = Mem.RPtr(0)
        Mem.RPtr(0) = pendingPtr
        '
        Set pendingDict = Nothing 'Traverse short list and deallocate as needed
        '
        Mem.RPtr(0) = secondLastPtr 'Restore to long list
        '
        .saP.pvData = pendingPtr + prevOffset
        If .lastTerminatedPtr <> Mem.RPtr(0) Then 'Memory was reclaimed
            .saP.pvData = Mem.RPtr(0) + .nextPtrOffset
            Mem.RPtr(0) = pendingPtr
        End If
        .lastTerminatedPtr = pendingPtr
        .lastTerminatedNextPtr = followPtr
        .saP.pvData = .dPtr
        '
        Set pendingDict = dictToDelay
        pendingPtr = v.Dealloc.ThisPtr
    End With
End Sub
