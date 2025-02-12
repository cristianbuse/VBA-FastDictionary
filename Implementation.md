## Implementation

This document outlines the design decisions made in creating an efficient and robust dictionary, providing a comprehensive overview of its functionality. Please refer to the table of contents below for easy navigation to specific design decisions.

This Dictionary does not require any DLL references or any kind of external libraries. Works on Mac and Windows on both x32 and x64.

## Table of Contents

- [Compatibility with Scripting.Dictionary](#compatibility-with-scriptingdictionary)
  - [Hashing Numbers incompatibility](#hashing-numbers-incompatibility)
  - [Error numbers incompatibility](#error-numbers-incompatibility)
  - [Item (Get) incompatibility](#item-get-incompatibility)
  - [Zero-length text, 0 (zero) and Empty incompatibility](#zero-length-text-0-zero-and-empty-incompatibility)
- [Hashing](#hashing)
  - [Number Hashing](#number-hashing)
  - [Object Hashing](#object-hashing)
  - [Text Hashing on Mac](#text-hashing-on-mac)
  - [Text Hashing on Windows](#text-hashing-on-windows)
    - [Scripting.Dictionary.HashVal usefulness](#scriptingdictionaryhashval-usefulness)
    - [Scripting.Dictionary memory layout](#scriptingdictionary-memory-layout)
      - [Scripting.Dictionary conclusions](#scriptingdictionary-conclusions)
    - [Faking a Scripting.Dictionary instance](#faking-a-scriptingdictionary-instance)
      - [Scripting.Dictionary heap issue](#scriptingdictionary-heap-issue)
- [Metadata](#metadata)
- [Hash Map/Table](#hash-maptable)
  - [Sub-hashing](#sub-hashing)
  - [Finding a key](#finding-a-key)
  - [Adding a key](#adding-a-key)
- [Rehashing](#rehashing)
- [NewEnum](#newenum)
  - [x64 implementation](#x64-implementation)
  - [x32 implementation](#x32-implementation)
  - [Enumerator management](#enumerator-management)
- [Additional functionality](#additional-functionality)
- [OLE Automation](#ole-automation)
- [x64 Assembly](#x64-assembly)
  - [VBA x64 class method call mechanism](#vba-class-method-call-mechanism)
    - [Class virtual table](#class-virtual-table)
    - [Class method code](#class-method-code)
    - [Class method call](#class-method-call)
  - [Stack Bug Fixes](#stack-bug-fixes)
    - [Item (Get) stack fix](#item-get-stack-fix)
    - [NewEnum stack fix](#newenum-stack-fix)
    - [Class_Terminate stack fix](#class_terminate-stack-fix)
***

## Compatibility with ```Scripting.Dictionary```

The Dictionary presented in this repository is designed to be a drop-in replacement for ```Scripting.Dictionary``` (Microsoft Scripting Runtime - scrrun.dll on Windows). However, there are a few differences, and their purpose is to make this Dictionary the better choice from a functionality perspective.

### Hashing Numbers incompatibility

The ```Scripting.Dictionary``` casts all the keys of type number to ```Single``` and only then hashes the values. This can be easily checked with the following code snippet:

```VBA
Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Sub TestNumberHash()
#If Mac Then
    MsgBox "Scripting.Dictionary is not available on Mac"
#Else
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim s As Single
    '
    i = 12345
    s = i
    '
    Debug.Print d.HashVal(i) 'Prints 1196
    '
    Const dictHashModulo As Long = 1201
    Dim l As Long
    '
    CopyMemory l, s, 4
    '
    Debug.Print l Mod dictHashModulo 'Prints 1196
#End If
End Sub
```

Thus, any number outside the ```Single``` range is not hashed. The following all return 0 (zero):
```VBA
Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
Debug.Print d.HashVal(10000000)
Debug.Print d.HashVal(1E+300)
Debug.Print d.HashVal(215363454)
```

So, Scripting.Dictionary hashes any number outside the -9,999,999 to 9,999,999 range to a value of 0 which explains why it is terribly slow to add such numbers. See [adding large numbers](benchmarking/result_screenshots/add_number_(long_large)_win_vba7_x64.png).

Because of this huge disadvantage, this repo's Dictionary hashes all numbers by casting them to ```Double``` instead of ```Single```. However, this creates an incompatibility:
```VBA
Sub TestHashPrecision()
    Dim sd As Object: Set sd = CreateObject("Scripting.Dictionary")
    Dim nd As New Dictionary
    '
    Dim d As Double: d = 12345.6789101112
    Dim s As Single: s = d 'Approx. 12345.68
    '
    'The new dictionary sees keys as different
    nd.Add d, 1
    nd.Add s, 2
    '
    'The scripting dictionary sees keys as same
    sd.Add d, 1
    sd.Add s, 2 'Throws error 457
End Sub
```
Scripting.Dictionary would always downgrade a ```Double``` to a ```Single``` to perform the comparison. This is of course in line with how VBA behaves as seen [here](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/comparison-operators#remarks):

> When a Single is compared to a Double, the Double is rounded to the precision of the Single

However, the new Dictionary (this repo) casts ```Single``` to ```Double```. This seems more of an improvement rather than an issue, not to mention that the number of collisions is greatly reduced thus improving speed by orders of magnitude.

### Error numbers incompatibility

This Dictionary only raises errors 5, 9, 450 and 457. For example Scripting.Dictionary raises error 32811 if calling ```Remove``` with a key that does not exist while this Dictionary raises error 9 (Subscript out of Range).

### Item (Get) incompatibility

When calling the ```Item``` (Get) property with a key that does not exist, the ```Scripting.Dictionary``` adds a new key-item pair where the key is the key that did not exist previously, and the item is ```Empty```. This kind of behaviour makes sense in the ```Let``` or ```Set``` counterparts of the ```Item``` property - which is why this Dictionary emulates the same behaviour. However, for the ```Get``` property this does not make much sense. On the contrary, it's misleading. Moreover, most likely no one would ever rely on this kind of functionality considering the ```Exists``` method does not throw an error if avoiding errors is the goal.

So, this Dictionary throws error 9 if ```Item``` (Get) is called with a key that is not part of the dictionary.

### Zero-length text, 0 (zero) and Empty incompatibility

The way Scripting.Dictionary compares these 3 values is very misleading. Consider the next code snippet:
```VBA
Dim d As Object: Set d = CreateObject("Scripting.Dictionary")

On Error Resume Next

d.Add Empty, Nothing
Debug.Assert Err.Number = 0

d.Add "", Nothing 'Not allowed because Empty exists
Debug.Assert Err.Number = 457: Err.Clear

d.Add 0, Nothing 'Not allowed because Empty exists
Debug.Assert Err.Number = 457: Err.Clear

'Dict contains Empty
d.RemoveAll

d.Add 0, Nothing
Debug.Assert Err.Number = 0

d.Add "", Nothing
Debug.Assert Err.Number = 0

'Dict contains 0 and "" keys
d.RemoveAll

On Error GoTo 0
```
We can see in the above code that:
- when comparing ```Empty``` to ```""``` (or ```vbNullString```) they are considered equal
- when comparing ```Empty``` to 0 (zero) they are considered equal
- when comparing 0 (zero) to ```""``` (or ```vbNullString```) they are NOT considered equal

This is the standard behaviour when comparing Variants in VBA, clearly outlined [here](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/comparison-operators#remarks). However, for the Scripting.Dictionary this yields different results depending on the order these special values are added to the dictionary.

To avoid this unfortunate and misleading behaviour, this Dictionary distinguishes between these values and allows all 3 at the same time. The following code is valid when using this Dictionary:
```VBA
dict("") = 1
dict(Empty) = 2
dict(0) = 3
Debug.Print dict("") '1
Debug.Print dict(Empty) '2
Debug.Print dict(0) '3
```

## Hashing

A few different hashing strategies were implemented in this Dictionary with the sole purpose that hashing is fast without having to worry about key data type or number of key-item pairs being added.

All hash values are stored (until key is removed or replaced). This requires that the hashes have a good distribution and do not rely on the hash table size. Thus, there is no rehashing in the real sense of the word - for more details see the [Rehashing](#rehashing) section.

All hash values are combined with data type metadata in the upper bits of the hash so that when comparing hash values we are comparing types in the same instruction - for more details see [Metadata](#metadata) section.

All hash values are calculated in the ```GetIndex``` method. This is to avoid any extra function call/stack frame required if having a separate method.

To achieve good hash distribution the following strategies were implemented:
- numbers are first cast to ```Double``` (8 bytes) and then 4 primes are used to get the best hash distribution
- objects are first cast to ```IUnknown``` so that any class instance is only added once to the dictionary i.e. cannot add the same instance as different interfaces. A prime number is used for best hash distribution - in fact, it seems to outperform anything available as seen [here](benchmarking/result_screenshots/add_object_(class1)_win_vba7_x64.png)
- on Mac, all texts are hashed by iterating each wide character (Integer) in a loop using a prime
- on Windows, the Mac strategy is only applied for texts of length 6 or below and for binary compare only. All other texts are hashed using the ```HashVal``` method on a fake instance of ```Scripting.Dictionary``` - with early-binding speed even though there is no dll reference

### Number Hashing

As mentioned above, numbers are first cast to ```Double```. See [Hashing Numbers incompatibility](#hashing-numbers-incompatibility) for details as to why this was chosen.

While initially a single prime number (13) was used to hash all numbers, this was changed in [7d58829](https://github.com/cristianbuse/VBA-FastDictionary/commit/7d58829410082f7899a6933495398868d2c56eab) to 4 prime numbers. The new approach cut the time in half for hashing large integer numbers and also brought small improvements for hashing smaller integers. Both strategies yield the same results for fractional numbers. The numbers are hashed as per [these lines](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L528-L560).

Quick example:
```VBA
Dim d As New Dictionary
d.Add CLng(1234567890), Empty
Debug.Assert d.Exists(CDbl(1234567890))
Debug.Assert d.HashVal(CCur(123.456)) = d.HashVal(CDbl(123.456))
Debug.Assert d.HashVal(CVErr(2042)) <> d.HashVal(CInt(2042)) 'Different because Errors are seen as not-numbers
Debug.Assert d.HashVal(CDbl(CVErr(2042))) = d.HashVal(CInt(2042))
```

### Object Hashing

Objects are first cast to ```IUnknown``` and then the ```IUnknown``` interface pointer is hashed. This ensures each instance is only added once to the dictionary regardless of the interface used.

Code for ```Class1```:
```VBA
Option Explicit
Implements Class2 'Class2 has no code
```

Quick example:
```VBA
Dim d As New Dictionary
Dim c1 As New Class1
Dim c2 As Class2: Set c2 = c1

d.Add c1, Empty
Debug.Assert d.Exists(c1)
Debug.Assert d.Exists(c2)
d.Add c2, Null 'Throws error 457
```

Objects pointers are well distributed anyway because:
- each class instance takes a certain amount of space and so even if adjacent in memory the pointers for 2 instances still have some bytes in between (not consecutive numbers)
- class instances are stored in memory depending on where the memory allocator finds space

So, there is no need to split the pointer into smaller integers to hash. Instead, a modulo prime number is used for best hash distribution. The prime value of 2701 was chosen after running speed tests for all the prime numbers up to 10k. The code is basically [this](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L508-L525).

This strategy seems to yield the best results as seen [here](benchmarking/result_screenshots/add_object_(class1)_win_vba7_x64.png) or [here](benchmarking/result_screenshots/add_object_(collection)_win_vba7_x64.png).

### Text Hashing on Mac

On Mac, all texts are hashed by iterating each wide character (Integer) in a loop. Each char code is added to the previous hash value and the result is multiplied with a prime number. This is repeated until all characters are iterated. A bitmask is used to avoid overflow. The code is [this](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L491-L503). The prime number value of 131 was carefully chosen after many speed tests with different prime values.

For text compare, the key is first passed to the ```VBA.LCase``` function and then the result is hashed.
```LCase``` is fast enough on Mac that there is no need to build a [cached map for each character code](https://github.com/cristianbuse/VBA-FastDictionary/blob/ae95c6e909625c3d95328f64bb3e01a2232485fc/benchmarking/third-party_code/cHashD/modHashD.bas#L42-L52) like ```cHashD``` does.

There is an integer accessor being used (same for Windows) so that reading the char codes in a ```String``` is done fast via a 'fake' array. More details on this in the [Text Hashing on Windows](#text-hashing-on-windows) section below.

### Text Hashing on Windows

The Mac strategy of iterating char codes is only applied in Windows for texts with lengths smaller than 7 and for binary compare only. All other texts are hashed using the ```HashVal``` method on a fake instance of ```Scripting.Dictionary```.

The only reason the Mac strategy for short texts (<7 length) is still used is that it's simply faster - and 7 seems to be the first length that runs faster on the fake instance. Please note that for text compare the iteration strategy is not used on Windows, so no calls to ```LCase``` are made.

#### Scripting.Dictionary.HashVal usefulness

As mentioned above, most texts are hashed using the ```HashVal``` function on a fake Scripting.Dictionary instance. The reason is again speed. For lengthy strings, it is much slower to iterate char codes (in native VBA) than to call this method. See how much better this Dictionary performs on lengthy text keys [here](benchmarking/result_screenshots/add_text_(len_40-60_text_compare_ascii)_win_vba7_x64.png) as opposed to shorter [here](benchmarking/result_screenshots/add_text_(len_5_binary_compare_unicode)_win_vba7_x64.png) solely because it's calling the compiled ```HashVal```.

This would not be needed if code could be compiled in VBA but unfortunately, it cannot. It could be compiled in something like [TwinBasic](https://twinbasic.com) but then it would require all users to reference a dll file which is a big impediment for most VBA users because of distribution problems and also because some users would have IT permission difficulties.

The following will describe how calling Scripting.Dictionary.HashVal is achieved with early binding speed without needing a reference all while avoiding the implementation issues of Scripting.Dictionary.

#### Scripting.Dictionary memory layout

By inspecting the memory layout of a random instance of ```Scripting.Dictionary``` we can conclude that it looks like this:
```VBA
Private Type ScrDictLayout
    vTable1 As LongPtr
    vTable2 As LongPtr
    vTable3 As LongPtr
    vTable4 As LongPtr
    unkPtr1 As LongPtr
    refCount As Long
    firstItemPtr As LongPtr
    lastItemPtr As LongPtr
#If Win64 = 0 Then
    Dummy As Long
#End If
    hashTablePtr As LongPtr
    hashTableSize As Long
    compMode As Long
    localeID As Long
    unkPtr2 As LongPtr
    unkPtr3 As LongPtr
End Type
```

If we run the following code:
```VBA
Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Type ScrDictLayout
    vTable(0 To 3) As LongPtr
    unkPtr1 As LongPtr
    refCount As Long
    firstItemPtr As LongPtr
    lastItemPtr As LongPtr
#If Win64 = 0 Then
    Dummy As Long
#End If
    hashTablePtr As LongPtr
    hashTableSize As Long
    compMode As Long
    localeID As Long
    unkPtr(0 To 1) As LongPtr
End Type

Sub TestScrDictLayout()
#If Mac Then
    MsgBox "Scripting.Dictionary is not available on Mac"
#Else
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim sdl As ScrDictLayout
    '
    CopyMemory sdl, ByVal ObjPtr(d), LenB(sdl)
    Stop
#End If
End Sub
```
when the code stops execution on the ```Stop``` line, we get something like this in the Locals window:
![image](https://github.com/cristianbuse/VBA-FastDictionary/assets/23198997/29877199-25ba-463e-b275-c29318cd9063)

What we see is that compare mode is set to 0 or ```vbBinaryCompare``` (by default) and the hash table size is 1201. All hashes are apparently applied a ```Hash Mod 1201``` before ```HashVal``` returns:
```VBA
Sub TestHashValDefaultRange()
#If Mac Then
    MsgBox "Scripting.Dictionary is not available on Mac"
#Else
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim h As Long
    '
    For i = 0 To 10000000
        h = d.HashVal(i)
        Debug.Assert h >= 0 And h < 1201
    Next i
#End If
End Sub
```

Via memory manipulation, we can change the value of 1201 to something else and we get different hashes:
```VBA
Sub TestHashValDefaultRange()
#If Mac Then
    MsgBox "Scripting.Dictionary is not available on Mac"
#Else
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim sdl As ScrDictLayout
    Dim sizeOffset As LongPtr
    '
    sizeOffset = VarPtr(sdl.hashTableSize) - VarPtr(sdl) 'For both x32 and x64
    CopyMemory ByVal ObjPtr(d) + sizeOffset, &H7FFFFFFF, 4
    
    Debug.Print d.HashVal(123)
    Debug.Print d.HashVal(12345)
    Debug.Print d.HashVal(1234567)
    
    CopyMemory ByVal ObjPtr(d) + sizeOffset, 1201, 4
#End If
End Sub
```

We get this:
```
1123418112 
1178657792 
1234613304 
```

However, if the value is not set back to the original 1201 then a crash will occur. That's because the ```hashTablePtr``` probably points to a table of 1201 size and when the instance is destroyed, the wrong size is being deallocated.

##### Scripting.Dictionary conclusions

Based on the above examples, we can now conclude the following:
- in case of a state loss, using a real Scripting.Dictionary instance for hashing will lead to a crash if we change the hash size to anything else than 1201. Please note ```hashTablePtr``` cannot be changed without leading to a crash, or at best, a memory leak. So, we use a fake instance - see [Faking a Scripting.Dictionary instance](#faking-a-scriptingdictionary-instance) below
- the Scripting.Dictionary never resizes its hash table beyond 1201 which explains the poor performance for more than 32k items even for text keys as seen [here](benchmarking/result_screenshots/add_text_(len_17-23_binary_compare_unicode)_win_vba7_x64.png). There are so many hash collisions that the linear search simply degrades performance
- the Scripting.Dictionary always applies the ```Mod``` operator before returning a hash value and for that, it must read the ```hashTableSize``` (1201 by default) from the heap. This causes real speed problems when spawning many Scripting.Dictionary instances even if each instance has only a few items. See [Scripting.Dictionary heap issue](#scriptingdictionary-heap-issue) below for more details

#### Faking a Scripting.Dictionary instance

Here is a standalone method that calls Scripting.Dictionary.HashVal:
```VBA
Option Explicit

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Type ScrDictLayout
    vTable(0 To 3) As LongPtr
    unkPtr1 As LongPtr
    refCount As Long
    firstItemPtr As LongPtr
    lastItemPtr As LongPtr
#If Win64 = 0 Then
    Dummy As Long
#End If
    hashTablePtr As LongPtr
    hashTableSize As Long
    compMode As Long
    localeID As Long
    unkPtr(0 To 1) As LongPtr
End Type

Public Function HashVal(ByRef v As Variant _
                      , Optional ByVal compMode As VbCompareMethod = vbBinaryCompare _
                      , Optional ByVal hashTableSize As Long = 1201) As Long
           
#If Mac Then
    Err.Raise 5, , "Scripting.Dictionary not available on Mac"
#Else
    Const dictVTableSize As Long = 22
    Const opNumDictHashVal As Long = 21
    Const opNumCollItem As Long = 7
    #If Win64 Then
        Const PTR_SIZE As Long = 8
        Const NULL_PTR As LongLong = 0^
    #Else
        Const PTR_SIZE As Long = 4
        Const NULL_PTR As Long = 0&
    #End If
    '
    Static fakeDict As Collection
    Static vTable(0 To dictVTableSize - 1) As LongPtr
    Static sdl As ScrDictLayout
    Static lcid As Long
    Static isSet As Boolean
    '
    If Not isSet Then
        'Copy entire memory layout for Scripting.Dictionary instance
        CopyMemory sdl, ByVal ObjPtr(CreateObject("Scripting.Dictionary")), LenB(sdl)
        '
        'Copy main virtual function table
        CopyMemory vTable(0), ByVal sdl.vTable(0), dictVTableSize * PTR_SIZE
        '
        lcid = sdl.localeID
        '
        'Replace main virtual table with our own
        sdl.vTable(0) = VarPtr(vTable(0))
        '
        'Map Collection.Item to Dictionary.HashVal
        vTable(opNumCollItem) = vTable(opNumDictHashVal)
        '
        'Set up fake instance
        CopyMemory ByVal VarPtr(fakeDict), VarPtr(sdl), PTR_SIZE
        '
        sdl.hashTablePtr = NULL_PTR
        isSet = True
    End If
    #If Win64 Then
        If VarType(v) = vbLongLong Then 'Check for Object with Default Property
            If Not IsObject(v) Then
                HashVal = fakeDict.Item(CDbl(v))
                Exit Function
            End If
        End If
    #End If
    If compMode < vbDatabaseCompare Then
        sdl.compMode = compMode
        sdl.localeID = lcid
    Else
        sdl.compMode = vbTextCompare
        sdl.localeID = compMode
    End If
    sdl.hashTableSize = hashTableSize
    '
    HashVal = fakeDict.Item(v) 'Dict.HashVal with early-binding speed
#End If
End Function
```
Note that ```fakeDict.Item(v)``` can be replaced with just ```fakeDict(v)``` because VBA "sees" the object as a ```Collection```. 

With the above code, calls can be made to the new ```HashVal``` method which in turn calls Scripting.Dictionary.HashVal with early-binding speed without needing a reference. For example ```Debug.Print HashVal("abc")```.

Of course, the method signature for ```Collection.Item```:  
```
HRESULT Item(
             [in] VARIANT* Index, 
             [out, retval] VARIANT* pvarRet);
```
perfectly matches the signature for ```Scripting.Dictionary.HashVal```:
```
HRESULT HashVal(
                [in] VARIANT* Key, 
                [out, retval] VARIANT* HashVal);
```
if we inspect their type libraries.

There are 2 reasons why such code was not used in this repository:
1) it would require an additional .bas module - the design goal was to have a single class with zero dependencies
2) it adds an extra function call (stack frame) which impacts performance, especially when dealing with millions of keys

##### Scripting.Dictionary heap issue

The initial approach was to have a fake Scripting.Dictionary instance for each instance of this repo's Dictionary. However, this unveiled another Scripting.Dictionary bug which was mentioned briefly before. Luckily, while testing this Dictionary on parsing a JSON file of 12MB size which required tens of thousands of Dictionary instances, it was found that there is a serious speed impact which is only noticeable when using multiple instances. After further investigation and testing, it seems that each Scripting.Dictionary instance (real or fake) must read the hash size from the heap (and also the compare mode) for any key being hashed, and this becomes a problem for multiple instances. This was easily confirmed by moving the storage of the fake layouts into a global array of UDTs inside a standard .bas module, which immediately solved the issue and improved speed about 6-7 times.

For this Dictionary, the fix was obvious - keep a single fake instance in the default ```Dictionary``` instance and access it from all the other instances. This was fixed in [2a39d61](https://github.com/cristianbuse/VBA-FastDictionary/commit/2a39d6183c4de3581d6d87794ec2dc25b3cf5dd4). Presumably if using only one instance, the same heap location is accessed each time and there must be some caching involved because the issue does not occur.

However, this cannot be fixed for Scripting.Dictionary and it now becomes clear why it is not suitable for work that involves multiple instances, like parsing a JSON.

Each instance of this Dictionary uses a fake array of ```Collection``` type (one element) which is set to read the single fake Scripting.Dictionary instance from the default/predeclared Dictionary instance. Also, there is a fake array of type ```Long``` (two elements) which allows each Dictionary instance to read/write compare mode and locale ID into the single fake instance. See the [```Private Type Hasher```](https://github.com/cristianbuse/VBA-FastDictionary/blob/ae95c6e909625c3d95328f64bb3e01a2232485fc/src/Dictionary.cls#L267-L281) struct and the [```InitHasher```](https://github.com/cristianbuse/VBA-FastDictionary/blob/ae95c6e909625c3d95328f64bb3e01a2232485fc/src/Dictionary.cls#L1096-L1169) method.

## Metadata

As briefly mentioned in the [Hashing](#hashing) section, all hash values are combined with data type metadata in the upper bits of the hash with the goal of minimizing comparisons and ultimately being more efficient.

The hash + meta layout is briefly shown in the text at [the top of the GetIndex method](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L432-L447). Here is another representation:
![image](https://github.com/cristianbuse/VBA-FastDictionary/assets/23198997/abd41e6e-7c69-4b5c-853e-562ced91b086)

With this chosen layout, hash values of up to 268,435,456 (0x10000000) can be stored (28 bits), while the next upper 3 bits store metadata about the type. All number data types are combined into a single type for compatibility with Scripting.Dictionary. Similarly, ```vbDataObject``` is combined with ```vbObject``` as per below:

| Key Var Type | Meta Type |
| --- | --- |
| vbEmpty | Empty |
| vbNull | Null |
| vbInteger | Number |
| vbLong | Number |
| vbSingle | Number |
| vbDouble | Number |
| vbCurrency | Number |
| vbDate | Number |
| vbString | Text |
| vbObject | Object |
| vbError | Error |
| vbBoolean | Number |
| vbDataObject | Object |
| vbDecimal | Number |
| vbByte | Number |
| vbLongLong | Number |

The sign bit is not used on x32 because there is separate storage available - see [NewEnum](#newenum) section for more details. However, on x64 the sign bit is used if the Item is an object. This removes the need to have separate storage - the idea is to avoid calling ```IsObject``` repeatedly when the item is retrieved.

For example, if a number key and a text key happen to have the same hash value of 21, when comparing the 2 hashes, they won't be the same thanks to the extra meta bits:
![image](https://github.com/cristianbuse/VBA-FastDictionary/assets/23198997/353edac3-0c5e-4b50-826b-8f697408bb35)

When searching the hash table for a key that was just hashed, we first compare the hash + meta values and only if they are equal, we then compare the actual values. This avoids unnecessary comparisons which are especially slow for texts. In fact, before we even compare the hash + meta values, we first compare a sub-hash called "control byte" but that is described in more detail in the [Hash Map/Table](#hash-maptable) section below.

The meta values are defined in the [HashMeta](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L129-L137) enum. There is also a special value ```hmRemove``` used to mark items/keys that have been removed.

Please note that for special float values like +Inf (positive infinity), -Inf (negative infinity), QNaN (quiet NaN) and SNaN (signaling NaN), the meta bits are all set to 0 (zero). This is to avoid comparison against these special values - the hash comparison will be enough in the same way it is for ```Empty``` and ```Null```.

## Hash Map/Table

After watching the [Designing a Fast, Efficient, Cache-friendly Hash Table, Step by Step](https://www.youtube.com/watch?v=ncHmEUmJZf4) video on YouTube, presented by Matt Kulukundis, the idea of using SSE instructions to compare multiple sub-hashes in a set of just 3 instructions sounded great. Since we cannot natively use SSE instructions in VBA, this Dictionary uses the next best thing - bitwise operations.

The structure looks like [this](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L139-L152). In short:
- the hash map/table is divided into groups (buckets) of fixed size
- each group can hold 4 (x32) or 8 (x64) elements of ```Long``` data type - which will store indexes into the keys/meta storage. This is an array fixed in size at compile time
- each group has an integer "Control" value which is either ```Long``` (x32) or ```LongLong``` (x64) to allow for bitwise operations. Each byte in the control corresponds to one element/index in the group's array. To avoid overflow problems, only the lower 7 bits in each byte are used. Each 7 bits in a control byte are in fact sub-hash values corresponding to the hash of each key pointed by the indexes in the group's array.

The following diagram illustrates the above:
![image](https://github.com/cristianbuse/VBA-FastDictionary/assets/23198997/9e348182-1f5b-4871-8127-5866a8be7c05)

### Sub-hashing

When a key is hashed, the following sub-hashes are computed, in the [GetIndex](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L566-L636) method:
- lower bits in the hash value are used to identify the group slot. The number of bits used depends on the number of groups (table size). A modulo operator is applied to compute this sub-hash. These are the 'pink' bits in the above diagram
- the next 7 bits in the hash value are the control byte. To compute this sub-hash, a bit mask and a right-shift is applied. These are the 'red' bits in the above diagram

### Finding a key

If a key needs to be found, then the steps are the following:
- the full hash+meta is computed and then the 2 sub-hashes
- the group slot sub-hash indicates the group/bucket to search in
- within the group, the control sub-hash byte is checked against the group's control integer using [these](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L588-L593) bitwise operations. This will find all matching positions within the group
- for each matching position, the full hash+meta value is compared
- if the full hashes match, then the actual keys are compared
- the search will continue until a match is found or until a group that was never full is encountered. Please note that during this process the group sub-hash is being incremented

Please note that by comparing sub-hashes and/or hashes minimizes the number of key comparisons, and this greatly improves performance.

### Adding a key

Before adding a key, the steps in the [Finding a key](#finding-a-key) section always run first - to avoid key duplication. So, when adding a key, we already have computed:
- the full hash
- the group sub-hash. Note this will always indicate a group that has available space because this value is incremented in the finding process i.e. it might not correspond with the actual bits in the full hash
- the control byte sub-hash

When adding the Key-Item (and hash+meta) pair to the data storage, in the group indicated by the group sub-hash, the following operations are performed:
- the index for the data is added in the first available position within the group
- the corresponding byte in the control integer is updated with the value of the control sub-hash so that it can be used later for fast find

Code [here](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L347-L352).

When adding a new key, the new corresponding index is added to the hash table, so there is a chance that the group sub-hash indicates a full group. As mentioned above, this is taken care of in the finding process (by incrementing the group slot value), and we always end up adding the new index into a bucket that has available space. Since hashes are not going to have perfect distribution, a new index can be added a few groups/buckets away from the intended position. This is why we track if the group was ever full via a boolean flag. The search for a key will always be done until the first group that was never empty is found. The nice benefit is that we can achieve a higher load factor without needing to resize individual groups/buckets. The current max load is set to 50% which seems to provide a good balance between storage and performance.

## Rehashing

As briefly mentioned in the [Hashing](#hashing) section, there is no rehashing in the real sense of the word. Instead, all hash values are stored along with the key. Still, the method called when the hash table needs to grow is named [```Rehash```](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L394-L421) because most people are familiar with the concept.

By avoiding rehashing the actual keys, this Dictionary can adapt efficiently the hash table size to any number of key-item pairs.

Of course, sub-hash values are re-computed based on the full hash and the new hash table size, when the hash table needs to grow in size. However, these operations are fast (modulo and bit-shifts) and we end up having a hash map resize with a small penalty in performance.

If key-item pairs are removed from the dictionary, then some of the groups in the hash table might remain with unused positions because the ```Add``` process only checks the ```WasEverFull``` flag. Please note this is not a problem for the ```Key``` method which always checks for ```.Count = GROUP_SIZE```. The idea is that a call to ```Rehash``` will actually fix the unused positions and obviously, adding items can easily lead to a resize. Meanwhile, ```Key``` (Let) will not lead to a resize and so the logic is different to avoid endlessly searching for a group slot.

## NewEnum

To enable a ```For Each..``` loop on a class instance in VBA, the instance class must have a defined public method with the special attribute ```Attribute NewEnum.VB_UserMemId = -4``` where ```NewEnum``` is the method name. That method must return an instance that implements the ```IEnumVariant``` interface and VBA will make calls to ```IEnumVARIANT::Next``` until all items are iterated. ```-4``` is simply the ```dispIdMember``` that [```IDispatch::Invoke```](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-invoke) uses to call the method (late-bind call).

Unfortunately, we cannot create a class enumerator (```IEnumVariant```) with native code in VBA. Other developers have tried using assembly injected at runtime - for example [clsTrickHashTable](https://www.vbforums.com/showthread.php?788247-VB6-Hash-table). However, there is no solution available for all platforms and even the ones available often lead to crashes.

A few alternatives were explored in [this](https://codereview.stackexchange.com/questions/287926/iterate-internal-array-for-a-vba-class) Code Review article. In short:
1) we can hijack calls to ```IEnumVARIANT::Next``` on a fake instance and use our own method inside a standard .bas module. This can be done in multiple ways; however, it is too slow (extra stack frames) compared to what a Scripting.Dictionary or a VBA.Collection can do. Moreover, this approach adds a module dependency
2) we can use instances of ```IEnumVARIANT``` as returned by a temporary ```Collection```. If we can make our data structure mimic the items in a Collection's linked list, then the same code that iterates collection items can iterate our own data

This Dictionary uses approach #2 mentioned above but implemented differently for 32-bit and 64-bit versions of VBA.

A Collection item looks like this:
```VBA
Private Type VbCollectionItem
    Data                As Variant
    KeyPtr              As LongPtr
    pPrevIndexedItem    As LongPtr
    pNextIndexedItem    As LongPtr
    pParentItem         As LongPtr
    pRightBranch        As LongPtr
    pLeftBranch         As LongPtr
    bFlag               As Boolean
End Type
```
This is well detailed on VB Forums or [here](https://gist.github.com/wqweto/39822f4fb7090fa086aeff1e2e06e630).

As described in the Code Review article linked above, we only really need the following structure:

| Variant | Unused Pointer | Unused pointer | Next Item Pointer |
| ------- | -------------- | -------------- | ----------------- |

In short, just the first 4 members of a real Collection item.

So, we could create an array of such type instead of an array of ```Variant``` and we then make sure the Next Pointer always links to the next ```Variant```. In fact this is the exact approach used for this Dictionary but for x32 only. For x64 there is a better approach which is less wasteful.

### x64 implementation

Many thanks to [sancarn](https://github.com/sancarn) for his help on figuring out VT_RECORD (vbUserDefined).

As mentioned above, we need the following structure:
| Variant | Unused Pointer | Unused pointer | Next Item Pointer |
| ------- | -------------- | -------------- | ----------------- |

Here is a more useful diagram:  
![image](https://github.com/cristianbuse/VBA-FastDictionary/assets/23198997/815662ff-402d-4070-8048-8e4717521cc2)

We can see on the left side of the diagram that:
- a ```Variant``` uses 24 Bytes on x64. The last 8 Bytes are only used for User Defined Types (UDTs). Keys in the Dictionary cannot be UDTs and so the 8 Bytes are not used
- we are wasting 16 Bytes with unused pointers just to have the Next Pointer at the correct offset

In total, 24 Bytes are not used - the exact size of a ```Variant```.

We can see on the right side of the diagram that:
- we can make use of the 8 unused Bytes in a ```Variant``` and use them for the Next Pointer
- the Next Pointer for Variant at position n would sit in the last Bytes of the Variant at position n+1

In total, we are using all Bytes. Moreover, we don't need a custom structure - we can just continue using an array of ```Variant``` type.

The code that updates the Next Pointers looks like [this](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L1224-L1246)

There is an additional benefit - we can still return the keys array via the ```Keys``` method, without having to iterate the keys, if there are no gaps in the array. See [code](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L841-L843).

### x32 implementation

Unfortunately, we cannot apply the same strategy used for x64. This is because pointers are only 4 Bytes (```Long``` data type) on x32, and the Next Pointer would not have the correct offset. Moreover, the last bytes in a ```Variant``` are used - for example, a ```Double``` would use 8 Bytes.

So, we must use the following custom structure:
| Variant | Unused Pointer | Unused pointer | Next Item Pointer |
| ------- | -------------- | -------------- | ----------------- |

However, we can make use of the space between the Variant and the Next Pointer (Extra information):
![image](https://github.com/cristianbuse/VBA-FastDictionary/assets/23198997/e20e1594-b994-4dbb-b752-9dc84f25e130)

[Here](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L154-L162) is how the custom Variant looks like on x32 and [here](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L246-L255) is the main data storage structure for keys and items, and [this](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L1248-L1264) is the code that updates the Next Pointers.

Although, the ```Keys``` method [has to iterate the keys](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L859-L868) in order to return the keys array, this is not a big drawback because:
- ```Items``` method is not affected
- ```For Each``` can be used on the Dictionary instance directly (```For Each v In dict```) which is anyway faster than iterating the keys array (```For Each v In dict.Keys```)

Using the ```EnumerableVariant``` structure for storing keys brings additional benefits:
- there is no need to have a ```Meta``` array, like we have on x64 to store the hash+meta values
- there is no need to use bitmasks to check if item or key is an object, like we have on x64

Compared to not implementing this functionality, there are only 8 extra Bytes being used per Variant: 4 for the Next Pointer and 4 for the 2 Boolean flags (is Item/Key Object). That's because the hash+meta uses 4 Bytes regardless. While the 2 Boolean flags are slightly faster than using bitmasks, the speed difference is negligeable. Overall, having the iterator functionality trumps the need for the additional space.

### Enumerator management

When using the class iterator described above, there are a few scenarios that we need to consider:
- a ```For Each``` loop can be called inside an existing ```For Each``` loop
- items/keys can be added/removed while inside of a ```For Each``` loop. This can lead to a [re-allocation of the data](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L311-L328)
- ```NewEnum``` can be called and stored to be used at a later stage

To account for all the scenarios above, this Dictionary has additional management in the [```RemoveUnusedEnums```](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L1268-L1294) and [```ShiftEnumPointer```](https://github.com/cristianbuse/VBA-FastDictionary/blob/2cce33fb21498720d992538e546d17e6822381f0/src/Dictionary.cls#L1296-L1335) methods.

## Additional functionality

Compared to a Scripting.Dictionary, this Dictionary has a few extra methods that can be useful:
- ```AllowDuplicateKeys```
- ```Factory``` - returns a new Dictionary instance
- ```Index``` - returns the index for a specified Key
- ```ItemAtIndex``` - returns or replaces the Item at the specified index
- ```KeyAtIndex``` - returns the Key at the specified index
- ```KeysItems2D``` - returns a 2D array of all the Keys and Items
- ```PredictCount``` - if the number of Key-Item pairs is known upfront or if a good guess is possible, then a call to ```PredictCount``` with the expected number of pairs will prepare the internal size of the hash map so that there are no calls made to ```Rehash```. This results in better performance
- ```Self``` - this method is useful in ```With New Dictionary``` code blocks

## OLE Automation

Although this dictionary does not rely on other libraries or references, it still requires that the basic ```OLE Automation``` reference in enabled. This is because the ```IUnknown``` interface is needed to properly hash object keys - for more details see this [discussion](https://github.com/cristianbuse/VBA-FastDictionary/discussions/5).

This is such a basic reference that even VBA itself uses it - if we look at the type library for VBA, we can see:
```
library VBA
{
    // TLib : OLE Automation : {00020430-0000-0000-C000-000000000046}
    importlib("stdole2.tlb");
```

## x64 Assembly

This class implements fixes by overwriting x64 assembly bytes for the following known VBA x64 bugs:
- [Bug with For Each enumeration on x64 Custom Classes](https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes)
- [VBA takes wrong branch at If-statement - severe compiler bug?](https://stackoverflow.com/questions/65041832/vba-takes-wrong-branch-at-if-statement-severe-compiler-bug) which is same as [issue #10](https://github.com/cristianbuse/VBA-FastDictionary/issues/10)

First, let's understand how x64 methods are called.

### VBA class method call mechanism

The first 8 bytes (4 bytes on x32) pointed by a class instance pointer hold the address of the class virtual table.

#### Class virtual table

Each VBA class is derived from [IDispatch](https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nn-oaidl-idispatch) which in turn is derived from [IUnknown](https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nn-unknwn-iunknown). In other words, the virtual table for a VBA class looks like this:
```VBA
IUnknown::QueryInterface     'Position 0
IUnknown::AddRef
IUnknown::Release
IDispatch::GetTypeInfoCount
IDispatch::GetTypeInfo
IDispatch::GetIDsOfNames
IDispatch::Invoke            'Position 6
PublicMethod1                'Position 7
...
PublicMethodN
PrivateMethod1
...
PrivateMethodN
```

We can read the pointer to the first function in a VBA class by using the following code. In an empty project add a ```Class1``` with the following code:
```VBA
Option Explicit

Public Function Test() As Variant
    Debug.Print "Test"
End Function
Public Sub Test2()
    Dim i As Long
    For i = 1 To 3
        Debug.Print i
    Next i
End Sub
```

Now add the following code in a standard .bas module and run the ```TestFuncPointer``` method. This will print the pointer for the ```Test``` method of ```Class1``` to the [Immediate](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/immediate-window) window.
```VBA
Option Explicit

#If Win64 Then
    Const PTR_SIZE = 8
#Else
    Const PTR_SIZE = 4
#End If
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Function MemLongPtr(ByVal addr As LongLong) As LongPtr
    CopyMemory MemLongPtr, ByVal addr, PTR_SIZE
End Function

Sub TestFuncPointer()
    Dim c As New Class1
    Dim vTable As LongPtr
    Dim testPtr As LongPtr
    Dim b() As Byte
    '
    vTable = MemLongPtr(ObjPtr(c))
    testPtr = MemLongPtr(vTable + PTR_SIZE * 7)
    '
    Debug.Print testPtr
End Sub
```

For those familiar with [COM](https://en.wikipedia.org/wiki/Component_Object_Model), there will be a virtual table for each interface that a class implements.

A class has many default interfaces (e.g. ```IMarshall```, ```IConnectionPointCointainer```, ```_DClass``` etc.) but ```IClassModuleEvt``` in particular is useful for finding the pointers for ```_Initialize``` and ```_Terminate```, which in turn call, if present, the ```Class_Initialize``` and ```Class_Terminate``` (more details on class footprint / layout can be found [here](https://codereview.stackexchange.com/questions/294682/faster-vb6-vba-class-deallocation)).

In addition to the default interfaces, a class will have a virtual table for each implemented interface via the ```Implements``` keyword. It will also have an interface for each use of the ```WithEvents``` keyword.

#### Class method code

If we inspect the bytes pointed by the ```testPtr``` in the above example, it seems there are some assembly instructions. **These are not the actual method instructions**. After testing, it seems this is simply code that is being called via late-binding (e.g. when the method is called on a variable declared ```As Object```) - a call via ```IDispatch::Invoke```. It is not called at all via early-binding (e.g. when the method is called on a variable declared ```As Class1```).

The following diagram shows how the bytes are organized:
![methodCall](https://github.com/user-attachments/assets/642d7473-b555-4cc4-886f-f7908e446cf6)

Let's see the assembly for a ```Function``` and a ```Sub``` on x64.
Add the following code in a standard .bas module and run the ```TestClassMethodASM``` method:
```VBA
Option Explicit
#If Win64 Then

Const PTR_SIZE = 8

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Function MemLongPtr(ByVal addr As LongLong) As LongPtr
    CopyMemory MemLongPtr, ByVal addr, PTR_SIZE
End Function

Sub TestClassMethodASM()
    Dim c As New Class1
    Dim vTable As LongPtr
    Dim testPtr As LongPtr
    Dim b() As Byte
    '
    vTable = MemLongPtr(ObjPtr(c))
    testPtr = MemLongPtr(vTable + PTR_SIZE * 7) '7 because IDispatch has methods 0 to 6
    '
    ReDim b(1 To 69)
    CopyMemory b(1), ByVal testPtr, UBound(b)
    Debug.Print StringToHex(CStr(b))
End Sub
Public Function StringToHex(ByRef s As String) As String
    Static map(0 To 255) As String
    Dim b() As Byte: b = s
    Dim i As Long

    If LenB(map(0)) = 0 Then
        For i = 0 To 255
            map(i) = Right$("0" & Hex$(i), 2)
        Next i
    End If

    StringToHex = Space$(LenB(s) * 2 + 2)
    Mid$(StringToHex, 1, 2) = "0x"

    For i = LBound(b) To UBound(b)
        Mid$(StringToHex, (i + 1) * 2 + 1, 2) = map(b(i))
    Next i
End Function
#End If
```
The resulting hex can be translated to this:
```assembly
; Typical assembly for a class Function with no arguments
;-------------------------------------------------------------------------------
66490F6EEC             ; MOVQ XMM5,R12                    ; Saves R12 value in XMM5     
48B8 987307FAFA7F0000  ; MOV RAX,00007FFAFA077398         ; Copies literal value into RAX - this value seems to always be the same
488B00                 ; MOV RAX,QWORD PTR [RAX]          ; Reads memory value pointed by RAX into RAX (Dereference)
4C8B20                 ; MOV R12,QWORD PTR [RAX]          ; Reads memory value pointed by RAX into R12 (Dereference)
4981EC 10000000        ; SUB R12,0000000000000010         ; Adds space (negative) for 2 pointers to the stack, one for instance pointer and one for function return
498BC4                 ; MOV RAX,R12                      ; Saves R12 value into RAX - seems useless because RAX gets overwritten just below in the second last instruction
49898C24 00000000      ; MOV QWORD PTR [R12+00000000],RCX ; Saves RCX value at the address pointed by R12
49899424 08000000      ; MOV QWORD PTR [R12+00000008],RDX ; Saves RDX value at the address pointed by R12 + 0x8
48BA F0D1254C0E020000  ; MOV RDX,0000020E4C25D1F0         ; Pushes literal value to RDX. The value matches the value at -8 offset from function pointer
48B8 8A1800FAFA7F0000  ; MOV RAX,00007FFAFA00188A         ; Pushes literal value to RAX. This is jumped to in the next instruction and is the same value for all methods in the class i.e. a wrapper
FFE0                   ; JMP RAX                          ; Jumps to asm pointer by RAX as mentioned above. This wrapper will use the address passed in RDX to call the required method
; ...


; Typical assembly for a class Sub with no arguments
;-------------------------------------------------------------------------------
66490F6EEC             ; MOVQ XMM5,R12                      
48B8 987307FAFA7F0000  ; MOV RAX,00007FFAFA077398         
488B00                 ; MOV RAX,QWORD PTR [RAX]          
4C8B20                 ; MOV R12,QWORD PTR [RAX]          
4981EC 08000000        ; SUB R12,0000000000000008         ; Adds space (negative) for instance pointer to the stack
498BC4                 ; MOV RAX,R12                      
49898C24 00000000      ; MOV QWORD PTR [R12+00000000],RCX 
                                                          ; RDX value is no longer saved at the address pointed by R12 + 0x8 (compared to a Function)
48BA 302820570E020000  ; MOV RDX,0000020E4C25D1F0         ; The literal value is changed every time code is compiled
48B8 8A1800FAFA7F0000  ; MOV RAX,00007FFAFA00188A         
FFE0                   ; JMP RAX                          
; ...
```

The asm code pointed by the entries in the class virtual table is only called via late-binding and in turn calls a wrapper code (pointed by the ```RAX``` register) that eventually calls the PCode (pointed by the ```RDX``` register).

The ```RDX``` value always matches the value at function pointer address less 8 bytes e.g.```[testPtr - 8]```. For early-binded calls, this pointer at -8 offset is used to access the PCode.

So how does the late-binding wrapper (pointed by RAX) look like? Here are just a few lines:
```assembly
4C0FB74210            ; MOVZX R8,WORD PTR [RDX+10]       ; Reads 2-Byte Int from RDX+16 which is related to early-binding PCode. This value seems 0 all the time. Then zero-extends to a 8-Byte (6 upper bytes zero) and moves value into R8
4D290424              ; SUB QWORD PTR [R12],R8           ; Subtract value of R8 from the value at address pointed by R12. This suggests the 2-Byte Int was some kind of offset
488B5208              ; MOV RDX,QWORD PTR [RDX+08]       ; Reads Pointer from RDX+8 which is basically pointing to the actual function PCode instructions. Then writes that pointer to RDX
48B91D7C75F6FA7F0000  ; MOV RCX,00007FFAF6757C1D         ; Writes literal into RCX. Seems to be same value even after recompilation (including app restart)
55                    ; PUSH RBP                         ; Pushes RBP onto stack (decrements RSP by 8 and then stores RBP at the top of the stack). Basically starts setting up a new stack frame
488BEC                ; MOV RBP,RSP                      ; Saves stack pointer value into RBP
4881EC50010000        ; SUB RSP,0000000000000150         ; Allocates 336 bytes on the stack

4C89BC2480000000      ; MOV QWORD PTR [RSP+00000080],R15 ; Saves R15 value at RSP+128
4C89742478            ; MOV QWORD PTR [RSP+78],R14       ; Saves R14 value at RSP+120
4C896C2470            ; MOV QWORD PTR [RSP+70],R13       ; Saves R13 value at RSP+112
66480F7E6C2468        ; MOVQ QWORD PTR [RSP+68],XMM5     ; Saves XMM5 value at RSP+104
48897C2460            ; MOV QWORD PTR [RSP+60],RDI       ; Saves RDI value at RSP+96
4889742458            ; MOV QWORD PTR [RSP+58],RSI       ; Saves RSI value at RSP+88
48895C2450            ; MOV QWORD PTR [RSP+50],RBX       ; Saves RBX value at RSP+80

4C8BEC                ; MOV R13,RSP                      ; Saves stack pointer value into R13
4C89A578FFFFFF        ; MOV QWORD PTR [RBP-00000088],R12 ; Saves R12 value at RBP-136
4D8BF4                ; MOV R14,R12                      ; Saves R12 into R14
488BDA                ; MOV RBX,RDX                      ; Saves RDX into RBX
48899D68FFFFFF        ; MOV QWORD PTR [RBP-00000098],RBX ; Saves RBX at RBP-152 (RBP is base pointer for current stack frame)
48898D38FFFFFF        ; MOV QWORD PTR [RBP-000000C8],RCX ; Saves RCX at RBP-200
498BC0                ; MOV RAX,R8                       ; Saves R8 into RAX

DBE2                  ; FNCLEX                           ; Clears the floating-point exception flags in the x87 FPU (Floating Point Unit) status word
488D15E20E0700        ; LEA RDX,[0000000000070F56]       ; Loads RIP+0x70EE2 (# 0x70F56) into RDX. RIP is Instruction Pointer Register

488B3B                ; MOV RDI,QWORD PTR [RBX]          ; Reads pointer from RBX which is basically the PCode pointer. Then writes that pointer to RDI (Destination Index Register - often used in string operations)
488B7760              ; MOV RSI,QWORD PTR [RDI+60]       ; Reads Pointer from RDI+96 then writes that pointer to RSI
4889B560FFFFFF        ; MOV QWORD PTR [RBP-000000A0],RSI ; Saves RSI value at RBP-160

488B7708              ; MOV RSI,QWORD PTR [RDI+08]       ; Reads Pointer from RDI+8 then writes that pointer to RSI
488B7628              ; MOV RSI,QWORD PTR [RSI+28]       ; Reads Pointer from RSI+40 then writes that pointer to RSI. Linked-list?
488B7610              ; MOV RSI,QWORD PTR [RSI+10]       ; Reads Pointer from RSI+16 then writes that pointer to RSI.
488975B8              ; MOV QWORD PTR [RBP-48],RSI       ; Saves RSI value at RBP-72
48895590              ; MOV QWORD PTR [RBP-70],RDX       ; Saves the previous address 0x70F56 (RIP+0x70EE2) at RBP-112
C7458800000000        ; MOV DWORD PTR [RBP-78],00000000  ; Saves 4-Byte Int 0x0 at RBP-120
4C8D05EF620000        ; LEA R8,[0000000000006393]        ; Loads RIP+0x62EF (# 0x6393) into RDX
493BC8                ; CMP RCX,R8                       ; 
0F85C1000000          ; JNE 000000000000016E

66F743141000          ; TEST WORD PTR [RBX+14],0010
744A                  ; JE 00000000000000FF

498B0E                ; MOV RCX,QWORD PTR [R14]
480BC9                ; OR RCX,RCX                       ; For updating flags, specifically Zero Flag and Sign Flag
0F8483000000          ; JE 0000000000000144
; ...
```

#### Class method call

Based on the above, we can heavily simplify a class method call to this:
- early-binded call
  - Push arguments to stack
  - Read Temp pointer at Function pointer - 8
  - Read PCODE Header pointer at Temp pointer + 8
  - Read PCode Arguments Size (including instance pointer and function return) at PCode Header pointer + 8 (WORD)
  - Read PCode Variables Size at PCode Header Pointer + 10 (WORD)
  - Read PCode Size at PCode Header Pointer + 12 (DWORD)
  - Run PCode at PCode Header Pointer - PCode Size
  - Decrease Stack Pointer by PCode Arguments Size + PCode Variables Size
  - Return
- late-binded call
  - Push arguments to stack
  - Call ```IDispatch::Invoke```
  - Create stack frame
  - Push arguments to stack (ByRef i.e. pointing to original arguments)
  - Call ASM at function pointer
  - Increase R12 to account for the ByRef arguments pushed by ```Invoke```
  - Copy Arguments starting with R12+0x0
  - Put Temp Pointer into RDX
  - Jump to global wrapper pointed by RAX
  - Read PCODE Header pointer at Temp pointer (RDX) + 8
  - Create new stack frame
  - Save R12 which accounts for arguments pushed by ```Invoke``` 
  - Read PCode Arguments Size (including instance pointer and function return) at PCode Header pointer + 8 (WORD)
  - Read PCode Variables Size at PCode Header Pointer + 10 (WORD)
  - Read PCode Size at PCode Header Pointer + 12 (DWORD)
  - Run PCode at PCode Header Pointer - PCode Size
  - Decrease Stack Pointer by PCode Arguments Size + PCode Variables Size
  - Return to global wrapper
  - Remove stack frame
  - Return to ```Invoke```
  - Remove stack frame
  - Return

### Stack Bug fixes

Unfortunately, for the above mentioned bugs with [For Each](https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes) and [Class_Terminate](https://stackoverflow.com/questions/65041832/vba-takes-wrong-branch-at-if-statement-severe-compiler-bug), VBA does not increase the stack size correctly and code in the last called method ends up overwriting values that are part of the previous stack frame. **This stack misalignment issue only happens via late-bound calls**.

This class simply adds additional, unused space to the stack. This makes sure that values in the previous stack frame are not overwritten when the bugs occur. We cannot possibly know how big the previous stack frame is, so we artificially increase the stack frame by adding a large amount e.g. 2048 bytes, of which most won't be used in normal circumstances.

The strategy applied to ```Class_Terminate``` is different from the one applied to ```NewEnum``` and ```Item``` (Get). This is because ```Class_Terminate``` is ```Private``` and the call coming from ```IClassModuleEvt::_Terminate``` is always calling the late-bound assembly bytes that in turn call the global wrapper that calls into PCode. However, for the other methods, the call can be made directly to PCode since they are ```Public```, and so the fix must be dynamic based on call type.

Stack fix code [here](https://github.com/cristianbuse/VBA-FastDictionary/blob/b019abe22fe93acd488c642c62509416aceadc75/src/Dictionary.cls#L1559-L1656).

#### Item (Get) stack fix

As seen in issue [#16](https://github.com/cristianbuse/VBA-FastDictionary/issues/16), if a class instance is stored in a variable of type ```Variant```, and the call is made implicitly to the default method, then the previous stack frame is not "detected" properly and it's being overwritten from the default method.

Minimal example would be:
```VBA
Public Function TestDictionary()
    Dim v As Variant: Set v = New Dictionary
    Set v(0) = New Dictionary
    v(0).Add Empty, "test" 'This will either cause a crash or pass an 'Unsupported variant type' to the 'Add' method
End Function
```
Please note this issue **can be reproduced with any VBA class** as seen in the above mentioned issue, not just with ```Dictionary```.

A more simplified view of a call to ```Item```, compared to [Class method call](#class-method-call), can be written as:
- early-bound
  - arguments are correctly pushed to stack
  - PCode is run
  - PCode argument size is popped from the stack when PCode finishes execution
- late-bound
  - assembly code increases R12 register by the size of the arguments
  - global wrapper is called which creates a new stack frame and takes R12 into account
  - PCode is run
  - PCode argument size is popped from the stack when PCode finishes execution

If we are to fix the late-bound calls by increasing the size of R12 with additional unused space, then we must also increase the PCode argument size if we don't want to quickly run out of stack space. E.g. Push 2048 bytes then also pop 2048 bytes. The problem is that we cannot do this just one time and be done with it, because an early-bound call would push the normal argument size to the stack and then it would remove a larger amount which would lead to a stack corruption and a crash. The solution is to "track" the call type and then adjust accordingly the amount being popped from the stack.

The following code is the automatically generated assembly for ```Item``` (Get):
```assembly
66490F6EEC           ; MOVQ XMM5,R12                    ; Saves R12 value in lower 8 bytes of XMM5    
48B8B873F4FFFC7F0000 ; MOV RAX,00007FFCFFF473B8         ; Copies literal value into RAX - this value seems to always be the same
488B00               ; MOV RAX,QWORD PTR [RAX]          ; Reads memory value pointed by RAX into RAX (Dereference)
4C8B20               ; MOV R12,QWORD PTR [RAX]          ; Reads memory value pointed by RAX into R12 (Dereference)
4981EC18000000       ; SUB R12,0000000000000018         ; Adds space (negative) for 3 pointers to the stack, one for instance pointer, one for function return and one for the argument
498BC4               ; MOV RAX,R12                      ; Saves R12 value into RAX - seems useless because RAX gets overwritten just below in the second last instruction
49898C2400000000     ; MOV QWORD PTR [R12+00000000],RCX ; Saves RCX value at the address pointed by R12
4989942408000000     ; MOV QWORD PTR [R12+00000008],RDX ; Saves RDX value at the address pointed by R12 + 0x8
4D89842410000000     ; MOV QWORD PTR [R12+00000010],R8  ; Saves R8 value at the address pointed by R12 + 0x10
48BAC87B6BADCB020000 ; MOV RDX,000002CBAD6B7BC8         ; Pushes literal value to RDX. The value matches the value at -8 offset from function pointer
48B88A1EEDFFFC7F0000 ; MOV RAX,00007FFCFFED1E8A         ; Pushes literal value to RAX. This is jumped to in the next instruction and is the same value for all methods in the class i.e. global wrapper
FFE0                 ; JMP RAX                          ; Jumps to asm pointer by RAX as mentioned above. This wrapper will use the address passed in RDX to call PCode
```

Notice that the ```MOV RAX,R12``` is useless because ```RAX``` gets overwritten in the second last instruction. Also, the offsets to ```R12``` don't need to be 4-Byte (DWORD). This can be updated like this:
```assembly
66490F6EEC           ; MOVQ XMM5,R12              ; Same as original
48B8B873F4FFFC7F0000 ; MOV RAX,00007FFCFFF473B8   ; Same as original
488B00               ; MOV RAX,QWORD PTR [RAX]    ; Same as original
4C8B20               ; MOV R12,QWORD PTR [RAX]    ; Same as original
;-------------------------------------------------------------------
4981EC00080000       ; SUB R12,0000000000000800   ; Replace 0x18 (24 bytes) with 0x800 (2048 bytes)
                                                  ; Save 3 bytes by removing 'MOV RAX,R12'
49890C24             ; MOV QWORD PTR [R12],RCX    ; Same as 'MOV QWORD PTR [R12+00000000],RCX' but written with 4 bytes instead of 8 i.e. save 4 bytes
4989542408           ; MOV QWORD PTR [R12+08],RDX ; Same as 'MOV QWORD PTR [R12+00000008],RDX' but written with 5 bytes instead of 8 i.e. save 3 bytes
4D89442410           ; MOV QWORD PTR [R12+10],R8  ; Same as 'MOV QWORD PTR [R12+00000010],R8'  but written with 5 bytes instead of 8 i.e. save 3 bytes
                                                  ; We now have 3 + 4 + 3 + 3 = 13 freed bytes that we can use freely
48B80495888DCB020000 ; MOV RAX,000002CB8D889504   ; Add address of the global variable that will hold the call type to RAX (10 bytes)
C60001               ; MOV BYTE PTR [RAX],01      ; Write byte value 0x01 to the variable that holds call type (3 bytes)
;-------------------------------------------------------------------
48BAC87B6BADCB020000 ; MOV RDX,000002CBAD6B7BC8   ; Same as original
48B88A1EEDFFFC7F0000 ; MOV RAX,00007FFCFFED1E8A   ; Same as original
FFE0                 ; JMP RAX                    ; Same as original
```

In short, we get the late-bound assembly to write a ```Byte``` value of ```1``` to a global variable that will indicate the call type. Since VBA uses little-endian memory layout, the variable can be of type ```Long``` (Enum actually).
The asm is updated [here](https://github.com/cristianbuse/VBA-FastDictionary/blob/b019abe22fe93acd488c642c62509416aceadc75/src/Dictionary.cls#L1636-L1651). This is done while preserving the exact number of assembly bytes with no loss of original functionality.

With the above in place, each call to ```Item``` (Get) can access the variable that holds the call type and can adjust the correct stack size that needs to be popped. Code [here](https://github.com/cristianbuse/VBA-FastDictionary/blob/b019abe22fe93acd488c642c62509416aceadc75/src/Dictionary.cls#L370-L380).

We find the ```Item``` (Get) function pointer because we strategically place it as the second method in the class (i.e. 9th method in the virtual table), immediately after ```NewEnum```.

#### NewEnum stack fix

The [For Each](https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes) bug usually leads to a crash. [Previous solution](https://github.com/cristianbuse/VBA-FastDictionary/blob/4b93590de56cec7e92bc1f741ee068d1e87e9527/src/Dictionary.cls#L1494-L1542) was simply solving the late-bound calls by artificially increasing stack size for ```NewEnum``` both in the late-bound asm and in the PCode stack pop, just once. However, with the addition of [Item (Get) stack fix](#item-get-stack-fix), the solution was upgraded to account for early-bound calls to ```NewEnum``` and so the stack size to be popped is also [dynamically adjusted](https://github.com/cristianbuse/VBA-FastDictionary/blob/b019abe22fe93acd488c642c62509416aceadc75/src/Dictionary.cls#L341-L351) in the exact way as ```Item``` (Get).

The main difference when compared to ```Item``` is that ```NewEnum``` did not have an argument and so, a dummy one was added: ```Optional ByVal dummyRegisterR8 As Long```. This does not affect functionality while providing the much needed ```MOV QWORD PTR [R12+00000010],R8``` instruction which allows us to save 3 bytes by rewriting to ```MOV QWORD PTR [R12+10],R8``` i.e. use a BYTE offset instead of a DWORD. Without this workaround, we would need to allocate separate memory space with code-execution privileges, in order to achieve the desired result.

We find the enumerator function pointer because we strategically place it as the first method in the class (i.e. 8th method in the virtual table).

#### Class_Terminate stack fix

Besides the well-known [Class_Terminate](https://stackoverflow.com/questions/65041832/vba-takes-wrong-branch-at-if-statement-severe-compiler-bug) bug, there are a few other scenarios that were raised in issues [#10](https://github.com/cristianbuse/VBA-FastDictionary/issues/10) and [#15](https://github.com/cristianbuse/VBA-FastDictionary/issues/15), and they are all related to the same stack misalignment.

The solution for ```Class_Terminate``` is simpler because we know there are no arguments, no function return, no local variables and the calls are always made via the late-bound assembly. However, it is not as simple as increasing both the ```R12``` register size and the argument size for PCode, because this leads to an 'Out of stack space' issue as seen in [#11](https://github.com/cristianbuse/VBA-FastDictionary/issues/11), when lots of instances are terminated from within the same stack frame e.g. in a loop or when a parent collection/container is terminated.

While the first call to ```Class_Terminate``` should have a larger stack frame to avoid the bug, all subsequent nested calls should have the original size. Tracking the nesting level was done anyway for managed nesting deallocation. So, we use the nesting level to adjust the stack size as seen [here](https://github.com/cristianbuse/VBA-FastDictionary/blob/b019abe22fe93acd488c642c62509416aceadc75/src/Dictionary.cls#L1879-L1903).

We find the ```_Terminate``` pointer by finding the 4th method in the virtual table of the ```IClassModuleEvt``` interface as seen [here](https://github.com/cristianbuse/VBA-FastDictionary/blob/b019abe22fe93acd488c642c62509416aceadc75/src/Dictionary.cls#L1585-L1589).
