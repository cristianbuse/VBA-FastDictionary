## Implementation

This document outlines the design decisions made in creating an efficient and robust dictionary, providing a comprehensive overview of its functionality. Please refer to the table of contents below for easy navigation to specific design decisions.

This Dictionary does not require any DLL references or any kind of external libraries. Works on Mac and Windows on both x32 and x64.

## Table of Contents

- [Compatibility with Scripting.Dictionary](#compatibility-with-scriptingdictionary)
  - [Hashing Numbers incompatibility](#hashing-numbers-incompatibility)
  - [Error numbers incompatibility](#error-numbers-incompatibility)
  - [Item (Get) incompatibility](#item-get-incompatibility)
- [Hashing](#hashing)
  - [Number Hashing](#number-hashing)
  - [Object Hashing](#object-hashing)
  - [Text Hashing on Mac](#text-hashing-on-mac)
  - [Text Hashing on Windows](#text-hashing-on-windows)
    - [Scripting.Dictionary.HashVal benefit](#scriptingdictionaryhashval-benefit)
      - [Scripting.Dictionary Conclusions](#scriptingdictionary-conclusions)
    - [Faking a Scripting.Dictionary instance](#faking-a-scriptingdictionary-instance)
***

## Compatibility with ```Scripting.Dictionary```

The Dictionary presented in this repository is designed to be a drop-in replacement for ```Scripting.Dictionary``` (Microsoft Scripting Runtime - scrrun.dll on Windows). However, there are a few differences and their purpose is to make this Dictionary the better choice from a functionality perspective.

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
W Scripting.Dictionary would always downgrade a ```Double``` to a ```Single``` to perform the comparison. This is of course is line with how VBA behaves as seen [](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/comparison-operators#remarks):

> When a Single is compared to a Double, the Double is rounded to the precision of the Single

However, the new Dictionary (this repo) casts ```Single``` to ```Double```. This seems more of an improvement rather than an issue not the mention that the number of collisions is greatly reduced thus improving speed by orders of magnitude.

### Error numbers incompatibility

This Dictionary only raises errors 5, 9, 450 and 457. For example Scripting.Dictionary raises error 32811 if calling ```Remove``` with a key that does not exist while this Dictionary raises error 9 (Subscript out of Range).

The main reason not to ad to the same error numbers is speed. For example in the [```Item```](https://github.com/cristianbuse/VBA-FastDictionary/blob/ae95c6e909625c3d95328f64bb3e01a2232485fc/src/Dictionary.cls#L666-L667) method, instead of using an extra ```If``` statement to check if the call to ```GetIndex``` returns ```NOT_FOUND```, the code simply continues and if the key was indeed missing, error 9 is raised anyway when trying to access the storage array with an invalid index. Other methods like ```Remove``` will simply return error 9 for consistency. The avoidal of the extra ```If``` statement does not impact speed for a few items but for millions of items there is a difference and speed for chosen over consistency here.

### Item (Get) incompatibility

When calling the ```Item``` (Get) property with a key that does not exist, the ```Scripting.Dictionary``` adds a new key-item pair where the key is the key that did not exist previously and the item is ```Empty```. This kind of behaviour makes sense in the ```Let``` or ```Set``` counterparts of the ```Item``` property - which is why this Dictionary emulates the same behaviour. However, for the ```Get``` property this does not make much sense. On the contrary, it's misleading. Moverover, most likely no one would ever rely on this kind of functionality considering the ```Exists``` method does not throw an error if avoiding errors is the goal.

So, this Dictionary throws error 9 if ```Item``` (Get) is called with a key that is not part of the dictionary.

## Hashing

A few different hashing strategies were implemented in this Dictionary with the sole purpose that hashing is fast without having to worry about key data type or number of key-item pairs being added. Although more details will follow below, here are the strategies in short:
- numbers are first casted to ```Double``` (8 bytes) and 4 primes are used to get the best hash distribution
- objects are first casted to ```IUnknown``` so that any class instance is only added once to the dictionary i.e. cannot add the same instance as different interfaces. A prime number is used for best hash distribution - in fact it seems to outperform anything available as seen [here](benchmarking/result_screenshots/add_object_(class1)_win_vba7_x64.png)
- on Mac, all texts are hashed by iterating each wide character (Integer) in a loop using a prime
- on Windows, the Mac strategy is only applied for texts with length of 6 or below and for binary compare only. All other texts are hashed using the ```HashVal``` method on a fake instance of ```Scripting.Dictionary``` - with early-binding speed even though there is no dll reference
- all hash values are combined with data type metadata and stored with the purpose of rehashing much faster. This requires that the hashes have a good spread to start with and are not reliant on the hash table size
- sub-hash values are computed based on the hash and the current hash table size. These are the ones used to find the correct hash group/bucket and also the position within the group
- the only place to do the hash is in the ```GetIndex``` method. This is to avoid any extra stack frames required if having a separate method

### Number Hashing

As mentioned above, numbers are first casted to ```Double```. See [Hashing Numbers incompatibility](#hashing-numbers-incompatibility) for details as to why this was chosen.

While initially a single prime number (13) was used to hash all numbers, this was changed in [7d58829](https://github.com/cristianbuse/VBA-FastDictionary/commit/7d58829410082f7899a6933495398868d2c56eab) to 4 prime numbers. The new approach cut the time in half for hashing large integer numbers and also brought small improvements for hashing smaller integers. Both strategies were yielding same results for fractional numbers. The numbers are hashed as per [these lines](https://github.com/cristianbuse/VBA-FastDictionary/blob/7d58829410082f7899a6933495398868d2c56eab/src/Dictionary.cls#L528-L541).

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

Objects are first casted to ```IUnknown``` and then the ```IUnknown``` interface pointer is hashed. This ensures each instance is only added once to the dictionary regardless of the interface being used.

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
- class instances are stored in memory depending of where the memory allocator finds space

So, there is no need to split the pointer into smaller integers to hash. Instead a modulo prime number is used for best hash distribution. The prime value of 2701 was chosen after running speed tests for all the prime numbers up to 10k. The code is basically [this](https://github.com/cristianbuse/VBA-FastDictionary/blob/7d58829410082f7899a6933495398868d2c56eab/src/Dictionary.cls#L516-L525).

This strategy seems to yield the best results as seen [here](benchmarking/result_screenshots/add_object_(class1)_win_vba7_x64.png) or [here](benchmarking/result_screenshots/add_object_(collection)_win_vba7_x64.png).

### Text Hashing on Mac

On Mac, all texts are hashed by iterating each wide character (Integer) in a loop. Each char code is added to the previous hash value and the result is multiplied with a prime number. This is repeated until all characters are iterated. A bitmask is used to avoid overflow. The code is [this](https://github.com/cristianbuse/VBA-FastDictionary/blob/ae95c6e909625c3d95328f64bb3e01a2232485fc/src/Dictionary.cls#L492-L508). The prime number value of 131 was carefully chosen after many speed tests with different prime values.

For text compare, the key is first passed to the ```VBA.LCase``` function and only then it is hashed.
```LCase``` is fast enough on Mac that there is no need to build a [cached map for each character code](https://github.com/cristianbuse/VBA-FastDictionary/blob/ae95c6e909625c3d95328f64bb3e01a2232485fc/benchmarking/third-party_code/cHashD/modHashD.bas#L42-L52) like ```cHashD``` does.

There is an integer accessor being used (same for Windows) so that reading the char codes in a ```String``` is done fast via a 'fake' array. More details on this in the [Text Hashing on Windows](#text-hashing-on-windows) section below.

### Text Hashing on Windows

The Mac strategy of iterating char codes is only applied for texts with length of 6 or below and for binary compare only. All other texts are hashed using the ```HashVal``` method on a fake instance of ```Scripting.Dictionary```.

Why still use the Mac strategy for short texts (<7 len)? It's simply faster and this is the only reason - also explains why 7 is not an arbitrary number. Please note that for text compare the iteration strategy is not used and so no calls to ```LCase``` are being made.

#### Scripting.Dictionary.HashVal benefit

As mentioned above, most texts are hashed using the ```HashVal``` function on a fake Scripting.Dictionary instance. The reason is again speed. For lenghty strings it is much slower to iterate char codes (in native VBA) than to call this method. See how much better this Dictionary performs on lengthy text keys [here](benchmarking/result_screenshots/add_text_(len_40-60_text_compare_ascii)_win_vba7_x64.png) as opposed to shorter [here](benchmarking/result_screenshots/add_text_(len_5_binary_compare_unicode)_win_vba7_x64.png) solely because it's calling the compiled ```HashVal```.

This would not be needed if code could be compiled in VBA but unfortunately it cannot. It could be compiled in something like [TwinBasic](https://twinbasic.com) but then it would require all users to reference a dll file which is a big impediment for most VBA users because of distribution problems but also because some users would have IT permission difficulties.

The following will describe how calling Scripting.Dictionary.HashVal is achieved with early binding without needing a reference all while avoiding the implementation issues of Scripting.Dictionary.

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
when the code stops execution on the ```Stop``` line we get something like this in the Locals window:
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

Via memory manipulation we can change the value of 1201 to something else and we get different hashes:
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

##### Scripting.Dictionary Conclusions

Based on the above examples, we can now conclude the following:
- in case of a state loss, using a real Scripting.Dictionary instance for hashing would lead to a crash. Please note ```hashTablePtr``` cannot be changed as it still leads to a crash and setting it to zero is a memory leak. So, we use a fake instance - see [Faking a Scripting.Dictionary instance](#faking-a-scriptingdictionary-instance) below
- the Scripting.Dictionary never resizes it's hash table beyond 1201 which explains the poor performance for more than 32k items even for text keys as seen [here](benchmarking/result_screenshots/add_text_(len_17-23_binary_compare_unicode)_win_vba7_x64.png). There are so many hash collisions that the linear search simply degrades performance
- the Scripting.Dictionary always applies the ```Mod``` operator before returning a hash value and for that it must read the ```hashTableSize``` (1201 by default) from the heap. This causes real speed problems when spawning many Scripting.Dictionary instances even if each instance has only a few items. See []() below for more details

#### Faking a Scripting.Dictionary instance

To be continued..

