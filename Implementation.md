## Implementation

This document outlines the design decisions made in creating an efficient and robust dictionary, providing a comprehensive overview of its functionality. Please refer to the table of contents below for easy navigation to specific design decisions.

## Table of Contents

- [Compatibility with Scripting.Dictionary](#compatibility-with-scriptingdictionary)
  - [Hashing Numbers](#hashing-numbers-incompatibility)
  - [Error numbers incompatibility](#error-numbers-incompatibility)
  - [Item (Get) incompatibility](#item-get-incompatibility)

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
Where Scripting.Dictionary would always downgrade a ```Double``` to a ```Single``` to perform the comparison. This is of course is line with how VBA behaves as seen [here](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/comparison-operators#remarks):

> When a Single is compared to a Double, the Double is rounded to the precision of the Single

However, the new Dictionary (this repo) casts ```Single``` to ```Double```. This seems more of an improvement rather than an issue not the mention that the number of collisions is greatly reduced thus improving speed by orders of magnitude.

### Error numbers incompatibility

This Dictionary only raises errors 5, 9, 450 and 457. For example Scripting.Dictionary raises error 32811 if calling ```Remove``` with a key that does not exist while this Dictionary raises error 9 (Subscript out of Range).

The main reason not to adhere to the same error numbers is speed. For example in the [```Item```](https://github.com/cristianbuse/VBA-FastDictionary/blob/ae95c6e909625c3d95328f64bb3e01a2232485fc/src/Dictionary.cls#L666-L667) method, instead of using an extra ```If``` statement to check if the call to ```GetIndex``` returns ```NOT_FOUND```, the code simply continues and if the key was indeed missing, error 9 is raised anyway when trying to access the storage array with an invalid index. Other methods like ```Remove``` will simply return error 9 for consistency. The avoidal of the extra ```If``` statement does not impact speed for a few items but for millions of items there is a difference and speed for chosen over consistency here.

### Item (Get) incompatibility

When calling the ```Item``` (Get) property with a key that does not exist, the ```Scripting.Dictionary``` adds a new key-item pair where the key is the key that did not exist previously and the item is ```Empty```. This kind of behaviour makes sense in the ```Let``` or ```Set``` counterparts of the ```Item``` property - which is why this Dictionary emulates the same behaviour. However, for the ```Get``` property this does not make much sense. On the contrary, it's misleading. Moverover, most likely no one would ever rely on this kind of functionality considering the ```Exists``` method does not throw an error if avoiding errors is the goal.

So, this Dictionary throws error 9 if ```Item``` (Get) is called with a key that is not part of the dictionary.