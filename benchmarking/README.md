# Benchmarking VBA-FastDictionary

This Dictionary has decent performance while compatible across all VBA platforms and applications. However, as you will see below, in most cases it is the fastest solution when compared to what is already available.

## Table of Contents

- [What's already available](#whats-already-available)
  - [Pros and Cons](#pros-and-cons)
    - [VBA.Collection](#vbacollection)
    - [Scripting.Dictionary](#scriptingdictionary)
    - [VBA-Dictionary](#vba-dictionary)
    - [cHashD](#chashd)
- [Benchmarking code](#benchmarking-code)
- [Classes tested](#classes-tested)
- [Results](#results)
- [Conclusions](#conclusions)
  - [Final thoughts](#final-thoughts)

## What's already available

The following will be used for comparison:
- the built-in ```VBA.Collection``` class
- the ```Scripting.Dictionary``` available under the Microsoft Scripting Runtime (scrrun.dll) on Windows only

The [Data Structures](https://github.com/sancarn/awesome-vba?tab=readme-ov-file#data-structures) section of @Sancarn 's [Awesome VBA](https://github.com/sancarn/awesome-vba) repo is a good source for finding alternatives. There surely are many other implementations out there but they are not discussed here.

Two of the publicly available dictionaries which are listed in the above repo will be used for speed comparison with this repository:
- [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary)
- [cHashD](https://www.vbforums.com/showthread.php?834515-Simple-and-fast-lightweight-HashList-Class-(no-APIs)) - final version will be used (available as zip) in post [#30](https://www.vbforums.com/showthread.php?834515-Simple-and-fast-lightweight-HashList-Class-(no-APIs)&p=5479053&viewfull=1#post5479053)

There are other classes listed in the Awesome VBA repo but those are just extensions of ```VBA.Collection``` or ```Scripting.Dictionary``` and so will not be used. The [clsTrickHashTable](https://www.vbforums.com/showthread.php?788247-VB6-Hash-table) has a nice assembly enumerator but it's not Mac compatible and is not suitable for VBA7 because of its reliance on API calls which are slow in VBA7 - this is tested and explained in [this Code Review question](https://codereview.stackexchange.com/questions/270258/evaluate-performance-of-dll-calls-from-vba).

### Pros and Cons

#### [VBA.Collection](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/collection-object)

**Pros**
- native - works in all VBA environments
- fast instantiation
- decent speed for up to 100k key-item pairs (when compared to the other solutions)

**Cons**
- keys can only be of ```String``` data type - any other type needs to be casted
- can only compare keys in text compare mode
- cannot retrieve keys unless using memory manipulation
- cannot enumerate keys using ```For Each..```
- slow speed for 100k+ items (when compared to the other solutions)

#### [Scripting.Dictionary](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object)

**Pros**
- in general, the fastest solution for up to 32k key-item pairs where keys are ```String``` or ```Object```
- the standard/go-to solution for most VBA users on Windows

**Cons**
- not available on Mac
- slow for 32k+ key-item pairs - it's hash table size is fixed to 1201 - see [Scripting.Dictionary Conclusions](/Implementation.md#scriptingdictionary-conclusions) for more details
- very slow for number keys especially outside the -9,999,999 to 9,999,999 range because all numbers are casted to ```Single``` before they are hashed  - see [this](/Implementation.md#hashing-numbers-incompatibility) for more details
- has speed issues when multiple instances are being used - the implementation is constantly reading the compare mode and the hash size (1201) from the heap - see [Scripting.Dictionary Conclusions](/Implementation.md#scriptingdictionary-conclusions) for more details

#### [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary)

This is a class that wraps around Scripting.Dictionary on Windows and uses a combination of Collections (for keys) and Arrays on Mac. 

**Pros**
- easy drop-in replacement for Scripting.Dictionary 
- works on Mac and Windows on both x32 and x64

**Cons**
- slower on Windows because it uses late-binding for the internal Scripting.Dictionary
- very slow on Mac because of the strategy used
- has serious bugs. For example, on Mac, an integer key of value ```2``` is considered equal to a text key of value ```2__2``` while in text comparison mode. Same would happen on Windows if choosing to set the compiler constant ```UseScriptingDictionaryIfAvailable``` to ```False```
- inherits all the cons of Scripting.Dictionary

#### [cHashD](https://www.vbforums.com/showthread.php?834515-Simple-and-fast-lightweight-HashList-Class-(no-APIs))

This is a class that does its own hashing and uses only arrays and some nice logic.

**Pros**
- has methods like ```Add```, ```Count```, ```Item``` (default), ```Exists```, ```Remove``` which makes it consistent with Scripting.Dictionary and thus easy to use for most people
- works on Mac and Windows on both x32 and x64. For this to be True and to avoid VBA7 API overhead issues, the [050f128](https://github.com/cristianbuse/VBA-FastDictionary/commit/050f12800274a7408ed2aea17c55f1bb1009b30c) commit does the necessary changes so that comparison is fair
- has good performance when you know the number of items in advance
- allows duplicates keys
- hash is fast for keys of ```String``` type when the keys are relatively short and that's because it iterates the Wide-Character codes as Integers

**Cons**
- has poor performance when you don't know the number of items in advance. The default value of ```16384``` as in ```ReInit 16384``` in the ```Class_Initialize``` has a huge performance impact when many instances of this class are created. Using a smaller initial number improves performance by making initialization faster but then works poorly if too many items are inserted which cause too many collisions and linear search
- doesn't really handle all data types for keys. For example, all Variant/Decimal keys would get assigned the exact same hash value and would all go into the same hash bucket (last bucket). ```LongLong``` is not covered as a data type and would end up in the same bucket as Decimal. Same for Variant/Error
- for object keys it ignores the interface i.e. same instance can be passed as 2 different interfaces and the keys are considered different
- the hash is slow if using keys of ```String``` type when the keys are lengthy. This can be fixed by compiling in something like TwinBasic but will require a dll reference to be used in VBA

## Benchmarking code

Special thanks to [Guido](https://github.com/guwidoe) for his excellent code modules:
- [LibStringTools.bas](https://github.com/guwidoe/VBA-StringTools)
- [LibTimer.bas](https://gist.github.com/guwidoe/5c74c64d79c0e1cd1be458b0632b279a)

Copies of both modules are available under the [third_party_code](src/third_party_code) folder.

The [Benchmarking.xlsm](Benchmarking.xlsm) Excel file contains all the code listed under [benchmarking/src](src) folder.
The results are being written to 8 worksheets (one for each operation being measured) and can be exported as screenshots.

The actual tests are in the [BenchTests.bas](src/BenchTests.bas) module and they call the main ```Benchmark``` method of the [Benchmarking.bas](src/Benchmarking.bas) module with various key inputs.

## Classes tested

Tim Hall's ```VBA-Dictionary```, ```VBA.Collection``` and ```Scripting.Dictionary``` are tested as-is.

For ```cHashD``` class we test 3 approaches:
1) the default size of 16384 for the hash table size - which as you will see does not perform very well
2) assuming the number of key-item pairs to add is known in advance then the hash table is sized prior to adding items with the goal of achieving approximately 10% load
3) assuming the number of key-item pairs to add is known in advance then the hash table is sized prior to adding items with the goal of achieving approximately 38.5% load

For the new Dictionary (this repo) we have 2 approaches for adding items:
1) default rehashing - the hash table resizes when the load reaches 50%
2) assuming the number of key-item pairs to add is known in advance then the hash table is sized prior to adding items with the goal of achieving approximately 50% load - this is slightly faster than the default rehashing but not by much

## Results

All screenshots are saved under the [result_screenshots](result_screenshots) folder. The results can be reproduced by running the speed tests in the [Benchmarking.xlsm](Benchmarking.xlsm) file under the [BenchTests.bas](src/BenchTests.bas) module.

Tests:
- T01 - Mixed Keys
- T02 - Double Fractional Keys
- T03 - Double Large Integer Keys
- T04 - Double Small Integer Keys
- T05 - Long Large Keys
- T06 - Long Small Keys
- T07 - Class1 Keys
- T08 - Collection Keys
- T09 - Text Keys (length 5, binary compare)
- T10 - Text Keys (length 5, text compare)
- T11 - Text Keys (length 8-12, binary compare)
- T12 - Text Keys (length 8-12, text compare)
- T13 - Text Keys (length 17-23, binary compare)
- T14 - Text Keys (length 17-23, text compare)
- T15 - Text Keys (length 40-60, binary compare)
- T16 - Text Keys (length 40-60, text compare)

### Win x64 VBA7

| Operation | T01 | T02 | T03 | T04 | T05 | T06 | T07 | T08 | T09 | T10 | T11 | T12 | T13 | T14 | T15 | T16 |
| --------- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Add       | ![T01](result_screenshots/add_mixed_(binary_compare)_win_vba7_x64.png) | ![T02](result_screenshots/add_number_(double_fractional)_win_vba7_x64.png) | ![T03](result_screenshots/add_number_(double_large_ints)_win_vba7_x64.png) | ![T04](result_screenshots/add_number_(double_small_ints)_win_vba7_x64.png) | ![T05](result_screenshots/add_number_(long_large)_win_vba7_x64.png) | ![T06](result_screenshots/add_number_(long_small)_win_vba7_x64.png) | ![T07](result_screenshots/add_object_(class1)_win_vba7_x64.png) | ![T08](result_screenshots/add_object_(collection)_win_vba7_x64.png) | ![T09](result_screenshots/add_text_(len_17-23_binary_compare_unicode)_win_vba7_x64.png) | ![T10](result_screenshots/add_text_(len_17-23_text_compare_unicode)_win_vba7_x64.png) | ![T11](result_screenshots/add_text_(len_40-60_binary_compare_ascii)_win_vba7_x64.png) | ![T12](result_screenshots/add_text_(len_40-60_text_compare_ascii)_win_vba7_x64.png) | ![T13](result_screenshots/add_text_(len_5_binary_compare_unicode)_win_vba7_x64.png) | ![T14](result_screenshots/add_text_(len_5_text_compare_unicode)_win_vba7_x64.png) | ![T15](result_screenshots/add_text_(len_8-12_binary_compare_ascii)_win_vba7_x64.png) | ![T16](result_screenshots/add_text_(len_8-12_text_compare_ascii)_win_vba7_x64.png) |
| Exists (True) | ![T01](result_screenshots/exists(true)_mixed_(binary_compare)_win_vba7_x64.png) | ![T02](result_screenshots/exists(true)_number_(double_fractional)_win_vba7_x64.png) | ![T03](result_screenshots/exists(true)_number_(double_large_ints)_win_vba7_x64.png) | ![T04](result_screenshots/exists(true)_number_(double_small_ints)_win_vba7_x64.png) | ![T05](result_screenshots/exists(true)_number_(long_large)_win_vba7_x64.png) | ![T06](result_screenshots/exists(true)_number_(long_small)_win_vba7_x64.png) | ![T07](result_screenshots/exists(true)_object_(class1)_win_vba7_x64.png) | ![T08](result_screenshots/exists(true)_object_(collection)_win_vba7_x64.png) | ![T09](result_screenshots/exists(true)_text_(len_17-23_binary_compare_unicode)_win_vba7_x64.png) | ![T10](result_screenshots/exists(true)_text_(len_17-23_text_compare_unicode)_win_vba7_x64.png) | ![T11](result_screenshots/exists(true)_text_(len_40-60_binary_compare_ascii)_win_vba7_x64.png) | ![T12](result_screenshots/exists(true)_text_(len_40-60_text_compare_ascii)_win_vba7_x64.png) | ![T13](result_screenshots/exists(true)_text_(len_5_binary_compare_unicode)_win_vba7_x64.png) | ![T14](result_screenshots/exists(true)_text_(len_5_text_compare_unicode)_win_vba7_x64.png) | ![T15](result_screenshots/exists(true)_text_(len_8-12_binary_compare_ascii)_win_vba7_x64.png) | ![T16](result_screenshots/exists(true)_text_(len_8-12_text_compare_ascii)_win_vba7_x64.png) |
| Exists (False) | ![T01](result_screenshots/exists(false)_mixed_(binary_compare)_win_vba7_x64.png) | ![T02](result_screenshots/exists(false)_number_(double_fractional)_win_vba7_x64.png) | ![T03](result_screenshots/exists(false)_number_(double_large_ints)_win_vba7_x64.png) | ![T04](result_screenshots/exists(false)_number_(double_small_ints)_win_vba7_x64.png) | ![T05](result_screenshots/exists(false)_number_(long_large)_win_vba7_x64.png) | ![T06](result_screenshots/exists(false)_number_(long_small)_win_vba7_x64.png) | ![T07](result_screenshots/exists(false)_object_(class1)_win_vba7_x64.png) | ![T08](result_screenshots/exists(false)_object_(collection)_win_vba7_x64.png) | ![T09](result_screenshots/exists(false)_text_(len_17-23_binary_compare_unicode)_win_vba7_x64.png) | ![T10](result_screenshots/exists(false)_text_(len_17-23_text_compare_unicode)_win_vba7_x64.png) | ![T11](result_screenshots/exists(false)_text_(len_40-60_binary_compare_ascii)_win_vba7_x64.png) | ![T12](result_screenshots/exists(false)_text_(len_40-60_text_compare_ascii)_win_vba7_x64.png) | ![T13](result_screenshots/exists(false)_text_(len_5_binary_compare_unicode)_win_vba7_x64.png) | ![T14](result_screenshots/exists(false)_text_(len_5_text_compare_unicode)_win_vba7_x64.png) | ![T15](result_screenshots/exists(false)_text_(len_8-12_binary_compare_ascii)_win_vba7_x64.png) | ![T16](result_screenshots/exists(false)_text_(len_8-12_text_compare_ascii)_win_vba7_x64.png) |
| Item (Get) | ![T01](result_screenshots/item(get)_mixed_(binary_compare)_win_vba7_x64.png) | ![T02](result_screenshots/item(get)_number_(double_fractional)_win_vba7_x64.png) | ![T03](result_screenshots/item(get)_number_(double_large_ints)_win_vba7_x64.png) | ![T04](result_screenshots/item(get)_number_(double_small_ints)_win_vba7_x64.png) | ![T05](result_screenshots/item(get)_number_(long_large)_win_vba7_x64.png) | ![T06](result_screenshots/item(get)_number_(long_small)_win_vba7_x64.png) | ![T07](result_screenshots/item(get)_object_(class1)_win_vba7_x64.png) | ![T08](result_screenshots/item(get)_object_(collection)_win_vba7_x64.png) | ![T09](result_screenshots/item(get)_text_(len_17-23_binary_compare_unicode)_win_vba7_x64.png) | ![T10](result_screenshots/item(get)_text_(len_17-23_text_compare_unicode)_win_vba7_x64.png) | ![T11](result_screenshots/item(get)_text_(len_40-60_binary_compare_ascii)_win_vba7_x64.png) | ![T12](result_screenshots/item(get)_text_(len_40-60_text_compare_ascii)_win_vba7_x64.png) | ![T13](result_screenshots/item(get)_text_(len_5_binary_compare_unicode)_win_vba7_x64.png) | ![T14](result_screenshots/item(get)_text_(len_5_text_compare_unicode)_win_vba7_x64.png) | ![T15](result_screenshots/item(get)_text_(len_8-12_binary_compare_ascii)_win_vba7_x64.png) | ![T16](result_screenshots/item(get)_text_(len_8-12_text_compare_ascii)_win_vba7_x64.png) |
| Item (Let) | ![T01](result_screenshots/item(let)_mixed_(binary_compare)_win_vba7_x64.png) | ![T02](result_screenshots/item(let)_number_(double_fractional)_win_vba7_x64.png) | ![T03](result_screenshots/item(let)_number_(double_large_ints)_win_vba7_x64.png) | ![T04](result_screenshots/item(let)_number_(double_small_ints)_win_vba7_x64.png) | ![T05](result_screenshots/item(let)_number_(long_large)_win_vba7_x64.png) | ![T06](result_screenshots/item(let)_number_(long_small)_win_vba7_x64.png) | ![T07](result_screenshots/item(let)_object_(class1)_win_vba7_x64.png) | ![T08](result_screenshots/item(let)_object_(collection)_win_vba7_x64.png) | ![T09](result_screenshots/item(let)_text_(len_17-23_binary_compare_unicode)_win_vba7_x64.png) | ![T10](result_screenshots/item(let)_text_(len_17-23_text_compare_unicode)_win_vba7_x64.png) | ![T11](result_screenshots/item(let)_text_(len_40-60_binary_compare_ascii)_win_vba7_x64.png) | ![T12](result_screenshots/item(let)_text_(len_40-60_text_compare_ascii)_win_vba7_x64.png) | ![T13](result_screenshots/item(let)_text_(len_5_binary_compare_unicode)_win_vba7_x64.png) | ![T14](result_screenshots/item(let)_text_(len_5_text_compare_unicode)_win_vba7_x64.png) | ![T15](result_screenshots/item(let)_text_(len_8-12_binary_compare_ascii)_win_vba7_x64.png) | ![T16](result_screenshots/item(let)_text_(len_8-12_text_compare_ascii)_win_vba7_x64.png) |
| Key (Let) | ![T01](result_screenshots/key(let)_mixed_(binary_compare)_win_vba7_x64.png) | ![T02](result_screenshots/key(let)_number_(double_fractional)_win_vba7_x64.png) | ![T03](result_screenshots/key(let)_number_(double_large_ints)_win_vba7_x64.png) | ![T04](result_screenshots/key(let)_number_(double_small_ints)_win_vba7_x64.png) | ![T05](result_screenshots/key(let)_number_(long_large)_win_vba7_x64.png) | ![T06](result_screenshots/key(let)_number_(long_small)_win_vba7_x64.png) | ![T07](result_screenshots/key(let)_object_(class1)_win_vba7_x64.png) | ![T08](result_screenshots/key(let)_object_(collection)_win_vba7_x64.png) | ![T09](result_screenshots/key(let)_text_(len_17-23_binary_compare_unicode)_win_vba7_x64.png) | ![T10](result_screenshots/key(let)_text_(len_17-23_text_compare_unicode)_win_vba7_x64.png) | ![T11](result_screenshots/key(let)_text_(len_40-60_binary_compare_ascii)_win_vba7_x64.png) | ![T12](result_screenshots/key(let)_text_(len_40-60_text_compare_ascii)_win_vba7_x64.png) | ![T13](result_screenshots/key(let)_text_(len_5_binary_compare_unicode)_win_vba7_x64.png) | ![T14](result_screenshots/key(let)_text_(len_5_text_compare_unicode)_win_vba7_x64.png) | ![T15](result_screenshots/key(let)_text_(len_8-12_binary_compare_ascii)_win_vba7_x64.png) | ![T16](result_screenshots/key(let)_text_(len_8-12_text_compare_ascii)_win_vba7_x64.png) |
| For Each / NewEnum | ![T01](result_screenshots/forEach_mixed_(binary_compare)_win_vba7_x64.png) | ![T02](result_screenshots/forEach_number_(double_fractional)_win_vba7_x64.png) | ![T03](result_screenshots/forEach_number_(double_large_ints)_win_vba7_x64.png) | ![T04](result_screenshots/forEach_number_(double_small_ints)_win_vba7_x64.png) | ![T05](result_screenshots/forEach_number_(long_large)_win_vba7_x64.png) | ![T06](result_screenshots/forEach_number_(long_small)_win_vba7_x64.png) | ![T07](result_screenshots/forEach_object_(class1)_win_vba7_x64.png) | ![T08](result_screenshots/forEach_object_(collection)_win_vba7_x64.png) | ![T09](result_screenshots/forEach_text_(len_17-23_binary_compare_unicode)_win_vba7_x64.png) | ![T10](result_screenshots/forEach_text_(len_17-23_text_compare_unicode)_win_vba7_x64.png) | ![T11](result_screenshots/forEach_text_(len_40-60_binary_compare_ascii)_win_vba7_x64.png) | ![T12](result_screenshots/forEach_text_(len_40-60_text_compare_ascii)_win_vba7_x64.png) | ![T13](result_screenshots/forEach_text_(len_5_binary_compare_unicode)_win_vba7_x64.png) | ![T14](result_screenshots/forEach_text_(len_5_text_compare_unicode)_win_vba7_x64.png) | ![T15](result_screenshots/forEach_text_(len_8-12_binary_compare_ascii)_win_vba7_x64.png) | ![T16](result_screenshots/forEach_text_(len_8-12_text_compare_ascii)_win_vba7_x64.png) |
| Remove | ![T01](result_screenshots/remove_mixed_(binary_compare)_win_vba7_x64.png) | ![T02](result_screenshots/remove_number_(double_fractional)_win_vba7_x64.png) | ![T03](result_screenshots/remove_number_(double_large_ints)_win_vba7_x64.png) | ![T04](result_screenshots/remove_number_(double_small_ints)_win_vba7_x64.png) | ![T05](result_screenshots/remove_number_(long_large)_win_vba7_x64.png) | ![T06](result_screenshots/remove_number_(long_small)_win_vba7_x64.png) | ![T07](result_screenshots/remove_object_(class1)_win_vba7_x64.png) | ![T08](result_screenshots/remove_object_(collection)_win_vba7_x64.png) | ![T09](result_screenshots/remove_text_(len_17-23_binary_compare_unicode)_win_vba7_x64.png) | ![T10](result_screenshots/remove_text_(len_17-23_text_compare_unicode)_win_vba7_x64.png) | ![T11](result_screenshots/remove_text_(len_40-60_binary_compare_ascii)_win_vba7_x64.png) | ![T12](result_screenshots/remove_text_(len_40-60_text_compare_ascii)_win_vba7_x64.png) | ![T13](result_screenshots/remove_text_(len_5_binary_compare_unicode)_win_vba7_x64.png) | ![T14](result_screenshots/remove_text_(len_5_text_compare_unicode)_win_vba7_x64.png) | ![T15](result_screenshots/remove_text_(len_8-12_binary_compare_ascii)_win_vba7_x64.png) | ![T16](result_screenshots/remove_text_(len_8-12_text_compare_ascii)_win_vba7_x64.png) |

## Conclusions

- For Add-ing keys of type Object, Fractional Numbers or lengthy Strings, this Dictionary is the fastest solution for almost any number of key-item pairs. Even for shorter Strings and Integer numbers keys, this Dictionary is still the fastest for large number of key-item pairs, while for small number of pairs the difference is so insignificant that it does not justify using any other solution. Keep in mind that this Dictionary is fast even without knowing the number of pairs in advance (rehashing) while solutions like ```cHashD``` simply cannot operate decently without knowing in advance
- For checking if a key Exists, this Dictionary is simply the best choice. Only the Scripting.Dictionary is slightly better for small number of key-item pairs and the difference is insignificant. For cases when the keys do not exist then this Dictionary is even faster than for cases when keys exist. That's because it checks a whole group (8 keys on x64 and 4 keys on x32) in a few bitwise operations (on sub-hashes) without ever needing to compare the keys themselves
- Retrieving Item(s) via Get or setting via Let/Set makes this Dictionary the best choice for any type of keys
- Setting Keys to a different value is only possible for ```Scripting.Dictionary```, ```VBA-Dictionary``` and this Dictionary with the latter being the fastest choice
- Iterating keys using a ```For Each..``` loop is only supported by ```Scripting.Dictionary``` and this Dictionary while the latter is just faster
- This Dictionary was not optimized for Remove and so it is not the fastest in this regard, with the exception of large number of key-item pairs or lengthy text keys. However, the trade-off is to have faster Add, Item and Exists operations

### Final thoughts

Although it might seem that ```Scripting.Dictionary``` is faster for up to 32k items and for specific key types, the difference is usually of microseconds or a few milliseconds when compared to this Dictionary. However, the lack of compatibility with Mac and some of the other issues mentioned ([reading heap is slow for multiple instances](/Implementation.md#scriptingdictionary-heap-issue)) makes this Dictionary a better choice over ```Scripting.Dictionary```.

Although it might seem that ```cHashD``` is faster for adding keys of type ```Long``` (up to a certain number of pairs), that comes with the downside of not being compatible with ```Scripting.Dictionary```. For example key ```CLng(1)``` is seen as different than ```CDbl(1)``` while ```Scripting.Dictionary``` and this repo's Dictionary 'see' them as the same number. Moreover, ```cHashD``` does not perform well without knowing the number of items in advance and so the comparison is just for illustrative purposes.

Overall, the Dictionary presented in this repository seems to be the best choice for any scenario, any key type, key length, platform or number of key-item pairs added.
