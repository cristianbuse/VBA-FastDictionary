# VBA-FastDictionary [![Mentioned in Awesome VBA](https://awesome.re/mentioned-badge.svg)](https://github.com/sancarn/awesome-vba)
Fast Native Dictionary for VBA. Compatible with Windows and Mac. Does not require any DLL references or any kind of external libraries.

Compatible with VBA7 / [twinBASIC](https://twinbasic.com/) / VB6 / VBA6.

The Dictionary presented in this repository is designed to be a drop-in replacement for Scripting.Dictionary (Microsoft Scripting Runtime - scrrun.dll on Windows). However, there are a few differences, and their purpose is to make this Dictionary the better choice from a functionality perspective. For more details see [Compatibility with Scripting.Dictionary](Implementation.md#compatibility-with-scriptingdictionary).

Special thanks to [Guido](https://github.com/guwidoe) for his contribution (see [Benchmarking code](benchmarking/README.md#benchmarking-code)) and for his invaluable feedback and ideas.

## Installation

Download the [latest release](https://github.com/cristianbuse/VBA-FastDictionary/releases/latest), extract and import the ```Dictionary.cls``` class into your project.

Although the ```OLE Automation``` project reference should be enabled by default (fundamental COM), please enable it if it's disabled. For more details see [OLE Automation](https://github.com/cristianbuse/VBA-FastDictionary/blob/master/Implementation.md#ole-automation) section.

## Why use VBA-FastDictionary

- Easy replacement for Scripting.Dictionary
- Quite fast for a native VBA Dictionary 
- Mac compatible
- Overcomes VB Class Deallocation design and has linear deallocation speed. See [Faster VB6 / VBA class deallocation](https://codereview.stackexchange.com/questions/294682/faster-vb6-vba-class-deallocation)
- Overcomes VB x64 Compiler Issues. See [x64 Assembly fixes](https://github.com/cristianbuse/VBA-FastDictionary/blob/master/Implementation.md#x64-assembly)
- Automatic Rehashing and Resizing
- Supports ```For Each```
- Has [Additional functionality](https://github.com/cristianbuse/VBA-FastDictionary/blob/master/Implementation.md#additional-functionality)

## Testing

Download the [latest release](https://github.com/cristianbuse/VBA-FastDictionary/releases/latest), extract and import the ```TestDictionary.bas``` module into your project.
Run ```RunAllDictionaryTests``` method. On failure, execution will stop on the first failed Assert.

## Benchmarking

In most cases, this Dictionary is the fastest solution when compared to what is already available. Please see [Benchmarking](benchmarking/README.md) for more details.

## Implementation

For those interested in how this Dictionary works and why some design decisions were made, please see [Implementation](Implementation.md) for more details.

## Demo

```VBA
Dim d As New Dictionary
Dim c As Collection
Dim v As Variant

d.Add "abc", 1
d.Add "Abc", New Collection

d("Abc").Add 1
Debug.Print d.Item("abc")
Debug.Print d.Exists("ABC") 'False

d.Remove "abc"
d.RemoveAll

d.CompareMode = vbTextCompare
d.Add "abc", 1
Debug.Print d.Exists("ABC") 'True

On Error Resume Next
d.Add "Abc", New Collection 'Throws error 457
Debug.Print Err.Number
On Error GoTo 0

d.Add 123, 456
Set c = New Collection
d.Add c, "Test"

d.Item(123) = 789
d.Key(c) = "Test"

Debug.Print
For Each v In d
    Debug.Print v
Next v

Debug.Print
For Each v In d.Items
    Debug.Print v
Next v
```
