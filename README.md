# VBA-FastDictionary
Fast Native Dictionary for VBA. Compatible with Windows and Mac.

Can be used as a replacement for ```Scripting.Dictionary``` on Windows.

Special thanks to [Guido](https://github.com/guwidoe) for his contribution (see [Benchmarking code](benchmarking/README.md#benchmarking-code)) and for his invaluable feedback and ideas.

## Installation

Download the latest [release](https://github.com/cristianbuse/VBA-FastDictionary/releases), extract and import the ```Dictionary.cls``` class into your project.

## Testing

Download the latest [release](https://github.com/cristianbuse/VBA-FastDictionary/releases), extract and import the ```TestDictionary.bas``` module into your project.
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