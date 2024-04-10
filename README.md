# VBA-FastDictionary
Fast Native Dictionary for VBA. Compatible with Windows and Mac.

Can be used as a replacement for ```Scripting.Dictionary``` on Windows.

## Installation

Download the latest [release](https://github.com/cristianbuse/VBA-FastDictionary/releases), extract and import the ```Dictionary.cls``` class into your project.

## Testing

Download the latest [release](https://github.com/cristianbuse/VBA-FastDictionary/releases), extract and import the ```TestDictionary.bas``` module into your project.
Run ```RunAllDictionaryTests``` method. On failure, execution will stop on the first failed Assert.

## Benchmarking

In most cases, this Dictionary is the fastest solution when compared to what is already available. Please see [Benchmarking](benchmarking/README.md) for more details.

## Implementation

For those interested in how this Dictionary works and why some design decisions were made, please see [Implementation](Implementation.md) for more details.