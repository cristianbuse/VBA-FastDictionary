Attribute VB_Name = "LibStringTools"
'===============================================================================
' VBA StringTools
' ------------------------------------------
' https://github.com/guwidoe/VBA-StringTools
' ------------------------------------------
' MIT License
'
' Copyright (c) 2024 Guido Witt-Dörring
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to
' deal in the Software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
' sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.
'===============================================================================

Option Explicit
Option Base 0
Option Compare Binary

#If Mac Then
    #If VBA7 Then 'https://developer.apple.com/library/archive/documentation/System/Conceptual/ManPages_iPhoneOS/man3/iconv.3.html
        Private Declare PtrSafe Function iconv_open Lib "/usr/lib/libiconv.dylib" (ByVal toCode As LongPtr, ByVal fromCode As LongPtr) As LongPtr
        Private Declare PtrSafe Function iconv_close Lib "/usr/lib/libiconv.dylib" (ByVal cd As LongPtr) As Long
        Private Declare PtrSafe Function iconv Lib "/usr/lib/libiconv.dylib" (ByVal cd As LongPtr, ByRef inBuf As LongPtr, ByRef inBytesLeft As LongPtr, ByRef outBuf As LongPtr, ByRef outBytesLeft As LongPtr) As LongPtr

        Private Declare PtrSafe Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As LongPtr) As LongPtr
        Private Declare PtrSafe Function errno_location Lib "/usr/lib/libSystem.B.dylib" Alias "__error" () As LongPtr
        
        'https://developer.apple.com/documentation/corefoundation/1541720-cfstringgetsystemencoding
        Private Declare PtrSafe Function CFStringGetSystemEncoding Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" () As Long
        Private Declare PtrSafe Function CFStringConvertEncodingToWindowsCodepage Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" (ByVal CFStringEncoding As Long) As Long
    #Else
        Private Declare Function iconv Lib "/usr/lib/libiconv.dylib" (ByVal cd As Long, ByRef inBuf As Long, ByRef inBytesLeft As Long, ByRef outBuf As Long, ByRef outBytesLeft As Long) As Long
        Private Declare Function iconv_open Lib "/usr/lib/libiconv.dylib" (ByVal toCode As Long, ByVal fromCode As Long) As Long
        Private Declare Function iconv_close Lib "/usr/lib/libiconv.dylib" (ByVal cd As Long) As Long

        Private Declare Function CopyMemory Lib "/usr/lib/libc.dylib" Alias "memmove" (Destination As Any, Source As Any, ByVal Length As Long) As Long
        Private Declare Function errno_location Lib "/usr/lib/libSystem.B.dylib" Alias "__error" () As Long
        
        Private Declare Function CFStringGetSystemEncoding Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" () As Long
        Private Declare Function CFStringConvertEncodingToWindowsCodepage Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" (ByVal CFStringEncoding As Long) As Long
    #End If
#Else 'Windows
    #If VBA7 Then
        Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
        Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long

        Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long
        Private Declare PtrSafe Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
        Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
        
        Private Declare PtrSafe Function GetACP Lib "kernel32" () As Long
        
        Private Declare PtrSafe Function GetCPInfoExW Lib "kernel32" (ByVal codePage As Long, ByVal dwFlags As Long, lpCPInfoExW As CPINFOEXW) As Long
    #Else
        Private Declare Function MultiByteToWideChar Lib "kernel32" Alias "MultiByteToWideChar" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
        Private Declare Function WideCharToMultiByte Lib "kernel32" Alias "WideCharToMultiByte" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

        Private Declare Function GetLastError Lib "kernel32" () As Long
        Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
        Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
        
        Private Declare Function GetACP Lib "kernel32" () As Long
        
        Private Declare Function GetCPInfoExW Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpCPInfoExW As CPINFOEXW) As Long
    #End If
#End If

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If

'Flag used to simulate ByRef Variants
Private Const VT_BYREF As Long = &H4000

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    #If Win64 Then
        dummyPadding As Long
        pvData As LongLong
    #Else
        pvData As Long
    #End If
    rgsabound0 As SAFEARRAYBOUND
End Type
Private Const FADF_HAVEVARTYPE As Long = &H80

Private Const BYTE_SIZE As Long = 1
Private Const INT_SIZE As Long = 2

Private Const MAX_DEFAULTCHAR = 2
Private Const MAX_LEADBYTES = 12  '  5 ranges, 2 bytes ea., 0 term.
Private Const MAX_PATH = 260

Private Type CPINFOEXW
    MaxCharSize As Long                    'max length (in bytes) of a character
    defaultChar(MAX_DEFAULTCHAR - 1) As Byte ' default character (MB)
    LeadByte(MAX_LEADBYTES - 1) As Byte      ' lead byte ranges
    UnicodeDefaultChar(0 To 1) As Byte       ' default character (Unicode)
    codePage As Long                         ' code page id
    CodePageName(MAX_PATH - 1) As Byte       ' code page name (Unicode)
End Type

Private Type CpInfo 'Custom extended CpInfo type for use in this library
    'From CpInfoExW:
    codePage As Long              ' code page id
    MaxCharSize As Long           ' max length (in bytes) of a character
    defaultChar As String         ' default character (MB)
    LeadByte As String            ' lead byte ranges
    UnicodeDefaultChar As String  ' default character (Unicode)
    CodePageName As String        ' code page name (Unicode)
    'Extra:
    AllowsFlags As Boolean
    AllowsQueryReversible As Boolean
    MacConvDescriptorName As String
    IsInitialized As Boolean
End Type

Private Type EscapeSequence
    ueFormat As UnicodeEscapeFormat
    ueSignature As String
    letSngSurrogate As Boolean
    buffPosition As Long
    currPosition As Long
    sigSize As Long
    escSize As Long
    codepoint As Long
    unEscSize As Long
End Type

Private Type TwoCharTemplate
    s As String * 2
End Type
Private Type LongTemplate
    l As Long
End Type

#If Win64 Then
    #If Mac Then
        Private Const vbLongLong As Long = 20 'Apparently missing for x64 on Mac
    #End If
    Private Const vbLongPtr As Long = vbLongLong
#Else
    Private Const vbLongLong As Long = 20 'Useful in Select Case logic
    Private Const vbLongPtr As Long = vbLong
#End If

Private Type StringificationSettings
    maxChars As Long
    escapeNonPrintable As Boolean
    Delimiter As String
    maxCharsPerElement As Long
    maxCharsPerLine As Long
    maxLines As Long
    inklColIndices As Boolean
    inklRowIndices As Boolean
End Type

Dim printfSettings As StringificationSettings
Dim printfSettingsAreInitialized As Boolean

'API error codes
Private Const WC_ERR_INVALID_CHARS As Long = &H80&
Private Const MB_ERR_INVALID_CHARS As Long = &H8&

Private Const ERROR_INVALID_PARAMETER      As Long = 87
Private Const ERROR_INSUFFICIENT_BUFFER    As Long = 122
Private Const ERROR_INVALID_FLAGS          As Long = 1004
Private Const ERROR_NO_UNICODE_TRANSLATION As Long = 1113

Private Const MAC_API_ERR_EILSEQ As Long = 92 'Illegal byte sequence
Private Const MAC_API_ERR_EINVAL As Long = 22 'Invalid argument
Private Const MAC_API_ERR_E2BIG  As Long = 7  'Argument list too long

Private Const vbErrInternalError As Long = 51

'Custom error codes:
Public Const STERROR_CPINFO_NOT_SET As Long = vbObjectError + 52

Public Enum UnicodeEscapeFormat
    [_efNone] = 0
    efPython = 1 '\uXXXX \u00XXXXXX (4 or 8 hex digits, 8 for chars outside BMP)
    efRust = 2   '\u{X} \U{XXXXXX}  (1 to 6 hex digits)
    efUPlus = 4  'u+XXXX u+XXXXXX   (4 or 6 hex digits)
    efMarkup = 8 '&#ddddddd;        (1 to 7 decimal digits)
    efAll = 15
    [_efMin] = efPython
    [_efMax] = efAll
End Enum

'https://learn.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
Public Enum CodePageIdentifier
    [_first] = -1 '(Is initialized)
  'Enum_Name   Identifier             '.NET Name               Additional information
    cpIBM037 = 37                     'IBM037                  IBM EBCDIC US-Canada
    cpIBM437 = 437                    'IBM437                  OEM United States
    cpIBM500 = 500                    'IBM500                  IBM EBCDIC International
    cpASMO_708 = 708                  'ASMO-708                Arabic (ASMO 708)
    cpASMO_449 = 709                  '                        Arabic (ASMO-449+, BCON V4)
    cpTransparent_Arabic = 710        '                        Arabic - Transparent Arabic
    cpDOS_720 = 720                   'DOS-720                 Arabic (Transparent ASMO); Arabic (DOS)
    cpIbm737 = 737                    'ibm737                  OEM Greek (formerly 437G); Greek (DOS)
    cpIbm775 = 775                    'ibm775                  OEM Baltic; Baltic (DOS)
    cpIbm850 = 850                    'ibm850                  OEM Multilingual Latin 1; Western European (DOS)
    cpIbm852 = 852                    'ibm852                  OEM Latin 2; Central European (DOS)
    cpIBM855 = 855                    'IBM855                  OEM Cyrillic (primarily Russian)
    cpIbm857 = 857                    'ibm857                  OEM Turkish; Turkish (DOS)
    cpIBM00858 = 858                  'IBM00858                OEM Multilingual Latin 1 + Euro symbol
    cpIBM860 = 860                    'IBM860                  OEM Portuguese; Portuguese (DOS)
    cpIbm861 = 861                    'ibm861                  OEM Icelandic; Icelandic (DOS)
    cpDOS_862 = 862                   'DOS-862                 OEM Hebrew; Hebrew (DOS)
    cpIBM863 = 863                    'IBM863                  OEM French Canadian; French Canadian (DOS)
    cpIBM864 = 864                    'IBM864                  OEM Arabic; Arabic (864)
    cpIBM865 = 865                    'IBM865                  OEM Nordic; Nordic (DOS)
    cpCp866 = 866                     'cp866                   OEM Russian; Cyrillic (DOS)
    cpIbm869 = 869                    'ibm869                  OEM Modern Greek; Greek, Modern (DOS)
    cpIBM870 = 870                    'IBM870                  IBM EBCDIC Multilingual/ROECE (Latin 2); IBM EBCDIC Multilingual Latin 2
    cpWindows_874 = 874               'windows-874             Thai (Windows)
    cpCp875 = 875                     'cp875                   IBM EBCDIC Greek Modern
    cpShift_jis = 932                 'shift_jis               ANSI/OEM Japanese; Japanese (Shift-JIS)
    cpGb2312 = 936                    'gb2312                  ANSI/OEM Simplified Chinese (PRC, Singapore); Chinese Simplified (GB2312)
    cpKs_c_5601_1987 = 949            'ks_c_5601-1987          ANSI/OEM Korean (Unified Hangul Code)
    cpBig5 = 950                      'big5                    ANSI/OEM Traditional Chinese (Taiwan; Hong Kong SAR, PRC); Chinese Traditional (Big5)
    cpIBM1026 = 1026                  'IBM1026                 IBM EBCDIC Turkish (Latin 5)
    cpIBM01047 = 1047                 'IBM01047                IBM EBCDIC Latin 1/Open System
    cpIBM01140 = 1140                 'IBM01140                IBM EBCDIC US-Canada (037 + Euro symbol); IBM EBCDIC (US-Canada-Euro)
    cpIBM01141 = 1141                 'IBM01141                IBM EBCDIC Germany (20273 + Euro symbol); IBM EBCDIC (Germany-Euro)
    cpIBM01142 = 1142                 'IBM01142                IBM EBCDIC Denmark-Norway (20277 + Euro symbol); IBM EBCDIC (Denmark-Norway-Euro)
    cpIBM01143 = 1143                 'IBM01143                IBM EBCDIC Finland-Sweden (20278 + Euro symbol); IBM EBCDIC (Finland-Sweden-Euro)
    cpIBM01144 = 1144                 'IBM01144                IBM EBCDIC Italy (20280 + Euro symbol); IBM EBCDIC (Italy-Euro)
    cpIBM01145 = 1145                 'IBM01145                IBM EBCDIC Latin America-Spain (20284 + Euro symbol); IBM EBCDIC (Spain-Euro)
    cpIBM01146 = 1146                 'IBM01146                IBM EBCDIC United Kingdom (20285 + Euro symbol); IBM EBCDIC (UK-Euro)
    cpIBM01147 = 1147                 'IBM01147                IBM EBCDIC France (20297 + Euro symbol); IBM EBCDIC (France-Euro)
    cpIBM01148 = 1148                 'IBM01148                IBM EBCDIC International (500 + Euro symbol); IBM EBCDIC (International-Euro)
    cpIBM01149 = 1149                 'IBM01149                IBM EBCDIC Icelandic (20871 + Euro symbol); IBM EBCDIC (Icelandic-Euro)
    cpUTF_16 = 1200                   'utf-16                  Unicode UTF-16, little endian byte order (BMP of ISO 10646); available only to managed applications
    cpUnicodeFFFE = 1201              'unicodeFFFE             Unicode UTF-16, big endian byte order; available only to managed applications
    cpWindows_1250 = 1250             'windows-1250            ANSI Central European; Central European (Windows)
    cpWindows_1251 = 1251             'windows-1251            ANSI Cyrillic; Cyrillic (Windows)
    cpWindows_1252 = 1252             'windows-1252            ANSI Latin 1; Western European (Windows)
    cpWindows_1253 = 1253             'windows-1253            ANSI Greek; Greek (Windows)
    cpWindows_1254 = 1254             'windows-1254            ANSI Turkish; Turkish (Windows)
    cpWindows_1255 = 1255             'windows-1255            ANSI Hebrew; Hebrew (Windows)
    cpWindows_1256 = 1256             'windows-1256            ANSI Arabic; Arabic (Windows)
    cpWindows_1257 = 1257             'windows-1257            ANSI Baltic; Baltic (Windows)
    cpWindows_1258 = 1258             'windows-1258            ANSI/OEM Vietnamese; Vietnamese (Windows)
    cpJohab = 1361                    'Johab                   Korean (Johab)
    cpMacintosh = 10000               'macintosh               MAC Roman; Western European (Mac)
    cpX_mac_japanese = 10001          'x-mac-japanese          Japanese (Mac)
    cpX_mac_chinesetrad = 10002       'x-mac-chinesetrad       MAC Traditional Chinese (Big5); Chinese Traditional (Mac)
    cpX_mac_korean = 10003            'x-mac-korean            Korean (Mac)
    cpX_mac_arabic = 10004            'x-mac-arabic            Arabic (Mac)
    cpX_mac_hebrew = 10005            'x-mac-hebrew            Hebrew (Mac)
    cpX_mac_greek = 10006             'x-mac-greek             Greek (Mac)
    cpX_mac_cyrillic = 10007          'x-mac-cyrillic          Cyrillic (Mac)
    cpX_mac_chinesesimp = 10008       'x-mac-chinesesimp       MAC Simplified Chinese (GB 2312); Chinese Simplified (Mac)
    cpX_mac_romanian = 10010          'x-mac-romanian          Romanian (Mac)
    cpX_mac_ukrainian = 10017         'x-mac-ukrainian         Ukrainian (Mac)
    cpX_mac_thai = 10021              'x-mac-thai              Thai (Mac)
    cpX_mac_ce = 10029                'x-mac-ce                MAC Latin 2; Central European (Mac)
    cpX_mac_icelandic = 10079         'x-mac-icelandic         Icelandic (Mac)
    cpX_mac_turkish = 10081           'x-mac-turkish           Turkish (Mac)
    cpX_mac_croatian = 10082          'x-mac-croatian          Croatian (Mac)
    cpUTF_32 = 12000                  'utf-32                  Unicode UTF-32, little endian byte order; available only to managed applications
    cpUTF_32BE = 12001                'utf-32BE                Unicode UTF-32, big endian byte order; available only to managed applications
    cpX_Chinese_CNS = 20000           'x-Chinese_CNS           CNS Taiwan; Chinese Traditional (CNS)
    cpX_cp20001 = 20001               'x-cp20001               TCA Taiwan
    cpX_Chinese_Eten = 20002          'x_Chinese-Eten          Eten Taiwan; Chinese Traditional (Eten)
    cpX_cp20003 = 20003               'x-cp20003               IBM5550 Taiwan
    cpX_cp20004 = 20004               'x-cp20004               TeleText Taiwan
    cpX_cp20005 = 20005               'x-cp20005               Wang Taiwan
    cpX_IA5 = 20105                   'x-IA5                   IA5 (IRV International Alphabet No. 5, 7-bit); Western European (IA5)
    cpX_IA5_German = 20106            'x-IA5-German            IA5 German (7-bit)
    cpX_IA5_Swedish = 20107           'x-IA5-Swedish           IA5 Swedish (7-bit)
    cpX_IA5_Norwegian = 20108         'x-IA5-Norwegian         IA5 Norwegian (7-bit)
    cpUs_ascii = 20127                'us-ascii                US-ASCII (7-bit)
    cpX_cp20261 = 20261               'x-cp20261               T.61
    cpX_cp20269 = 20269               'x-cp20269               ISO 6937 Non-Spacing Accent
    cpIBM273 = 20273                  'IBM273                  IBM EBCDIC Germany
    cpIBM277 = 20277                  'IBM277                  IBM EBCDIC Denmark-Norway
    cpIBM278 = 20278                  'IBM278                  IBM EBCDIC Finland-Sweden
    cpIBM280 = 20280                  'IBM280                  IBM EBCDIC Italy
    cpIBM284 = 20284                  'IBM284                  IBM EBCDIC Latin America-Spain
    cpIBM285 = 20285                  'IBM285                  IBM EBCDIC United Kingdom
    cpIBM290 = 20290                  'IBM290                  IBM EBCDIC Japanese Katakana Extended
    cpIBM297 = 20297                  'IBM297                  IBM EBCDIC France
    cpIBM420 = 20420                  'IBM420                  IBM EBCDIC Arabic
    cpIBM423 = 20423                  'IBM423                  IBM EBCDIC Greek
    cpIBM424 = 20424                  'IBM424                  IBM EBCDIC Hebrew
    cpX_EBCDIC_KoreanExtended = 20833 'x-EBCDIC-KoreanExtended IBM EBCDIC Korean Extended
    cpIBM_Thai = 20838                'IBM-Thai                IBM EBCDIC Thai
    cpKoi8_r = 20866                  'koi8-r                  Russian (KOI8-R); Cyrillic (KOI8-R)
    cpIBM871 = 20871                  'IBM871                  IBM EBCDIC Icelandic
    cpIBM880 = 20880                  'IBM880                  IBM EBCDIC Cyrillic Russian
    cpIBM905 = 20905                  'IBM905                  IBM EBCDIC Turkish
    cpIBM00924 = 20924                'IBM00924                IBM EBCDIC Latin 1/Open System (1047 + Euro symbol)
    cpEuc_jp = 20932                  'EUC-JP                  Japanese (JIS 0208-1990 and 0212-1990)
    cpX_cp20936 = 20936               'x-cp20936               Simplified Chinese (GB2312); Chinese Simplified (GB2312-80)
    cpX_cp20949 = 20949               'x-cp20949               Korean Wansung
    cpCp1025 = 21025                  'cp1025                  IBM EBCDIC Cyrillic Serbian-Bulgarian
    cpDeprecated = 21027                       '                        (deprecated)
    cpKoi8_u = 21866                  'koi8-u                  Ukrainian (KOI8-U); Cyrillic (KOI8-U)
    cpIso_8859_1 = 28591              'iso-8859-1              ISO 8859-1 Latin 1; Western European (ISO)
    cpIso_8859_2 = 28592              'iso-8859-2              ISO 8859-2 Central European; Central European (ISO)
    cpIso_8859_3 = 28593              'iso-8859-3              ISO 8859-3 Latin 3
    cpIso_8859_4 = 28594              'iso-8859-4              ISO 8859-4 Baltic
    cpIso_8859_5 = 28595              'iso-8859-5              ISO 8859-5 Cyrillic
    cpIso_8859_6 = 28596              'iso-8859-6              ISO 8859-6 Arabic
    cpIso_8859_7 = 28597              'iso-8859-7              ISO 8859-7 Greek
    cpIso_8859_8 = 28598              'iso-8859-8              ISO 8859-8 Hebrew; Hebrew (ISO-Visual)
    cpIso_8859_9 = 28599              'iso-8859-9              ISO 8859-9 Turkish
    cpIso_8859_13 = 28603             'iso-8859-13             ISO 8859-13 Estonian
    cpIso_8859_15 = 28605             'iso-8859-15             ISO 8859-15 Latin 9
    cpX_Europa = 29001                'x-Europa                Europa 3
    cpIso_8859_8_i = 38598            'iso-8859-8-i            ISO 8859-8 Hebrew; Hebrew (ISO-Logical)
    cpIso_2022_jp = 50220             'iso-2022-jp             ISO 2022 Japanese with no halfwidth Katakana; Japanese (JIS)
    cpCsISO2022JP = 50221             'csISO2022JP             ISO 2022 Japanese with halfwidth Katakana; Japanese (JIS-Allow 1 byte Kana)
    cpIso_2022_jp_w_1b_Kana = 50222   'iso-2022-jp             ISO 2022 Japanese JIS X 0201-1989; Japanese (JIS-Allow 1 byte Kana - SO/SI)
    cpIso_2022_kr = 50225             'iso-2022-kr             ISO 2022 Korean
    cpX_cp50227 = 50227               'x-cp50227               ISO 2022 Simplified Chinese; Chinese Simplified (ISO 2022)
    cpISO_2022_Trad_Chinese = 50229   '                        ISO 2022 Traditional Chinese
    cpEBCDIC_Jap_Katakana_Ext = 50930 '                        EBCDIC Japanese (Katakana) Extended
    cpEBCDIC_US_Can_and_Jap = 50931   '                        EBCDIC US-Canada and Japanese
    cpEBCDIC_Kor_Ext_and_Kor = 50933  '                        EBCDIC Korean Extended and Korean
    cpEBCDIC_Simp_Chin_Ext = 50935    '                        EBCDIC Simplified Chinese Extended and Simplified Chinese
    cpEBCDIC_Simp_Chin = 50936        '                        EBCDIC Simplified Chinese
    cpEBCDIC_US_Can_Trad_Chin = 50937 '                        EBCDIC US-Canada and Traditional Chinese
    cpEBCDIC_Jap_Latin_Ext = 50939    '                        EBCDIC Japanese (Latin) Extended and Japanese
    euc_jp = 51932                    'euc-jp                  EUC Japanese
    cpEUC_CN = 51936                  'EUC-CN                  EUC Simplified Chinese; Chinese Simplified (EUC)
    cpEuc_kr = 51949                  'euc-kr                  EUC Korean
    cpEUC_Traditional_Chinese = 51950 '                        EUC Traditional Chinese
    cpHz_gb_2312 = 52936              'hz-gb-2312              HZ-GB2312 Simplified Chinese; Chinese Simplified (HZ)
    cpGB18030 = 54936                 'GB18030                 Windows XP and later: GB18030 Simplified Chinese (4 byte); Chinese Simplified (GB18030)
    cpSMS_GSM_7bit = 55000            '                        SMS GSM 7bit
    cpSMS_GSM_7bit_Spanish = 55001    '                        SMS GSM 7bit Spanish
    cpSMS_GSM_7bit_Portuguese = 55002 '                        SMS GSM 7bit Portuguese
    cpSMS_GSM_7bit_Turkish = 55003    '                        SMS GSM 7bit Turkish
    cpSMS_GSM_7bit_Greek = 55004      '                        SMS GSM 7bit Greek
    cpX_iscii_de = 57002              'x-iscii-de              ISCII Devanagari
    cpX_iscii_be = 57003              'x-iscii-be              ISCII Bangla
    cpX_iscii_ta = 57004              'x-iscii-ta              ISCII Tamil
    cpX_iscii_te = 57005              'x-iscii-te              ISCII Telugu
    cpX_iscii_as = 57006              'x-iscii-as              ISCII Assamese
    cpX_iscii_or = 57007              'x-iscii-or              ISCII Odia
    cpX_iscii_ka = 57008              'x-iscii-ka              ISCII Kannada
    cpX_iscii_ma = 57009              'x-iscii-ma              ISCII Malayalam
    cpX_iscii_gu = 57010              'x-iscii-gu              ISCII Gujarati
    cpX_iscii_pa = 57011              'x-iscii-pa              ISCII Punjabi
    cpUTF_7 = 65000                   'utf-7                   Unicode (UTF-7)
    cpUTF_8 = 65001                   'utf-8                   Unicode (UTF-8)
    [_last]
End Enum

'According to documentation:
'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-widechartomultibyte
'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-multibytetowidechar
'Note: The documentation doesn't seem to list all codepages for which certain
'      flags are disallowed. This can lead to 'Library implementation erroneous'
'      errors when calling Encode, Decode or Transcode with 'raiseErrors = True'
Private Static Function CodePageAllowsFlags(ByVal cpId As Long) As Boolean
    Dim arr(CodePageIdentifier.[_first] To CodePageIdentifier.[_last]) As Boolean

    If cpId < CodePageIdentifier.[_first] _
    Or cpId > CodePageIdentifier.[_last] Then
        CodePageAllowsFlags = False
        Exit Function
    End If
    
    If arr(CodePageIdentifier.[_first]) Then
        CodePageAllowsFlags = arr(cpId)
        Exit Function
    End If

    Dim i As Long
    For i = CodePageIdentifier.[_first] To CodePageIdentifier.[_last]
        arr(i) = True
    Next i

    'According to docs:
    arr(cpIso_2022_jp) = False
    arr(cpCsISO2022JP) = False
    arr(cpIso_2022_jp_w_1b_Kana) = False
    arr(cpIso_2022_kr) = False
    arr(cpX_cp50227) = False
    arr(cpISO_2022_Trad_Chinese) = False
    For i = cpX_iscii_de To cpX_iscii_pa
        arr(i) = False
    Next i
    arr(cpUTF_7) = False

    'According to trial and error, it is easier to whitelist:
    For i = CodePageIdentifier.[_first] + 1 To CodePageIdentifier.[_last]
        arr(i) = False
    Next i
    arr(cpUTF_32) = True   'Not sure about this one
    arr(cpUTF_32BE) = True 'Not sure about this one
    arr(cpGB18030) = True  'This one is definitely allowed
    arr(cpUTF_8) = True    'This one is definitely allowed

    CodePageAllowsFlags = arr(cpId)
End Function

'According to documentation:
'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-widechartomultibyte
'https://learn.microsoft.com/en-us/windows/win32/api/stringapiset/nf-stringapiset-multibytetowidechar
Private Static Function CodePageAllowsQueryReversible(ByVal cpId As Long) As Boolean
    Dim arr(CodePageIdentifier.[_first] To CodePageIdentifier.[_last]) As Boolean

    If cpId < CodePageIdentifier.[_first] _
    Or cpId > CodePageIdentifier.[_last] Then
        CodePageAllowsQueryReversible = False
        Exit Function
    End If
    
    If arr(CodePageIdentifier.[_first]) Then
        CodePageAllowsQueryReversible = arr(cpId)
        Exit Function
    End If

    Dim i As Long
    For i = CodePageIdentifier.[_first] To CodePageIdentifier.[_last]
        arr(i) = True
    Next i

    'According to docs:
    arr(cpUTF_7) = False
    arr(cpUTF_8) = False

    'According to trial and error there are a bunch more:
    arr(cpIso_2022_jp) = False
    arr(cpCsISO2022JP) = False
    arr(cpIso_2022_jp_w_1b_Kana) = False
    arr(cpIso_2022_kr) = False
    arr(cpX_cp50227) = False
    arr(cpISO_2022_Trad_Chinese) = False
    arr(cpHz_gb_2312) = False
    arr(cpGB18030) = False
    arr(cpX_iscii_de) = False
    arr(cpX_iscii_be) = False
    arr(cpX_iscii_ta) = False
    arr(cpX_iscii_te) = False
    arr(cpX_iscii_as) = False
    arr(cpX_iscii_or) = False
    arr(cpX_iscii_ka) = False
    arr(cpX_iscii_ma) = False
    arr(cpX_iscii_gu) = False
    arr(cpX_iscii_pa) = False

    CodePageAllowsQueryReversible = arr(cpId)
End Function

'Returns an array for converting CodePageIDs to ConversionDescriptorNames
Private Static Function ConvDescriptorName(ByVal cpId As Long) As String
    Dim arr(CodePageIdentifier.[_first] To CodePageIdentifier.[_last]) As String
    
    If cpId < CodePageIdentifier.[_first] _
    Or cpId > CodePageIdentifier.[_last] Then
        ConvDescriptorName = StrConv("-", vbFromUnicode)
        Exit Function
    End If
    
    If arr(CodePageIdentifier.[_first]) = "-" Then
        ConvDescriptorName = StrConv(arr(cpId), vbFromUnicode)
        Exit Function
    End If

    Dim i As Long
    For i = CodePageIdentifier.[_first] To CodePageIdentifier.[_last]
        arr(i) = "-"
    Next i

    'Source:
    'https://developer.apple.com/library/archive/documentation/System/Conceptual/ManPages_iPhoneOS/man3/iconv_open.3.html#//apple_ref/doc/man/3/iconv_open
    'European languages
    arr(cpIso_8859_1) = "ISO-8859-1"
    arr(cpIso_8859_2) = "ISO-8859-2"
    arr(cpIso_8859_3) = "ISO-8859-3"
    arr(cpIso_8859_4) = "ISO-8859-4"
    arr(cpIso_8859_5) = "ISO-8859-5"
    arr(cpIso_8859_7) = "ISO-8859-7"
    arr(cpIso_8859_9) = "ISO-8859-9"
    arr(28600) = "ISO-8859-10"
    arr(cpIso_8859_13) = "ISO-8859-13"
    arr(28604) = "ISO-8859-14"
    arr(cpIso_8859_15) = "ISO-8859-15"
    arr(28606) = "ISO-8859-16"
    arr(20866) = "KOI8-R"
    arr(cpKoi8_u) = "KOI8-U"
    'arr( ) =  "KOI8-RU" 'No equivalent ID, variation of KOI8-R
    arr(cpWindows_1250) = "CP1250"
    arr(cpWindows_1251) = "CP1251"
    arr(cpWindows_1252) = "CP1252"
    arr(cpWindows_1253) = "CP1253"
    arr(cpWindows_1254) = "CP1254"
    arr(cpWindows_1257) = "CP1257"
    arr(cpIbm850) = "CP850"
    arr(cpCp866) = "CP866"
    arr(cpMacintosh) = "MacRoman"   'duplicate
    arr(cpX_mac_ce) = "MacCentralEurope"
    arr(cpX_mac_icelandic) = "MacIceland"
    arr(cpX_mac_croatian) = "MacCroatian"
    arr(cpX_mac_romanian) = "MacRomania"
    arr(cpX_mac_cyrillic) = "MacCyrillic"
    arr(cpX_mac_ukrainian) = "MacUkraine"
    arr(cpX_mac_greek) = "MacGreek"
    arr(cpX_mac_turkish) = "MacTurkish"
    arr(cpMacintosh) = "Macintosh"

    'Semitic languages
    arr(cpIso_8859_6) = "ISO-8859-6"
    arr(cpIso_8859_8) = "ISO-8859-8"
    arr(cpWindows_1255) = "CP1255"
    arr(cpWindows_1256) = "CP1256"
    arr(cpDOS_862) = "CP862"
    arr(cpX_mac_hebrew) = "MacHebrew"
    arr(cpX_mac_arabic) = "MacArabic"

    'Japanese
    arr(euc_jp) = "EUC-JP"
    arr(cpShift_jis) = "SHIFT_JIS"
    arr(cpShift_jis) = "CP932" '(duplicate)
    arr(cpIso_2022_jp) = "ISO-2022-JP"
    arr(cpCsISO2022JP) = "ISO-2022-JP-2"
    arr(cpIso_2022_jp_w_1b_Kana) = "ISO-2022-JP-1"

    'Chinese
    arr(cpEUC_CN) = "EUC-CN"
    'arr( ) =  "HZ" 'No equivalent ID, 7-bit encoding method for GB2312
    arr(cpGb2312) = "GBK" 'duplicate
    arr(cpGb2312) = "CP936"
    arr(cpGB18030) = "GB18030"
    'arr( ) =  "EUC-TW" 'No equivalent ID, extended UNIX Code for Traditional Chinese
    arr(cpBig5) = "BIG5"
    arr(cpBig5) = "CP950" '(duplicate)
    arr(951) = "BIG5-HKSCS"
    arr(951) = "BIG5-HKSCS:2001"
    arr(951) = "BIG5-HKSCS:1999"
    arr(cpX_cp50227) = "ISO-2022-CN"
    'arr( ) =  "ISO-2022-CN-EXT" 'No equivalent ID, extended version of ISO-2022-CN

    'Korean
    arr(cpEuc_kr) = "EUC-KR"
    arr(cpKs_c_5601_1987) = "CP949"
    arr(cpIso_2022_kr) = "ISO-2022-KR"
    arr(cpJohab) = "JOHAB"

    'Armenian
    'arr( ) =  "ARMSCII-8" '8-bit Armenian character encoding

    'Georgian
    'arr( ) =  "Georgian-Academy" 'No equivalent ID
    'arr( ) =  "Georgian-PS" 'No equivalent ID

    'Tajik
    'arr( ) =  "KOI8-T" 'No equivalent ID

    'Kazakh
    'arr( ) =  "PT154" 'No equivalent ID, Paratype KZ

    'Thai
    arr(cpWindows_874) = "TIS-620" 'duplicate
    arr(cpWindows_874) = "CP874"
    arr(cpX_mac_thai) = "MacThai"

    'Laotian
    'arr( ) =  "MuleLao-1" 'No equivalent ID, MULE (MULtilingual Enhancement to GNU Emacs) internal encoding for the Lao script
    arr(1133) = "CP1133"

    'Vietnamese
    'arr( ) =  "VISCII" 'No equivalent ID, 8-bit encoding for the Vietnamese alphabet
    'arr( ) =  "TCVN" 'No equivalent ID, Vietnamese national standard for character encoding
    arr(cpWindows_1258) = "CP1258"

    'Platform specifics
    'arr( ) =  "HP-ROMAN8" 'No equivalent ID, 8-bit character encoding used by Hewlett-Packard for their workstations and printers.
    'arr( ) =  "NEXTSTEP" 'No equivalent ID, encoding is associated with the NeXTSTEP operating system developed by NeXT, the company founded by Steve Jobs after leaving Apple in the 1980s.

    'Full Unicode
    'arr( ) =  "UCS-2"
    arr(cpUnicodeFFFE) = "UCS-2BE" '(duplicate)
    arr(cpUTF_16) = "UCS-2LE" '(duplicate)
    'arr( ) =  "UCS-4"
    arr(cpUTF_32BE) = "UCS-4BE" '(duplicate)
    arr(cpUTF_32) = "UCS-4LE" '(duplicate)
    'arr( ) =  "UTF-16"
    arr(cpUnicodeFFFE) = "UTF-16BE"
    arr(cpUTF_16) = "UTF-16LE"
    'arr( ) =  "UTF-32"
    arr(cpUTF_32BE) = "UTF-32BE"
    arr(cpUTF_32) = "UTF-32LE"
    arr(cpUTF_7) = "UTF-7"
    arr(cpUTF_8) = "UTF-8"
    'arr( ) =  "C99"
    'arr( ) =  "JAVA"

    'Full Unicode in terms of uint16_t or uint32_t
    '(with machine dependent endianness and alignment)
    'arr( ) =  "UCS-2-INTERNAL"
    'arr( ) =  "UCS-4-INTERNAL"

    'Locale dependent in terms of char or wchar_t
    '(with  machine  dependent  endianness  and  alignment and with
    'semantics depending on the OS and the  current  LC_CTYPE  locale facet)
    'arr( ) =  "char"
    'arr( ) =  "wchar_t"

    'When  configured with the option --enable-extra-encodings
    'it also pro-vides provides vides support for a few extra encodings:

    'European languages
    arr(cpIBM437) = "CP437"
    arr(cpIbm737) = "CP737"
    arr(cpIbm775) = "CP775"
    arr(cpIbm852) = "CP852"
    arr(853) = "CP853"
    arr(cpIBM855) = "CP855"
    arr(cpIbm857) = "CP857"
    arr(cpIBM00858) = "CP858"
    arr(cpIBM860) = "CP860"
    arr(cpIbm861) = "CP861"
    arr(cpIBM863) = "CP863"
    arr(cpIBM865) = "CP865"
    arr(cpIbm869) = "CP869"
    arr(1125) = "CP1125"

    'Semitic languages
    arr(cpIBM864) = "CP864"

    'Japanese
    'arr( ) =  "EUC-JISX0213" 'No equivalent ID
    'arr( ) =  "Shift_JISX0213" 'No equivalent ID
    'arr( ) =  "ISO-2022-JP-3" 'No equivalent ID

    'Chinese
    'arr( ) = "BIG5-2003" '(experimental 'No equivalent ID

    'Turkmen
    'arr( ) =  "TDS565" 'No equivalent ID

    'Platform specifics
    'arr( ) =  "ATARIST" 'No equivalent ID, 8-bit character encoding used on Atari ST computers, which were a series of personal computers released in the 1980s.
    'arr( ) =  "RISCOS-LATIN1" 'No equivalent ID, 8-bit character encoding used on the RISC OS operating system, which was developed by Acorn Computers in the late 1980s.

    'The empty encoding name is equivalent to "char":
    'it denotes the locale dependent character encoding.
    ConvDescriptorName = ConvDescriptorName(cpId)
End Function

Private Function GetCpInfo(ByVal cpId As CodePageIdentifier) As CpInfo
    Const methodName As String = "GetCpInfo"
    Static cpInfos(CodePageIdentifier.[_first] To _
                   CodePageIdentifier.[_last]) As CpInfo
    
    If cpId < CodePageIdentifier.[_first] _
    Or cpId > CodePageIdentifier.[_last] Then
        Exit Function
    End If
    
    #If Mac Then
        If Not cpInfos(cpId).IsInitialized Then GoSub InitializeCpInfosMac
    #Else 'Windows
        Dim cpi As CPINFOEXW
        With cpInfos(cpId)
            If Not .IsInitialized Then
                .codePage = cpId
                .AllowsFlags = CodePageAllowsFlags(cpId)
                .AllowsQueryReversible = CodePageAllowsQueryReversible(cpId)
                .MacConvDescriptorName = ConvDescriptorName(cpId)
                If GetCPInfoExW(cpId, 0, cpi) Then 'This is not really needed:
                    .MaxCharSize = cpi.MaxCharSize
                    .defaultChar = cpi.defaultChar
                    .LeadByte = TrimX(CStr(cpi.LeadByte), vbNullChar)
                    .UnicodeDefaultChar = cpi.UnicodeDefaultChar
                    .CodePageName = TrimX(CStr(cpi.CodePageName), vbNullChar)
                End If
                .IsInitialized = True
            End If
        End With
    #End If
    GetCpInfo = cpInfos(cpId)
    Exit Function
    
InitializeCpInfosMac:
    Const errMsg As String = "CpInfoData for codepage %CP% not available. " & _
        "You can fix this error by manually supplying the data in the Select" _
        & " Case stement in the function " & methodName & "."
        
    'The following code was generated on a Windows machine by running the
    ''GenerateGetCpInfoCode' procedure from the module 'LibStringToolsCodeGen'
    'with manual adjustments for 1200 (UTF-16LE), 1201 (UTF-16BE),
    '12000 (UTF-32LE), 12001 (UTF-32BE)
    With cpInfos(cpId)
        .codePage = cpId
        .AllowsFlags = CodePageAllowsFlags(cpId)
        .AllowsQueryReversible = CodePageAllowsQueryReversible(cpId)
        .MacConvDescriptorName = ConvDescriptorName(cpId)
        
        Select Case cpId
            'cpIBM037 (37    (IBM EBCDIC - U.S./Canada))
            Case cpIBM037
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "37    (IBM EBCDIC - U.S./Canada)"

            'cpIBM437 (437   (OEM - United States))
            Case cpIBM437
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "437   (OEM - United States)"

            'cpIBM500 (500   (IBM EBCDIC - International))
            Case cpIBM500
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "500   (IBM EBCDIC - International)"

            'cpASMO_708 (708   (Arabic - ASMO))
            Case cpASMO_708
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "708   (Arabic - ASMO)"

            'cpASMO_449 ()
            Case cpASMO_449
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", cpASMO_449)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpTransparent_Arabic ()
            Case cpTransparent_Arabic
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", cpTransparent_Arabic)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpDOS_720 (720   (Arabic - Transparent ASMO))
            Case cpDOS_720
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "720   (Arabic - Transparent ASMO)"

            'cpIbm737 (737   (OEM - Greek 437G))
            Case cpIbm737
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "737   (OEM - Greek 437G)"

            'cpIbm775 (775   (OEM - Baltic))
            Case cpIbm775
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "775   (OEM - Baltic)"

            'cpIbm850 (850   (OEM - Multilingual Latin I))
            Case cpIbm850
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "850   (OEM - Multilingual Latin I)"

            'cpIbm852 (852   (OEM - Latin II))
            Case cpIbm852
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "852   (OEM - Latin II)"

            '853 ()
            Case 853
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", 853)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpIBM855 (855   (OEM - Cyrillic))
            Case cpIBM855
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "855   (OEM - Cyrillic)"

            'cpIbm857 (857   (OEM - Turkish))
            Case cpIbm857
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "857   (OEM - Turkish)"

            'cpIBM00858 (858   (OEM - Multilingual Latin I + Euro))
            Case cpIBM00858
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "858   (OEM - Multilingual Latin I + Euro)"

            'cpIBM860 (860   (OEM - Portuguese))
            Case cpIBM860
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "860   (OEM - Portuguese)"

            'cpIbm861 (861   (OEM - Icelandic))
            Case cpIbm861
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "861   (OEM - Icelandic)"

            'cpDOS_862 (862   (OEM - Hebrew))
            Case cpDOS_862
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "862   (OEM - Hebrew)"

            'cpIBM863 (863   (OEM - Canadian French))
            Case cpIBM863
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "863   (OEM - Canadian French)"

            'cpIBM864 (864   (OEM - Arabic))
            Case cpIBM864
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "864   (OEM - Arabic)"

            'cpIBM865 (865   (OEM - Nordic))
            Case cpIBM865
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "865   (OEM - Nordic)"

            'cpCp866 (866   (OEM - Russian))
            Case cpCp866
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "866   (OEM - Russian)"

            'cpIbm869 (869   (OEM - Modern Greek))
            Case cpIbm869
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "869   (OEM - Modern Greek)"

            'cpIBM870 (870   (IBM EBCDIC - Multilingual/ROECE (Latin-2)))
            Case cpIBM870
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "870   (IBM EBCDIC - Multilingual/ROECE (Latin-2))"

            'cpWindows_874 (874   (ANSI/OEM - Thai))
            Case cpWindows_874
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "874   (ANSI/OEM - Thai)"

            'cpCp875 (875   (IBM EBCDIC - Modern Greek))
            Case cpCp875
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "875   (IBM EBCDIC - Modern Greek)"

            'cpShift_jis (932   (ANSI/OEM - Japanese Shift-JIS))
            Case cpShift_jis
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0x819FE0FC")
                .UnicodeDefaultChar = HexToString("0xFB30")
                .CodePageName = "932   (ANSI/OEM - Japanese Shift-JIS)"

            'cpGb2312 (936   (ANSI/OEM - Simplified Chinese GBK))
            Case cpGb2312
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0x81FE")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "936   (ANSI/OEM - Simplified Chinese GBK)"

            'cpKs_c_5601_1987 (949   (ANSI/OEM - Korean))
            Case cpKs_c_5601_1987
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0x81FE")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "949   (ANSI/OEM - Korean)"

            'cpBig5 (950   (ANSI/OEM - Traditional Chinese Big5))
            Case cpBig5
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0x81FE")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "950   (ANSI/OEM - Traditional Chinese Big5)"

            '951 ()
            Case 951
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", 951)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpIBM1026 (1026  (IBM EBCDIC - Turkish (Latin-5)))
            Case cpIBM1026
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1026  (IBM EBCDIC - Turkish (Latin-5))"

            'cpIBM01047 (1047  (IBM EBCDIC - Latin-1/Open System))
            Case cpIBM01047
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1047  (IBM EBCDIC - Latin-1/Open System)"

            '1125 ()
            Case 1125
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", 1125)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            '1133 ()
            Case 1133
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", 1133)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpIBM01140 (1140  (IBM EBCDIC - U.S./Canada (37 + Euro)))
            Case cpIBM01140
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1140  (IBM EBCDIC - U.S./Canada (37 + Euro))"

            'cpIBM01141 (1141  (IBM EBCDIC - Germany (20273 + Euro)))
            Case cpIBM01141
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1141  (IBM EBCDIC - Germany (20273 + Euro))"

            'cpIBM01142 (1142  (IBM EBCDIC - Denmark/Norway (20277 + Euro)))
            Case cpIBM01142
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1142  (IBM EBCDIC - Denmark/Norway (20277 + Euro))"

            'cpIBM01143 (1143  (IBM EBCDIC - Finland/Sweden (20278 + Euro)))
            Case cpIBM01143
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1143  (IBM EBCDIC - Finland/Sweden (20278 + Euro))"

            'cpIBM01144 (1144  (IBM EBCDIC - Italy (20280 + Euro)))
            Case cpIBM01144
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1144  (IBM EBCDIC - Italy (20280 + Euro))"

            'cpIBM01145 (1145  (IBM EBCDIC - Latin America/Spain (20284 + Euro)))
            Case cpIBM01145
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1145  (IBM EBCDIC - Latin America/Spain (20284 + Euro))"

            'cpIBM01146 (1146  (IBM EBCDIC - United Kingdom (20285 + Euro)))
            Case cpIBM01146
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1146  (IBM EBCDIC - United Kingdom (20285 + Euro))"

            'cpIBM01147 ()
            Case cpIBM01147
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = ""

            'cpIBM01148 (1148  (IBM EBCDIC - International (500 + Euro)))
            Case cpIBM01148
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1148  (IBM EBCDIC - International (500 + Euro))"

            'cpIBM01149 (1149  (IBM EBCDIC - Icelandic (20871 + Euro)))
            Case cpIBM01149
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1149  (IBM EBCDIC - Icelandic (20871 + Euro))"

            'cpUTF_16 ()
            Case cpUTF_16
'                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
'                          Replace(errMsg, "%CP%", cpUTF_16)
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = ""
                .UnicodeDefaultChar = HexToString("0xFDFF")
                .CodePageName = "1200 (UTF-16LE)"

            'cpUnicodeFFFE ()
            Case cpUnicodeFFFE
'                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
'                          Replace(errMsg, "%CP%", cpUnicodeFFFE)
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = ""
                .UnicodeDefaultChar = HexToString("0xFDFF")
                .CodePageName = "1201 (UTF-16BE)"

            'cpWindows_1250 (1250  (ANSI - Central Europe))
            Case cpWindows_1250
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1250  (ANSI - Central Europe)"

            'cpWindows_1251 (1251  (ANSI - Cyrillic))
            Case cpWindows_1251
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1251  (ANSI - Cyrillic)"

            'cpWindows_1252 (1252  (ANSI - Latin I))
            Case cpWindows_1252
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1252  (ANSI - Latin I)"

            'cpWindows_1253 (1253  (ANSI - Greek))
            Case cpWindows_1253
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1253  (ANSI - Greek)"

            'cpWindows_1254 (1254  (ANSI - Turkish))
            Case cpWindows_1254
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1254  (ANSI - Turkish)"

            'cpWindows_1255 (1255  (ANSI - Hebrew))
            Case cpWindows_1255
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1255  (ANSI - Hebrew)"

            'cpWindows_1256 (1256  (ANSI - Arabic))
            Case cpWindows_1256
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1256  (ANSI - Arabic)"

            'cpWindows_1257 (1257  (ANSI - Baltic))
            Case cpWindows_1257
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1257  (ANSI - Baltic)"

            'cpWindows_1258 (1258  (ANSI/OEM - Viet Nam))
            Case cpWindows_1258
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1258  (ANSI/OEM - Viet Nam)"

            'cpJohab (1361  (Korean - Johab))
            Case cpJohab
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0x84D3D8DEE0F9")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "1361  (Korean - Johab)"

            'cpMacintosh (10000 (MAC - Roman))
            Case cpMacintosh
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10000 (MAC - Roman)"

            'cpX_mac_japanese (10001 (MAC - Japanese))
            Case cpX_mac_japanese
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0x819FE0FC")
                .UnicodeDefaultChar = HexToString("0xFB30")
                .CodePageName = "10001 (MAC - Japanese)"

            'cpX_mac_chinesetrad (10002 (MAC - Traditional Chinese Big5))
            Case cpX_mac_chinesetrad
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0x81FC")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10002 (MAC - Traditional Chinese Big5)"

            'cpX_mac_korean (10003 (MAC - Korean))
            Case cpX_mac_korean
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0xA1ACB0C8CAFD")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10003 (MAC - Korean)"

            'cpX_mac_arabic (10004 (MAC - Arabic))
            Case cpX_mac_arabic
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10004 (MAC - Arabic)"

            'cpX_mac_hebrew (10005 (MAC - Hebrew))
            Case cpX_mac_hebrew
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10005 (MAC - Hebrew)"

            'cpX_mac_greek (10006 (MAC - Greek I))
            Case cpX_mac_greek
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10006 (MAC - Greek I)"

            'cpX_mac_cyrillic (10007 (MAC - Cyrillic))
            Case cpX_mac_cyrillic
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10007 (MAC - Cyrillic)"

            'cpX_mac_chinesesimp (10008 (MAC - Simplified Chinese GB 2312))
            Case cpX_mac_chinesesimp
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0xA1A9B0F7")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10008 (MAC - Simplified Chinese GB 2312)"

            'cpX_mac_romanian (10010 (MAC - Romania))
            Case cpX_mac_romanian
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10010 (MAC - Romania)"

            'cpX_mac_ukrainian (10017 (MAC - Ukraine))
            Case cpX_mac_ukrainian
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10017 (MAC - Ukraine)"

            'cpX_mac_thai (10021 (MAC - Thai))
            Case cpX_mac_thai
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10021 (MAC - Thai)"

            'cpX_mac_ce (10029 (MAC - Latin II))
            Case cpX_mac_ce
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10029 (MAC - Latin II)"

            'cpX_mac_icelandic (10079 (MAC - Icelandic))
            Case cpX_mac_icelandic
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10079 (MAC - Icelandic)"

            'cpX_mac_turkish (10081 (MAC - Turkish))
            Case cpX_mac_turkish
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10081 (MAC - Turkish)"

            'cpX_mac_croatian (10082 (MAC - Croatia))
            Case cpX_mac_croatian
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "10082 (MAC - Croatia)"

            'cpUTF_32 ()
            Case cpUTF_32
                'Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", cpUTF_32)
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = ""
                .UnicodeDefaultChar = HexToString("0xFDFF")
                .CodePageName = "12000 (UTF-32LE)"

            'cpUTF_32BE ()
            Case cpUTF_32BE
'                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
'                          Replace(errMsg, "%CP%", cpUTF_32BE)
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = ""
                .UnicodeDefaultChar = HexToString("0xFDFF")
                .CodePageName = "12001 (UTF-32BE)"

            'cpX_Chinese_CNS (20000 (CNS - Taiwan))
            Case cpX_Chinese_CNS
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0xA1FE")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20000 (CNS - Taiwan)"

            'cpX_cp20001 (20001 (TCA - Taiwan))
            Case cpX_cp20001
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0x818491D8DFFC")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20001 (TCA - Taiwan)"

            'cpX_Chinese_Eten (20002 (Eten - Taiwan))
            Case cpX_Chinese_Eten
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0x81AFDDFE")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20002 (Eten - Taiwan)"

            'cpX_cp20003 (20003 (IBM5550 - Taiwan))
            Case cpX_cp20003
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0x8184878789E8F9FB")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20003 (IBM5550 - Taiwan)"

            'cpX_cp20004 (20004 (TeleText - Taiwan))
            Case cpX_cp20004
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0xA1FE")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20004 (TeleText - Taiwan)"

            'cpX_cp20005 (20005 (Wang - Taiwan))
            Case cpX_cp20005
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0x8DF5F9FC")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20005 (Wang - Taiwan)"

            'cpX_IA5 (20105 (IA5 IRV International Alphabet No.5))
            Case cpX_IA5
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20105 (IA5 IRV International Alphabet No.5)"

            'cpX_IA5_German (20106 (IA5 German))
            Case cpX_IA5_German
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20106 (IA5 German)"

            'cpX_IA5_Swedish (20107 (IA5 Swedish))
            Case cpX_IA5_Swedish
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20107 (IA5 Swedish)"

            'cpX_IA5_Norwegian (20108 (IA5 Norwegian))
            Case cpX_IA5_Norwegian
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20108 (IA5 Norwegian)"

            'cpUs_ascii (20127 (US-ASCII))
            Case cpUs_ascii
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20127 (US-ASCII)"

            'cpX_cp20261 (20261 (T.61))
            Case cpX_cp20261
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0xC1CF")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20261 (T.61)"

            'cpX_cp20269 (20269 (ISO 6937 Non-Spacing Accent))
            Case cpX_cp20269
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20269 (ISO 6937 Non-Spacing Accent)"

            'cpIBM273 (20273 (IBM EBCDIC - Germany))
            Case cpIBM273
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20273 (IBM EBCDIC - Germany)"

            'cpIBM277 (20277 (IBM EBCDIC - Denmark/Norway))
            Case cpIBM277
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20277 (IBM EBCDIC - Denmark/Norway)"

            'cpIBM278 (20278 (IBM EBCDIC - Finland/Sweden))
            Case cpIBM278
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20278 (IBM EBCDIC - Finland/Sweden)"

            'cpIBM280 (20280 (IBM EBCDIC - Italy))
            Case cpIBM280
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20280 (IBM EBCDIC - Italy)"

            'cpIBM284 (20284 (IBM EBCDIC - Latin America/Spain))
            Case cpIBM284
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20284 (IBM EBCDIC - Latin America/Spain)"

            'cpIBM285 (20285 (IBM EBCDIC - United Kingdom))
            Case cpIBM285
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20285 (IBM EBCDIC - United Kingdom)"

            'cpIBM290 (20290 (IBM EBCDIC - Japanese Katakana Extended))
            Case cpIBM290
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20290 (IBM EBCDIC - Japanese Katakana Extended)"

            'cpIBM297 (20297 (IBM EBCDIC - France))
            Case cpIBM297
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20297 (IBM EBCDIC - France)"

            'cpIBM420 (20420 (IBM EBCDIC - Arabic))
            Case cpIBM420
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20420 (IBM EBCDIC - Arabic)"

            'cpIBM423 (20423 (IBM EBCDIC - Greek))
            Case cpIBM423
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20423 (IBM EBCDIC - Greek)"

            'cpIBM424 (20424 (IBM EBCDIC - Hebrew))
            Case cpIBM424
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20424 (IBM EBCDIC - Hebrew)"

            'cpX_EBCDIC_KoreanExtended (20833 (IBM EBCDIC - Korean Extended))
            Case cpX_EBCDIC_KoreanExtended
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20833 (IBM EBCDIC - Korean Extended)"

            'cpIBM_Thai (20838 (IBM EBCDIC - Thai))
            Case cpIBM_Thai
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20838 (IBM EBCDIC - Thai)"

            'cpKoi8_r (20866 (Russian - KOI8))
            Case cpKoi8_r
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20866 (Russian - KOI8)"

            'cpIBM871 (20871 (IBM EBCDIC - Icelandic))
            Case cpIBM871
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20871 (IBM EBCDIC - Icelandic)"

            'cpIBM880 (20880 (IBM EBCDIC - Cyrillic (Russian)))
            Case cpIBM880
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20880 (IBM EBCDIC - Cyrillic (Russian))"

            'cpIBM905 (20905 (IBM EBCDIC - Turkish))
            Case cpIBM905
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20905 (IBM EBCDIC - Turkish)"

            'cpIBM00924 (20924 (IBM EBCDIC - Latin-1/Open System (1047 + Euro)))
            Case cpIBM00924
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20924 (IBM EBCDIC - Latin-1/Open System (1047 + Euro))"

            'cpEuc_jp (20932 (JIS X 0208-1990 & 0212-1990))
            Case cpEuc_jp
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0x8E8EA1FE")
                .UnicodeDefaultChar = HexToString("0xFB30")
                .CodePageName = "20932 (JIS X 0208-1990 & 0212-1990)"

            'cpX_cp20936 (20936 (Simplified Chinese GB2312))
            Case cpX_cp20936
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0xA1A9B0F7")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "20936 (Simplified Chinese GB2312)"

            'cpX_cp20949 ()
            Case cpX_cp20949
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0xA1ACB0C8CAFD")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = ""

            'cpCp1025 (21025 (IBM EBCDIC - Cyrillic (Serbian, Bulgarian)))
            Case cpCp1025
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "21025 (IBM EBCDIC - Cyrillic (Serbian, Bulgarian))"

            'cpDeprecated (21027 (Ext Alpha Lowercase))
            Case cpDeprecated
                .MaxCharSize = 1
                .defaultChar = "o"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "21027 (Ext Alpha Lowercase)"

            'cpKoi8_u (21866 (Ukrainian - KOI8-U))
            Case cpKoi8_u
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "21866 (Ukrainian - KOI8-U)"

            'cpIso_8859_1 (28591 (ISO 8859-1 Latin I))
            Case cpIso_8859_1
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "28591 (ISO 8859-1 Latin I)"

            'cpIso_8859_2 (28592 (ISO 8859-2 Central Europe))
            Case cpIso_8859_2
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "28592 (ISO 8859-2 Central Europe)"

            'cpIso_8859_3 (28593 (ISO 8859-3 Latin 3))
            Case cpIso_8859_3
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "28593 (ISO 8859-3 Latin 3)"

            'cpIso_8859_4 (28594 (ISO 8859-4 Baltic))
            Case cpIso_8859_4
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "28594 (ISO 8859-4 Baltic)"

            'cpIso_8859_5 (28595 (ISO 8859-5 Cyrillic))
            Case cpIso_8859_5
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "28595 (ISO 8859-5 Cyrillic)"

            'cpIso_8859_6 (28596 (ISO 8859-6 Arabic))
            Case cpIso_8859_6
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "28596 (ISO 8859-6 Arabic)"

            'cpIso_8859_7 (28597 (ISO 8859-7 Greek))
            Case cpIso_8859_7
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "28597 (ISO 8859-7 Greek)"

            'cpIso_8859_8 (28598 (ISO 8859-8 Hebrew: Visual Ordering))
            Case cpIso_8859_8
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "28598 (ISO 8859-8 Hebrew: Visual Ordering)"

            'cpIso_8859_9 (28599 (ISO 8859-9 Latin 5))
            Case cpIso_8859_9
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "28599 (ISO 8859-9 Latin 5)"

            '28600 ()
            Case 28600
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", 28600)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpIso_8859_13 (28603 (ISO 8859-13 Latin 7))
            Case cpIso_8859_13
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "28603 (ISO 8859-13 Latin 7)"

            '28604 ()
            Case 28604
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", 28604)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpIso_8859_15 (28605 (ISO 8859-15 Latin 9))
            Case cpIso_8859_15
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "28605 (ISO 8859-15 Latin 9)"

            '28606 ()
            Case 28606
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", 28606)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpX_Europa ()
            Case cpX_Europa
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", cpX_Europa)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpIso_8859_8_i (38598 (ISO 8859-8 Hebrew: Logical Ordering))
            Case cpIso_8859_8_i
                .MaxCharSize = 1
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "38598 (ISO 8859-8 Hebrew: Logical Ordering)"

            'cpIso_2022_jp (50220 (ISO-2022 Japanese with no halfwidth Katakana))
            Case cpIso_2022_jp
                .MaxCharSize = 5
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "50220 (ISO-2022 Japanese with no halfwidth Katakana)"

            'cpCsISO2022JP (50221 (ISO-2022 Japanese with halfwidth Katakana))
            Case cpCsISO2022JP
                .MaxCharSize = 5
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "50221 (ISO-2022 Japanese with halfwidth Katakana)"

            'cpIso_2022_jp_w_1b_Kana (50222 (ISO-2022 Japanese JIS X 0201-1989))
            Case cpIso_2022_jp_w_1b_Kana
                .MaxCharSize = 5
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "50222 (ISO-2022 Japanese JIS X 0201-1989)"

            'cpIso_2022_kr (50225 (ISO-2022 Korean))
            Case cpIso_2022_kr
                .MaxCharSize = 5
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "50225 (ISO-2022 Korean)"

            'cpX_cp50227 (50227 (ISO-2022 Simplified Chinese))
            Case cpX_cp50227
                .MaxCharSize = 5
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "50227 (ISO-2022 Simplified Chinese)"

            'cpISO_2022_Trad_Chinese (50229 (ISO-2022 Traditional Chinese))
            Case cpISO_2022_Trad_Chinese
                .MaxCharSize = 5
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "50229 (ISO-2022 Traditional Chinese)"

            'cpEBCDIC_Jap_Katakana_Ext ()
            Case cpEBCDIC_Jap_Katakana_Ext
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", cpEBCDIC_Jap_Katakana_Ext)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpEBCDIC_US_Can_and_Jap ()
            Case cpEBCDIC_US_Can_and_Jap
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", cpEBCDIC_US_Can_and_Jap)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpEBCDIC_Kor_Ext_and_Kor ()
            Case cpEBCDIC_Kor_Ext_and_Kor
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", cpEBCDIC_Kor_Ext_and_Kor)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpEBCDIC_Simp_Chin_Ext ()
            Case cpEBCDIC_Simp_Chin_Ext
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", cpEBCDIC_Simp_Chin_Ext)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpEBCDIC_Simp_Chin ()
            Case cpEBCDIC_Simp_Chin
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", cpEBCDIC_Simp_Chin)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpEBCDIC_US_Can_Trad_Chin ()
            Case cpEBCDIC_US_Can_Trad_Chin
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", cpEBCDIC_US_Can_Trad_Chin)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpEBCDIC_Jap_Latin_Ext ()
            Case cpEBCDIC_Jap_Latin_Ext
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", cpEBCDIC_Jap_Latin_Ext)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'euc_jp ()
            Case euc_jp
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", euc_jp)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpEUC_CN ()
            Case cpEUC_CN
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", cpEUC_CN)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpEuc_kr (51949 (EUC-Korean))
            Case cpEuc_kr
                .MaxCharSize = 2
                .defaultChar = "?"
                .LeadByte = HexToString("0xA1ACB0C8CAFD")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "51949 (EUC-Korean)"

            'cpEUC_Traditional_Chinese ()
            Case cpEUC_Traditional_Chinese
                'TODO:
                Err.Raise STERROR_CPINFO_NOT_SET, methodName, _
                          Replace(errMsg, "%CP%", cpEUC_Traditional_Chinese)
                '.MaxCharSize =
                '.DefaultChar = ""
                '.LeadByte = ""
                '.UnicodeDefaultChar = ""
                '.CodePageName = ""

            'cpHz_gb_2312 (52936 (HZ-GB2312 Simplified Chinese))
            Case cpHz_gb_2312
                .MaxCharSize = 5
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "52936 (HZ-GB2312 Simplified Chinese)"

            'cpGB18030 (54936 (GB18030 Simplified Chinese))
            Case cpGB18030
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "54936 (GB18030 Simplified Chinese)"

            'cpSMS_GSM_7bit (55000 (SMS GSM 7bit))
            Case cpSMS_GSM_7bit
                .MaxCharSize = 2
                .defaultChar = " "
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "55000 (SMS GSM 7bit)"

            'cpSMS_GSM_7bit_Spanish (55001 (SMS GSM 7bit Spanish))
            Case cpSMS_GSM_7bit_Spanish
                .MaxCharSize = 2
                .defaultChar = " "
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "55001 (SMS GSM 7bit Spanish)"

            'cpSMS_GSM_7bit_Portuguese (55002 (SMS GSM 7bit Portuguese))
            Case cpSMS_GSM_7bit_Portuguese
                .MaxCharSize = 2
                .defaultChar = " "
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "55002 (SMS GSM 7bit Portuguese)"

            'cpSMS_GSM_7bit_Turkish (55003 (SMS GSM 7bit Turkish))
            Case cpSMS_GSM_7bit_Turkish
                .MaxCharSize = 2
                .defaultChar = " "
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "55003 (SMS GSM 7bit Turkish)"

            'cpSMS_GSM_7bit_Greek (55004 (SMS GSM 7bit Greek))
            Case cpSMS_GSM_7bit_Greek
                .MaxCharSize = 2
                .defaultChar = " "
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "55004 (SMS GSM 7bit Greek)"

            'cpX_iscii_de (57002 (ISCII - Devanagari))
            Case cpX_iscii_de
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "57002 (ISCII - Devanagari)"

            'cpX_iscii_be (57003 (ISCII - Bangla))
            Case cpX_iscii_be
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "57003 (ISCII - Bangla)"

            'cpX_iscii_ta (57004 (ISCII - Tamil))
            Case cpX_iscii_ta
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "57004 (ISCII - Tamil)"

            'cpX_iscii_te (57005 (ISCII - Telugu))
            Case cpX_iscii_te
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "57005 (ISCII - Telugu)"

            'cpX_iscii_as (57006 (ISCII - Assamese))
            Case cpX_iscii_as
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "57006 (ISCII - Assamese)"

            'cpX_iscii_or (57007 (ISCII - Odia (Oriya)))
            Case cpX_iscii_or
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "57007 (ISCII - Odia (Oriya))"

            'cpX_iscii_ka (57008 (ISCII - Kannada))
            Case cpX_iscii_ka
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "57008 (ISCII - Kannada)"

            'cpX_iscii_ma (57009 (ISCII - Malayalam))
            Case cpX_iscii_ma
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "57009 (ISCII - Malayalam)"

            'cpX_iscii_gu (57010 (ISCII - Gujarati))
            Case cpX_iscii_gu
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "57010 (ISCII - Gujarati)"

            'cpX_iscii_pa (57011 (ISCII - Punjabi (Gurmukhi)))
            Case cpX_iscii_pa
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0x3F00")
                .CodePageName = "57011 (ISCII - Punjabi (Gurmukhi))"

            'cpUTF_7 (65000 (UTF-7))
            Case cpUTF_7
                .MaxCharSize = 5
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0xFDFF")
                .CodePageName = "65000 (UTF-7)"

            'cpUTF_8 (65001 (UTF-8))
            Case cpUTF_8
                .MaxCharSize = 4
                .defaultChar = "?"
                .LeadByte = HexToString("0x")
                .UnicodeDefaultChar = HexToString("0xFDFF")
                .CodePageName = "65001 (UTF-8)"
        End Select
        
        .IsInitialized = True
    End With
    Return
End Function

Private Function GetApiErrorNumber() As Long
    #If Mac Then
        CopyMemory GetApiErrorNumber, ByVal errno_location(), 4
    #Else
        GetApiErrorNumber = Err.LastDllError 'GetLastError
    #End If
End Function

Private Function SetApiErrorNumber(ByVal errNumber As Long) As Long
    #If Mac Then
        CopyMemory ByVal errno_location(), errNumber, 4
    #Else
        SetLastError errNumber
    #End If
End Function

Public Function GetNonUnicodeSystemCodepage() As CodePageIdentifier
    #If Mac Then
        GetNonUnicodeSystemCodepage = _
            CFStringConvertEncodingToWindowsCodepage(CFStringGetSystemEncoding())
    #Else
        GetNonUnicodeSystemCodepage = GetACP
    #End If
End Function

#If Mac = 0 Then
Public Function GetBstrFromWideStringPtr(ByVal lpwString As LongPtr) As String
    Dim Length As Long
    If lpwString Then Length = lstrlenW(lpwString)
    If Length Then
        GetBstrFromWideStringPtr = Space$(Length)
        CopyMemory ByVal StrPtr(GetBstrFromWideStringPtr), ByVal lpwString, Length * 2
    End If
End Function
#End If

'This function attempts to transcode 'str' from codepage 'fromCodePage' to
'codepage 'toCodePage' using the appropriate API functions on the platform.
'Calling this function with 'raiseErrors = False' will:
'   - If 'customDefaultChar = ""':
'         Replace all occurrences of invalid bytes in the input string with the
'         so-called standard-default-character of the target codepage in the
'         output. This character is usually a "?" but could also be something
'         else.
'   - If 'customDefaultChar' is any other character:
'         Replace all occurrences of invalid bytes in the input string with the
'         character specified in 'customDefaultChar' in the target codepage in
'         the output.
'   Note: This override of the standard-default-character does NOT work for the
'         target codepages UTF-8 and UTF-16, here U+FFFD will always be used,
'         regardless of the value of 'customDefaultChar'
'Calling it with 'raiseErrors = True' will raise an error if either:
'   - the string 'str' contains byte sequences that do not represent a valid
'     string of codepage 'fromCodePage', or
'   - the string contains codepoints that can not be represented in 'toCodePage'
'     and will lead to the insertion of a "default character".
'E.g.: Transcode("°", cpUTF_16, cpUs_ascii, True) will raise an error, because
'      "°" is not an ASCII character.
'Note that even calling the function with 'raiseErrors = True' doesn't guarantee
'that the conversion is reversible, because sometimes codepoints are replaced
'with more generic characters that aren't the default character (raise no error)
'E.g.:Decode(Transcode("³", cpUTF_16, cpUs_ascii, True), cpUs_ascii) returns "3"
Public Function Transcode(ByRef str As String, _
                          ByVal fromCodePage As CodePageIdentifier, _
                          ByVal toCodePage As CodePageIdentifier, _
                 Optional ByVal raiseErrors As Boolean = False, _
                 Optional ByVal customDefaultChar As String = vbNullString) _
                          As String
    Const methodName As String = "Transcode"
    'https://developer.apple.com/library/archive/documentation/System/Conceptual/ManPages_iPhoneOS/man3/iconv.3.html
    If Not (LenB(customDefaultChar) = 2 Or LenB(customDefaultChar) = 0) Then _
        Err.Raise 5, methodName, "'customDefaultChar' must of length 1."
        
    #If Mac Then
        Dim cpi As CpInfo:           cpi = GetCpInfo(toCodePage)
        Dim inBytesLeft As LongPtr:  inBytesLeft = LenB(str)
        Dim outBytesLeft As LongPtr: outBytesLeft = inBytesLeft * cpi.MaxCharSize
        Dim cd As LongPtr: cd = GetConversionDescriptor(fromCodePage, toCodePage)
        Dim buffer As String
        buffer = Space$((CLng(inBytesLeft) * cpi.MaxCharSize + 1) \ 2)
        Dim inBuf As LongPtr:        inBuf = StrPtr(str)
        Dim outBuf As LongPtr:       outBuf = StrPtr(buffer)
        Dim irrevConvCount As LongPtr
        Dim defaultChar As String

        Do While inBytesLeft > 0
            SetApiErrorNumber 0
            irrevConvCount = iconv(cd, inBuf, inBytesLeft, outBuf, outBytesLeft)

            If irrevConvCount = -1 Then 'Error occurred
                If defaultChar <> vbNullString Then
                    If toCodePage = cpUTF_16 Then
                        defaultChar = HexToString("0xFDFF")
                    ElseIf toCodePage = cpUTF_8 Then
                        defaultChar = Encode(HexToString("0xFDFF"), cpUTF_8, _
                                             False, vbNullString)
                    ElseIf customDefaultChar <> vbNullString Then
                        defaultChar = Encode(customDefaultChar, toCodePage, _
                                             False, vbNullString)
                    Else
                        defaultChar = Encode(cpi.UnicodeDefaultChar, toCodePage, _
                                             False, vbNullString)
                    End If
                    If defaultChar = vbNullString Then _
                        Err.Raise vbErrInternalError, methodName, _
                            "Invalid default character specified by library! " _
                             & "Library implementation erroneous!"
                End If
                Select Case GetApiErrorNumber
                    Case MAC_API_ERR_EILSEQ
                        If raiseErrors Then Err.Raise 5, methodName, _
                            "Input is invalid byte sequence of " & _
                            "CodePage " & fromCodePage

                        CopyMemory ByVal outBuf, StrPtr(defaultChar), _
                                   LenB(defaultChar)
                        outBuf = outBuf + LenB(defaultChar)
                        outBytesLeft = outBytesLeft - LenB(defaultChar)
                        inBuf = inBuf + 1
                        inBytesLeft = inBytesLeft - 1
                    Case MAC_API_ERR_EINVAL
                        If raiseErrors Then Err.Raise 5, methodName, _
                            "Input is incomplete byte sequence of" & _
                            "CodePage " & fromCodePage

                        CopyMemory ByVal outBuf, StrPtr(defaultChar), _
                                   LenB(defaultChar)
                        outBuf = outBuf + outBytesLeft
                        inBuf = inBuf + inBytesLeft
                        outBytesLeft = 0
                        inBytesLeft = 0
                End Select
            End If
        Loop

        If irrevConvCount > 0 And raiseErrors Then Err.Raise 5, _
            methodName, "Default char would be used, encoding would be irreversible"

        Transcode = LeftB$(buffer, LenB(buffer) - CLng(outBytesLeft))

        'These errors are bugs and should be raised even if raiseErrors = False:
        'For some reason GetApiErrorNumber sometimes seems to return something
        'that has nothing to do with this library, therefore skip this for now
'        Select Case GetApiErrorNumber
'            Case MAC_API_ERR_E2BIG
'                Err.Raise vbErrInternalError, methodName, _
'                    "Output buffer overrun while transcoding from CodePage " _
'                    & fromCodePage & " to CodePage " & toCodePage
'            Case Is <> 0
'                Err.Raise vbErrInternalError, methodName, "Unknown error " & _
'                    "occurred during transcoding with 'iconv'. API Error" & _
'                    "Code: " & GetApiErrorNumber
'        End Select
        If iconv_close(cd) <> 0 Then
            Err.Raise vbErrInternalError, methodName, "Unknown error occurred" _
                & " when calling 'iconv_close'. API ErrorCode: " & _
                GetApiErrorNumber
        End If
    #Else
        If toCodePage = cpUTF_16 Then
            Transcode = Decode(str, fromCodePage, raiseErrors)
        ElseIf fromCodePage = cpUTF_16 Then
            Transcode = Encode(str, toCodePage, raiseErrors, customDefaultChar)
        Else
            Transcode = Encode(Decode(str, fromCodePage, raiseErrors), _
                               toCodePage, raiseErrors, customDefaultChar)
        End If
    #End If
End Function

#If Mac Then
Private Function GetConversionDescriptor( _
                            ByVal fromCodePage As CodePageIdentifier, _
                            ByVal toCodePage As CodePageIdentifier) As LongPtr
    Const methodName As String = "GetConversionDescriptor"
    Dim toCpCdName As String:   toCpCdName = ConvDescriptorName(toCodePage)
    Dim fromCpCdName As String: fromCpCdName = ConvDescriptorName(fromCodePage)
    'Todo: potentially implement custom error numbers
    If LenB(toCpCdName) = 0 Then Err.Raise 5, methodName, _
        "No conversion descriptor name assigned to CodePage " & toCodePage
    If LenB(fromCpCdName) = 0 Then Err.Raise 5, methodName, _
        "No conversion descriptor name assigned to CodePage " & fromCodePage

    SetApiErrorNumber 0  'Clear previous errors
    GetConversionDescriptor = iconv_open(StrPtr(toCpCdName), StrPtr(fromCpCdName))

    If GetConversionDescriptor = -1 Then
        Select Case GetApiErrorNumber
            Case MAC_API_ERR_EINVAL
                Err.Raise 5, methodName, "The conversion from CodePage " & _
                    fromCodePage & " to CodePage " & toCodePage & " is not " & _
                    "supported by the implementation of 'iconv' on this platform"
            Case Is <> 0
                Err.Raise vbErrInternalError, methodName, "Unknown error " & _
                    "trying to create a conversion descriptor. API Error" & _
                    "Code: " & GetApiErrorNumber
        End Select
    End If
End Function
#End If

'This function tries to encode utf16leStr from vba-internal codepage UTF-16LE to
'codepage 'toCodePage' using the appropriate API functions on the platform.
'Calling this function with 'raiseErrors = False' will:
'   - If 'customDefaultChar = ""':
'         Replace all occurrences of invalid bytes in the input string with the
'         so-called standard-default-character of the target codepage in the
'         output. This character is usually a "?" but could also be something
'         else.
'   - If 'customDefaultChar' is any other character:
'         Replace all occurrences of invalid bytes in the input string with the
'         character specified in 'customDefaultChar' in the target codepage in
'         the output.
'   Note: This override of the standard-default-character does NOT work for the
'         target codepage UTF-8, here U+FFFD will always be used, regardless of
'         the value of 'customDefaultChar'
'Calling this function with 'raiseErrors = True' will raise an error if either:
'   - the string 'utf16leStr' contains byte sequences that do not represent a
'     valid UTF-16LE string, or
'   - the string contains codepoints that can not be represented in 'toCodePage'
'     and will lead to the insertion of a "default character".
'E.g.: Encode("°", cpUs_ascii, True) will raise an error, because
'      "°" is not an ASCII character.
'Note that even calling the function with 'raiseErrors = True' doesn't guarantee
'that the conversion is reversible, because sometimes codepoints are replaced
'with more generic characters that aren't the default character (raise no error)
'E.g.: Decode(Encode("³", cpUTF_16, cpUs_ascii, True), cpUs_ascii) returns "3"
Public Function Encode(ByRef utf16leStr As String, _
                       ByVal toCodePage As CodePageIdentifier, _
              Optional ByVal raiseErrors As Boolean = False, _
              Optional ByVal customDefaultChar As String = vbNullString) _
                       As String
    Const methodName As String = "Encode"

    If Not (LenB(customDefaultChar) = 2 Or LenB(customDefaultChar) = 0) Then _
        Err.Raise 5, methodName, "'customDefaultChar' must of length 1."

    If toCodePage = cpUTF_16 Then Err.Raise 5, methodName, _
        "Input string should already be UTF-16. Can't encode UTF-16 to UTF-16."

    If utf16leStr = vbNullString Then Exit Function
    #If Mac Then
        Encode = Transcode(utf16leStr, cpUTF_16, toCodePage, raiseErrors)
    #Else
        Dim cpi As CpInfo: cpi = GetCpInfo(toCodePage)
    
        If raiseErrors Then
            Dim usedDefaultChar As Boolean
            Dim lpUsedDefaultChar As LongPtr
            If cpi.AllowsQueryReversible Then _
                lpUsedDefaultChar = VarPtr(usedDefaultChar)
            Dim dwFlags As Long
            If cpi.AllowsFlags Then dwFlags = WC_ERR_INVALID_CHARS
        Else
            If StrPtr(customDefaultChar) <> 0 Then _
                customDefaultChar = Encode(customDefaultChar, toCodePage, _
                                           False, vbNullString)
        End If
        
        Dim byteCount As Long
        SetApiErrorNumber 0
        byteCount = WideCharToMultiByte(toCodePage, dwFlags, StrPtr(utf16leStr), _
            Len(utf16leStr), 0, 0, StrPtr(customDefaultChar), lpUsedDefaultChar)
            
        If byteCount = 0 Then
            Select Case GetApiErrorNumber
                Case ERROR_NO_UNICODE_TRANSLATION
                    Err.Raise 5, methodName, _
                        "Input is invalid byte sequence of CodePage " & cpUTF_16
                Case ERROR_INVALID_PARAMETER
                    If StrPtr(customDefaultChar) = 0 Then
                        Err.Raise 5, methodName, _
                            "Conversion to CodePage " & toCodePage & " is" & _
                            " not supported by the API on this platform."
                    Else 'In this case we know that the customDefaultChar is the
                    'problem, because the conversion already worked without,
                    'when we transcoded the customDefaultChar itself.
                        Err.Raise 5, methodName, _
                            "Conversion to CodePage " & toCodePage & " is" & _
                            " not supported with 'customDefaultChar = """ & _
                            Decode(customDefaultChar, toCodePage) & """'."
                    End If
                Case ERROR_INSUFFICIENT_BUFFER, ERROR_INVALID_FLAGS
                    Err.Raise vbErrInternalError, methodName, _
                        "Library implementation erroneous. API Error: " & _
                        GetApiErrorNumber
                Case Else
                    Err.Raise vbErrInternalError, methodName, _
                        "Completely unexpected error. API Error: " & _
                        GetApiErrorNumber
            End Select
        End If

        If raiseErrors And usedDefaultChar Then _
            Err.Raise 5, methodName, "Default char would be used, encoding " & _
                "would be irreversible."

        Dim b() As Byte: ReDim b(0 To byteCount - 1)
        Encode = b
        WideCharToMultiByte toCodePage, dwFlags, StrPtr(utf16leStr), _
                            Len(utf16leStr), StrPtr(Encode), byteCount, _
                            StrPtr(customDefaultChar), lpUsedDefaultChar

        Select Case GetApiErrorNumber
            Case Is <> 0
                Err.Raise vbErrInternalError, methodName, _
                        "Completely unexpected error. API Error: " & _
                        GetApiErrorNumber
        End Select
    #End If
End Function

'This function tries to decode 'str' from codepage 'fromCodePage' to the vba-
'internal codepage UTF-16LE using the appropriate API functions on the platform.
'Calling it with 'raiseErrors = True' will raise an error if the string 'str'
'contains byte sequences that does not represent a valid encoding in codepage
'fromCodePage.
'E.g.: If 'str' is an UTF-8 encoded string that was read from an external file
'      using 'Open' and 'Get', you can convert it to the VBA-internal UTF-16LE
'      like this:
'      Decode(str, cpUTF_8)
'      By default, the function will replace invalid bytes in the input string
'      with the unicode standard replacement character U+FFFD. If you want to
'      validate that 'str' does not contain invalid bytes, use it like this:
'      Decode(str, cpUTF_8, True)
'      The function will now raise an error if invalid data is encountered.
Public Function Decode(ByRef str As String, _
                       ByVal fromCodePage As CodePageIdentifier, _
              Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "Decode"

    If fromCodePage = cpUTF_16 Then Err.Raise 5, methodName, _
        "VBA strings are UTF-16 by default. No need to decode string from UTF-16."

    If str = vbNullString Then Exit Function
    #If Mac Then
        Decode = Transcode(str, fromCodePage, cpUTF_16, raiseErrors)
    #Else
        Dim charCount As Long
        Dim dwFlags As Long
        
        If raiseErrors And CodePageAllowsFlags(fromCodePage) Then _
            dwFlags = MB_ERR_INVALID_CHARS
        
        SetApiErrorNumber 0
        charCount = MultiByteToWideChar(fromCodePage, dwFlags, StrPtr(str), _
                                        LenB(str), 0, 0)
        If charCount = 0 Then
            Select Case GetApiErrorNumber
                Case ERROR_NO_UNICODE_TRANSLATION
                    Err.Raise 5, methodName, _
                        "Input is invalid byte sequence of CodePage " & fromCodePage
                Case ERROR_INVALID_PARAMETER
                    Err.Raise 5, methodName, _
                        "Conversion from CodePage " & fromCodePage & " is not" _
                        & " supported by the API on this platform."
                Case ERROR_INSUFFICIENT_BUFFER, ERROR_INVALID_FLAGS
                    Err.Raise vbErrInternalError, methodName, _
                        "Library implementation erroneous. API Error: " & _
                        GetApiErrorNumber
                Case Else
                    Err.Raise vbErrInternalError, methodName, _
                        "Completely unexpected error. API Error: " & _
                        GetApiErrorNumber
            End Select
        End If

        Decode = Space$(charCount)
        MultiByteToWideChar fromCodePage, dwFlags, StrPtr(str), LenB(str), _
                            StrPtr(Decode), charCount

        Select Case GetApiErrorNumber
            Case Is <> 0
                Err.Raise vbErrInternalError, methodName, _
                        "Completely unexpected error. API Error: " & _
                        GetApiErrorNumber
        End Select
    #End If
End Function

'Returns strings defined as hex literal as string
'Accepts the following formattings:
'   0xXXXXXX...
'   &HXXXXXX...
'   XXXXXX...
'Where:
'   - prefixes 0x and &H are case sensitive
'   - there's an even number of Xes, X = 0-9 or a-f or A-F (case insensitive)
'Raises error 5 if:
'   - Length is not even / partial bytes
'   - Invalid characters are found (outside prefix and 0-9 / a-f / A-F ranges)
'Examples:
'   - HexToString("0x610062006300") returns "abc"
'   - StrConv(HexToString("0x616263"), vbUnicode) returns "abc"
'   - HexToString("0x61626t") or HexToString("0x61626") both raise error 5
Public Function HexToString(ByRef hexStr As String) As String
    Const methodName As String = "HexToString"
    Const errPrefix As String = "Invalid Hex string literal. "
    Dim size As Long: size = Len(hexStr)

    If size = 0 Then Exit Function
    If size Mod 2 = 1 Then Err.Raise 5, methodName, errPrefix & "Uneven length"

    Static nibbleMap(0 To 255) As Long 'Nibble: 0 to F. Byte: 00 to FF
    Static charMap(0 To 255) As String
    Dim i As Long

    If nibbleMap(0) = 0 Then
        For i = 0 To 255
            nibbleMap(i) = -256 'To force invalid character code
            charMap(i) = ChrB$(i)
        Next i
        For i = 0 To 9
            nibbleMap(Asc(CStr(i))) = i
        Next i
        For i = 10 To 15
            nibbleMap(i + 55) = i 'Asc("A") to Asc("F")
            nibbleMap(i + 87) = i 'Asc("a") to Asc("f")
        Next i
    End If

    Dim prefix As String: prefix = Left$(hexStr, 2)
    Dim startPos As Long: startPos = -4 * CLng(prefix = "0x" Or prefix = "&H")
    Dim b() As Byte:      b = hexStr
    Dim j As Long
    Dim CharCode As Long

    HexToString = MidB$(hexStr, 1, size / 2 - Sgn(startPos))
    For i = startPos To UBound(b) Step 4
        j = j + 1
        CharCode = nibbleMap(b(i)) * &H10& + nibbleMap(b(i + 2))
        If CharCode < 0 Or b(i + 1) > 0 Or b(i + 3) > 0 Then
            Err.Raise 5, methodName, errPrefix & "Expected a-f/A-F or 0-9"
        End If
        MidB$(HexToString, j, 1) = charMap(CharCode)
    Next i
End Function

'Converts the input string into a string of hex literals.
'e.g.: "abc" will be turned into "0x610062006300" (UTF-16LE)
'e.g.: StrConv("ABC", vbFromUnicode) will be turned into "0x414243"
'Note:
'StringToHex and HexToString could also be implemented using an api function:
'https://learn.microsoft.com/en-us/windows/win32/api/wincrypt/nf-wincrypt-cryptbinarytostringw
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

'Replaces all occurences of unicode characters outside the codePoint range
'defined by maxNonEscapedCharCode with literals of the following formats
'specified by `escapeFormat`:
' efPython = 1 ... \uXXXX \u00XXXXXX   (4 or 8 hex digits, 8 for chars outside BMP)
' efRust   = 2 ... \u{XXXX} \U{XXXXXX} (1 to 6 hex digits)
' efUPlus  = 4 ... u+XXXX u+XXXXXX     (4 or 6 hex digits)
' efMarkup = 8 ... &#ddddddd;          (1 to 7 decimal digits)
'Where:
'   - prefixes \u is case insensitive
'   - Xes are the digits of the codepoint in hexadecimal. (X = 0-9 or A-F/a-f)
'Note:
'   - Avoid u+XXXX syntax if string contains literals without delimiters as it
'     can be misinterpreted if adjacent to text starting with 0-9 or a-f.
'   - This function accepts all combinations of UnicodeEscapeFormats:
'     If called with, e.g. `escapeFormat = efRust Or efPython`, every character
'     in the scope will be escaped with in either format, efRust or efPython,
'     chosen at random for each replacement.
'   - If `escapeFormat` is set to efAll, it will replace every character in the
'     scope with a randomly chosen format of all available fotrmats.
'   - To escape every character, set `maxNonEscapedCharCode = -1`
Public Function EscapeUnicode(ByRef str As String, _
                     Optional ByVal maxNonEscapedCharCode As Long = &HFF, _
                     Optional ByVal escapeFormat As UnicodeEscapeFormat _
                                                = efPython) As String
    Const methodName As String = "EscapeUnicode"
    If maxNonEscapedCharCode < -1 Then Err.Raise 5, methodName, _
        "`maxNonEscapedCharCode` must be greater or equal -1."
    If escapeFormat < [_efMin] Or escapeFormat > [_efMax] Then _
        Err.Raise 5, methodName, "Invalid escape type."
    If Len(str) = 0 Then Exit Function
    Dim i As Long
    Dim j As Long:                j = 1
    Dim result() As String:       ReDim result(1 To Len(str))
    Dim copyChunkSize As Long
    Dim rndEscapeFormat As Boolean
    rndEscapeFormat = ((escapeFormat And (escapeFormat - 1)) <> 0) 'eFmt <> 2^n
    Dim numescapeFormats As Long
    If rndEscapeFormat Then
        Dim escapeFormats() As Long
        For i = 0 To (Log(efAll + 1) / Log(2)) - 1
            If 2 ^ i And escapeFormat Then
                ReDim Preserve escapeFormats(0 To numescapeFormats)
                escapeFormats(numescapeFormats) = 2 ^ i
                numescapeFormats = numescapeFormats + 1
            End If
        Next i
    End If
    For i = 1 To Len(str)
        Dim codepoint As Long: codepoint = AscU(Mid$(str, i, 2))
        If codepoint > maxNonEscapedCharCode Then
            If copyChunkSize > 0 Then
                result(j) = Mid$(str, i - copyChunkSize, copyChunkSize)
                copyChunkSize = 0
                j = j + 1
            End If
            If rndEscapeFormat Then
                escapeFormat = escapeFormats(Int(numescapeFormats * Rnd))
            End If
            Select Case escapeFormat
                Case efPython
                    If codepoint > &HFFFF& Then 'Outside BMP
                        result(j) = "\u" & "00" & Right$("0" & Hex(codepoint), 6)
                    Else 'BMP
                        result(j) = "\u" & Right$("000" & Hex(codepoint), 4)
                    End If
                Case efRust
                    result(j) = "\u{" & Hex(codepoint) & "}"
                Case efUPlus
                    If codepoint < &H1000& Then
                        result(j) = "u+" & Right$("000" & Hex(codepoint), 4)
                    Else
                        result(j) = "u+" & Hex(codepoint)
                    End If
                Case efMarkup
                    result(j) = "&#" & codepoint & ";"
            End Select
            If rndEscapeFormat Then
                If Int(2 * Rnd) = 1 Then result(j) = UCase(result(j))
            End If
            j = j + 1
        Else
            If codepoint < &H10000 Then
                copyChunkSize = copyChunkSize + 1
            Else
                copyChunkSize = copyChunkSize + 2
            End If
        End If
        If codepoint > &HFFFF& Then i = i + 1
    Next i
    If copyChunkSize > 0 Then _
        result(j) = Mid$(str, i - copyChunkSize, copyChunkSize)
    EscapeUnicode = Join(result, "")
End Function

'Replaces all occurences of unicode literals
'Accepts the following formattings `escapeFormat`:
'   efPython = 1 ... \uXXXX \u000XXXXX    (4 or 8 hex digits, 8 for chars outside BMP)
'   efRust   = 2 ... \u{XXXX} \U{XXXXXXX} (1 to 6 hex digits)
'   efUPlus  = 4 ... u+XXXX u+XXXXXX      (4 or 6 hex digits)
'   efMarkup = 8 ... &#ddddddd;           (1 to 7 decimal digits)
'Where:
'   - prefixes \u is case insensitive
'   - Xes are the digits of the codepoint in hexadecimal. (X = 0-9 or A-F/a-f)
'Example:
'   - "abcd &#97;u+0062\U0063xy\u{64}", efAll returns "abcd abcxyd"
'Notes:
'   - Avoid u+XXXX syntax if string contains literals without delimiters as it
'     can be misinterpreted if adjacent to text starting with 0-9 or a-f.
'   - This function also accepts all combinations of UnicodeEscapeFormats:
'       E.g.:
'UnescapeUnicode("abcd &#97;u+0062\U0063xy\u{64}", efMarkup Or efRust)
'       will return:
'"abcd au+0062\U0063xyd"
'   - By default, this function will not invalidate UTF-16 strings if they are
'     currently valid, but this can happen if `allowSingleSurrogates = True`
'     E.g.: EscapeUnicode(ChrU(&HD801&, True)) returns "\uD801", but this string
'     can no longer be un-escaped with UnescapeUnicode because "\uD801"
'     represents a surrogate halve which is invalid unicode on its own.
'     So UnescapeUnicode("\uD801") returns "\uD801" again, unless called with
'     the optional parameter `allowSingleSurrogates = False` like this
'     `UnescapeUnicode("\uD801", , True)`. This will return invalid UTF-16.
Public Function UnescapeUnicode(ByRef str As String, _
                       Optional ByVal escapeFormat As UnicodeEscapeFormat = efAll, _
                       Optional ByVal allowSingleSurrogates As Boolean = False) _
                                As String
    If escapeFormat < [_efMin] Or escapeFormat > [_efMax] Then
        Err.Raise 5, "EscapeUnicode", "Invalid escape format"
    End If

    Dim escapes() As EscapeSequence: escapes = NewEscapes()
    Dim lb As Long: lb = LBound(escapes)
    Dim ub As Long: ub = UBound(escapes)
    Dim i As Long

    For i = lb To ub 'Find first signature for each wanted format
        With escapes(i)
            If escapeFormat And .ueFormat Then
                .buffPosition = InStr(1, str, .ueSignature, vbBinaryCompare)
                .letSngSurrogate = allowSingleSurrogates
            End If
        End With
    Next i
    UnescapeUnicode = str 'Allocate buffer
    
    Const posByte As Byte = &H80
    Const buffSize As Long = 1024
    Dim buffSignaturePos(1 To buffSize) As Byte
    Dim buffFormat(1 To buffSize) As UnicodeEscapeFormat
    Dim buffEscIndex(1 To buffSize) As Long
    Dim posOffset As Long
    Dim diff As Long
    Dim highSur As Long
    Dim lowSur As Long
    Dim remainingLen As Long: remainingLen = Len(str)
    Dim posChar As String:    posChar = ChrB$(posByte)
    Dim outPos As Long:       outPos = 1
    Dim inPos As Long:        inPos = 1

    Do
        Dim upperLimit As Long: upperLimit = posOffset + buffSize
        For i = lb To ub 'Find all signatures within buffer size
            With escapes(i)
                Do Until .buffPosition = 0 Or .buffPosition > upperLimit
                    .buffPosition = .buffPosition - posOffset
                    buffSignaturePos(.buffPosition) = posByte
                    buffFormat(.buffPosition) = .ueFormat
                    buffEscIndex(.buffPosition) = i
                    .buffPosition = .buffPosition + .sigSize + posOffset
                    .buffPosition = InStr(.buffPosition, str, .ueSignature)
                Loop
            End With
        Next i

        Dim temp As String:  temp = buffSignaturePos
        Dim nextPos As Long: nextPos = InStrB(1, temp, posChar)

        Do Until nextPos = 0 'Unescape all found signatures from buffer
            i = buffEscIndex(nextPos)
            escapes(i).currPosition = nextPos + posOffset
            Select Case buffFormat(nextPos)
                Case efPython: TryPythonEscape escapes(i), str
                Case efRust:   TryRustEscape escapes(i), str
                Case efUPlus:  TryUPlusEscape escapes(i), str
                Case efMarkup: TryMarkupEscape escapes(i), str
            End Select
            With escapes(i)
                If .unEscSize > 0 Then
                    diff = .currPosition - inPos
                    If outPos > 1 Then
                        Mid$(UnescapeUnicode, outPos) = Mid$(str, inPos, diff)
                    End If
                    outPos = outPos + diff
                    If .unEscSize = 1 Then
                        Mid$(UnescapeUnicode, outPos) = ChrW$(.codepoint)
                    Else
                        .codepoint = .codepoint - &H10000
                        highSur = &HD800& Or (.codepoint \ &H400&)
                        lowSur = &HDC00& Or (.codepoint And &H3FF&)
                        Mid$(UnescapeUnicode, outPos) = ChrW$(highSur)
                        Mid$(UnescapeUnicode, outPos + 1) = ChrW$(lowSur)
                    End If
                    outPos = outPos + .unEscSize
                    inPos = .currPosition + .escSize
                    nextPos = nextPos + .escSize - .sigSize
                End If
                nextPos = InStrB(nextPos + .sigSize, temp, posChar)
            End With
        Loop
        remainingLen = remainingLen - buffSize
        posOffset = posOffset + buffSize
        Erase buffSignaturePos
    Loop Until remainingLen < 1
    
    If outPos > 1 Then
        diff = Len(str) - inPos + 1
        If diff > 0 Then
            Mid$(UnescapeUnicode, outPos, diff) = Mid$(str, inPos, diff)
        End If
        UnescapeUnicode = Left$(UnescapeUnicode, outPos + diff - 1)
    End If
End Function
Private Function NewEscapes() As EscapeSequence()
    Static escapes(0 To 6) As EscapeSequence
    If escapes(0).ueFormat = [_efNone] Then
        InitEscape escapes(0), efPython, "\U"
        InitEscape escapes(1), efPython, "\u"
        InitEscape escapes(2), efRust, "\U{"
        InitEscape escapes(3), efRust, "\u{"
        InitEscape escapes(4), efUPlus, "U+"
        InitEscape escapes(5), efUPlus, "u+"
        InitEscape escapes(6), efMarkup, "&#"
    End If
    NewEscapes = escapes
End Function
Private Sub InitEscape(ByRef escape As EscapeSequence, _
                       ByVal ueFormat As UnicodeEscapeFormat, _
                       ByRef ueSignature As String)
    With escape
        .ueFormat = ueFormat
        .ueSignature = ueSignature
        .sigSize = Len(ueSignature)
    End With
End Sub

Private Sub TryPythonEscape(ByRef escape As EscapeSequence, ByRef str As String)
    Const h As String = "[0-9A-Fa-f]"
    Const PYTHON_ESCAPE_PATTERN_NOT_BMP = "00[01]" & h & h & h & h & h
    Const PYTHON_ESCAPE_PATTERN_BMP As String = h & h & h & h & "*"
    Dim potentialEscape As String

    With escape
        .unEscSize = 0
        potentialEscape = Mid$(str, .currPosition + 2, 8) 'Exclude leading \[Uu]
        If potentialEscape Like PYTHON_ESCAPE_PATTERN_NOT_BMP Then
            .escSize = 10 '\[Uu]00[01]HHHHH
            .codepoint = CLng("&H" & potentialEscape) 'No extra Mid$ needed
            If .codepoint < &H10000 Then
                If IsValidBMP(.codepoint, .letSngSurrogate) Then
                    .unEscSize = 1
                    Exit Sub
                End If
            ElseIf .codepoint < &H110000 Then
                .unEscSize = 2
                Exit Sub
            End If
        End If
        If potentialEscape Like PYTHON_ESCAPE_PATTERN_BMP Then
            .escSize = 6 '\[Uu]HHHH
            .codepoint = CLng("&H" & Left$(potentialEscape, 4))
            If IsValidBMP(.codepoint, .letSngSurrogate) Then .unEscSize = 1
        End If
    End With
End Sub
Private Function IsValidBMP(ByVal codepoint As Long, _
                            ByVal letSingleSurrogate As Boolean) As Boolean
    IsValidBMP = (codepoint < &HD800& Or codepoint >= &HE000& Or letSingleSurrogate)
End Function

Private Sub TryRustEscape(ByRef escape As EscapeSequence, ByRef str As String)
    Static rustEscPattern(1 To 6) As String
    Static isPatternInit As Boolean
    Dim potentialEscape As String
    Dim nextBrace As Long
    
    If Not isPatternInit Then
        Dim i As Long
        rustEscPattern(1) = "[0-9A-Fa-f]}*"
        For i = 2 To 6
            rustEscPattern(i) = "[0-9A-Fa-f]" & rustEscPattern(i - 1)
        Next i
        isPatternInit = True
    End If
    With escape
        .unEscSize = 0
        potentialEscape = Mid$(str, .currPosition + 3, 7) 'Exclude leading \[Uu]{
        nextBrace = InStr(2, potentialEscape, "}", vbBinaryCompare)
        
        If nextBrace = 0 Then Exit Sub
        If Not potentialEscape Like rustEscPattern(nextBrace - 1) Then Exit Sub
        
        .codepoint = CLng("&H" & Left$(potentialEscape, nextBrace - 1))
        .escSize = nextBrace + 3
        If .codepoint < &H10000 Then
            If IsValidBMP(.codepoint, .letSngSurrogate) Then .unEscSize = 1
        ElseIf .codepoint < &H110000 Then
            .unEscSize = 2
        End If
    End With
End Sub

Private Sub TryUPlusEscape(ByRef escape As EscapeSequence, _
                           ByRef str As String)
    Const h As String = "[0-9A-Fa-f]"
    Const UPLUS_ESCAPE_PATTERN_4_DIGITS = h & h & h & h & "*"
    Const UPLUS_ESCAPE_PATTERN_5_DIGITS = h & h & h & h & h & "*"
    Const UPLUS_ESCAPE_PATTERN_6_DIGITS = h & h & h & h & h & h
    Dim potentialEscape As String
    
    With escape
        .unEscSize = 0
        potentialEscape = Mid$(str, .currPosition + 2, 6) 'Exclude leading [Uu]+
        If potentialEscape Like UPLUS_ESCAPE_PATTERN_6_DIGITS Then
            .escSize = 8
            .codepoint = CLng("&H" & potentialEscape)
            If .codepoint < &H10000 Then
                If IsValidBMP(.codepoint, .letSngSurrogate) Then
                    .unEscSize = 1
                    Exit Sub
                End If
            ElseIf .codepoint < &H110000 Then
                .unEscSize = 2
                Exit Sub
            End If
        End If
        If potentialEscape Like UPLUS_ESCAPE_PATTERN_5_DIGITS Then
            .escSize = 7
            .codepoint = CLng("&H" & Left$(potentialEscape, 5))
            If .codepoint < &H10000 Then
                If IsValidBMP(.codepoint, .letSngSurrogate) Then
                    .unEscSize = 1
                    Exit Sub
                End If
            Else
                .unEscSize = 2
                Exit Sub
            End If
        End If
        If potentialEscape Like UPLUS_ESCAPE_PATTERN_4_DIGITS Then
            .escSize = 6
            .codepoint = CLng("&H" & Left$(potentialEscape, 4))
            If IsValidBMP(.codepoint, .letSngSurrogate) Then .unEscSize = 1
        End If
    End With
End Sub
Private Sub TryMarkupEscape(ByRef escape As EscapeSequence, _
                            ByRef str As String)
    Static mEscPattern(1 To 7) As String
    Static isPatternInit As Boolean
    Dim potentialEscape As String
    Dim nextSemicolon As Long
    
    If Not isPatternInit Then
        Dim i As Long
        For i = 1 To 6
            mEscPattern(i) = String$(i, "#") & ";*"
        Next i
        mEscPattern(7) = "1######;"
        isPatternInit = True
    End If
    With escape
        .unEscSize = 0
        potentialEscape = Mid$(str, .currPosition + 2, 8) 'Exclude leading &[#]
        nextSemicolon = InStr(2, potentialEscape, ";", vbBinaryCompare)
        
        If nextSemicolon = 0 Then Exit Sub
        If Not potentialEscape Like mEscPattern(nextSemicolon - 1) Then Exit Sub
        
        .codepoint = CLng(Left$(potentialEscape, nextSemicolon - 1))
        .escSize = nextSemicolon + 2
        If .codepoint < &H10000 Then
            If IsValidBMP(.codepoint, .letSngSurrogate) Then .unEscSize = 1
        ElseIf .codepoint < &H110000 Then
            .unEscSize = 2
        End If
    End With
End Sub

'Returns the given unicode codepoint as standard VBA UTF-16LE string
Public Function ChrU(ByVal codepoint As Long, _
            Optional ByVal allowSingleSurrogates As Boolean = False) As String
    Const methodName As String = "ChrU"
    Static st As TwoCharTemplate
    Static lt As LongTemplate

    If codepoint < &H8000 Then Err.Raise 5, methodName, "Codepoint < -32768"
    If codepoint < 0 Then codepoint = codepoint And &HFFFF& 'Incase of uInt input

    If codepoint < &HD800& Then
        ChrU = ChrW$(codepoint)
    ElseIf codepoint < &HE000& And Not allowSingleSurrogates Then
        Err.Raise 5, methodName, "Range reserved for surrogate pairs. " & _
            "Call with 'allowSingleSurrogates = True' if this is intentional."
    ElseIf codepoint < &H10000 Then
        ChrU = ChrW$(codepoint)
    ElseIf codepoint < &H110000 Then
        lt.l = (&HD800& Or (codepoint \ &H400& - &H40&)) _
            Or (&HDC00 Or (codepoint And &H3FF&)) * &H10000 '&HDC00 with no &
        LSet st = lt
        ChrU = st.s
    Else
        Err.Raise 5, methodName, "Codepoint outside of valid Unicode range."
    End If
End Function

'Returns a given characters unicode codepoint as long.
'Note: One unicode character can consist of two VBA "characters", a so-called
'      "surrogate pair" (input string of length 2, so Len(char) = 2!)
Public Function AscU(ByRef char As String) As Long
    AscU = AscW(char) And &HFFFF&
    If Len(char) > 1 Then
        Dim lo As Long: lo = AscW(Mid$(char, 2, 1)) And &HFFFF&
        If &HDC00& > lo Or lo > &HDFFF& Then Exit Function
        AscU = (AscU - &HD800&) * &H400& + (lo - &HDC00&) + &H10000
    End If
End Function

'Function transcoding a VBA-native UTF-16LE encoded string to an ASCII string
'Note: Information will be lost for codepoints > 127!
Public Function EncodeASCII(ByRef utf16leStr As String) As String
    Dim i As Long
    Dim j As Long:         j = 0
    Dim utf16le() As Byte: utf16le = utf16leStr
    Dim ascii() As Byte

    ReDim ascii(1 To Len(utf16leStr))
    For i = LBound(ascii) To UBound(ascii)
        If utf16le(j) < 128 And utf16le(j + 1) = 0 Then
            ascii(i) = utf16le(j)
            j = j + 2
        Else
            ascii(i) = &H3F 'Chr(&H3F) = "?"
            j = j + 2
        End If
    Next i
    EncodeASCII = ascii
End Function

'Function transcoding an ASCII encoded string to the VBA-native UTF-16LE
Public Function DecodeASCII(ByRef asciiStr As String) As String
    Dim i As Long
    Dim j As Long:         j = 0
    Dim ascii() As Byte:   ascii = asciiStr
    Dim utf16le() As Byte: ReDim utf16le(0 To LenB(asciiStr) * 2 - 1)

    For i = LBound(ascii) To UBound(ascii)
        utf16le(j) = ascii(i)
        j = j + 2
    Next i
    DecodeASCII = utf16le
End Function

'Function transcoding a VBA-native UTF-16LE encoded string to an ANSI string
'Note: Information will be lost for codepoints > 255!
Public Function EncodeANSI(ByRef utf16leStr As String) As String
    Dim i As Long
    Dim j As Long:         j = 0
    Dim utf16le() As Byte: utf16le = utf16leStr
    Dim ansi() As Byte

    ReDim ansi(1 To Len(utf16leStr))
    For i = LBound(ansi) To UBound(ansi)
        If utf16le(j + 1) = 0 Then
            ansi(i) = utf16le(j)
            j = j + 2
        Else
            ansi(i) = &H3F 'Chr(&H3F) = "?"
            j = j + 2
        End If
    Next i
    EncodeANSI = ansi
End Function

'Function transcoding an ANSI encoded string to the VBA-native UTF-16LE
Public Function DecodeANSI(ByRef ansiStr As String) As String
    Dim i As Long
    Dim j As Long:         j = 0
    Dim ansi() As Byte:    ansi = ansiStr
    Dim utf16le() As Byte: ReDim utf16le(0 To LenB(ansiStr) * 2 - 1)

    For i = LBound(ansi) To UBound(ansi)
        utf16le(j) = ansi(i)
        j = j + 2
    Next i
    DecodeANSI = utf16le
End Function

'Function transcoding an VBA-native UTF-16LE encoded string to UTF-8
Public Function EncodeUTF8(ByRef utf16leStr As String, _
                  Optional ByVal raiseErrors As Boolean = False) _
                                  As String
    Const methodName As String = "EncodeUTF8"
    Dim codepoint As Long
    Dim lowSurrogate As Long
    Dim i As Long:            i = 1
    Dim j As Long:            j = 0
    Dim utf8() As Byte:       ReDim utf8(Len(utf16leStr) * 4 - 1)

    Do While i <= Len(utf16leStr)
        codepoint = AscW(Mid$(utf16leStr, i, 1)) And &HFFFF&

        If codepoint >= &HD800& And codepoint <= &HDBFF& Then 'high surrogate
            lowSurrogate = AscW(Mid$(utf16leStr, i + 1, 1)) And &HFFFF&

            If &HDC00& <= lowSurrogate And lowSurrogate <= &HDFFF& Then
                codepoint = (codepoint - &HD800&) * &H400& + _
                            (lowSurrogate - &HDC00&) + &H10000
                i = i + 1
            Else
                If raiseErrors Then _
                    Err.Raise 5, methodName, _
                        "Invalid Unicode codepoint. (Lonely high surrogate)"
                codepoint = &HFFFD&
            End If
        End If

        If codepoint < &H80& Then
            utf8(j) = codepoint
            j = j + 1
        ElseIf codepoint < &H800& Then
            utf8(j) = &HC0& Or ((codepoint And &H7C0&) \ &H40&)
            utf8(j + 1) = &H80& Or (codepoint And &H3F&)
            j = j + 2
        ElseIf codepoint < &HDC00& Then
            utf8(j) = &HE0& Or ((codepoint And &HF000&) \ &H1000&)
            utf8(j + 1) = &H80& Or ((codepoint And &HFC0&) \ &H40&)
            utf8(j + 2) = &H80& Or (codepoint And &H3F&)
            j = j + 3
        ElseIf codepoint < &HE000& Then
            If raiseErrors Then _
                Err.Raise 5, methodName, _
                    "Invalid Unicode codepoint. (Lonely low surrogate)"
            codepoint = &HFFFD&
            utf8(j) = &HE0& Or ((codepoint And &HF000&) \ &H1000&)
            utf8(j + 1) = &H80& Or ((codepoint And &HFC0&) \ &H40&)
            utf8(j + 2) = &H80& Or (codepoint And &H3F&)
            j = j + 3
        ElseIf codepoint < &H10000 Then
            utf8(j) = &HE0& Or ((codepoint And &HF000&) \ &H1000&)
            utf8(j + 1) = &H80& Or ((codepoint And &HFC0&) \ &H40&)
            utf8(j + 2) = &H80& Or (codepoint And &H3F&)
            j = j + 3
        Else
            utf8(j) = &HF0& Or ((codepoint And &H1C0000) \ &H40000)
            utf8(j + 1) = &H80& Or ((codepoint And &H3F000) \ &H1000&)
            utf8(j + 2) = &H80& Or ((codepoint And &HFC0&) \ &H40&)
            utf8(j + 3) = &H80& Or (codepoint And &H3F&)
            j = j + 4
        End If

        i = i + 1
    Loop
    EncodeUTF8 = MidB$(utf8, 1, j)
End Function

'Function transcoding an UTF-8 encoded string to the VBA-native UTF-16LE
'TODO: Make error character insertion 100% identical to API function
Public Function DecodeUTF8(ByRef utf8Str As String, _
                  Optional ByVal raiseErrors As Boolean = False) As String

    Const methodName As String = "DecodeUTF8"
    Dim i As Long
    Dim numBytesOfCodePoint As Byte

    Static numBytesOfCodePoints(0 To 255) As Byte
    Static mask(2 To 4) As Long
    Static minCp(2 To 4) As Long

    If numBytesOfCodePoints(0) = 0 Then
        For i = &H0& To &H7F&: numBytesOfCodePoints(i) = 1: Next i '0xxxxxxx
        '110xxxxx - C0 and C1 are invalid (overlong encoding)
        For i = &HC2& To &HDF&: numBytesOfCodePoints(i) = 2: Next i
        For i = &HE0& To &HEF&: numBytesOfCodePoints(i) = 3: Next i '1110xxxx
       '11110xxx - 11110100, 11110101+ (= &HF5+) outside of valid Unicode range
        For i = &HF0& To &HF4&: numBytesOfCodePoints(i) = 4: Next i
        For i = 2 To 4: mask(i) = (2 ^ (7 - i) - 1): Next i
        minCp(2) = &H80&: minCp(3) = &H800&: minCp(4) = &H10000
    End If

    Dim codepoint As Long
    Dim currByte As Byte
    Dim utf8() As Byte:  utf8 = utf8Str
    Dim utf16() As Byte: ReDim utf16(0 To (UBound(utf8) - LBound(utf8) + 1) * 2)
    Dim j As Long:       j = 0
    Dim k As Long

    i = LBound(utf8)
    Do While i <= UBound(utf8)
        codepoint = utf8(i)
        numBytesOfCodePoint = numBytesOfCodePoints(codepoint)

        If numBytesOfCodePoint = 0 Then
            If raiseErrors Then Err.Raise 5, methodName, "Invalid byte"
            numBytesOfCodePoint = 1
            GoTo insertErrChar
        ElseIf numBytesOfCodePoint = 1 Then
            utf16(j) = codepoint
            j = j + 2
        ElseIf i + numBytesOfCodePoint - 1 > UBound(utf8) Then
            If raiseErrors Then Err.Raise 5, methodName, _
                    "Incomplete UTF-8 codepoint at end of string."
            GoTo insertErrChar
        Else
            codepoint = utf8(i) And mask(numBytesOfCodePoint)

            For k = 1 To numBytesOfCodePoint - 1
                currByte = utf8(i + k)

                If (currByte And &HC0&) = &H80& Then
                    codepoint = (codepoint * &H40&) + (currByte And &H3F)
                Else
                    If raiseErrors Then _
                        Err.Raise 5, methodName, "Invalid continuation byte"
                    numBytesOfCodePoint = k
                    GoTo insertErrChar
                End If
            Next k
            'Convert the Unicode codepoint to UTF-16LE bytes
            If codepoint < minCp(numBytesOfCodePoint) Then
                If raiseErrors Then Err.Raise 5, methodName, "Overlong encoding"
                GoTo insertErrChar
            ElseIf codepoint < &HD800& Then
                utf16(j) = codepoint And &HFF&
                utf16(j + 1) = codepoint \ &H100&
                j = j + 2
            ElseIf codepoint < &HE000& Then
                If raiseErrors Then Err.Raise 5, methodName, _
                "Invalid Unicode codepoint.(Range reserved for surrogate pairs)"
                GoTo insertErrChar
            ElseIf codepoint < &H10000 Then
                If codepoint = &HFEFF& Then GoTo nextCp '(BOM - will be ignored)
                utf16(j) = codepoint And &HFF&
                utf16(j + 1) = codepoint \ &H100&
                j = j + 2
            ElseIf codepoint < &H110000 Then 'Calculate surrogate pair
                Dim m As Long:           m = codepoint - &H10000
                Dim loSurrogate As Long: loSurrogate = &HDC00& Or (m And &H3FF)
                Dim hiSurrogate As Long: hiSurrogate = &HD800& Or (m \ &H400&)

                utf16(j) = hiSurrogate And &HFF&
                utf16(j + 1) = hiSurrogate \ &H100&
                utf16(j + 2) = loSurrogate And &HFF&
                utf16(j + 3) = loSurrogate \ &H100&
                j = j + 4
            Else
                If raiseErrors Then Err.Raise 5, methodName, _
                        "Codepoint outside of valid Unicode range"
insertErrChar:  utf16(j) = &HFD
                utf16(j + 1) = &HFF
                j = j + 2
            End If
        End If
nextCp: i = i + numBytesOfCodePoint 'Move to the next UTF-8 codepoint
    Loop
    DecodeUTF8 = MidB$(utf16, 1, j)
End Function

'Function transcoding an VBA-native UTF-16LE encoded string to UTF-32
Public Function EncodeUTF32LE(ByRef utf16leStr As String, _
                     Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "EncodeUTF32LE"

    If utf16leStr = "" Then Exit Function

    Dim codepoint As Long
    Dim lowSurrogate As Long
    Dim utf32() As Byte:      ReDim utf32(Len(utf16leStr) * 4 - 1)
    Dim i As Long:            i = 1
    Dim j As Long:            j = 0

    Do While i <= Len(utf16leStr)
        codepoint = AscW(Mid$(utf16leStr, i, 1)) And &HFFFF&

        If codepoint >= &HD800& And codepoint <= &HDBFF& Then 'high surrogate
            lowSurrogate = AscW(Mid$(utf16leStr, i + 1, 1)) And &HFFFF&

            If &HDC00& <= lowSurrogate And lowSurrogate <= &HDFFF& Then
                codepoint = (codepoint - &HD800&) * &H400& + _
                            (lowSurrogate - &HDC00&) + &H10000
                i = i + 1
            Else
                If raiseErrors Then Err.Raise 5, methodName, _
                    "Invalid Unicode codepoint. (Lonely high surrogate)"
                codepoint = &HFFFD&
            End If
        End If
        
        If codepoint >= &HD800& And codepoint < &HE000& Then
            If raiseErrors Then Err.Raise 5, methodName, _
                "Invalid Unicode codepoint. (Lonely low surrogate)"
            codepoint = &HFFFD&
        ElseIf codepoint > &H10FFFF Then
            If raiseErrors Then Err.Raise 5, methodName, _
                "Codepoint outside of valid Unicode range"
            codepoint = &HFFFD&
        End If

        utf32(j) = codepoint And &HFF&
        utf32(j + 1) = (codepoint \ &H100&) And &HFF&
        utf32(j + 2) = (codepoint \ &H10000) And &HFF&
        i = i + 1: j = j + 4
    Loop
    EncodeUTF32LE = MidB$(utf32, 1, j)
End Function

'Function transcoding an UTF-32 encoded string to the VBA-native UTF-16LE
Public Function DecodeUTF32LE(ByRef utf32str As String, _
                     Optional ByVal raiseErrors As Boolean = False) As String
    Const methodName As String = "DecodeUTF32LE"

    If utf32str = "" Then Exit Function

    Dim codepoint As Long
    Dim utf32() As Byte:   utf32 = utf32str
    Dim utf16() As Byte:   ReDim utf16(LBound(utf32) To UBound(utf32))
    Dim i As Long: i = LBound(utf32)
    Dim j As Long: j = i

    Do While i < UBound(utf32)
        If utf32(i + 2) = 0 And utf32(i + 3) = 0 Then
            utf16(j) = utf32(i): utf16(j + 1) = utf32(i + 1): j = j + 2
        Else
            If utf32(i + 3) <> 0 Then
                If raiseErrors Then _
                    Err.Raise 5, methodName, _
                    "Codepoint outside of valid Unicode range"
                codepoint = &HFFFD&
            Else
                codepoint = utf32(i + 2) * &H10000 + _
                            utf32(i + 1) * &H100& + utf32(i)
                If codepoint >= &HD800& And codepoint < &HE000& Then
                    If raiseErrors Then _
                        Err.Raise 5, methodName, _
                        "Invalid Unicode codepoint. " & _
                        "(Range reserved for surrogate pairs)"
                    codepoint = &HFFFD&
                ElseIf codepoint > &H10FFFF Then
                    If raiseErrors Then _
                        Err.Raise 5, methodName, _
                        "Codepoint outside of valid Unicode range"
                    codepoint = &HFFFD&
                End If
            End If

            Dim n As Long:             n = codepoint - &H10000
            Dim highSurrogate As Long: highSurrogate = &HD800& Or (n \ &H400&)
            Dim lowSurrogate As Long:  lowSurrogate = &HDC00& Or (n And &H3FF)

            utf16(j) = highSurrogate And &HFF&
            utf16(j + 1) = highSurrogate \ &H100&
            utf16(j + 2) = lowSurrogate And &HFF&
            utf16(j + 3) = lowSurrogate \ &H100&
            j = j + 4
        End If
        i = i + 4
    Loop
    DecodeUTF32LE = MidB$(utf16, 1, j)
End Function

'Returns a UTF-16 string containing alphanumeric characters randomly equally
'distributed. (0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz)
Public Function RandomStringAlphanumeric(ByVal Length As Long, _
                                Optional ByVal useRndWH As Boolean = False) As String
    Const methodName As String = "RandomStringAlphanumeric"
    Const INKL_CHARS As String = _
        "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    Static chars() As Byte
    Static numPossChars As Long
    Static IsInitialized As Boolean
    If Not IsInitialized Then
        chars = StrConv(INKL_CHARS, vbFromUnicode)
        numPossChars = UBound(chars) - LBound(chars) + 1
        IsInitialized = True
    End If
    
    If Length = 0 Then Exit Function
    If Length < 0 Then Err.Raise 5, methodName, "Length must be >= 0"
    Dim b() As Byte: ReDim b(0 To Length * 2 - 1)
    Dim i As Long
    If useRndWH Then
        For i = 0 To Length * 2 - 1 Step 2
            b(i) = chars(Int(RndWH * numPossChars))
        Next i
    Else
        For i = 0 To Length * 2 - 1 Step 2
            b(i) = chars(Int(Rnd * numPossChars))
        Next i
    End If
    RandomStringAlphanumeric = b
End Function

'Returns a UTF-16 string containing random ASCII characters equally,
'randomly distributed.
Public Function RandomStringASCII(ByVal Length As Long, _
                         Optional ByVal useRndWH As Boolean = False) As String
    Const methodName As String = "RandomStringASCII"
    Const MAX_ASC As Long = &H7F&
    If Length = 0 Then Exit Function
    If Length < 0 Then Err.Raise 5, methodName, "Length must be >= 0"
    Dim i As Long
    Dim b() As Byte: ReDim b(0 To Length * 2 - 1)
    If useRndWH Then
        For i = 0 To Length * 2 - 1 Step 2
            b(i) = Int(MAX_ASC * RndWH) + 1
        Next i
    Else
        For i = 0 To Length * 2 - 1 Step 2
            b(i) = Int(MAX_ASC * Rnd) + 1
        Next i
    End If
    RandomStringASCII = b
End Function

'Function returning a UTF-16 string containing random characters from the BMP
'(Basic Multilingual Plane, so from all 2 byte UTF-16 chars) equally, randomly
'distributed. Excludes surrogate range and BOM.
Public Function RandomStringBMP(ByVal Length As Long, _
                       Optional ByVal useRndWH As Boolean = False) As String
    Const methodName As String = "RandomStringBMP"
    Const MAX_UINT As Long = &HFFFF&
    If Length = 0 Then Exit Function
    If Length < 0 Then Err.Raise 5, methodName, "Length must be >= 0"

    Dim i As Long
    Dim char As Long
    Dim b() As Byte:  ReDim b(0 To Length * 2 - 1)

    For i = 0 To Length * 2 - 1 Step 2
        Do
            If useRndWH Then
                char = Int(MAX_UINT * RndWH) + 1
            Else
                char = Int(MAX_UINT * Rnd) + 1
            End If
        Loop Until (char < &HD800& Or char > &HDFFF&) _
               And (char <> &HFEFF&)
        b(i) = char And &HFF
        b(i + 1) = char \ &H100& And &HFF
    Next i
    RandomStringBMP = b
End Function

'Returns a UTF-16 string containing random valid unicode characters equally,
'randomly distributed. Excludes surrogate range and BOM.
'Length in UTF-16 codepoints, (Len(result) = length)
Public Function RandomStringUnicode(ByVal Length As Long, _
                           Optional ByVal useRndWH As Boolean = False) As String
    Const methodName As String = "RandomStringUnicode"
    Const MAX_UNICODE As Long = &H10FFFF
    If Length = 0 Then Exit Function
    If Length < 0 Then Err.Raise 5, methodName, "Length must be >= 0"

    Dim i As Long
    Dim char As Long
    Dim b() As Byte: ReDim b(0 To Length * 2 - 1)

    If Length > 1 Then
        For i = 0 To Length * 2 - 3 Step 2
            Do
                If useRndWH Then
                    char = Int(MAX_UNICODE * RndWH) + 1
                Else
                    char = Int(MAX_UNICODE * Rnd) + 1
                End If
            Loop Until (char < &HD800& Or char > &HDFFF&) _
                   And (char <> &HFEFF&)
            If char < &H10000 Then
                b(i) = char And &HFF
                b(i + 1) = char \ &H100& And &HFF
            Else
                Dim m As Long: m = char - &H10000
                Dim highSurrogate As Long: highSurrogate = &HD800& + (m \ &H400&)
                Dim lowSurrogate As Long: lowSurrogate = &HDC00& + (m And &H3FF)
                b(i) = highSurrogate And &HFF&
                b(i + 1) = highSurrogate \ &H100&
                i = i + 2
                b(i) = lowSurrogate And &HFF&
                b(i + 1) = lowSurrogate \ &H100&
            End If
        Next i
    End If
    RandomStringUnicode = b
    
    Const MAX_UINT As Long = &HFFFF&
    If CInt(b(UBound(b) - 1)) + b(UBound(b)) = 0 Then
        Do
            If useRndWH Then
                char = Int(MAX_UINT * RndWH) + 1
            Else
                char = Int(MAX_UINT * Rnd) + 1
            End If
        Loop Until (char < &HD800& Or char > &HDFFF&) _
               And (char <> &HFEFF&)
        Mid$(RandomStringUnicode, Len(RandomStringUnicode), 1) = ChrW(char)
    End If
End Function

'Returns a string containing random byte data
Public Function RandomBytes(ByVal numBytes As Long, _
                   Optional ByVal useRndWH As Boolean = False) As String
    Const methodName As String = "RandomBytes"
    If numBytes = 0 Then Exit Function
    If numBytes < 0 Then Err.Raise 5, methodName, "numBytes must be >= 0"

    Dim bytes() As Byte: ReDim bytes(0 To numBytes - 1)
    Dim i As Long
    If useRndWH Then
        For i = 0 To numBytes - 1
            bytes(i) = Int(RndWH * &H100)
        Next i
    Else
        For i = 0 To numBytes - 1
            bytes(i) = Int(Rnd * &H100)
        Next i
    End If
    RandomBytes = bytes
End Function

'Returns a UTF-16 string containing random characters from the codepoint range
'between 'minCodepoint' and 'maxCodepoint'.
'E.g.: RandomString(10, 48, 57) will return a string of length 100 containing
'      all the digit characters randomly, e.g. "3239107914"
Public Function RandomString(ByVal Length As Long, _
                    Optional ByVal minCodepoint As Long = 1, _
                    Optional ByVal maxCodepoint As Long = &H10FFFF, _
                    Optional ByVal useRndWH As Boolean = False) As String
    Const methodName As String = "RandomString"
    Const MAX_UNICODE As Long = &H10FFFF
    Const MAX_UINT As Long = &HFFFF&
    If Length = 0 Then Exit Function
    If Length < 0 Then Err.Raise 5, methodName, "Length must be >= 0"
    If maxCodepoint > MAX_UNICODE Or maxCodepoint < 0 Then Err.Raise 5, _
        methodName, "'maxCodepoint' outside of valid unicode range."
    If minCodepoint > MAX_UNICODE Or minCodepoint < 0 Then Err.Raise 5, _
        methodName, "'minCodepoint' outside of valid unicode range."
    If minCodepoint > maxCodepoint Then Err.Raise 5, methodName, _
        "'minCodepoint' can't be greater than 'maxCodepoint'."
    If minCodepoint > MAX_UINT And Length Mod 2 = 1 Then Err.Raise 5, methodName, _
        "Can't build string of uneven length from only Surrogate Pairs."
        
    Dim cpRange As Long: cpRange = maxCodepoint - minCodepoint + 1

    Dim i As Long
    Dim char As Long
    Dim b() As Byte: ReDim b(0 To Length * 2 - 1)

    If Length > 1 Then
        For i = 0 To Length * 2 - 3 Step 2
            Do
                If useRndWH Then
                    char = Int(cpRange * RndWH) + minCodepoint
                Else
                    char = Int(cpRange * Rnd) + minCodepoint
                End If
            Loop Until (char < &HD800& Or char > &HDFFF&) _
                   And (char <> &HFEFF&)

            If char < &H10000 Then
                b(i) = char And &HFF
                b(i + 1) = char \ &H100& And &HFF
            Else
                Dim m As Long: m = char - &H10000
                Dim highSurrogate As Long: highSurrogate = &HD800& + (m \ &H400&)
                Dim lowSurrogate As Long: lowSurrogate = &HDC00& + (m And &H3FF)
                b(i) = highSurrogate And &HFF&
                b(i + 1) = highSurrogate \ &H100&
                i = i + 2
                b(i) = lowSurrogate And &HFF&
                b(i + 1) = lowSurrogate \ &H100&
            End If
        Next i
    End If
    RandomString = b
    
    If CInt(b(UBound(b) - 1)) + b(UBound(b)) = 0 Then
        Do
            If useRndWH Then
                char = Int(cpRange * RndWH) + minCodepoint
            Else
                char = Int(cpRange * Rnd) + minCodepoint
            End If
        Loop Until (char < &HD800& Or char > &HDFFF&) _
               And (char <> &HFEFF&) _
               And (char <= MAX_UINT)
        Mid$(RandomString, Len(RandomString), 1) = ChrW(char)
    End If
End Function

'Returns a UTF-16 string containing characters from `sourceChars` randomly,
'equally distributed.
'E.g. if 'sourceChars = "aab"', the returned string will, on average, contain
'     about twice as many "a"s as "b"s
Public Function RandomStringFromChars(ByVal Length As Long, _
                             Optional ByRef sourceChars As String = _
    "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz", _
                             Optional ByVal useRndWH As Boolean = False) _
                                      As String
    Const methodName As String = "RandomStringFromChars"
    If Length = 0 Then Exit Function
    If Len(sourceChars) = 0 Then Err.Raise 5, methodName, _
        "No characters to build a string from specified in 'sourceChars'"
    If Length < 0 Then Err.Raise 5, methodName, "Length must be >= 0"
    
    Dim chars() As String:    chars = StringToCodepointStrings(sourceChars)
    Dim codepoints() As Long: codepoints = StringToCodepointNums(sourceChars)
    Dim numChars As Long:     numChars = UBound(chars) - LBound(chars) + 1
    If numChars * 2 = Len(sourceChars) And Length Mod 2 = 1 Then Err.Raise 5, _
    methodName, "Can't build string of uneven length from only Surrogate Pairs."
        
    RandomStringFromChars = Space$(Length)

    Dim i As Long
    For i = 1 To Length - 1
        Dim idx As Long
        If useRndWH Then
            idx = Int(RndWH * numChars)
        Else
            idx = Int(Rnd * numChars)
        End If
        Mid$(RandomStringFromChars, i) = chars(idx)
        If codepoints(idx) > &HFFFF& Then i = i + 1
    Next i
    If Mid$(RandomStringFromChars, Length) = " " Then
        Do
            If useRndWH Then
                idx = Int(RndWH * numChars)
            Else
                idx = Int(Rnd * numChars)
            End If
        Loop Until codepoints(idx) < &H10000
        Mid$(RandomStringFromChars, Length) = chars(idx)
    End If
End Function

'Returns an array of strings containing the individual UTF-16 characters
'Surrogate pairs remain together.
Public Function StringToCodepointStrings(ByRef str As String) As Variant
    If Len(str) = 0 Then Exit Function
    Dim arr() As String: ReDim arr(0 To Len(str) - 1)
    Dim i As Long, j As Long
    For i = 1 To Len(str)
        If AscU(Mid$(str, i, 2)) > &HFFFF& Then
            arr(j) = Mid$(str, i, 2)
            i = i + 1
        Else
            arr(j) = Mid$(str, i, 1)
        End If
        j = j + 1
    Next i
    ReDim Preserve arr(0 To j - 1)
    StringToCodepointStrings = arr
End Function

'Returns an array of numbers representing the individual UTF-16 codepoints from
'the string 'str'
Public Function StringToCodepointNums(ByRef str As String) As Variant
    If Len(str) = 0 Then Exit Function
    Dim arr() As Long: ReDim arr(0 To Len(str) - 1)
    Dim i As Long, j As Long
    Dim codepoint As Long
    For i = 1 To Len(str)
        codepoint = AscU(Mid$(str, i, 2))
        arr(j) = codepoint
        If codepoint > &HFFFF& Then i = i + 1
        j = j + 1
    Next i
    ReDim Preserve arr(0 To j - 1)
    StringToCodepointNums = arr
End Function

'Returns a string of length `Length` containing strings from `sourceStringsArr`
'randomly concatenated, equally distributed.
'E.g. if 'sourceStringsArr = ("ab", "cd")' and 'Length = 5', the returned string
'     could look like this: "ababc", or like this: "cdabc"
Public Function RandomStringFromStrings(ByVal Length As Long, _
                                        ByRef sourceStringsArr As Variant, _
                               Optional ByVal useRndWH As Boolean = False) _
                                        As String
    Const methodName As String = "RandomStringFromStrings"
    Dim strings As Variant
    If IsArray(sourceStringsArr) Then strings = sourceStringsArr _
                                 Else: strings = VBA.Array(sourceStringsArr)
    If Not LBound(strings) = 0 Then _
        ReDim Preserve strings(0 To UBound(strings) - LBound(strings))

    Dim i As Long, j As Long
    Dim tmpStr As String
    On Error GoTo catch

    For i = LBound(strings) To UBound(strings)
        tmpStr = CStr(strings(i))
        If LenB(tmpStr) > 0 Then
            strings(j) = tmpStr
            j = j + 1
        End If
    Next i
    On Error GoTo 0
    If j = 0 Then Err.Raise 5, methodName, "No strings with LenB > 0 in " & _
                                           "'sourceStringsArr'"
    GoTo continue
catch:
    Err.Raise 5, methodName, "Argument 'strings' contains invalid elements " & _
                             "that can't be converted to String type."
continue:
    If j <> i Then ReDim Preserve strings(0 To j - 1)
    RandomStringFromStrings = Space$(Length)
    
    Dim numStrings As Long: numStrings = j
    i = 1
    Do Until i > Length * 2
        If useRndWH Then
            j = Int(RndWH * numStrings)
        Else
            j = Int(Rnd * numStrings)
        End If
        MidB$(RandomStringFromStrings, i, LenB(strings(j))) = strings(j)
        i = i + LenB(strings(j))
    Loop
End Function

'Returns an array of 'numElements' random UTF-16 strings
Public Function RandomStringArray(ByVal numElements As Long, _
                         Optional ByVal maxElementLength As Long = 10, _
                         Optional ByVal minElementLength As Long = 0, _
                         Optional ByVal minCodepoint As Long = 1, _
                         Optional ByVal maxCodepoint As Long = &H10FFFF, _
                         Optional ByVal useRndWH As Boolean = False) _
                                  As String()
    Const methodName As String = "RandomStringArray"
    Const MAX_UNICODE As Long = &H10FFFF
    Const MAX_UINT As Long = &HFFFF&
    If numElements < 0 Then Err.Raise 5, methodName, "numElements must be >= 0"
    If maxCodepoint > MAX_UNICODE Or maxCodepoint < 0 Then Err.Raise 5, _
        methodName, "'maxCodepoint' outside of valid unicode range."
    If minCodepoint > MAX_UNICODE Or minCodepoint < 0 Then Err.Raise 5, _
        methodName, "'minCodepoint' outside of valid unicode range."
    If minCodepoint > maxCodepoint Then Err.Raise 5, methodName, _
        "'minCodepoint' can't be greater than 'maxCodepoint'."
    If minElementLength > maxElementLength Then Err.Raise 5, methodName, _
        "'minElementLength' can't be greater than 'maxElementLength'."
    If minCodepoint > &HFFFF& And maxElementLength = minElementLength _
    And maxElementLength Mod 2 = 1 Then Err.Raise 5, methodName, _
        "Can't build string of uneven length from only Surrogate Pairs."
    
    If numElements = 0 Then
        RandomStringArray = Split("", , 0)
        Exit Function
    End If
    
    Dim stringArray() As String: ReDim stringArray(0 To numElements - 1)
    Dim i As Long
    Dim strLength As Long
    
    For i = 0 To numElements - 1
        Do
            strLength = Int((maxElementLength - minElementLength + 1) _
                            * Rnd + minElementLength) 'No need to use RndWH here
            If minCodepoint > &HFFFF& Then
                If strLength Mod 2 = 0 Then Exit Do
            Else
                Exit Do
            End If
        Loop
        stringArray(i) = RandomString(strLength, minCodepoint, maxCodepoint, useRndWH)
    Next i
    RandomStringArray = stringArray
End Function

'Removes all characters from a string (str) that are not in the string inklChars
'Default inklChars are all alphanumeric characters including dot and space
Public Function CleanString(ByRef str As String, _
                   Optional ByVal inklChars As String = _
    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789. ") _
                            As String
    Dim sChr As String
    Dim i As Long
    Dim j As Long: j = 1

    For i = 1 To Len(str)
        sChr = Mid$(str, i, 1)

        If InStr(1, inklChars, sChr, vbBinaryCompare) Then
            Mid$(str, j, 1) = sChr
            j = j + 1
        End If
    Next i
    CleanString = Left$(str, j - 1)
End Function

'Removes all non-numeric characters from a string.
'Keeps only codepoints U+0030 - U+0039 AND ALSO
'keeps the Unicode "Fullwidth Digits" (U+FF10 - U+FF19)!
Public Function RemoveNonNumeric(ByVal str As String) As String
    Dim sChr As String
    Dim i As Long
    Dim j As Long: j = 1
    For i = 1 To Len(str)
        sChr = Mid$(str, i, 1)
        If sChr Like "#" Then _
            Mid$(str, j, 1) = sChr: j = j + 1
    Next i
    RemoveNonNumeric = Left$(str, j - 1)
End Function

'Inserts a string into another string at a specified position
'Insert("abcd", "ff", 0) = "ffabcd"
'Insert("abcd", "ff", 1) = "affbcd"
'Insert("abcd", "ff", 3) = "abcffd"
'Insert("abcd", "ff", 4) = "abcdff"
'Insert("abcd", "ff", 9) = "abcdff"
'Todo: function may be optimizable by avoiding double string concatenation
Public Function Insert(ByRef str As String, _
                       ByRef strToInsert As String, _
                       ByRef afterPos As Long) As String
    Const methodName As String = "Insert"
    If afterPos < 0 Then Err.Raise 5, methodName, _
        "Argument 'afterPos' = " & afterPos & " < 0, invalid"

    Insert = Mid$(str, 1, afterPos) & strToInsert & Mid$(str, afterPos + 1)
End Function

'Works like Insert but interprets 'afterPos' as byte-index, not char-index
'Inserting at uneven byte positions likely invalidates an utf-16 string!
Public Function InsertB(ByRef str As String, _
                        ByRef strToInsert As String, _
                        ByRef afterPos As Long) As String
    Const methodName As String = "InsertB"
    If afterPos < 0 Then Err.Raise 5, methodName, _
        "Argument 'afterPos' = " & afterPos & " < 0, invalid"

    InsertB = MidB$(str, 1, afterPos) & strToInsert & MidB$(str, afterPos + 1)
End Function

'Counts the number of times a substring exists in a string. Does not count
'overlapping occurrences of substring.
''lLimit' can define a number after which the function should stop counting,
'enables premature exiting of the procedure if for example the calling code
'just needs the information if count >= lLimit
'E.g.: CountSubstring("abababab", "abab") -> 2
Public Function CountSubstring(ByRef str As String, _
                               ByRef subStr As String, _
                      Optional ByVal lStart As Long = 1, _
                      Optional ByVal lLimit As Long = -1, _
                      Optional ByVal lCompare As VbCompareMethod _
                                                 = vbBinaryCompare) As Long
    Const methodName As String = "CountSubstring"
    If lStart < 1 Then Err.Raise 5, methodName, _
        "Argument 'Start' = " & lStart & " < 1, invalid"
    If lLimit < -1 Then Err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"
    If subStr = vbNullString Then Exit Function
    
    If lCompare = vbTextCompare And Len(str) > 1000 Then
        'In the case of vbTextCompare, InStr's runtime will always be
        'proportional to the find position relative to the beginning of the
        'string, not relative to its 'Start' parameter, therefore, this method
        'using 'Replace' should usually be much faster
        CountSubstring = (Len(str) - Len(Replace(str, subStr, vbNullString, _
                          lStart, lLimit, vbTextCompare))) \ Len(subStr)
    Else
        Dim lenSubStr As Long: lenSubStr = Len(subStr)
        Dim i As Long:         i = InStr(lStart, str, subStr, lCompare)
    
        CountSubstring = 0
        Do Until i = 0 Or lLimit = CountSubstring
            CountSubstring = CountSubstring + 1
            i = InStr(i + lenSubStr, str, subStr, lCompare)
        Loop
    End If
End Function

'Like CountSubstring but scans a string bytewise.
'Example illustrating the difference to CountSubstring:
'                       |c1||c2|
'bytes = HexToString("0x00610061")
'                         |c3|
'sFind =   HexToString("0x6100")
'CountSubstring(bytes, sFind) -> 0
'CountSubstringB(bytes, sFind) -> 1
Public Function CountSubstringB(ByRef bytes As String, _
                                ByRef subStr As String, _
                       Optional ByVal lStart As Long = 1, _
                       Optional ByVal lLimit As Long = -1, _
                       Optional ByVal lCompare As VbCompareMethod _
                                               = vbBinaryCompare) As Long
    Const methodName As String = "CountSubstringB"
    If lStart < 1 Then Err.Raise 5, methodName, _
        "Argument 'Start' = " & lStart & " < 1, invalid"
    If lLimit < -1 Then Err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"
    If subStr = vbNullString Then Exit Function
    
    Dim lenBSubStr As Long: lenBSubStr = LenB(subStr)
    Dim i As Long:          i = InStrB(lStart, bytes, subStr, lCompare)

    CountSubstringB = 0
    Do Until i = 0 Or lLimit = CountSubstringB
        CountSubstringB = CountSubstringB + 1
        i = InStrB(i + lenBSubStr, bytes, subStr, lCompare)
    Loop
End Function

'Counts the number of times a substring exists in a string unless they are
'escaped' (appear twice in a row). Does not count overlapping occurrences of
'substring.
'E.g.: CountSubstringUnlessEscaped("abababababab", "abab") -> 1
Public Function CountSubstringUnlessEscaped(ByRef str As String, _
                                            ByRef subStr As String, _
                                   Optional ByVal lStart As Long = 1, _
                                   Optional ByVal lLimit As Long = -1, _
                                   Optional ByVal lCompare As VbCompareMethod _
                                                            = vbBinaryCompare) _
                                            As Long
    Const methodName As String = "CountSubstringUnlessEscaped"
    If lStart < 1 Then Err.Raise 5, methodName, _
        "Argument 'Start' = " & lStart & " < 1, invalid"
    If lLimit < -1 Then Err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"
        
    Dim lenSubStr As Long: lenSubStr = Len(subStr)
    Dim i As Long:         i = InStr(lStart, str, subStr, lCompare)

    CountSubstringUnlessEscaped = 0
    Do Until i = 0 Or lLimit = CountSubstringUnlessEscaped
        If StrComp(subStr, Mid$(str, i + lenSubStr, lenSubStr), lCompare) = 0 Then
            i = i + lenSubStr
        Else
            CountSubstringUnlessEscaped = CountSubstringUnlessEscaped + 1
        End If
        i = InStr(i + lenSubStr, str, subStr, lCompare)
    Loop
End Function

'Like CountSubstringUnlessEscaped but scans a string bytewise.
'Example illustrating the difference to CountSubstring:
'                       |c1||c2||c3||c4|
'bytes = HexToString("0x0061006100610061")
'                         |escape||ct|
'sFind =   HexToString("0x6100")
'CountSubstringUnlessEscaped(bytes, sFind) -> 0
'CountSubstringUnlessEscapedB(bytes, sFind) -> 1
Public Function CountSubstringUnlessEscapedB(ByRef bytes As String, _
                                             ByRef subStr As String, _
                                    Optional ByVal lStart As Long = 1, _
                                    Optional ByVal lLimit As Long = -1, _
                                    Optional ByVal lCompare As VbCompareMethod _
                                                            = vbBinaryCompare) _
                                             As Long
    Const methodName As String = "CountSubstringUnlessEscaped"
    If lStart < 1 Then Err.Raise 5, methodName, _
        "Argument 'Start' = " & lStart & " < 1, invalid"
    If lLimit < -1 Then Err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"

    Dim lenBSubStr As Long: lenBSubStr = LenB(subStr)
    Dim i As Long:          i = InStrB(lStart, bytes, subStr, lCompare)

    CountSubstringUnlessEscapedB = 0
    Do Until i = 0 Or lLimit = CountSubstringUnlessEscapedB
        If StrComp(subStr, MidB(bytes, i + lenBSubStr, lenBSubStr), _
                   lCompare) = 0 Then
            i = i + lenBSubStr
        Else
            CountSubstringUnlessEscapedB = CountSubstringUnlessEscapedB + 1
        End If
        i = InStrB(i + lenBSubStr, bytes, subStr, lCompare)
    Loop
End Function

'Works exactly like the inbuilt 'Replace', but is much, much faster on large
'strings with many replacements when vbBinaryCompare is used
Public Function ReplaceFast(ByRef str As String, _
                            ByRef sFind As String, _
                            ByRef sReplace As String, _
                   Optional ByVal lStart As Long = 1, _
                   Optional ByVal lCount As Long = -1, _
                   Optional ByVal lCompare As VbCompareMethod _
                                           = vbBinaryCompare) As String
    Const methodName As String = "ReplaceFast"
    If lStart < 1 Then Err.Raise 5, methodName, _
        "Argument 'lStart' = " & lStart & " < 1, invalid"
    If lCount < -1 Then Err.Raise 5, methodName, _
        "Argument 'lCount' = " & lCount & " < -1, invalid"
    lCount = lCount And &H7FFFFFFF
    
    If lCompare <> vbBinaryCompare Or Len(str) < 10000 Or lCount < 10000 Then
        'In the case of vbTextCompare, InStr's runtime will always be
        'proportional to the find position relative to the beginning of the
        'string, not relative to its 'Start' parameter, therefore, the algorithm
        'used in this function is not feasible and native 'Replace' is always
        'much faster in this case.
        'Also, blow 10k replacements, native Replace has a speed advantage
        ReplaceFast = Replace(str, sFind, sReplace, lStart, lCount, lCompare)
        Exit Function
    End If

    If Len(str) = 0 Or Len(sFind) = 0 Then
        ReplaceFast = Mid$(str, lStart)
        Exit Function
    End If

    Dim lenFind As Long:         lenFind = Len(sFind)
    Dim lenReplace As Long:      lenReplace = Len(sReplace)
    If lenFind = 0 Then Exit Function
    
    Static sFindPositions() As Long
    If (Not Not sFindPositions) = 0 Then ReDim sFindPositions(0 To 32767)
    Dim numFinds As Long
    Dim k As Long:         k = InStr(lStart, str, sFind, lCompare)
    
    On Error GoTo catch
    Do Until k = 0 Or lCount = numFinds
        sFindPositions(numFinds) = k
        numFinds = numFinds + 1
        k = InStr(k + lenFind, str, sFind, lCompare)
    Loop
    On Error GoTo 0
    GoTo continue
catch:
    ReDim Preserve sFindPositions(LBound(sFindPositions) To _
                                  UBound(sFindPositions) * 4)
    Resume
continue:
    Dim bufferSizeChange As Long
    bufferSizeChange = numFinds * (lenReplace - lenFind) - lStart + 1

    If Len(str) + bufferSizeChange < 1 Then Exit Function

    ReplaceFast = Space$(Len(str) + bufferSizeChange)

    Dim j As Long:              j = 1
    Dim lastOccurrence As Long: lastOccurrence = lStart
    Dim count As Long:          count = 1

    For k = 0 To numFinds - 1
        If count > lCount Then Exit For
        Dim i As Long:    i = sFindPositions(k)
        Dim diff As Long: diff = i - lastOccurrence
        If diff > 0 Then _
            Mid$(ReplaceFast, j, diff) = Mid$(str, lastOccurrence, diff)
        j = j + diff
        If lenReplace <> 0 Then
            Mid$(ReplaceFast, j, lenReplace) = sReplace
            j = j + lenReplace
        End If
        count = count + 1
        lastOccurrence = i + lenFind
    Next k
    If j <= Len(ReplaceFast) Then Mid$(ReplaceFast, j) = Mid$(str, lastOccurrence)
End Function

'Works like the inbuilt 'Replace', but parses the string bytewise, not charwise.
'This function uses the same algorithm used by ReplaceFast by default, therefore
'no ReplaceFastB exists or is required in this library.
'Example illustrating the difference:
'bytes = HexToString("0x00610061")
'sFind = HexToString("0x6100")
'? StringToHex(ReplaceB(bytes, sFind, "")) -> "0x0061"
'? StringToHex(Replace(bytes, sFind, "")) -> "0x00610061"
Public Function ReplaceB(ByRef bytes As String, _
                             ByRef sFind As String, _
                             ByRef sReplace As String, _
                    Optional ByVal lStart As Long = 1, _
                    Optional ByVal lCount As Long = -1, _
                    Optional ByVal lCompare As VbCompareMethod _
                                            = vbBinaryCompare) As String
    Const methodName As String = "ReplaceB"
    If lStart < 1 Then Err.Raise 5, methodName, _
        "Argument 'lStart' = " & lStart & " < 1, invalid"
    If lCount < -1 Then Err.Raise 5, methodName, _
        "Argument 'lCount' = " & lCount & " < -1, invalid"
    lCount = lCount And &H7FFFFFFF

    If LenB(bytes) = 0 Or LenB(sFind) = 0 Then
        ReplaceB = MidB$(bytes, lStart)
        Exit Function
    End If

    Dim lenBFind As Long:    lenBFind = LenB(sFind)
    Dim lenBReplace As Long: lenBReplace = LenB(sReplace)
    If lenBFind = 0 Then Exit Function
    
    Static sFindPositions() As Long
    If (Not Not sFindPositions) = 0 Then ReDim sFindPositions(0 To 32767)
    Dim numFinds As Long
    Dim k As Long:         k = InStrB(lStart, bytes, sFind, lCompare)
    
    On Error GoTo catch
    Do Until k = 0 Or lCount = numFinds
        sFindPositions(numFinds) = k
        numFinds = numFinds + 1
        k = InStrB(k + lenBFind, bytes, sFind, lCompare)
    Loop
    On Error GoTo 0
    GoTo continue
catch:
    ReDim Preserve sFindPositions(LBound(sFindPositions) To _
                                      UBound(sFindPositions) * 4)
    Resume
continue:
    Dim bufferSizeChange As Long
    bufferSizeChange = numFinds * (lenBReplace - lenBFind) - lStart + 1

    If LenB(bytes) + bufferSizeChange < 1 Then Exit Function
    
    'This array is 1-based to make calculation of 'bufferSizeChange' and the
    'prior if-statement identical to the algorithm in 'ReplaceFast'
    Dim buffer() As Byte:  ReDim buffer(1 To LenB(bytes) + bufferSizeChange)
    ReplaceB = buffer
    
    Dim j As Long:              j = 1
    Dim lastOccurrence As Long: lastOccurrence = lStart
    Dim count As Long:          count = 1

    For k = 0 To numFinds - 1
        If count > lCount Then Exit For
        Dim i As Long:    i = sFindPositions(k)
        Dim diff As Long: diff = i - lastOccurrence
        If diff > 0 Then _
            MidB$(ReplaceB, j, diff) = MidB$(bytes, lastOccurrence, diff)
        j = j + diff
        If lenBReplace <> 0 Then
            MidB$(ReplaceB, j, lenBReplace) = sReplace
            j = j + lenBReplace
        End If
        count = count + 1
        lastOccurrence = i + lenBFind
    Next k
    If j <= LenB(ReplaceB) Then MidB$(ReplaceB, j) = MidB$(bytes, lastOccurrence)
End Function

'Replaces consecutive occurrences of 'substring' that repeat more than 'limit'
'times with exactly 'limit' consecutive occurrences
'E.g.: LimitConsecutiveSubstringRepetition("aaaabaaac", "a", 1)  -> "abac"
'      LimitConsecutiveSubstringRepetition("aaaabaaac", "aa", 1) -> "aabaaac"
'      LimitConsecutiveSubstringRepetition("aaaabaaac", "a", 2)  -> "aabaac"
'      LimitConsecutiveSubstringRepetition("aaaabaaac", "ab", 0) -> "aaaaaac"
'Note that LimitConsecutiveSubstringRepetition(str, subStr, 0)
'is NOT the same as Replace(str, subStr, 0), e.g.:
'LimitConsecutiveSubstringRepetition("xaaaabbbby", "ab", 0) -> "xy"
'Note that due to the algorithm used by this function, in extremely rare cases,
'this function can produce different results than the naive solution of replace
'in a loop. This can only happen when the `subString` partially contains itself,
'for example, if it starts and ends with the same letter, but contains more than
'one different letter in total and is at least 3 letters long. This doesn't mean
'that this function doesn't work correctly, it simply means there is no
'objectively "correct" solution because in some cases different order of
'deletions of excess repetitions, of the subString we try to limit, can lead to
'different results, and in some cases this function will use a different order
'of deletion to avoid O(n^2) runtime complexity that could lead to indefinite
'freezing for some inputs.
'An example where the order of operations influences the final outcome can could
'look like this:
'str:="aabaaaababa", subStr:="aaba", limit:=0
'Depending on the order of deletion, the result could be either "aab", or "aba"
'Note that in this simple case LimitConsecutiveSubstringRepetition will procude
'the same result as the trivial Replace loop solution. Only some very weirdly
'constructed inputs can cause a different behavior.
'An example of an input that will cause a different result looks like this:
'str:="aaaaaaaababababaaaaabaababa", subStr:="aaba", limit:=0
'With everyday data such cases will most likely never occur.
Public Function LimitConsecutiveSubstringRepetition( _
                                           ByRef str As String, _
                                  Optional ByRef subStr As String = vbNewLine, _
                                  Optional ByVal limit As Long = 1, _
                                  Optional ByVal Compare As VbCompareMethod _
                                                          = vbBinaryCompare) _
                                           As String
    Const methodName As String = "LimitConsecutiveSubstringRepetition"
    Static recursionDepth As Long
    
    If limit < 0 Then Err.Raise 5, methodName, _
        "Argument 'limit' = " & limit & " < 0, invalid"

    Dim lenSubStr As Long:      lenSubStr = Len(subStr)
    Dim lenStr As Long:         lenStr = Len(str)
    
    LimitConsecutiveSubstringRepetition = str
    If lenStr = 0 Or lenSubStr = 0 Or lenStr < lenSubStr * (limit + 1) Then
        Exit Function
    End If
    
    'Generate an "above-limit sub-string" (subString repeated limit + 1 times):
    If lenSubStr = 1 Then
        Dim alSubStr As String: alSubStr = String$(limit + 1, subStr)
    Else
        alSubStr = Space$(lenSubStr * (limit + 1))
        Mid$(alSubStr, 1) = subStr
        If limit + 1 > 1 Then Mid$(alSubStr, lenSubStr + 1) = alSubStr
    End If
    Dim lenAlSubStr As Long:    lenAlSubStr = Len(alSubStr)
    
    'Normal algorithm, fastest for most cases
    Dim i As Long:              i = InStr(1, str, alSubStr, Compare)
    Dim j As Long:              j = 1
    Dim lastOccurrence As Long: lastOccurrence = 1 - lenSubStr
    Dim copyChunkSize As Long
    
    If i = 0 Then Exit Function

    Do Until i = 0
        i = i + lenSubStr * limit
        lastOccurrence = lastOccurrence + lenSubStr
        copyChunkSize = i - lastOccurrence
        Mid$(LimitConsecutiveSubstringRepetition, j, copyChunkSize) = _
            Mid$(str, lastOccurrence, copyChunkSize)
        j = j + copyChunkSize
        Do
            lastOccurrence = i
            i = InStr(lastOccurrence + lenSubStr, str, subStr, Compare)
        Loop Until i - lastOccurrence <> lenSubStr
        If i = 0 Then Exit Do
        If limit > 0 Then i = InStr(i, str, alSubStr, Compare)
    Loop
    copyChunkSize = lenStr - lastOccurrence - lenSubStr + 1
    Mid$(LimitConsecutiveSubstringRepetition, j, copyChunkSize) = _
        Mid$(str, lastOccurrence + lenSubStr)
    If j + copyChunkSize - 1 < Len(LimitConsecutiveSubstringRepetition) Then _
        LimitConsecutiveSubstringRepetition = _
            Left$(LimitConsecutiveSubstringRepetition, j + copyChunkSize - 1)

    'If normal algorithm was successful, the next loop will not be entered
    'This algorithm should be able to handle all other cases in O(n*Log(n)) time
    Do Until InStr(1, LimitConsecutiveSubstringRepetition, alSubStr, Compare) = 0
        Dim s As String: s = LimitConsecutiveSubstringRepetition
        lenStr = Len(s)
        If lenSubStr = 2 And limit = 0 _
        And StrComp(Left$(subStr, 1), Right$(subStr, 1), Compare) <> 0 Then
            'This algorithm should handle special cases like this:
            's = "aaabbb", subStr = "ab", limit = 0
            i = InStr(1, s, alSubStr, Compare)
            j = 1
            lastOccurrence = 1
            Dim leftChar As String:  leftChar = Left$(subStr, 1)
            Dim rightChar As String: rightChar = Right$(subStr, 1)
            Do Until i = 0
                Dim l As Long: l = i
                Dim r As Long: r = i + 1
                Do 'If statement inside the loop could be avoided using error
                    l = l - 1 'handling, but since that would only make the loop
                    r = r + 1 'about 3% faster, it is avoided for debugging
                    If l < 1 Then Exit Do 'convenience. It's an edge case anyway
                Loop Until StrComp(Mid$(s, l, 1), leftChar, Compare) <> 0 _
                        Or StrComp(Mid$(s, r, 1), rightChar, Compare) <> 0
                copyChunkSize = l + 1 - lastOccurrence
                If copyChunkSize > 0 Then _
                    Mid$(LimitConsecutiveSubstringRepetition, j, copyChunkSize) = _
                        Mid$(s, lastOccurrence, copyChunkSize)
                j = j + copyChunkSize
                lastOccurrence = r
                i = InStr(r, s, alSubStr, Compare)
            Loop
            copyChunkSize = lenStr - r + 1
            Mid$(LimitConsecutiveSubstringRepetition, j, copyChunkSize) = _
                        Mid$(s, lastOccurrence)
        Else
            Dim lSubStr As String:  lSubStr = Left$(alSubStr, lenSubStr * limit)
            Dim lenlSubStr As Long: lenlSubStr = Len(lSubStr) '= lenSubStr*limit
            Dim minL As Long:       minL = 1 'minL and maxR exist because we
            Dim maxR As Long 'must process the string strictly from left to
            i = InStr(1, s, alSubStr, Compare) 'right to copy the exact behavior
            j = 1 'of the naive solution. Existing alSubStrs must be prioritized
            lastOccurrence = 1 'Deal with chunks that cause runaway recursion:
            Do Until i = 0     's = "bababababaaaaaa", subStr = "baa", limit = 0
                Dim susChunk As String
                susChunk = Space$(lenAlSubStr * 2 - 2 + lenlSubStr)
                l = i                                'pos indices: l  r
                r = i + lenAlSubStr                   's:      ..babaaaa...
                maxR = InStr(r, s, alSubStr, Compare) - 1 '        |  |
                If maxR = -1 Then maxR = lenStr       'susChunk: ba> <aa -> baaa
                Dim lenLeft As Long, lenRight As Long
                Do
                    If l - lenSubStr + 1 < minL Then lenLeft = l - minL _
                                                Else lenLeft = lenSubStr - 1
                    If r + lenSubStr - 2 > maxR Then lenRight = maxR - r + 1 _
                                                Else: lenRight = lenSubStr - 1
                    If lenLeft + lenRight < lenSubStr Then Exit Do
                    Mid$(susChunk, 1, lenLeft) = Mid$(s, l - lenLeft, lenLeft)
                    If lenlSubStr > 0 Then _
                        Mid$(susChunk, lenLeft + 1, lenlSubStr) = lSubStr
                    Mid$(susChunk, lenLeft + lenlSubStr + 1, lenRight) = _
                        Mid$(s, r, lenRight)
                    susChunk = Left$(susChunk, lenLeft + lenRight + lenlSubStr)
                    Dim n As Long: n = InStr(1, susChunk, alSubStr, Compare)
                    If n = 0 Then Exit Do      'Here this holds: n <= lenRight
                    l = l + n - lenLeft - 1                     'r <= lenStr + 1
                    r = r + n + lenSubStr - 1 - lenLeft
                Loop
                copyChunkSize = l - lastOccurrence 'Edge handling is finished
                If copyChunkSize > 0 Then _
                    Mid$(LimitConsecutiveSubstringRepetition, j, copyChunkSize) = _
                        Mid$(s, lastOccurrence, copyChunkSize)
                j = j + copyChunkSize
                If limit > 0 Then
                    Mid$(LimitConsecutiveSubstringRepetition, j, lenlSubStr) = _
                        Left(alSubStr, lenlSubStr)
                    j = j + lenlSubStr
                    Mid$(s, r - lenlSubStr, lenlSubStr) = lSubStr
                End If
                minL = maxR + 1 'r - lenlSubStr
                lastOccurrence = r
                i = InStr(r - lenlSubStr, s, alSubStr, Compare)
            Loop
            copyChunkSize = lenStr - r + 1
            Mid$(LimitConsecutiveSubstringRepetition, j, copyChunkSize) = _
                        Mid$(s, lastOccurrence)
        End If
        
        'Copy the remaining chunk after any of the last two algorithms
        If j + copyChunkSize - 1 < Len(LimitConsecutiveSubstringRepetition) Then _
            LimitConsecutiveSubstringRepetition = _
                Left$(LimitConsecutiveSubstringRepetition, j + copyChunkSize - 1)
    Loop
End Function

'TODO: This function needs a rework to update the algorithm to the vesion
'LimitConsecutiveSubstringRepetition currently uses, this uses an old and
'bugged version
'Same as LimitConsecutiveSubstringRepetition, but scans the string bytewise.
'Example illustrating the difference:
'Dim bytes As String: bytes = HexToString("0x006100610061")
'Dim subStr As String: subStr = HexToString("0x6100")
'StringToHex(LimitConsecutiveSubstringRepetition(bytes, subStr, 1) _
'    -> "0x006100610061"
'StringToHex(LimitConsecutiveSubstringRepetitionB(bytes, subStr, 1) _
'    -> "0x00610061"
Public Function LimitConsecutiveSubstringRepetitionB( _
                                           ByRef bytes As String, _
                                  Optional ByRef subStr As String = vbNewLine, _
                                  Optional ByVal limit As Long = 1, _
                                  Optional ByVal Compare As VbCompareMethod _
                                                          = vbBinaryCompare) _
                                           As String
    Const methodName As String = "LimitConsecutiveSubstringRepetitionB"

    If limit < 0 Then Err.Raise 5, methodName, _
        "Argument 'limit' = " & limit & " < 0, invalid"
    If limit = 0 Then
        LimitConsecutiveSubstringRepetitionB = ReplaceB(bytes, subStr, _
                                                      vbNullString, , , Compare)
        Exit Function
    Else
        LimitConsecutiveSubstringRepetitionB = bytes
    End If
    If LenB(bytes) = 0 Then Exit Function
    If LenB(subStr) = 0 Then Exit Function

    Dim i As Long:                i = InStrB(1, bytes, subStr, Compare)
    Dim j As Long:                j = 1
    Dim lenBSubStr As Long:       lenBSubStr = LenB(subStr)
    Dim lastOccurrence As Long:   lastOccurrence = 1 - lenBSubStr
    Dim copyChunkSize As Long
    Dim consecutiveCount As Long
    Dim occurrenceDiff As Long

    Do Until i = 0
        occurrenceDiff = i - lastOccurrence
        If occurrenceDiff = lenBSubStr Then
            consecutiveCount = consecutiveCount + 1
            If consecutiveCount <= limit Then
                copyChunkSize = copyChunkSize + occurrenceDiff
            ElseIf consecutiveCount = limit + 1 Then
                MidB$(LimitConsecutiveSubstringRepetitionB, j, copyChunkSize) = _
                    MidB$(bytes, i - copyChunkSize, copyChunkSize)
                j = j + copyChunkSize
                copyChunkSize = 0
            End If
        Else
            copyChunkSize = copyChunkSize + occurrenceDiff
            consecutiveCount = 1
        End If
        lastOccurrence = i
        i = InStrB(i + lenBSubStr, bytes, subStr, Compare)
    Loop

    copyChunkSize = copyChunkSize + LenB(bytes) - lastOccurrence - lenBSubStr + 1
    MidB$(LimitConsecutiveSubstringRepetitionB, j, copyChunkSize) = _
        MidB$(bytes, LenB(bytes) - copyChunkSize + 1)

    LimitConsecutiveSubstringRepetitionB = _
        LeftB$(LimitConsecutiveSubstringRepetitionB, j + copyChunkSize - 1)
End Function

'Repeats the string str, repeatTimes times.
'Works with byte strings of uneven LenB
'E.g.: RepeatString("a", 3) -> "aaa"
'      StrConv(RepeatString(MidB("a", 1, 1), 3), vbUnicode) -> "aaa"
Public Function RepeatString(ByRef str As String, _
                    Optional ByVal repeatTimes As Long = 2) As String
    Const methodName As String = "RepeatString"
    If repeatTimes < 0 Then Err.Raise 5, methodName, _
        "Argument 'repeatTimes' = " & repeatTimes & " < 0, invalid"
    If repeatTimes = 0 Or LenB(str) = 0 Then
        Exit Function
    ElseIf LenB(str) = 2 Then
        RepeatString = String$(repeatTimes, str)
        Exit Function
    End If

    Dim newLength As Long: newLength = LenB(str) * repeatTimes
    RepeatString = Space$((newLength + 1) \ 2)
    If newLength Mod 2 = 1 Then RepeatString = MidB$(RepeatString, 2)
    
    MidB$(RepeatString, 1) = str
    If repeatTimes > 1 Then MidB$(RepeatString, LenB(str) + 1) = RepeatString
End Function

'Adds fillerStr to the right side of a string repeatedly until the resulting
'string reaches length 'Length'
'E.g.: PadRight("asd", 11, "xyz") -> "asdxyzxyzxy"
Public Function PadRight(ByRef str As String, _
                         ByVal Length As Long, _
                Optional ByVal fillerStr As String = " ") As String
    PadRight = PadRightB(str, Length * 2, fillerStr)
End Function

'Adds fillerStr to the left side of a string repeatedly until the resulting
'string reaches length 'Length'
'E.g.: PadLeft("asd", 11, "xyz") -> "yzxyzxyzasd"
Public Function PadLeft(ByRef str As String, _
                        ByVal Length As Long, _
               Optional ByVal fillerStr As String = " ") As String
    PadLeft = PadLeftB(str, Length * 2, fillerStr)
End Function

'Adds fillerStr to the right side of a string repeatedly until the resulting
'string reaches length 'Length' in bytes!
'E.g.: PadRightB("asd", 16, "xyz") -> "asdxyzxy"
Public Function PadRightB(ByRef str As String, _
                          ByVal Length As Long, _
                 Optional ByVal fillerStr As String = " ") As String
    Const methodName As String = "PadRightB"
    If Length < 0 Then Err.Raise 5, methodName, _
        "Argument 'Length' = " & Length & " < 0, invalid"
    If LenB(fillerStr) = 0 Then Err.Raise 5, methodName, _
        "Argument 'fillerStr' = vbNullString, invalid"

    If Length > LenB(str) Then
        If LenB(fillerStr) = 2 Then
            PadRightB = str & String((Length - LenB(str) + 1) \ 2, fillerStr)
            If Length Mod 2 = 1 Then _
                PadRightB = LeftB$(PadRightB, LenB(PadRightB) - 1)
        Else
            PadRightB = str & LeftB$(RepeatString(fillerStr, (((Length - _
                LenB(str))) + 1) \ LenB(fillerStr) + 1), Length - LenB(str))
        End If
    Else
        PadRightB = LeftB$(str, Length)
    End If
End Function

'Adds fillerStr to the left side of a string repeatedly until the resulting
'string reaches length 'Length' in bytes!
'Note that this can result in an invalid UTF-16 output for uneven lengths!
'E.g.: PadLeftB("asd", 16, "xyz") -> "yzxyzasd"
'      PadLeftB("asd", 11, "xyz") -> "?????"
Public Function PadLeftB(ByRef str As String, _
                         ByVal Length As Long, _
                Optional ByVal fillerStr As String = " ") As String
    Const methodName As String = "PadLeftB"
    If Length < 0 Then Err.Raise 5, methodName, _
        "Argument 'Length' = " & Length & " < 0, invalid"
    If LenB(fillerStr) = 0 Then Err.Raise 5, methodName, _
        "Argument 'fillerStr' = vbNullString, invalid"

    If Length > LenB(str) Then
        If LenB(fillerStr) = 2 Then
            PadLeftB = String((Length - LenB(str) + 1) \ 2, fillerStr) & str
            If Length Mod 2 = 1 Then _
                PadLeftB = RightB$(PadLeftB, LenB(PadLeftB) - 1)
        Else
            PadLeftB = RightB$(RepeatString(fillerStr, (((Length - LenB(str))) _
                          + 1) \ LenB(fillerStr) + 1), Length - LenB(str)) & str
        End If
    Else
        PadLeftB = RightB$(str, Length)
    End If
End Function

'Works like the inbuilt 'Split', but parses string bytewise, so it splits at
'all occurrences of 'Delimiter', even at uneven byte-index positions.
'Example illustrating the difference:
'bytes = HexToString("0x00610061")
'sDelim = HexToString("0x6100")
'SplitB(bytes, sDelim)) -> "0x00", "0x61"
'Split(bytes, sDelim, "")) -> "0x00610061"
Public Function SplitB(ByRef bytes As String, _
              Optional ByRef sDelimiter As String = " ", _
              Optional ByVal lLimit As Long = -1, _
              Optional ByVal lCompare As VbCompareMethod = vbBinaryCompare) _
                       As Variant
    Const methodName As String = "SplitB"
    If lLimit < -1 Then Err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"
    lLimit = lLimit And &H7FFFFFFF
    
    If lLimit = 0 Or bytes = vbNullString Then
        SplitB = Split("", , 0) 'Return empty but allocated string array
        Exit Function
    ElseIf LenB(bytes) = 0 Or LenB(sDelimiter) = 0 Or lLimit < 2 Then
        Dim arr() As String: ReDim arr(0 To 0)
        arr(0) = bytes
        SplitB = arr
        Exit Function
    End If

    Dim lenBDelim As Long:  lenBDelim = LenB(sDelimiter)
    ReDim arr(0 To CountSubstringB(bytes, sDelimiter, 1, lLimit, lCompare))
    Dim i As Long:              i = InStrB(1, bytes, sDelimiter, lCompare)
    Dim lastOccurrence As Long: lastOccurrence = 1
    Dim count As Long:          count = 0

    Do Until i = 0 Or count + 1 >= lLimit
        Dim diff As Long: diff = i - lastOccurrence
        arr(count) = MidB$(bytes, lastOccurrence, diff)
        count = count + 1
        lastOccurrence = i + lenBDelim
        i = InStrB(lastOccurrence, bytes, sDelimiter, lCompare)
    Loop
    arr(count) = MidB$(bytes, lastOccurrence)
    SplitB = arr
End Function

'Works like the inbuilt 'Split', but if delimiter is escaped (appears twice in
'a row) the string will not be split at that position and instead the double
'delimiter will be replaced by a single one
Public Function SplitUnlessEscaped(ByRef str As String, _
                          Optional ByRef sDelimiter As String = " ", _
                          Optional ByVal lLimit As Long = -1, _
                          Optional ByVal lCompare As VbCompareMethod = _
                                                     vbBinaryCompare) _
                                   As Variant
    Const methodName As String = "SplitUnlessEscaped"
    If lLimit < -1 Then Err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"
    lLimit = lLimit And &H7FFFFFFF

    If lLimit = 0 Then
        SplitUnlessEscaped = Split("", , 0) 'Return empty but allocated str arr
        Exit Function
    ElseIf Len(str) = 0 Or Len(sDelimiter) = 0 Or lLimit < 2 Then
        Dim arr() As String:  ReDim arr(0 To 0)
        arr(0) = str
        SplitUnlessEscaped = arr
        Exit Function
    End If
    
    Dim lenDelim As Long:   lenDelim = Len(sDelimiter)
    ReDim arr(0 To CountSubstringUnlessEscaped(str, sDelimiter, 1, lLimit, _
                                               lCompare))
    Dim partStart As Long:      partStart = 1
    Dim count As Long:          count = 0
    Dim lastOccurrence As Long: lastOccurrence = 1
    Dim i As Long:          i = InStr(lastOccurrence, str, sDelimiter, lCompare)
    
    Do Until i = 0 Or count + 1 >= lLimit
        If Mid$(str, i + lenDelim, lenDelim) = sDelimiter Then
            lastOccurrence = i + 2 * lenDelim
        Else
            arr(count) = Replace(Mid$(str, partStart, i - partStart), _
                                 sDelimiter & sDelimiter, sDelimiter)
            count = count + 1
            partStart = i + lenDelim
            lastOccurrence = partStart
        End If
        i = InStr(lastOccurrence, str, sDelimiter, lCompare)
    Loop
    
    If count < lLimit Then arr(count) = Replace(Mid$(str, partStart), _
                                            sDelimiter & sDelimiter, sDelimiter)
    SplitUnlessEscaped = arr
End Function

'Works like 'SplitB', but if delimiter is escaped (appears twice in
'a row) the string will not be split at that position and instead the double
'delimiter will be replaced by a single one
Public Function SplitUnlessEscapedB(ByRef bytes As String, _
                           Optional ByRef sDelimiter As String = " ", _
                           Optional ByVal lLimit As Long = -1, _
                           Optional ByVal lCompare As VbCompareMethod = _
                                                      vbBinaryCompare) _
                                    As Variant
    Const methodName As String = "SplitUnlessEscapedB"
    If lLimit < -1 Then Err.Raise 5, methodName, _
        "Argument 'lLimit' = " & lLimit & " < -1, invalid"
    lLimit = lLimit And &H7FFFFFFF

    If lLimit = 0 Then
        SplitUnlessEscapedB = Split("", , 0) 'Return empty but allocated str arr
        Exit Function
    ElseIf LenB(bytes) = 0 Or LenB(sDelimiter) = 0 Or lLimit < 2 Then
        Dim arr() As String:  ReDim arr(0 To 0)
        arr(0) = bytes
        SplitUnlessEscapedB = arr
        Exit Function
    End If
    
    Dim lenBDelim As Long:   lenBDelim = LenB(sDelimiter)
    ReDim arr(0 To CountSubstringUnlessEscapedB(bytes, sDelimiter, 1, lLimit, _
                                                lCompare))
    Dim partStart As Long:      partStart = 1
    Dim count As Long:          count = 0
    Dim lastOccurrence As Long: lastOccurrence = 1
    Dim i As Long:       i = InStrB(lastOccurrence, bytes, sDelimiter, lCompare)
    
    Do Until i = 0 Or count + 1 >= lLimit
        If MidB(bytes, i + lenBDelim, lenBDelim) = sDelimiter Then
            lastOccurrence = i + 2 * lenBDelim
        Else
            arr(count) = ReplaceB(MidB(bytes, partStart, i - partStart), _
                                  sDelimiter & sDelimiter, sDelimiter)
            count = count + 1
            partStart = i + lenBDelim
            lastOccurrence = partStart
        End If
        i = InStrB(lastOccurrence, bytes, sDelimiter, lCompare)
    Loop
    
    If count < lLimit Then arr(count) = ReplaceB(MidB(bytes, partStart), _
                                            sDelimiter & sDelimiter, sDelimiter)
    SplitUnlessEscapedB = arr
End Function

'Splits a string at every occurrence of the specified delimiter "delim", unless
'that delimiter occurs between non-escaped quotes. e.g. (" asf delim asdf ")
'will not be split. Quotes will not be removed.
'Quotes can be escaped by repetition.
'E.g.: SplitUnlessInQuotes("Hello "" ""World" "Goodbye World") returns
'      "Hello "" "" World", and "Goodbye World"
'If " is chosen as delimiter, splits at the outermost two occurrences of ", or
'if only one " exists in the string, splits the string into two parts.
'E.g. SplitUnlessInQuotes("asdf""asdf""asdf""asdf", """") returns
'    "asdf", "asdf""asdf", and "asdf"
Public Function SplitUnlessInQuotes(ByRef str As String, _
                           Optional ByRef delim As String = " ", _
                           Optional limit As Long = -1) As Variant
    Dim i As Long
    Dim s As String
    Dim ub As Long:         ub = -1
    Dim parts As Variant:   ReDim parts(0 To 0)
    Dim doSplit As Boolean: doSplit = True

    If delim = """" Then 'Handle this special case
        i = InStr(1, str, """", vbBinaryCompare)
        If i <> 0 Then
            Dim j As Long: j = InStrRev(str, """", , vbBinaryCompare)
            If i = j Then
                SplitUnlessInQuotes = Split(str, """", , vbBinaryCompare)
                Exit Function
            Else
                ReDim parts(0 To 2)
                parts(0) = Left$(str, i - 1)
                parts(1) = Mid$(str, i + 1, j)
                parts(2) = Mid$(str, j + 1)
            End If
        Else
            parts(0) = str
        End If
        SplitUnlessInQuotes = parts
        Exit Function
    End If

    For i = 1 To Len(str)
        If ub = limit - 2 Then
            ub = ub + 1
            ReDim Preserve parts(0 To ub)
            parts(ub) = Mid$(str, i)
            Exit For
        End If

        If Mid$(str, i, 1) = """" Then
            doSplit = Not doSplit
            If Not doSplit Then _
                doSplit = InStr(i + 1, str, """", vbBinaryCompare) = 0
        End If

        If Mid$(str, i, Len(delim)) = delim And doSplit Or i = Len(str) Then
            If i = Len(str) Then s = s & Mid$(str, i, 1)
            ub = ub + 1
            ReDim Preserve parts(0 To ub)
            parts(ub) = s
            s = vbNullString
            i = i + Len(delim) - 1
        Else
            s = s & Mid$(str, i, 1)
        End If
    Next i
    SplitUnlessInQuotes = parts
End Function

'Reads the memory of a String to an Array of Integers
'Notes:
'   - Ignores the last byte if input has an odd number of bytes
'   - If 'outLength' is -1 (default) then the remaining length is returned
'   - Excess length is ignored
Public Function StringToIntegers(ByRef s As String, _
                        Optional ByVal startIndex As Long = 1, _
                        Optional ByVal outLength As Long = -1, _
                        Optional ByVal outLowBound As Long = 0) As Integer()
    Static sArr As SAFEARRAY_1D
    Static memValue As Variant
    Static remoteVT As Variant
    Const methodName As String = "StringToIntegers"
    Dim cLen As Long: cLen = Len(s)

    If startIndex < 1 Or startIndex > cLen Then
        Err.Raise 9, methodName, "Invalid Start Index"
    ElseIf outLength < -1 Then
        Err.Raise 5, methodName, "Invalid Length for output"
    ElseIf outLength = -1 Or startIndex + outLength - 1 > cLen Then
        outLength = cLen - startIndex + 1
    End If
    If IsEmpty(memValue) Then
        remoteVT = VarPtr(memValue)
        CopyMemory remoteVT, vbInteger + VT_BYREF, 2
        With sArr
            .cDims = 1
            .fFeatures = FADF_HAVEVARTYPE
            .cbElements = INT_SIZE
        End With
        memValue = VarPtr(sArr)
    End If
    With sArr
        .pvData = StrPtr(s) + (startIndex - 1) * INT_SIZE
        .rgsabound0.lLbound = outLowBound
        .rgsabound0.cElements = outLength
    End With
    RemoteAssign remoteVT, vbArray + vbInteger, StringToIntegers, memValue
End Function

'This method assures the required redirection for both the remote varType and
'   the remote value at the same time thus removing any additional stack frames
'It can be used to both read from and write to memory by swapping the order of
'   the last 2 parameters
Private Sub RemoteAssign(ByRef remoteVT As Variant, _
                         ByVal newVT As VbVarType, _
                         ByRef targetVariable As Variant, _
                         ByRef newValue As Variant)
    remoteVT = newVT
    targetVariable = newValue
    remoteVT = vbLongPtr 'Stop linking to remote address, for safety
End Sub

'Reads the memory of an Array of Integers into a String
'Notes:
'   - If 'outLength' is -1 (default) then the remaining length is returned
'   - Excess length is ignored
Public Function IntegersToString(ByRef ints() As Integer, _
                        Optional ByVal startIndex As Long = 0, _
                        Optional ByVal outLength As Long = -1) As String
    Static sArr As SAFEARRAY_1D
    Static memValue As Variant
    Static remoteVT As Variant
    Const methodName As String = "IntegersToString"

    If GetArrayDimsCount(ints) <> 1 Then
        Err.Raise 5, methodName, "Expected 1D Array of Integers"
    ElseIf startIndex < LBound(ints) Or startIndex > UBound(ints) Then
        Err.Raise 9, methodName, "Invalid Start Index"
    ElseIf outLength < -1 Then
        Err.Raise 5, methodName, "Invalid Length for output"
    ElseIf outLength = -1 Or startIndex + outLength - 1 > UBound(ints) Then
        outLength = UBound(ints) - startIndex + 1
    End If
    If IsEmpty(memValue) Then
        remoteVT = VarPtr(memValue)
        CopyMemory remoteVT, vbInteger + VT_BYREF, 2
        With sArr
            .cDims = 1
            .fFeatures = FADF_HAVEVARTYPE
            .cbElements = BYTE_SIZE
            .rgsabound0.lLbound = 0
        End With
        memValue = VarPtr(sArr)
    End If
    With sArr
        .pvData = VarPtr(ints(startIndex))
        .rgsabound0.cElements = outLength * INT_SIZE
    End With
    RemoteAssign remoteVT, vbArray + vbByte, IntegersToString, memValue
End Function

'Returns the Number of dimensions for an input array
'Returns 0 if array is uninitialized or input not an array
'Note that a zero-length array has 1 dimension! Ex. Array() bounds are (0 to -1)
Private Function GetArrayDimsCount(ByRef arr As Variant) As Long
    Const MAX_DIMENSION As Long = 60 'VB limit
    Dim dimension As Long
    Dim tempBound As Long

    On Error GoTo FinalDimension
    For dimension = 1 To MAX_DIMENSION
        tempBound = LBound(arr, dimension)
    Next dimension
FinalDimension:
    GetArrayDimsCount = dimension - 1
End Function

'Works like the inbuilt `Replace` but also accepts arrays as 'sFindOrFinds' or
''sReplaceOrReplaces'. The arrays for the find and replace values do not need
'to have the same number of elements, if they differ in the number of elements
'the following logic will be used for the replacing:
'Let sFindOrFinds contain n elements, and sReplaceOrReplaces contain m elements
'and be 0 based arrays.
'(The function converts non 0 based input arrays to 0 based arrays)
'Case 1: n > m
'   sFindOrFinds(i) gets replaced by sReplaceOrReplaces(i Mod m)
'   E.g.: sFindOrFinds = ("1", "2", "3"), sReplaceOrReplaces = ("4", "5")
'         ReplaceMultiple("123123123", ...) returns "454454454"
'         -> Each "1" and "3" get replaced by "4" while "2" gets replaced by "5"
'
'Case 2: n < m
'   E.g.: sFindOrFinds = ("1", "2"), sReplaceOrReplaces = ("3", "4", "5")
'         Every odd numbered occurrence of "1" gets replaced by "3", every even
'         numbered occurrence of "1" gets replaced by "5" and every occurrence
'         of "2" gets replaced by "4"
'         ReplaceMultiple("123123123", ...) returns "343543343"
'Notes:
'All replacements will be performed in a single pass. That means if an element
'of 'sFindOrFinds' with a greater index contains any of the prior elements in
''sFindOrFinds' as a substring, it won't be replaced because of the order of
'precedence, which goes from first = highest to last = lowest.
'Example:
'ReplaceMultiple("HelloHello", Array("ell", "Hello"), Array("ell", "World"))
'   Here "Hello" will not be found and replaced because it contains "ell" as a
'   substring and "ell" has higher priority because if comes earlier in
'   'sFindOrFinds'. The output of this code is therefore "HelloHello"
''ReplaceMultipleMultiPass' is generally faster than this function for smaller
'   input strings (< 50k characters) or a large amount of 'finds' and 'replaces'
'   (> 1000 of each) However, it
'   does not support the n < m case and also the behavior can differ and be less
'   predictable, if for example finds appear in the string through prior
'   replacements that can then get replaced again in the next iteration.
'This function should not be used with more than a few thousand finds at once
'   because the runtime is proportional to n^2 because of an important check
'   as commented in the code!
Public Function ReplaceMultiple(ByRef str As String, _
                                ByRef sFindOrFinds As Variant, _
                                ByRef sReplaceOrReplaces As Variant, _
                       Optional ByVal lStart As Long = 1, _
                       Optional ByVal lCount As Long = -1, _
                       Optional ByVal lCompare As VbCompareMethod _
                                                = vbBinaryCompare) As String
    Const methodName As String = "ReplaceMultiple"
    If lStart < 1 Then Err.Raise 5, methodName, _
                               "Argument 'lStart' = " & lStart & " < 1, invalid"
    If lCount < -1 Then Err.Raise 5, methodName, _
                              "Argument 'lCount' = " & lCount & " < -1, invalid"
    
    Dim finds As Variant
    If IsArray(sFindOrFinds) Then finds = sFindOrFinds _
                             Else finds = VBA.Array(sFindOrFinds)
    If Not LBound(finds) = 0 Then _
        ReDim Preserve finds(0 To UBound(finds) - LBound(finds))
    Dim replaces As Variant
    If IsArray(sReplaceOrReplaces) Then replaces = sReplaceOrReplaces _
                                   Else replaces = VBA.Array(sReplaceOrReplaces)
    If Not LBound(replaces) = 0 Then _
        ReDim Preserve replaces(0 To UBound(replaces) - LBound(replaces))
    
    Dim i As Long
    On Error GoTo catch
    For i = LBound(finds) To UBound(finds)
        finds(i) = CStr(finds(i))
    Next i
    For i = LBound(replaces) To UBound(replaces)
        replaces(i) = CStr(replaces(i))
    Next i
    On Error GoTo 0: GoTo continue
catch:
    Err.Raise 5, methodName, "Argument 'sFindOrFinds' or 'sReplaceOrReplaces'" _
        & " contains invalid elements that can't be converted to String type."
continue:
    If lStart > Len(str) Then Exit Function
    
    lCount = lCount And &H7FFFFFFF
    
    If UBound(finds) = 0 And Len(finds(0)) = 0 _
    And UBound(replaces) = 0 And Len(replaces(0)) = 0 Then
        ReplaceMultiple = Mid$(str, lStart)
        Exit Function
    End If
    
    'Clean input arrays to deal with cases where one "find" contains another one
    'Unfortunately, this part of the algorithm introduces an O(n^2 * m)
    'complexity (n = number of finds, m average length of finds) which can make
    'this function very slow for more than a few 1000 finds. This can
    'theoretically be improved by using a Trie instead in certain cases to avoid
    'n^2 runtime complexity: https://en.wikipedia.org/wiki/Trie
    'A simple implementation of the trie algorithm has been tested, the
    'procedure is available in the test module ('ProcessFindsUsingTrie')
    'Unfortunately it performs at least 20 times slower than the naïve approach
    'implemented here:
    Dim j As Long
    For i = 0 To UBound(finds)
        If Len(finds(i)) <> 0 Then
            For j = i + 1 To UBound(finds)
                If InStr(1, finds(j), finds(i), lCompare) <> 0 Then
                    finds(j) = vbNullString
                End If
            Next j
        End If
    Next i
    
    'ProcessFindsUsingTrie finds, lCompare '<-- at least 20 times slower

    'Allocate buffer
    'Buffer calculation doesn't really take into account the parameter lCount.
    'We'll take it into account at the end of the function instead
    Dim n As Long:          n = UBound(finds) + 1
    Dim m As Long:          m = UBound(replaces) + 1
    Dim lenBBuffer As Long: lenBBuffer = LenB(str) - (lStart - 1) * 2
    If m > n Then
        For i = 0 To UBound(finds)
            Dim numReplPerFind As Long
            numReplPerFind = IIf(i < m Mod n, (m \ n) + 1, (m \ n))
            Dim numOcc As Long
            numOcc = CountSubstring(str, CStr(finds(i)), lStart, lCount, _
                                    lCompare)
            For j = i To m - 1 Step n
                lenBBuffer = lenBBuffer + (LenB(replaces(j)) - LenB(finds(i))) _
                             * IIf((j - i) \ n < numOcc Mod numReplPerFind _
                         , numOcc \ numReplPerFind + 1, numOcc \ numReplPerFind)
            Next j
        Next i
    Else
        For i = 0 To UBound(finds)
            numOcc = CountSubstring(str, CStr(finds(i)), lStart, lCount, _
                                    lCompare)
            lenBBuffer = lenBBuffer + _
                         (LenB(replaces(i Mod m)) - LenB(finds(i))) * numOcc
        Next i
    End If
    
    If lenBBuffer = 0 Then Exit Function
    
    Dim buffer() As Byte: ReDim buffer(1 To lenBBuffer)
    ReplaceMultiple = buffer
    
    'Keep track of the next position for replacing by using a min-heap
    Dim nextOccsHeap() As Long: ReDim nextOccsHeap(0 To UBound(finds), 0 To 2)
    Dim index1 As Long
    Dim index2 As Long
    Dim insertElement(0 To 2) As Long
    Dim heapSize As Long
    
    Dim lenBReplaces() As Long: ReDim lenBReplaces(0 To UBound(replaces))
    For i = 0 To UBound(replaces)
        lenBReplaces(i) = LenB(replaces(i))
    Next i
    Dim lenBFinds() As Long: ReDim lenBFinds(0 To UBound(finds))
    For i = 0 To UBound(finds)
        lenBFinds(i) = LenB(finds(i))
    Next i
    
    For i = 0 To UBound(finds)
        Dim nextOcc As Long
        If lenBFinds(i) < 2 Then
            nextOcc = 0
        Else
            nextOcc = InStr(lStart, str, finds(i), lCompare) * 2 - 1
        End If
        If nextOcc > 0 Then
            insertElement(0) = nextOcc
            insertElement(1) = i
            insertElement(2) = i Mod m
            GoSub HeapInsert
        End If
    Next i
    
    Dim lastOccurrence As Long: lastOccurrence = lStart * 2 - 1
    Dim currOccurrence As Long: currOccurrence = nextOccsHeap(0, 0)
    Dim currReplaceIdx As Long: currReplaceIdx = nextOccsHeap(0, 2)
    Dim builtStrPos As Long:    builtStrPos = 1
    Dim count As Long:          count = 1
    
    'Do the replacing using InStr() and Mid$ in a loop
    Do Until heapSize = 0 Or count > lCount
        Dim diff As Long: diff = currOccurrence - lastOccurrence
        If diff > 0 Then _
            MidB$(ReplaceMultiple, builtStrPos, diff) = _
                MidB$(str, lastOccurrence, diff)
        builtStrPos = builtStrPos + diff
        If lenBReplaces(currReplaceIdx) <> 0 Then
            MidB$(ReplaceMultiple, builtStrPos, lenBReplaces(currReplaceIdx)) = _
                replaces(currReplaceIdx)
            builtStrPos = builtStrPos + lenBReplaces(currReplaceIdx)
        End If
        count = count + 1
        lastOccurrence = currOccurrence + lenBFinds(nextOccsHeap(0, 1))
        If lenBFinds(nextOccsHeap(0, 1)) < 2 Then
            insertElement(0) = 0
        Else
            insertElement(0) = InStr((lastOccurrence + 1) \ 2, str, _
                                    finds(nextOccsHeap(0, 1)), lCompare) * 2 - 1
        End If
        insertElement(1) = nextOccsHeap(0, 1)
        If m > n Then
            currReplaceIdx = currReplaceIdx + n
            If currReplaceIdx >= m Then currReplaceIdx = nextOccsHeap(0, 1) Mod m
        End If
        insertElement(2) = currReplaceIdx
        GoSub HeapRemoveMin
        If insertElement(0) > 0 Then GoSub HeapInsert
        currOccurrence = nextOccsHeap(0, 0)
        currReplaceIdx = nextOccsHeap(0, 2)
    Loop
    
    Dim remainderStr As String: remainderStr = MidB$(str, lastOccurrence)

    If builtStrPos <= LenB(ReplaceMultiple) Then _
        MidB$(ReplaceMultiple, builtStrPos, LenB(remainderStr)) = remainderStr
        
    'Because we didn't take lCount into account when calculating the buffer:
    If count > lCount Then _
        ReplaceMultiple = _
            LeftB$(ReplaceMultiple, builtStrPos + LenB(remainderStr))
'    Dim numReplacements As Long
'    Debug.Print numReplacements
    Exit Function
HeapSwapElements:
    Dim temp(0 To 2) As Long
    temp(0) = nextOccsHeap(index1, 0)
    temp(1) = nextOccsHeap(index1, 1)
    temp(2) = nextOccsHeap(index1, 2)
    nextOccsHeap(index1, 0) = nextOccsHeap(index2, 0)
    nextOccsHeap(index1, 1) = nextOccsHeap(index2, 1)
    nextOccsHeap(index1, 2) = nextOccsHeap(index2, 2)
    nextOccsHeap(index2, 0) = temp(0)
    nextOccsHeap(index2, 1) = temp(1)
    nextOccsHeap(index2, 2) = temp(2)
    Return
HeapInsert:
    'numReplacements = numReplacements + 1
    nextOccsHeap(heapSize, 0) = insertElement(0)
    nextOccsHeap(heapSize, 1) = insertElement(1)
    nextOccsHeap(heapSize, 2) = insertElement(2)
    Dim currentIndex As Long: currentIndex = heapSize
    Do While currentIndex > 0 _
         And nextOccsHeap(currentIndex, 0) < _
             nextOccsHeap((currentIndex - 1) \ 2, 0) 'ParentNode
        index1 = currentIndex
        index2 = (currentIndex - 1) \ 2 'ParentIndex
        GoSub HeapSwapElements
        currentIndex = (currentIndex - 1) \ 2 'ParentIndex
    Loop
    heapSize = heapSize + 1
    Return
HeapRemoveMin:
    nextOccsHeap(0, 0) = nextOccsHeap(heapSize - 1, 0)
    nextOccsHeap(0, 1) = nextOccsHeap(heapSize - 1, 1)
    nextOccsHeap(0, 2) = nextOccsHeap(heapSize - 1, 2)
    heapSize = heapSize - 1
    
    currentIndex = 0
    Do
        Dim iLeft As Long: iLeft = 2 * currentIndex + 1 'LeftChildIndex
        Dim iRight As Long: iRight = 2 * currentIndex + 2 'RightChildIndex
        Dim smallest As Long: smallest = currentIndex
        
        If iLeft < heapSize Then
            If nextOccsHeap(iLeft, 0) < nextOccsHeap(smallest, 0) Then _
                smallest = iLeft
        End If
        If iRight < heapSize Then
            If nextOccsHeap(iRight, 0) < nextOccsHeap(smallest, 0) Then _
                smallest = iRight
        End If
        
        If smallest <> currentIndex Then
            index1 = currentIndex
            index2 = smallest
            GoSub HeapSwapElements
            currentIndex = smallest
        Else
            Exit Do
        End If
    Loop
    Return
End Function

'Same as ReplaceMultiple but parses the string bytewise
Public Function ReplaceMultipleB(ByRef bytes As String, _
                                 ByRef sFindOrFinds As Variant, _
                                 ByRef sReplaceOrReplaces As Variant, _
                        Optional ByVal lStart As Long = 1, _
                        Optional ByVal lCount As Long = -1, _
                        Optional ByVal lCompare As VbCompareMethod _
                                                 = vbBinaryCompare) As String
    Const methodName As String = "ReplaceMultipleB"
    If lStart < 1 Then Err.Raise 5, methodName, _
                               "Argument 'lStart' = " & lStart & " < 1, invalid"
    If lCount < -1 Then Err.Raise 5, methodName, _
                              "Argument 'lCount' = " & lCount & " < -1, invalid"
    Dim finds As Variant
    If IsArray(sFindOrFinds) Then finds = sFindOrFinds _
                             Else finds = VBA.Array(sFindOrFinds)
    If Not LBound(finds) = 0 Then _
        ReDim Preserve finds(0 To UBound(finds) - LBound(finds))
    Dim replaces As Variant
    If IsArray(sReplaceOrReplaces) Then replaces = sReplaceOrReplaces _
                                   Else replaces = VBA.Array(sReplaceOrReplaces)
    If Not LBound(replaces) = 0 Then _
        ReDim Preserve replaces(0 To UBound(replaces) - LBound(replaces))
    
    Dim i As Long
    On Error GoTo catch
    For i = LBound(finds) To UBound(finds)
        finds(i) = CStr(finds(i))
    Next i
    For i = LBound(replaces) To UBound(replaces)
        replaces(i) = CStr(replaces(i))
    Next i
    On Error GoTo 0: GoTo continue
catch:
    Err.Raise 5, methodName, "Argument 'sFindOrFinds' or 'sReplaceOrReplaces'" _
        & " contains invalid elements that can't be converted to String type."
continue:
    If lStart > LenB(bytes) Then Exit Function
    
    lCount = lCount And &H7FFFFFFF
    
    If UBound(finds) = 0 And Len(finds(0)) = 0 _
    And UBound(replaces) = 0 And Len(replaces(0)) = 0 Then
        ReplaceMultipleB = MidB$(bytes, lStart)
        Exit Function
    End If
    
    'Clean input arrays to deal with cases where one "find" contains another one
    Dim j As Long
    For i = 0 To UBound(finds)
        If LenB(finds(i)) <> 0 Then
            For j = i + 1 To UBound(finds)
                If InStrB(1, finds(j), finds(i), lCompare) <> 0 Then
                    finds(j) = vbNullString
                End If
            Next j
        End If
    Next i
    
    'Allocate buffer
    'Buffer calculation doesn't really take into account the parameter lCount.
    'We'll take it into account at the end of the function instead
    Dim n As Long:          n = UBound(finds) + 1
    Dim m As Long:          m = UBound(replaces) + 1
    Dim lenBBuffer As Long: lenBBuffer = LenB(bytes) - (lStart - 1)
    If m > n Then
        For i = 0 To UBound(finds)
            Dim numReplPerFind As Long
            numReplPerFind = IIf(i < m Mod n, (m \ n) + 1, (m \ n))
            Dim numOcc As Long
            numOcc = CountSubstringB(bytes, CStr(finds(i)), lStart, lCount, _
                                     lCompare)
            For j = i To m - 1 Step n
                lenBBuffer = lenBBuffer + (LenB(replaces(j)) - LenB(finds(i))) _
                             * IIf((j - i) \ n < numOcc Mod numReplPerFind _
                         , numOcc \ numReplPerFind + 1, numOcc \ numReplPerFind)
            Next j
        Next i
    Else
        For i = 0 To UBound(finds)
            numOcc = CountSubstringB(bytes, CStr(finds(i)), lStart, lCount, _
                                     lCompare)
            lenBBuffer = lenBBuffer + _
                         (LenB(replaces(i Mod m)) - LenB(finds(i))) * numOcc
        Next i
    End If
    
    If lenBBuffer = 0 Then Exit Function
    
    Dim buffer() As Byte: ReDim buffer(1 To lenBBuffer)
    ReplaceMultipleB = buffer
    
    'Keep track of the next position for replacing by using a min-heap
    Dim nextOccsHeap() As Long: ReDim nextOccsHeap(0 To UBound(finds), 0 To 2)
    Dim index1 As Long
    Dim index2 As Long
    Dim insertElement(0 To 2) As Long
    Dim heapSize As Long
    
    Dim lenBReplaces() As Long: ReDim lenBReplaces(0 To UBound(replaces))
    For i = 0 To UBound(replaces)
        lenBReplaces(i) = LenB(replaces(i))
    Next i
    Dim lenBFinds() As Long: ReDim lenBFinds(0 To UBound(finds))
    For i = 0 To UBound(finds)
        lenBFinds(i) = LenB(finds(i))
    Next i
    
    For i = 0 To UBound(finds)
        Dim nextOcc As Long
        If lenBFinds(i) = 0 Then
            nextOcc = 0
        Else
            nextOcc = InStr(lStart, bytes, finds(i), lCompare) * 2 - 1
        End If
        If nextOcc > 0 Then
            insertElement(0) = nextOcc
            insertElement(1) = i
            insertElement(2) = i Mod m
            GoSub HeapInsert
        End If
    Next i
    
    Dim lastOccurrence As Long: lastOccurrence = lStart
    Dim currOccurrence As Long: currOccurrence = nextOccsHeap(0, 0)
    Dim currReplaceIdx As Long: currReplaceIdx = nextOccsHeap(0, 2)
    Dim builtStrPos As Long:    builtStrPos = 1
    Dim count As Long:          count = 1
    
    'Do the replacing using InStrB() and MidB$ in a loop
    Do Until heapSize = 0 Or count > lCount
        Dim diff As Long: diff = currOccurrence - lastOccurrence
        If diff > 0 Then _
            MidB$(ReplaceMultipleB, builtStrPos, diff) = _
                MidB$(bytes, lastOccurrence, diff)
        builtStrPos = builtStrPos + diff
        If lenBReplaces(currReplaceIdx) <> 0 Then
            MidB$(ReplaceMultipleB, builtStrPos, lenBReplaces(currReplaceIdx)) = _
                replaces(currReplaceIdx)
            builtStrPos = builtStrPos + lenBReplaces(currReplaceIdx)
        End If
        count = count + 1
        lastOccurrence = currOccurrence + lenBFinds(nextOccsHeap(0, 1))
        If lenBFinds(nextOccsHeap(0, 1)) = 0 Then
            insertElement(0) = 0
        Else
            insertElement(0) = InStrB(lastOccurrence, bytes, _
                                      finds(nextOccsHeap(0, 1)), lCompare)
        End If
        insertElement(1) = nextOccsHeap(0, 1)
        If m > n Then
            currReplaceIdx = currReplaceIdx + n
            If currReplaceIdx >= m Then currReplaceIdx = nextOccsHeap(0, 1) Mod m
        End If
        insertElement(2) = currReplaceIdx
        GoSub HeapRemoveMin
        If insertElement(0) > 0 Then GoSub HeapInsert
        currOccurrence = nextOccsHeap(0, 0)
        currReplaceIdx = nextOccsHeap(0, 2)
    Loop

    Dim remainderStr As String: remainderStr = MidB$(bytes, lastOccurrence)
    
    If builtStrPos <= LenB(ReplaceMultipleB) Then _
        MidB$(ReplaceMultipleB, builtStrPos, LenB(remainderStr)) = remainderStr
        
    'Because we didn't take lCount into account when calculating the buffer:
    If count > lCount Then _
        ReplaceMultipleB = _
            LeftB$(ReplaceMultipleB, builtStrPos + LenB(remainderStr))
       
    Exit Function
HeapSwapElements:
    Dim temp(0 To 2) As Long
    temp(0) = nextOccsHeap(index1, 0)
    temp(1) = nextOccsHeap(index1, 1)
    temp(2) = nextOccsHeap(index1, 2)
    nextOccsHeap(index1, 0) = nextOccsHeap(index2, 0)
    nextOccsHeap(index1, 1) = nextOccsHeap(index2, 1)
    nextOccsHeap(index1, 2) = nextOccsHeap(index2, 2)
    nextOccsHeap(index2, 0) = temp(0)
    nextOccsHeap(index2, 1) = temp(1)
    nextOccsHeap(index2, 2) = temp(2)
    Return
HeapInsert:
    nextOccsHeap(heapSize, 0) = insertElement(0)
    nextOccsHeap(heapSize, 1) = insertElement(1)
    nextOccsHeap(heapSize, 2) = insertElement(2)
    Dim currentIndex As Long: currentIndex = heapSize
    Do While currentIndex > 0 _
         And nextOccsHeap(currentIndex, 0) < _
             nextOccsHeap((currentIndex - 1) \ 2, 0) 'ParentNode
        index1 = currentIndex
        index2 = (currentIndex - 1) \ 2 'ParentIndex
        GoSub HeapSwapElements
        currentIndex = (currentIndex - 1) \ 2 'ParentIndex
    Loop
    heapSize = heapSize + 1
    Return
HeapRemoveMin:
    nextOccsHeap(0, 0) = nextOccsHeap(heapSize - 1, 0)
    nextOccsHeap(0, 1) = nextOccsHeap(heapSize - 1, 1)
    nextOccsHeap(0, 2) = nextOccsHeap(heapSize - 1, 2)
    heapSize = heapSize - 1
    
    currentIndex = 0
    Do
        Dim iLeft As Long: iLeft = 2 * currentIndex + 1 'LeftChildIndex
        Dim iRight As Long: iRight = 2 * currentIndex + 2 'RightChildIndex
        Dim smallest As Long: smallest = currentIndex
        
        If iLeft < heapSize Then
            If nextOccsHeap(iLeft, 0) < nextOccsHeap(smallest, 0) Then _
                smallest = iLeft
        End If
        If iRight < heapSize Then
            If nextOccsHeap(iRight, 0) < nextOccsHeap(smallest, 0) Then _
                smallest = iRight
        End If
        
        If smallest <> currentIndex Then
            index1 = currentIndex
            index2 = smallest
            GoSub HeapSwapElements
            currentIndex = smallest
        Else
            Exit Do
        End If
    Loop
    Return
End Function

'Similar to ReplaceMultiple, but here, replacements will be performed one after
'another in the order of the input arrays and not all in a single pass like in
'the regular ReplaceMultiple.
'In this version lCount specifies the maximum number of replacements per find
'value, and not in total like in the regular ReplaceMultiple.
'Also, this version does not support multiple different replace values for a
'given find value, this means numFinds must be >= numReplaces.
Public Function ReplaceMultipleMultiPass(ByRef str As String, _
                                         ByRef sFindOrFinds As Variant, _
                                         ByRef sReplaceOrReplaces As Variant, _
                                Optional ByVal lStart As Long = 1, _
                                Optional ByVal lCount As Long = -1, _
                                Optional ByVal lCompare As VbCompareMethod _
                                                         = vbBinaryCompare) _
                                         As String
    Const methodName As String = "ReplaceMultipleMultiPass"
    If lStart < 1 Then Err.Raise 5, methodName, _
                               "Argument 'lStart' = " & lStart & " < 1, invalid"
    If lCount < -1 Then Err.Raise 5, methodName, _
                              "Argument 'lCount' = " & lCount & " < -1, invalid"
    
    Dim finds As Variant
    If IsArray(sFindOrFinds) Then finds = sFindOrFinds _
                             Else finds = VBA.Array(sFindOrFinds)
    If Not LBound(finds) = 0 Then _
        ReDim Preserve finds(0 To UBound(finds) - LBound(finds))
    Dim replaces As Variant
    If IsArray(sReplaceOrReplaces) Then replaces = sReplaceOrReplaces _
                                   Else replaces = VBA.Array(sReplaceOrReplaces)
    If Not LBound(replaces) = 0 Then _
        ReDim Preserve replaces(0 To UBound(replaces) - LBound(replaces))
        
    If UBound(replaces) > UBound(finds) Then Err.Raise 5, methodName, "'sFind" _
        & "OrFinds' must have at least as many elements as 'sReplaceOrReplaces'"
    
    If lStart > Len(str) Then Exit Function
    
    ReplaceMultipleMultiPass = Mid$(str, lStart)
    If UBound(finds) = 0 And Len(finds(0)) = 0 _
    And UBound(replaces) = 0 And Len(replaces(0)) = 0 Then Exit Function
    
    Dim numReplaces As Long: numReplaces = UBound(replaces) + 1
    
    Dim i As Long
    For i = 0 To UBound(finds)
        ReplaceMultipleMultiPass = ReplaceFast(ReplaceMultipleMultiPass, _
                  (finds(i)), (replaces(i Mod numReplaces)), , lCount, lCompare)
    Next i
End Function

'This function can replace multiple values with multiple different replace
'values in each element of an array or just in a regular string.
'E.g.: ArrayReplaceMultiple("ab", Array("a", "b"), Array("c", "d")) returns "cd"
'Or: ArrayReplaceMultiple(Array("ab", "ab"), Array("a", "b"), Array("c", "d"))
'returns an array with two elements: ("cd", "cd")
Public Function ArrayReplaceMultiple(ByVal strOrStrArr As Variant, _
                                     ByVal findOrFinds As Variant, _
                                     ByVal replaceOrReplaces As Variant, _
                            Optional ByVal compareMethod As VbCompareMethod _
                                                          = vbBinaryCompare) _
                                     As Variant
    Const methodName As String = "ReplaceMultiple"
    If IsArray(findOrFinds) Then If Not IsArray(replaceOrReplaces) Then _
        Err.Raise 5, methodName, "Finds and Replaces must both be array or not."
    If IsArray(findOrFinds) Then
        If Not UBound(findOrFinds) - LBound(findOrFinds) = _
               UBound(replaceOrReplaces) - LBound(replaceOrReplaces) Then
            Err.Raise 5, methodName, _
                "There must be the same number of find and replace values"
        End If
    Else
        Dim tmpArr As Variant: ReDim tmpArr(0 To 0)
        tmpArr(0) = findOrFinds
        findOrFinds = tmpArr
        tmpArr(0) = replaceOrReplaces
        replaceOrReplaces = tmpArr
    End If
    Dim i As Long, j As Long, k As Long
    If IsArray(strOrStrArr) Then
        For i = LBound(strOrStrArr) To UBound(strOrStrArr)
            k = LBound(replaceOrReplaces)
            For j = LBound(findOrFinds) To UBound(findOrFinds)
                strOrStrArr(i) = Replace(strOrStrArr(i), findOrFinds(j), _
                                        replaceOrReplaces(k), , , compareMethod)
                k = k + 1
            Next j
        Next i
    Else
        k = LBound(replaceOrReplaces)
        For j = LBound(findOrFinds) To UBound(findOrFinds)
            strOrStrArr = Replace(strOrStrArr, findOrFinds(j), _
                                  replaceOrReplaces(k), , , compareMethod)
            k = k + 1
        Next j
    End If
    ArrayReplaceMultiple = strOrStrArr
End Function

'This function splits a string into a given number of chunks of a given length.
'If a number of chunks and a chunkLength is specified, it creates chunks of the
'given length until either the number of chunks is reached, or the string ends.
'If no number of chunks is specified, the entire string will chunkified into
'chunks of length chunkLength.
'If no chunkLength is specified, the chunk length will be automatically
'calculated as Len(str) / numberOfChunks.
'If no chunkLength and no number of chunks is specified, the string will be
'split into parts of length 1.
'If discardIncompleteChunks = False, tha last chunk can have a different length
'than specified if Len(str) is not divisible by chunkLength
'If splitUTF16Surrogates = False, UTF-16 surrogate pairs will always remain
'together, this means that Len(chunk) can differ from the specified chunkLength
'by up to 1 in that case.
Public Function ChunkifyString(ByRef str As String, _
                      Optional ByVal numberOfChunks As Long = 0, _
                      Optional ByVal chunkLength As Long = 0, _
                      Optional ByVal discardIncompleteChunks As Boolean = False, _
                      Optional ByVal splitUTF16Surrogates As Boolean = True) _
                               As String()
    Dim lenStr As Long: lenStr = Len(str)

    If chunkLength = 0 And numberOfChunks = 0 Then
        chunkLength = 1
    ElseIf chunkLength = 0 And numberOfChunks > 0 Then
        chunkLength = Len(str) \ numberOfChunks
        If chunkLength = 0 Then chunkLength = 1
    End If
    
    If numberOfChunks = 0 Then
        numberOfChunks = Len(str) \ chunkLength
        If Len(str) Mod chunkLength > 0 Then numberOfChunks = numberOfChunks + 1
    End If

    Dim chunks() As String: ReDim chunks(0 To numberOfChunks - 1)
    
    Dim currChunkLength As Long
    Dim chunkIndex As Long
    Dim Position As Long:           Position = 1
    
    For chunkIndex = 0 To numberOfChunks - 1
        If Position > lenStr Then Exit For
        
        currChunkLength = chunkLength
        If Not splitUTF16Surrogates _
        And Position + currChunkLength - 1 < lenStr Then
            If AscU(Mid$(str, Position + currChunkLength - 1, 2)) > &HFFFF& Then
                currChunkLength = currChunkLength - 1
            End If
        End If
        
        If Position + currChunkLength - 1 > lenStr Then
            If discardIncompleteChunks Then Exit For
            currChunkLength = lenStr - Position + 1
        End If
        
        chunks(chunkIndex) = Mid$(str, Position, currChunkLength)
        Position = Position + currChunkLength
    Next chunkIndex
    
    'If the last chunk was not used, shrink the array
    If Position >= lenStr And discardIncompleteChunks Then
        ReDim Preserve chunks(0 To chunkIndex - 1)
    End If
    
    ChunkifyString = chunks
End Function

'This function can convert any variable into a convenient string for printing.
'The last 6 parameters are only relevant for arrays, the last 2 only for
'two dimensional arrays.
''maxChars' is the maximum length of the returned value
''escapeNonPrintable = True' will convert all non ANSI chars to escape sequences
''delimiter' can be used to manually define a delimiter for the returned arrays.
''maxCharsPerLine' will make the output contain newLine characters at least
'                  every 'maxCharsPerLine' characters.
''inklColIndices = True' will print column indices above the columns if the
'                        input is a two dimensional array.
''inklRowIndices = True' does the same for the row indices.
'TODO: `ToString` doesn't comply to its settings with the desired rigour.
'Problem Example:
'Multiple recursion of nested 1d arrays will ignor the line length limit.
Public Function ToString(ByVal Value As Variant, _
                Optional ByVal maxChars As Long = 0, _
                Optional ByVal escapeNonPrintable As Boolean = True, _
                Optional ByRef Delimiter As String = vbNullString, _
                Optional ByVal maxCharsPerElement As Long = 25, _
                Optional ByVal maxCharsPerLine As Long = 80, _
                Optional ByVal maxLines As Long = 10, _
                Optional ByVal inklColIndices As Boolean = True, _
                Optional ByVal inklRowIndices As Boolean = True)
    Const methodName As String = "ToString"
    
    If maxChars < 0 Then _
        Err.Raise 5, methodName, "'maxChars' can't be < 0"
    If maxCharsPerElement < 0 Then _
        Err.Raise 5, methodName, "'maxCharsPerElement' can't be < 0"
    If maxCharsPerLine < 0 Then _
        Err.Raise 5, methodName, "'maxCharsPerLine' can't be < 0"
    If maxLines < 0 Then _
        Err.Raise 5, methodName, "'maxLines' can't be < 0"
    
    If maxChars = 0 Then maxChars = &H7FFFFFFF
    If maxCharsPerElement = 0 Then maxCharsPerElement = &H7FFFFFFF
    If maxCharsPerLine = 0 Then maxCharsPerLine = &H7FFFFFFF
    If maxLines = 0 Then maxLines = &H7FFFFFFF
    
    Dim settings As StringificationSettings
    With settings
        .maxChars = maxChars
        .escapeNonPrintable = escapeNonPrintable
        .Delimiter = Delimiter
        .maxCharsPerElement = maxCharsPerElement
        .maxCharsPerLine = maxCharsPerLine
        .maxLines = maxLines
        .inklColIndices = inklColIndices
        .inklRowIndices = inklRowIndices
    End With
    ToString = BToString(Value, settings)
End Function

'Recursive "Backend" function for 'ToString'
Private Function BToString(ByVal Value As Variant, _
                           ByRef settings As StringificationSettings) As String
    'Don't use exit function in this Function! Instead use: GoTo CleanExit
    Static isRecursiveCall As Boolean
    Dim wasRecursiveCall As Boolean: wasRecursiveCall = isRecursiveCall
    isRecursiveCall = True

    Dim s As String
    Dim i As Long
    If IsArray(Value) Then
        Select Case GetArrayDimsCount(Value)
            Case 0 'Array uninitialized
                s = TypeName(Value)
            Case 1
                If wasRecursiveCall Then settings.maxCharsPerLine = &H7FFFFFFF
                s = ToString1dArray(Value, settings)
            Case 2
                If wasRecursiveCall Then
                    s = ToStringMultiDimArray(Value) 'No settings required
                Else
                    s = ToString2dimArray(Value, settings)
                End If
            Case Else
                s = ToStringMultiDimArray(Value) 'No settings required
        End Select
    ElseIf IsObject(Value) Then
        Select Case True
            'Can add custom logic to ToString any object here
            'Case TypeOf value Is Collection
                's = ToStringCollection(...
            Case Else
                s = TypeName(Value) & "(0x" & Hex(ObjPtr(Value)) & ")"
        End Select
    ElseIf IsEmpty(Value) Then
        s = "Empty"
    Else
        s = CStr(Value)
    End If
    
    With settings
        If Len(s) > .maxChars Then
            BToString = Left(s, Max(.maxChars - 3, 0)) & Left("...", .maxChars)
        Else
            BToString = s
        End If
        
        If .escapeNonPrintable Then BToString = EscapeUnicode(BToString, 255)
    End With
    If VarType(Value) = vbString Then BToString = "'" & BToString & "'"
    
CleanExit:
    If Not wasRecursiveCall Then isRecursiveCall = False
End Function

'Utility function for 'BToString'
'Note:
''maxChars' is only passed to exit the loop sooner in some cases
Private Function ToString1dArray(ByRef arr As Variant, _
                                 ByRef settings As StringificationSettings) _
                                 As String
    Dim s As String
    Dim Delimiter As String: Delimiter = settings.Delimiter
    If StrPtr(Delimiter) = 0 Then Delimiter = ", "
    
    If UBound(arr) - LBound(arr) = -1 Then
        ToString1dArray = "[]"
        Exit Function
    ElseIf UBound(arr) - LBound(arr) = 0 Then
        ToString1dArray = "[" & BToString(arr(UBound(arr)), settings) & "]"
        Exit Function
    End If

    Dim line As String: line = "["
    Dim i As Long
    For i = LBound(arr) To UBound(arr) - 1
        line = line & BToString(arr(i), settings) & Delimiter

       'Check if max characters per line would be exceeded with the next element
        If Len(line & BToString(arr(i + 1), settings) & _
               Delimiter) >= settings.maxCharsPerLine Then
            s = s & line & vbNewLine
            line = ""
            Dim lineCount As Long: lineCount = lineCount + 1
            If lineCount >= settings.maxLines _
            Or Len(s) > settings.maxChars Then Exit For
        End If
    Next i
    ToString1dArray = s & line & BToString(arr(UBound(arr)), settings) & "]"
End Function

'Utility function for 'BToString'
'Note: Will only be called in non-recursive calls of 'BToString'
Private Function ToString2dimArray(ByRef arr As Variant, _
                                   ByRef settings As StringificationSettings) _
                                   As String
    Dim Delimiter As String: Delimiter = settings.Delimiter
    If StrPtr(settings.Delimiter) = 0 Then Delimiter = "  "
    
    With settings
        Dim colWidths() As Long
        colWidths = CalculateColumnWidths(arr, settings)
        Dim numCols As Long
        numCols = CalculateNumColumnsToFit(colWidths, .maxCharsPerLine, _
                                           Len(Delimiter))
    
        Dim numRows As Long: numRows = UBound(arr, 1) - LBound(arr, 1) + 1
        Dim firstRows As Long: firstRows = Min(.maxLines \ 2, numRows)
        Dim lastRows As Long
        lastRows = Min(.maxLines - firstRows, numRows - firstRows)
    End With
    Dim s As String

    If settings.inklColIndices Then _
        s = BuildColHeadersLine(arr, colWidths, numCols, Delimiter, settings) _
            & vbNewLine
                          
    Dim i As Long
    For i = LBound(arr, 1) To LBound(arr, 1) + firstRows - 1
        If Len(s) > settings.maxChars Then Exit For
        s = s & BuildLine(arr, colWidths, numCols, Delimiter, i, settings) _
            & vbNewLine
    Next i
    
    If numRows > settings.maxLines And Len(s) < settings.maxChars Then _
        s = s & BuildDotsLine(arr, colWidths, numCols, Delimiter, settings) _
            & vbNewLine
    
    For i = UBound(arr, 1) - lastRows + 1 To UBound(arr, 1)
        If Len(s) > settings.maxChars Then Exit For
        s = s & BuildLine(arr, colWidths, numCols, Delimiter, i, settings) _
            & vbNewLine
    Next i
    ToString2dimArray = s & "(" & numRows & "*" & _
                         UBound(arr, 2) - LBound(arr, 2) + 1 & ", " & _
                         " " & Min(numRows, settings.maxLines) & _
                         "*" & numCols & " output)"
End Function

Private Function ToStringMultiDimArray(ByRef arr As Variant) As String
    Dim s As String
    s = "Array("
    Dim i As Long
    For i = 1 To GetArrayDimsCount(arr)
        s = s & LBound(arr, i) & " to " & UBound(arr, i) & ", "
    Next i
    ToStringMultiDimArray = Left(s, Len(s) - 2) & ")"
End Function

'Utility function for 'ToString2dimArray'
Private Function BuildColHeadersLine(ByRef arr As Variant, _
                                     ByRef colWidths() As Long, _
                                     ByVal numCols As Long, _
                                     ByVal Delimiter As String, _
                                     ByRef settings As StringificationSettings) _
                                     As String
    Dim rowNumPadding As Long
    rowNumPadding = Max(Len(CStr(UBound(arr, 1))), Len(CStr(LBound(arr, 1)))) + 2
    
    Dim lenDelim As Long: lenDelim = Len(Delimiter)
    Dim j As Long
    If numCols = UBound(arr, 2) - LBound(arr, 2) + 1 Then
        For j = LBound(arr, 2) To UBound(arr, 2)
            BuildColHeadersLine = BuildColHeadersLine & _
                                  PadLeft(CStr(j), colWidths(j)) _
                                  & Delimiter
        Next j
        If settings.inklRowIndices Then _
            BuildColHeadersLine = Space(rowNumPadding) & BuildColHeadersLine

        BuildColHeadersLine = Left(BuildColHeadersLine, _
                                   Len(BuildColHeadersLine) - lenDelim)
        Exit Function
    End If
    
    Dim leftPart As String, rightPart As String
    For j = 0 To numCols \ 2 - 1 'numCols is always even
        leftPart = leftPart & PadLeft(CStr(LBound(arr, 2) + j), _
            colWidths(LBound(arr, 2) + j)) & Space(lenDelim)
        rightPart = Space(lenDelim) & PadLeft(CStr(UBound(arr, 2) - j), _
                    colWidths(UBound(arr, 2) - j)) & rightPart
    Next j

    If settings.inklRowIndices Then _
        BuildColHeadersLine = Space(rowNumPadding)

    BuildColHeadersLine = BuildColHeadersLine & _
                          Left(leftPart, Len(leftPart) - lenDelim) & _
                          " ... " & Right(rightPart, Len(rightPart) - lenDelim)
End Function

'Utility function for 'ToString2dimArray'
Private Function BuildDotsLine(ByRef arr As Variant, _
                               ByRef colWidths() As Long, _
                               ByVal numCols As Long, _
                               ByVal Delimiter As String, _
                               ByRef settings As StringificationSettings) _
                               As String
    Dim rowNumPadding As Long
    rowNumPadding = Max(Len(CStr(UBound(arr, 1))), Len(CStr(LBound(arr, 1)))) + 2
    Dim lenDelim As Long: lenDelim = Len(Delimiter)
    Dim j As Long
    If numCols = UBound(arr, 2) - LBound(arr, 2) + 1 Then
        For j = LBound(arr, 2) To UBound(arr, 2)
            BuildDotsLine = BuildDotsLine & PadLeft("...", colWidths(j)) _
                                          & Delimiter
        Next j
        If settings.inklRowIndices Then _
                BuildDotsLine = PadRight("..", rowNumPadding) & BuildDotsLine

        BuildDotsLine = Left(BuildDotsLine, Len(BuildDotsLine) - lenDelim)
        Exit Function
    End If
    
    Dim leftPart As String, rightPart As String
    For j = 0 To numCols \ 2 - 1 'numCols is always even
        leftPart = leftPart & PadLeft("...", colWidths(LBound(arr, 2) + j)) _
                   & Space(lenDelim)
        rightPart = Space(lenDelim) & PadLeft("...", _
                    colWidths(UBound(arr, 2) - j)) & rightPart
    Next j

    If settings.inklRowIndices Then _
        BuildDotsLine = PadRight("..", rowNumPadding)

    BuildDotsLine = BuildDotsLine & Left(leftPart, Len(leftPart) - lenDelim) & _
                    " ... " & Right(rightPart, Len(rightPart) - lenDelim)
End Function

'Utility function for 'ToString2dimArray'
Private Function BuildLine(ByRef arr As Variant, _
                           ByRef colWidths() As Long, _
                           ByVal numCols As Long, _
                           ByVal Delimiter As String, _
                           ByVal rowIndex As Long, _
                           ByRef settings As StringificationSettings) As String
    Dim rowNumPadding As Long
    rowNumPadding = Max(Len(CStr(UBound(arr, 1))), Len(CStr(LBound(arr, 1)))) + 2
    
    Dim lenDelim As Long: lenDelim = Len(Delimiter)
    Dim j As Long
    If numCols = UBound(arr, 2) - LBound(arr, 2) + 1 Then
        For j = LBound(arr, 2) To UBound(arr, 2)
            BuildLine = BuildLine & _
                        PadLeft(BToString(arr(rowIndex, j), settings), _
                                colWidths(j)) & Delimiter
        Next j
        If settings.inklRowIndices Then _
            BuildLine = PadRight(CStr(rowIndex), rowNumPadding) & BuildLine

        BuildLine = Left(BuildLine, Len(BuildLine) - lenDelim)
        Exit Function
    End If
    
    Dim leftPart As String, rightPart As String
    For j = 0 To numCols \ 2 - 1 'numCols is always even
        leftPart = leftPart & PadLeft(BToString(arr(rowIndex, _
                   LBound(arr, 2) + j), settings), _
                   colWidths(LBound(arr, 2) + j)) & Delimiter
        rightPart = Delimiter & PadLeft(BToString(arr(rowIndex, _
                    UBound(arr, 2) - j), settings), _
                    colWidths(UBound(arr, 2) - j)) & rightPart
    Next j

    If settings.inklRowIndices Then _
        BuildLine = PadRight(CStr(rowIndex), rowNumPadding)

    BuildLine = BuildLine & Left(leftPart, Len(leftPart) - lenDelim) & _
                " ... " & Right(rightPart, Len(rightPart) - lenDelim)
End Function

'Utility function for 'ToString2dimArray'
Private Function CalculateColumnWidths(ByRef arr As Variant, _
                                       ByRef settings As StringificationSettings) _
                                       As Long()
    Dim colWidths() As Long: ReDim colWidths(LBound(arr, 2) To UBound(arr, 2))
    Dim i As Long, j As Long
    
    ' Calculate how many rows will be printed before and after the dots
    Dim numRows As Long: numRows = UBound(arr, 1) - LBound(arr, 1) + 1
    Dim numCols As Long: numCols = UBound(arr, 2) - LBound(arr, 2) + 1
    Dim firstRows As Long: firstRows = Min(settings.maxLines \ 2, numRows)
    Dim lastRows As Long
    lastRows = Min(settings.maxLines - firstRows, numRows - firstRows)
    
    Dim sumWidths As Long: sumWidths = 0
    For j = 0 To numCols \ 2
        Dim col1 As Long: col1 = LBound(arr, 2) + j
        Dim col2 As Long: col2 = UBound(arr, 2) - j
        'if col1>
        colWidths(col1) = Max(colWidths(col1), Len(CStr(col1))) 'Column labels
        colWidths(col2) = Max(colWidths(col2), Len(CStr(col2))) 'Column labels
        For i = LBound(arr, 1) To LBound(arr, 1) + firstRows - 1
            colWidths(col1) = Max(Len(BToString(arr(i, col1), settings)), _
                                  colWidths(col1))
            colWidths(col2) = Max(Len(BToString(arr(i, col2), settings)), _
                                  colWidths(col2))
        Next i
        sumWidths = sumWidths + colWidths(col1) + colWidths(col2)
        If sumWidths > settings.maxCharsPerLine Then Exit For
    Next j
    
    If settings.maxLines < numRows Then
        sumWidths = 0
        For j = 0 To numCols \ 2
            col1 = LBound(arr, 2) + j
            col2 = UBound(arr, 2) - j
            For i = UBound(arr, 1) - lastRows + 1 To UBound(arr, 1)
                colWidths(col1) = Max(Len(BToString(arr(i, col1), settings)), _
                                      colWidths(col1))
                colWidths(col2) = Max(Len(BToString(arr(i, col2), settings)), _
                                      colWidths(col2))
            Next i
            sumWidths = sumWidths + colWidths(col1) + colWidths(col2)
            If sumWidths > settings.maxCharsPerLine Then Exit For
        Next j
    End If

    CalculateColumnWidths = colWidths
End Function

'Utility function for 'ToString2dimArray'
Private Function CalculateNumColumnsToFit(ByRef colWidths() As Long, _
                                          ByVal maxCharsPerLine As Long, _
                                          ByVal delimLength As Long) As Long
    Dim totalWidth As Long
    Dim numCols As Long:     numCols = UBound(colWidths) - LBound(colWidths) + 1
    Dim extraWidthForDots As Long: extraWidthForDots = 3 ' Width for the "..."
    
    If Sum(colWidths) + delimLength * (numCols - 1) <= maxCharsPerLine Then
        CalculateNumColumnsToFit = numCols
        Exit Function
    End If
    
    'Else, omitt at least one column:
    Dim j As Long
    For j = LBound(colWidths) To UBound(colWidths) \ 2
        totalWidth = totalWidth + colWidths(j) + _
                     colWidths(UBound(colWidths) - (j - LBound(colWidths))) + _
                     delimLength * 2
        If totalWidth + extraWidthForDots <= maxCharsPerLine Then
            CalculateNumColumnsToFit = CalculateNumColumnsToFit + 2
        Else
            Exit For
        End If
    Next j
End Function

Private Function Sum(ParamArray p() As Variant) As Variant
    Dim v As Variant
    For Each v In p
        If IsArray(v) Then
            Dim e As Variant
            For Each e In v
                If IsArray(e) Then
                    Sum = Sum + Sum(e)
                Else
                    If IsNumeric(e) Then Sum = Sum + e
                End If
            Next e
        End If
        If IsNumeric(v) Then Sum = Sum + v
    Next v
End Function

Private Function Max(a As Long, b As Long) As Long
    If a > b Then Max = a Else Max = b
End Function

Private Function Min(a As Long, b As Long) As Long
    If a < b Then Min = a Else Min = b
End Function

'Sets the formatting rules adhered to by 'Printf'
Public Sub SetPrintfSettings(Optional ByVal maxChars As Long = 0, _
                          Optional ByVal escapeNonPrintable As Boolean = True, _
                            Optional ByRef Delimiter As String = vbNullString, _
                             Optional ByVal maxCharsPerElement As Long = 25, _
                             Optional ByVal maxCharsPerLine As Long = 80, _
                             Optional ByVal maxLines As Long = 10, _
                             Optional ByVal inklColIndices As Boolean = True, _
                             Optional ByVal inklRowIndices As Boolean = True)
    Const methodName As String = "SetPrintfSettings"
    
    If maxChars < 0 Then _
        Err.Raise 5, methodName, "'maxChars' can't be < 0"
    If maxCharsPerElement < 0 Then _
        Err.Raise 5, methodName, "'maxCharsPerElement' can't be < 0"
    If maxCharsPerLine < 0 Then _
        Err.Raise 5, methodName, "'maxCharsPerLine' can't be < 0"
    If maxLines < 0 Then _
        Err.Raise 5, methodName, "'maxLines' can't be < 0"
    
    If maxChars = 0 Then maxChars = &H7FFFFFFF
    If maxCharsPerElement = 0 Then maxCharsPerElement = &H7FFFFFFF
    If maxCharsPerLine = 0 Then maxCharsPerLine = &H7FFFFFFF
    If maxLines = 0 Then maxLines = &H7FFFFFFF
    
    With printfSettings
        .maxChars = maxChars
        .escapeNonPrintable = escapeNonPrintable
        .Delimiter = Delimiter
        .maxCharsPerElement = maxCharsPerElement
        .maxCharsPerLine = maxCharsPerLine
        .maxLines = maxLines
        .inklColIndices = inklColIndices
        .inklRowIndices = inklRowIndices
    End With
End Sub

'Utility function for 'Printf'
Private Function GetPrintfSettings() As StringificationSettings
    If Not printfSettingsAreInitialized Then
        SetPrintfSettings 'With only default arguments
        printfSettingsAreInitialized = True
    End If
    GetPrintfSettings = printfSettings
End Function

'Prints any variables passed to this function. Uses formatting rules previously
'set with 'SetPrintfSettings'
Public Function Printf(ParamArray args() As Variant) As String
    Dim arg As Variant
    Dim sArg As Variant
    Dim s As String
    Dim settings As StringificationSettings: settings = GetPrintfSettings
    For Each arg In args
        If VarType(arg) = vbString Then
            sArg = arg
        Else
            sArg = BToString(arg, settings)
        End If
        If InStr(1, sArg, vbNewLine, vbBinaryCompare) <> 0 Then
            s = s & vbNewLine & sArg & vbNewLine
        ElseIf Len(s) - InStrRev(s, vbNewLine, , vbBinaryCompare) + Len(sArg) _
               > settings.maxCharsPerLine - 2 Then
            s = s & vbNewLine & sArg & "  "
        Else
            s = s & sArg & "  "
        End If
        If Len(s) > settings.maxChars Then
            s = s & "..."
            Exit For
        End If
    Next arg
    Printf = TrimX(s, vbCrLf)
    Debug.Print Printf
End Function

'Prints an one or two dimensional array to the immediate window.
Public Sub PrintVar(ByRef arr As Variant, _
           Optional ByRef Delimiter As String = vbNullString, _
           Optional ByVal maxCharsPerElement As Long = 25, _
           Optional ByVal maxCharsPerLine As Long = 80, _
           Optional ByVal maxLines As Long = 10, _
           Optional ByVal escapeNonPrintable As Boolean = True, _
           Optional ByVal printColIndices As Boolean = True, _
           Optional ByVal printRowIndices As Boolean = True)
    Debug.Print ToString(arr, , escapeNonPrintable, Delimiter, _
                          maxCharsPerElement, maxCharsPerLine, maxLines, _
                          printColIndices, printRowIndices)
End Sub

'Works like the inbuilt trim but instead of just spaces, it will trim any
'characters occurring in 'charactersToTrim' from the edges of 'str'
'If 'charactersToTrim' is an empty string, nothing will be trimmed.
Public Function TrimX(ByRef str As String, _
             Optional ByRef charactersToTrim As String = " " & vbCrLf & vbTab, _
             Optional ByVal compareMethod As VbCompareMethod = vbBinaryCompare) _
                      As String
    If Len(str) = 0 Or Len(charactersToTrim) = 0 Then Exit Function
    Dim strLen As Long:   strLen = Len(str)
    Dim startIdx As Long: startIdx = 1
    Dim endIdx As Long:   endIdx = strLen

    Do While startIdx <= strLen _
        And InStr(1, charactersToTrim, Mid$(str, startIdx, 1), compareMethod) > 0
        startIdx = startIdx + 1
    Loop

    Do While endIdx >= 1
        If InStr(1, charactersToTrim, Mid$(str, endIdx, 1), compareMethod) > 0 Then
            endIdx = endIdx - 1
        Else
            Exit Do
        End If
    Loop

    If startIdx <= endIdx Then
        TrimX = Mid$(str, startIdx, endIdx - startIdx + 1)
    Else
        TrimX = vbNullString
    End If
End Function

'Returns the index of a given column letter in excel
Public Function ColLetterToNumber(ByRef colLetterOrNumber As Variant) As Long
    Const methodName As String = "ColLetterToNumber"
    
    If IsNumeric(colLetterOrNumber) Then
        ColLetterToNumber = CLng(colLetterOrNumber)
        If ColLetterToNumber > 16384 Then _
            Err.Raise 5, methodName, _
                "Only values up to 'XFD', or 16384 are allowed."
    ElseIf Len(colLetterOrNumber) <= 3 _
    And (CStr(UCase(colLetterOrNumber)) Like "[A-Z]" _
    Or CStr(UCase(colLetterOrNumber)) Like "[A-Z][A-Z]" _
    Or CStr(UCase(colLetterOrNumber)) Like "[A-Z][A-Z][A-Z]") Then
        Dim i As Long
        For i = 1 To Len(colLetterOrNumber)
            ColLetterToNumber = ColLetterToNumber * 26 + _
                (Asc(UCase(Mid$(colLetterOrNumber, i, 1))) - 65 + 1)
        Next i
        If ColLetterToNumber > 16384 Then _
            Err.Raise 5, methodName, _
                "Only values up to 'XFD', or 16384 are allowed."
    Else
        Err.Raise 5, methodName, "The input must be up to 3 letters or 5 " & _
            "digits long. Only values up to 'XFD', or 16384 are allowed and" & _
            " no mixture of letters and digits is allowed."
    End If
End Function

'Converts a string to date based on a format specified in 'dateFormat'
'E.g. ParseDate("27.04.1993", "DD.MM.YYYY") = ParseDate("042793", "MMDDYY")
'Follows the idea by Scott Craner: https://stackoverflow.com/a/64813581/12287457
Public Function ParseDate(ByRef str As String, _
                          ByRef dateFormat As String) As Date
    Const methodName As String = "ParseDate"
    
    If Len(str) <> Len(dateFormat) Then Err.Raise 5, methodName, _
        "The input string must be of same length as the format string."
    
    Dim i As Long
    For i = 1 To Len(str)
        If UCase(Mid$(dateFormat, i, 1)) = "D" Then
            Dim lDay As Long: lDay = lDay * 10 + CLng(Mid$(str, i, 1))
        ElseIf UCase(Mid$(dateFormat, i, 1)) = "Y" Then
            Dim lYear As Long: lYear = lYear * 10 + CLng(Mid$(str, i, 1))
        ElseIf UCase(Mid$(dateFormat, i, 1)) = "M" Then
            Dim sMonth As String: sMonth = sMonth & Mid$(str, i, 1)
        End If
    Next i
    
    If IsNumeric(sMonth) Then
        Dim lMonth As Long: lMonth = CLng(sMonth)
    Else
        lMonth = Month(CDate("01 " & sMonth & " 2023"))
    End If
    
    ParseDate = DateSerial(lYear, lMonth, lDay)
End Function

'https://en.wikipedia.org/wiki/Wichmann-Hill
'https://www.vbforums.com/showthread.php?499661-Wichmann-Hill-Pseudo-Random-Number-Generator-an-alternative-for-VB-Rnd()-function&p=3076123&viewfull=1#post3076123
Public Function RndWH(Optional ByVal Number As Long) As Double
    Static lngX As Long, lngY As Long, lngZ As Long, blnInit As Boolean
    Dim dblRnd As Double
    'If initialized and no input number given
    If blnInit And Number = 0 Then
        ' lngX, lngY and lngZ will never be 0
        lngX = (171& * lngX) Mod 30269&
        lngY = (172& * lngY) Mod 30307&
        lngZ = (170& * lngZ) Mod 30323&
    Else
        'If no initialization, use Timer, otherwise ensure positive Number
        If Number = 0 Then Number = Timer * 60 Else Number = Number And &H7FFFFFFF
        lngX = (Number Mod 30269&)
        lngY = (Number Mod 30307&)
        lngZ = (Number Mod 30323&)
        'lngX, lngY and lngZ must be bigger than 0
        If lngX = 0 Then lngX = 171&
        If lngY = 0 Then lngY = 172&
        If lngZ = 0 Then lngZ = 170&
        'Mark initialization state
        blnInit = True
    End If
    'Generate a random number
    dblRnd = CDbl(lngX) / 30269# + CDbl(lngY) / 30307# + CDbl(lngZ) / 30323#
    'Return a value between 0 and 1
    RndWH = dblRnd - Int(dblRnd)
End Function


