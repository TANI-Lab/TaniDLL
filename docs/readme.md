## 📘 Documentation

Full API reference is available here:

👉 [Download API Reference (PDF)](https://raw.githubusercontent.com/TANI-Lab/TaniDLL/main/docs/TANI_DLL_specification_en.pdf)

*Note: The PDF may not render correctly on GitHub. Please download to view.*


# COM Port DLL Specification (TANI DLL)

---

## 1. Overview

This DLL is a library for easily and safely handling COM ports (serial communication) from Excel (VBA), Delphi, and C/C++.  
It supports both 32-bit and 64-bit environments and is designed for communication with Arduino / ESP32.

---

## 2. Constant Definitions

### ■ Return Values

```c
TANI_OK          =  0
TANI_OK_LICENSE  =  1
TANI_ERR         = -1
RESULT_BLOCKED   = -2
TANI_ERR_LICENSE = -10
```

---

### ■ Communication Settings

```c
TANI_PARITY_NONE = 0
TANI_PARITY_ODD  = 1
TANI_PARITY_EVEN = 2

TANI_STOP_1 = 1
TANI_STOP_2 = 2

TANI_FLOW_NONE    = 0
TANI_FLOW_RTSCTS  = 1
TANI_FLOW_XONXOFF = 2
```

---

### ■ License Control

```c
LICENSE_CACHE_MIN = 60
COMM_FAIL_MAX     = 5
COMM_FAIL_WINDOW  = 5 / 1440.0
COMM_BLOCK_TIME   = 60 / 1440.0
```

---

## 3. Error Code Details

| Code | Description | Action |
|------|-------------|--------|
| 0 | Success | None |
| 1 | License authentication successful | First time only |
| -1 | General error | Check COM settings |
| -2 | Temporary block | Wait |
| -10 | License not authenticated | Run CheckLicense |

---

## 4. DLL Function Specifications

---

# ■ TANI_CheckLicense

### Signature (Delphi)
```delphi
function TANI_CheckLicense(LicenseKey: PAnsiChar): Integer; stdcall;
```

### C/C++
```cpp
typedef int (__stdcall *TANI_CheckLicense_t)(const char* key);
```

### Arguments

| Name | Type | Description |
|------|------|-------------|
| LicenseKey | const char* | Subscription key |

### Return Values

| Value | Description |
|-------|-------------|
| 1 | Authentication successful |
| Other | Failure |

---

### VBA Header

```vb
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function TANI_CheckLicense Lib "TaniDLL64.dll" ( _
            ByVal LicenseKey As String) As Long
    #Else
        Private Declare Function TANI_CheckLicense Lib "TaniDLL32.dll" ( _
            ByVal LicenseKey As String) As Long
    #End If
#Else
    Private Declare Function TANI_CheckLicense Lib "TaniDLL32.dll" ( _
        ByVal LicenseKey As String) As Long
#End If
```

### VBA Example

```vb
Dim ret As Long

ret = TANI_CheckLicense("XXXXXXXX")

If ret <> 1 Then
    MsgBox "License error", vbCritical
    Exit Sub
End If
```

### Delphi Example

```delphi
var ret: Integer;
begin
  ret := TANI_CheckLicense('XXXXXXXX');
end;
```

### C/C++

```cpp
int ret = TANI_CheckLicense("XXXXXXXX");
```

---

# ■ TANI_OpenEx

### Signature (Delphi)
```delphi
function TANI_OpenEx(
  id: Integer;
  name: PAnsiChar;
  BAUD: Integer;
  databits: Integer;
  parity: Integer;
  stopbits: Integer;
  flow: Integer
): Integer; stdcall;
```

### C/C++
```cpp
typedef int (__stdcall *TANI_OpenEx_t)(int,const char*,int,int,int,int,int);
```

### Arguments

| Name | Type | Description |
|------|------|-------------|
| id | int | Port ID (0–7) |
| name | const char* | COM port name |
| BAUD | int | Baud rate |
| databits | int | Data bits |
| parity | int | Parity |
| stopbits | int | Stop bits |
| flow | int | Flow control |

### Return Values

| Value | Description |
|-------|-------------|
| 0 | Success |
| -1 | Failure |

---

### VBA Header

```vb
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function TANI_OpenEx Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByVal name As String, _
            ByVal BAUD As Long, _
            ByVal databits As Long, _
            ByVal parity As Long, _
            ByVal stopbits As Long, _
            ByVal flow As Long) As Long
    #Else
        Private Declare Function TANI_OpenEx Lib "TaniDLL32.dll" ( _
            ByVal id As Long, _
            ByVal name As String, _
            ByVal BAUD As Long, _
            ByVal databits As Long, _
            ByVal parity As Long, _
            ByVal stopbits As Long, _
            ByVal flow As Long) As Long
    #End If
#Else
    Private Declare Function TANI_OpenEx Lib "TaniDLL32.dll" ( _
        ByVal id As Long, _
        ByVal name As String, _
        ByVal BAUD As Long, _
        ByVal databits As Long, _
        ByVal parity As Long, _
        ByVal stopbits As Long, _
        ByVal flow As Long) As Long
#End If
```

### VBA Example

```vb
Dim ret As Long

ret = TANI_OpenEx(0, "COM5", 115200, 8, 0, 1, 0)

If ret <> 0 Then
    MsgBox "Open failed", vbCritical
End If
```

### Delphi Example

```delphi
if TANI_OpenEx(0,'COM5',115200,8,0,1,0) <> 0 then Exit;
```

### C/C++

```cpp
if (TANI_OpenEx(0,"COM5",115200,8,0,1,0) != 0) return;
```

---

# ■ TANI_OpenSimple

### Signature (Delphi)
```delphi
function TANI_OpenSimple(
  id: Integer;
  name: PAnsiChar;
  baud: Integer
): Integer; stdcall;
```

### C/C++
```cpp
typedef int (__stdcall *TANI_OpenSimple_t)(int,const char*,int);
```

### Arguments

| Name | Type | Description |
|------|------|-------------|
| id | int | Port ID (0–7) |
| name | const char* | COM port name |
| baud | int | Baud rate |

### Return Values

| Value | Description |
|-------|-------------|
| 0 | Success |
| -1 | Failure |

---

### VBA Header

```vb
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function TANI_OpenSimple Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByVal name As String, _
            ByVal baud As Long) As Long
    #Else
        Private Declare Function TANI_OpenSimple Lib "TaniDLL32.dll" ( _
            ByVal id As Long, _
            ByVal name As String, _
            ByVal baud As Long) As Long
    #End If
#Else
    Private Declare Function TANI_OpenSimple Lib "TaniDLL32.dll" ( _
        ByVal id As Long, _
        ByVal name As String, _
        ByVal baud As Long) As Long
#End If
```

### VBA Example

```vb
Dim ret As Long

ret = TANI_OpenSimple(0, "COM5", 115200)

If ret <> 0 Then
    MsgBox "Open failed", vbCritical
End If
```

### Delphi Example

```delphi
if TANI_OpenSimple(0,'COM5',115200) <> 0 then Exit;
```

### C/C++

```cpp
if (TANI_OpenSimple(0,"COM5",115200) != 0) return;
```

---

# ■ TANI_Close

### Signature (Delphi)
```delphi
procedure TANI_Close(id: Integer); stdcall;
```

### C/C++
```cpp
typedef void (__stdcall *TANI_Close_t)(int);
```

### Arguments

| Name | Type | Description |
|------|------|-------------|
| id | int | Port ID |

### Return Values  
None

---

### VBA Header

```vb
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Sub TANI_Close Lib "TaniDLL64.dll" ( _
            ByVal id As Long)
    #Else
        Private Declare Sub TANI_Close Lib "TaniDLL32.dll" ( _
            ByVal id As Long)
    #End If
#Else
    Private Declare Sub TANI_Close Lib "TaniDLL32.dll" ( _
        ByVal id As Long)
#End If
```

### Example

```vb
TANI_Close 0
```

```delphi
TANI_Close(0);
```

```cpp
TANI_Close(0);
```

---

# ■ TANI_SendText

### Signature (Delphi)
```delphi
function TANI_SendText(id: Integer; txt: PAnsiChar; term: PAnsiChar): Integer; stdcall;
```

### C/C++
```cpp
typedef int (__stdcall *TANI_SendText_t)(int,const char*,const char*);
```

### Arguments

| Name | Type | Description |
|------|------|-------------|
| id | int | Port ID |
| txt | const char* | String to send |
| term | const char* | Terminator character |

### Return Values

| Value | Description |
|-------|-------------|
| >0 | Bytes sent |
| <=0 | Failure |

---

### VBA Header

```vb
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function TANI_SendText Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByVal txt As String, _
            ByVal TERM As String) As Long
    #Else
        Private Declare Function TANI_SendText Lib "TaniDLL32.dll" ( _
            ByVal id As Long, _
            ByVal txt As String, _
            ByVal TERM As String) As Long
    #End If
#Else
    Private Declare Function TANI_SendText Lib "TaniDLL32.dll" ( _
        ByVal id As Long, _
        ByVal txt As String, _
        ByVal TERM As String) As Long
#End If
```

### Example

```vb
TANI_SendText 0, "HELLO", vbCr
```

```delphi
TANI_SendText(0,'HELLO',#13);
```

```cpp
TANI_SendText(0,"HELLO","\r");
```

---

# ■ TANI_SendBytes

### Signature (Delphi)
```delphi
function TANI_SendBytes(id: Integer; data: PByte; length: Integer): Integer; stdcall;
```

### C/C++
```cpp
typedef int (__stdcall *TANI_SendBytes_t)(int,unsigned char*,int);
```

### Arguments

| Name | Type | Description |
|------|------|-------------|
| id | int | Port ID |
| data | unsigned char* | Data |
| length | int | Size |

### Return Values

| Value | Description |
|-------|-------------|
| >0 | Bytes sent |
| <=0 | Failure |

---

### VBA Header

```vb
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function TANI_SendBytes Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByRef data As Byte, _
            ByVal length As Long) As Long
    #Else
        Private Declare Function TANI_SendBytes Lib "TaniDLL32.dll" ( _
            ByVal id As Long, _
            ByRef data As Byte, _
            ByVal length As Long) As Long
    #End If
#Else
    Private Declare Function TANI_SendBytes Lib "TaniDLL32.dll" ( _
        ByVal id As Long, _
        ByRef data As Byte, _
        ByVal length As Long) As Long
#End If
```

---

### Example

```vb
Dim buf(2) As Byte
buf(0) = 1
buf(1) = 2
buf(2) = 3

Dim ret As Long
ret = TANI_SendBytes(0, buf(0), 3)

If ret <= 0 Then
    MsgBox "Send error", vbCritical
End If
```

```delphi
var buf: array[0..2] of Byte = (1,2,3);
TANI_SendBytes(0,@buf[0],3);
```

```cpp
unsigned char buf[3]={1,2,3};
TANI_SendBytes(0,buf,3);
```

---

# ■ TANI_RecvLine

### Signature (Delphi)
```delphi
function TANI_RecvLine(id: Integer; buf: PByte; bufSize: Integer; timeout: Integer; term: Byte): Integer; stdcall;
```

### C/C++
```cpp
typedef int (__stdcall *TANI_RecvLine_t)(int,unsigned char*,int,int,unsigned char);
```

### Arguments

| Name | Type | Description |
|------|------|-------------|
| id | int | Port ID |
| buf | unsigned char* | Buffer |
| bufSize | int | Size |
| timeout | int | ms |
| term | unsigned char | Terminator |

### Return Values

| Value | Description |
|-------|-------------|
| >0 | Bytes received |
| 0 | Timeout |
| <0 | Error |

---

### VBA Header

```vb
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function TANI_RecvLine Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByRef buf As Byte, _
            ByVal bufSize As Long, _
            ByVal timeout As Long, _
            ByVal TERM As Byte) As Long
    #Else
        Private Declare Function TANI_RecvLine Lib "TaniDLL32.dll" ( _
            ByVal id As Long, _
            ByRef buf As Byte, _
            ByVal bufSize As Long, _
            ByVal timeout As Long, _
            ByVal TERM As Byte) As Long
    #End If
#Else
    Private Declare Function TANI_RecvLine Lib "TaniDLL32.dll" ( _
        ByVal id As Long, _
        ByRef buf As Byte, _
        ByVal bufSize As Long, _
        ByVal timeout As Long, _
        ByVal TERM As Byte) As Long
#End If
```

---

### Example

```vb
Dim buf(255) As Byte
Dim ret As Long

ret = TANI_RecvLine(0, buf(0), 256, 1000, 13)

If ret > 0 Then
    Dim s As String
    s = StrConv(buf, vbUnicode)
    MsgBox "Received: " & Left(s, ret)
ElseIf ret = 0 Then
    MsgBox "Timeout"
Else
    MsgBox "Receive error", vbCritical
End If
```

```delphi
var buf: array[0..255] of Byte;
TANI_RecvLine(0,@buf[0],256,1000,13);
```

```cpp
unsigned char buf[256];
TANI_RecvLine(0,buf,256,1000,13);
```

---

# ■ TANI_RecvBytes

### Signature (Delphi)
```delphi
function TANI_RecvBytes(id: Integer; buf: PByte; size: Integer; timeout: Integer): Integer; stdcall;
```

### C/C++
```cpp
typedef int (__stdcall *TANI_RecvBytes_t)(int,unsigned char*,int,int);
```

### Arguments

| Name | Type | Description |
|------|------|-------------|
| id | int | Port ID |
| buf | unsigned char* | Buffer |
| size | int | Read size |
| timeout | int | ms |

### Return Values

| Value | Description |
|-------|-------------|
| >0 | Bytes received |
| 0 | Timeout |
| <0 | Error |

---

### VBA Header

```vb
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function TANI_RecvBytes Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByRef buf As Byte, _
            ByVal size As Long, _
            ByVal timeout As Long) As Long
    #Else
        Private Declare Function TANI_RecvBytes Lib "TaniDLL32.dll" ( _
            ByVal id As Long, _
            ByRef buf As Byte, _
            ByVal size As Long, _
            ByVal timeout As Long) As Long
    #End If
#Else
    Private Declare Function TANI_RecvBytes Lib "TaniDLL32.dll" ( _
        ByVal id As Long, _
        ByRef buf As Byte, _
        ByVal size As Long, _
        ByVal timeout As Long) As Long
#End If
```

---

### VBA Example

```vb
Dim buf(255) As Byte
Dim ret As Long

ret = TANI_RecvBytes(0, buf(0), 10, 1000)

If ret > 0 Then
    MsgBox "Bytes received: " & ret
ElseIf ret = 0 Then
    MsgBox "Timeout"
Else
    MsgBox "Receive error", vbCritical
End If
```

```delphi
var buf: array[0..255] of Byte;
TANI_RecvBytes(0,@buf[0],10,1000);
```

```cpp
unsigned char buf[256];
TANI_RecvBytes(0,buf,10,1000);
```

---

# ■ TANI_RecvBytesUntil

### Signature (Delphi)
```delphi
function TANI_RecvBytesUntil(id: Integer; buf: PByte; bufSize: Integer; termByte: Byte; inclTerm: Integer; timeout: Integer): Integer; stdcall;
```

### C/C++
```cpp
typedef int (__stdcall *TANI_RecvBytesUntil_t)(int,unsigned char*,int,unsigned char,int,int);
```

### Arguments

| Name | Type | Description |
|------|------|-------------|
| id | int | Port ID |
| buf | unsigned char* | Buffer |
| bufSize | int | Size |
| termByte | unsigned char | Terminator byte |
| inclTerm | int | Include terminator |
| timeout | int | ms |

### Return Values

| Value | Description |
|-------|-------------|
| >0 | Bytes received |
| 0 | Timeout |
| <0 | Error |

---

### VBA Header

```vb
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function TANI_RecvBytesUntil Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByRef buf As Byte, _
            ByVal bufSize As Long, _
            ByVal termByte As Byte, _
            ByVal inclTerm As Long, _
            ByVal timeout As Long) As Long
    #Else
        Private Declare Function TANI_RecvBytesUntil Lib "TaniDLL32.dll" ( _
            ByVal id As Long, _
            ByRef buf As Byte, _
            ByVal bufSize As Long, _
            ByVal termByte As Byte, _
            ByVal inclTerm As Long, _
            ByVal timeout As Long) As Long
    #End If
#Else
    Private Declare Function TANI_RecvBytesUntil Lib "TaniDLL32.dll" ( _
        ByVal id As Long, _
        ByRef buf As Byte, _
        ByVal bufSize As Long, _
        ByVal termByte As Byte, _
        ByVal inclTerm As Long, _
        ByVal timeout As Long) As Long
#End If
```

---

### Example

```vb
Dim buf(255) As Byte
Dim ret As Long

ret = TANI_RecvBytesUntil(0, buf(0), 256, &H23, 1, 1000)

If ret > 0 Then
    MsgBox "Bytes received: " & ret
ElseIf ret = 0 Then
    MsgBox "Timeout"
Else
    MsgBox "Receive error", vbCritical
End If
```

```delphi
var buf: array[0..255] of Byte;
TANI_RecvBytesUntil(0,@buf[0],256,$23,1,1000);
```

```cpp
unsigned char buf[256];
TANI_RecvBytesUntil(0,buf,256,0x23,1,1000);
```

---

## 5. Full Communication Sample

```cpp
unsigned char buf[256];

if (TANI_OpenEx(0,"COM5",115200,8,0,1,0) != 0) return;

TANI_SendText(0,"HELLO","\r");

TANI_RecvLine(0,buf,256,2000,13);

TANI_Close(0);
```

---

## 6. Notes

- ASCII assumed
- ID (0–15) management
- Pointer-safe design
- Offline support

---
