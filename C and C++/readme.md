# TANI DLL - Serial Communication Library for Windows

A simple and powerful DLL for serial (COM port) communication from Excel VBA, C/C++, and Delphi.  
Supports both 32-bit and 64-bit environments.

---

## Features

- Easy COM port communication from VBA / C / Delphi
- Full 32-bit / 64-bit support
- Works with Arduino / ESP32
- Text and binary communication
- Up to **16 COM ports simultaneously** (ID: 0–15)
- Timeout and terminator-based receiving
- Offline-capable license system

---

## Header File (C/C++)

```c
#ifndef TANI_DLL_H
#define TANI_DLL_H

#ifdef __cplusplus
extern "C" {
#endif

#include <windows.h>

#ifdef TANI_DLL_EXPORTS
#define TANI_API __declspec(dllexport)
#else
#define TANI_API __declspec(dllimport)
#endif

// Return Codes
#define TANI_OK             0
#define TANI_OK_LICENSE     1
#define TANI_ERR           -1
#define RESULT_BLOCKED     -2
#define TANI_ERR_LICENSE  -10

// Communication Settings
#define TANI_PARITY_NONE    0
#define TANI_PARITY_ODD     1
#define TANI_PARITY_EVEN    2

#define TANI_STOP_1         1
#define TANI_STOP_2         2

#define TANI_FLOW_NONE      0
#define TANI_FLOW_RTSCTS    1
#define TANI_FLOW_XONXOFF   2

// License
#define LICENSE_CACHE_MIN   60

// Function Prototypes

// License
TANI_API int __stdcall TANI_CheckLicense(const char* LicenseKey);

// Open / Close
TANI_API int __stdcall TANI_OpenSimple(
    int id,
    const char* name,
    int baud
);

TANI_API int __stdcall TANI_OpenEx(
    int id,
    const char* name,
    int baud,
    int databits,
    int parity,
    int stopbits,
    int flow
);

TANI_API void __stdcall TANI_Close(int id);

// Send
TANI_API int __stdcall TANI_SendText(
    int id,
    const char* txt,
    const char* term
);

TANI_API int __stdcall TANI_SendBytes(
    int id,
    const unsigned char* data,
    int length
);

// Receive
TANI_API int __stdcall TANI_RecvLine(
    int id,
    unsigned char* buf,
    int bufSize,
    int timeout,
    unsigned char term
);

TANI_API int __stdcall TANI_RecvBytes(
    int id,
    unsigned char* buf,
    int size,
    int timeout
);

TANI_API int __stdcall TANI_RecvBytesUntil(
    int id,
    unsigned char* buf,
    int bufSize,
    unsigned char termByte,
    int inclTerm,
    int timeout
);

#ifdef __cplusplus
}
#endif

#endif



