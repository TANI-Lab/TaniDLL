# TaniDLL

Easily control COM ports from Excel VBA (32/64-bit compatible).
No complex setup required.

---

## 🔧 Overview

TaniDLL is a Windows DLL library that allows simple and reliable serial (COM port) communication from:

* Excel VBA (32-bit / 64-bit)
* Delphi
* C/C++

Designed for both beginners and engineers, including Arduino / ESP32 hobbyists.

---

## 🚀 Quick Start

### 1. Purchase License

Get your license key from Gumroad:
https://imomushi.gumroad.com/l/hwwpuz

---

### 2. Install DLL

Copy the DLL to the Windows system folder:

* `TaniDLL32.dll` → `C:\Windows\SysWOW64`
* `TaniDLL64.dll` → `C:\Windows\System32`

---

### 3. VBA Declaration

```vba
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function TANI_OpenSimple Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByVal name As String, _
            ByVal BAUD As Long) As Long
    
        Private Declare PtrSafe Function TANI_OpenEx Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByVal name As String, _
            ByVal BAUD As Long, _
            ByVal databits As Long, _
            ByVal parity As Long, _
            ByVal stopbits As Long, _
            ByVal flow As Long) As Long
    
        Private Declare PtrSafe Sub TANI_Close Lib "TaniDLL64.dll" ( _
            ByVal id As Long)
    
        Private Declare PtrSafe Function TANI_SendText Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByVal txt As String, _
            ByVal TERM As String) As Long
    
        Private Declare PtrSafe Function TANI_SendBytes Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByRef data As Byte, _
            ByVal length As Long) As Long
    
        Private Declare PtrSafe Function TANI_RecvLine Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByRef buf As Byte, _
            ByVal bufSize As Long, _
            ByVal timeout As Long, _
            ByVal TERM As Byte) As Long
    
        Private Declare PtrSafe Function TANI_RecvBytesUntil Lib "TaniDLL64.dll" _
            (ByVal id As Long, _
             ByRef buf As Byte, _
             ByVal bufSize As Long, _
             ByVal termByte As Byte, _
             ByVal inclTerm As Long, _
             ByVal timeout As Long) As Long
             
        Private Declare PtrSafe Function TANI_RecvBytes Lib "TaniDLL64.dll" ( _
            ByVal id As Long, _
            ByRef buf As Byte, _
            ByVal size As Long, _
            ByVal timeout As Long) As Long
            
        Private Declare PtrSafe Function TANI_CheckLicense Lib "TaniDLL64.dll" ( _
         ByVal LicenseKey As String) As Long
    #Else
        ' 32bit版（省略）
    #End If
#End If
```

---

### 4. Activate License（First time only）

```vba
Sub Test_License()
    Dim ret As Long
    ret = TANI_CheckLicense("YOUR_LICENSE_KEY")
    
    If ret <> 1 Then
        MsgBox "license error", vbCritical
    End If
End Sub
```

* This only needs to be done once when acquiring or renewing your license.

---

### 5. Serial Communication Example

```vba
Sub TestSerial()

  Dim ret As Long
  Dim buf(255) As Byte

  ' Open
  ret = TANI_OpenEx(0, "COM5", 115200, 8, 0, 1, 0)
  If ret <> 0 Then
    MsgBox "open error", vbCritical
    Exit Sub
  End If

  ' Send
  ret = TANI_SendText(0, "HELLO", vbCr)
  If ret <= 0 Then
    MsgBox "send error", vbCritical
    GoTo EXIT_PROC
  End If

  ' Receive
  ret = TANI_RecvLine(0, buf(0), 256, 2000, 13)
  If ret > 0 Then
    Dim s As String
    s = StrConv(buf, vbUnicode)
    MsgBox "Recv: " & Left(s, ret)
  ElseIf ret = 0 Then
    MsgBox "timeout"
  Else
    MsgBox "Recv error", vbCritical
  End If

EXIT_PROC:
  TANI_Close 0

End Sub
```

---

## 💡 Features

* Easy COM port control from VBA
* 32-bit / 64-bit support
* Simple API design
* Text and binary communication supported
* Suitable for Arduino / ESP32 projects

---

## ⚠️ Troubleshooting

* **License error**
  → Check your license key

* **COM port open error**
  → Verify COM port name (e.g., "COM5")

* **No response**
  → Check baud rate and device connection

---

## 🔒 License

This software is proprietary and requires a valid subscription.

* Redistribution: Not allowed
* Reverse engineering: Prohibited

© TANI-Lab. All rights reserved.
