# TaniDLL
Simple and reliable serial communication for Excel VBA.
TaniDLL is a Windows DLL that enables simple and stable serial communication (COM port) from Excel VBA and other languages.

---

## 💡 What You Can Do

* Control Arduino / ESP32 from Excel VBA
* Send and receive text data easily
* Handle binary communication
* Build simple automation tools with VBA

---

## 📘 Full API Documentation

Complete API reference (all functions, parameters, and examples):

👉 [TaniDLL API Reference (PDF)](docs/TaniDLL_API_Reference_v1.0.pdf)

---

## 📂 Examples

Start here:

👉 `/examples/TaniDLLCheckProgram`
Full automatic test with ESP32 (recommended)

---

## ⚡ Quick Setup (Recommended)

> ⚠️ This quick test requires a serial device.
> We recommend using the provided ESP32 example program.

### 🔌 Prepare Device

Upload the following program to your ESP32:

👉 `/examples/TaniDLLCheckProgram/TaniDLL.ino.ino`

---

### ▶ Run Test

1. Import `TaniDLL.bas` into VBA
2. Run `Test_License`
3. Run `TaniDLL_AutoTest`

Done.

---
### ✅ Expected Result

You should see messages like:

[PASS] Open  
[PASS] TextRecv  
...

---

## 🔧 Installation

### Download

Purchase license and download from:

👉 https://imomushi.gumroad.com/l/hwwpuz

- Instant license key delivery
- Subscription-based access

---

### DLL Placement

#### Option 1 (Recommended)

Place the DLL in the same folder as your Excel file.

#### Option 2 (Advanced)

Copy to system folder:

* 32-bit → `C:\Windows\SysWOW64`
* 64-bit → `C:\Windows\System32`

---

### VBA Setup

Import:

```id="y2i0xp"
/vba/TaniDLL.bas
```

---

## 🔒 License

* Internet connection is required for license activation
* License verification is performed via a remote server
* Activation is required only once (or when renewing)
* After activation, the license is stored locally

---

## 🚀 Quick Example
This example sends "HELLO" and receives a response from the device.

```vba id="ywnjgi"
Sub TestSerial()

  Dim ret As Long
  Dim buf(255) As Byte

  ret = TANI_OpenEx(0, "COM5", 115200, 8, 0, 1, 0)
  If ret <> 0 Then
    MsgBox "open error", vbCritical
    Exit Sub
  End If

  ret = TANI_SendText(0, "HELLO", vbCr)

  ret = TANI_RecvLine(0, buf(0), 256, 2000, 13)
  If ret > 0 Then
    Dim s As String
    s = StrConv(buf, vbUnicode)
    MsgBox "Recv: " & Left(s, ret)
  End If

  TANI_Close 0

End Sub
```

---

## ⚠️ Notes

* Make sure to update COM port (e.g., COM5)
* Baud rate must match your device (e.g., 115200)
* Serial device must be connected before running
* License activation requires internet connection

---

## 📄 License

This software is distributed under a proprietary license.
Source code is not included.

---

## 👨‍💻 Author

TANI-Lab
