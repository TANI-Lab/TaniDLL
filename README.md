# TaniDLL

A stable multi-COM port communication DLL for Windows, designed to prevent common issues such as data loss, freezing, and unstable connections in Excel VBA and serial communication applications.

---

## 🚀 Why TaniDLL?

- Stable communication even in long-running applications  
- Reliable handling of multiple COM ports  
- Simple API for fast integration  
- Works seamlessly with Excel VBA  
- Lightweight and easy to deploy  

---

## 🎥 Demo

See TaniDLL in action (Excel + COM communication):

▶ https://youtube.com/shorts/4sBXty-INws

- Excel VBA Serial Communication Demo
- Multi COM Port Control from Excel

📊 See also: Stability test (7+ days continuous operation)

---

## ⏱️ Stability Test

The system has been running continuously for over 7 days with:

- No memory leaks observed
- Stable COM port communication
- No data loss detected

Test environment:
- Multiple ESP32 devices (3 units)
- Continuous real-time data logging
- Excel VBA + TaniDLL

Long-term testing is still ongoing and will be updated.

---

## ❗ Common Problems This Solves

- Serial communication becomes unstable over time  
- Multiple COM ports cause unexpected behavior  
- Reconnection logic is complex and error-prone  
- Data loss occurs in long-running applications  

---

## 🎯 Who is this for?

- Excel VBA developers working with serial devices  
- Engineers controlling Arduino / ESP32 from Windows  
- Developers building automation or measurement tools  

---

## 💡 What You Can Do

- Control Arduino / ESP32 from Excel VBA  
- Send and receive text data easily  
- Handle binary communication  
- Build automation tools with VBA  

---

## ⚡ Quick Start

> You can test with any serial device (USB-UART, Arduino, ESP32, etc.)

### 1. Import VBA Module

```
/vba/TaniDLL.bas
```
### 2.🔑 License Activation

Before using this software, you must activate your license.

If the license is not activated, the COM port cannot be opened.

```vba
Sub CheckLicense()

    Dim ret As Long

    ' If you enter the wrong key 5 times within 5 minutes,
    ' authentication will be locked for 1 hour.
    ret = TANI_CheckLicense("YOUR_LICENSE_KEY")

    Select Case ret
        Case 1
            Debug.Print "License OK"

        Case 0
            MsgBox "License authentication failed." & vbCrLf & _
                   "Please check your license key."
            Exit Sub

        Case -1
            MsgBox "Failed to connect to the server." & vbCrLf & _
                   "Please check your network connection."
            Exit Sub
    End Select

End Sub
```

### 3. Run Sample Code

```vba
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

## 📘 Documentation

👉 [Download API Reference (PDF)](https://raw.githubusercontent.com/TANI-Lab/TaniDLL/main/docs/TANI_DLL_specification_en.pdf)

*Note: The PDF may not render correctly on GitHub. Please download to view.*

---

## 📂 Examples

👉 [TaniDLLCheckProgram](https://github.com/TANI-Lab/TaniDLL/tree/main/examples/TaniDLLCheckProgram)

Includes full automatic test using ESP32

---

## 🔧 Installation

### Download

👉 [https://imomushi.gumroad.com/l/ixdvo](https://imomushi.gumroad.com/l/ixdvo)

- Free version available (2 COM port)  
- One-time purchase licenses (no subscription)  
- Instant license key delivery via email  

---

### License Options

- **Free (Personal)**  
  - 2 COM port  
  - For evaluation and personal use only  

- **Personal License ($29)**  
  - Up to 4 COM ports  
  - Non-commercial use only
  - 👉[https://imomushi.gumroad.com/l/wbaiq](https://imomushi.gumroad.com/l/wbaiq)

- **Business License ($129)**  
  - Up to 16 COM ports  
  - Commercial use allowed
  - 👉[https://imomushi.gumroad.com/l/fzumby](https://imomushi.gumroad.com/l/fzumby)  

---

### DLL Placement

**Recommended**  
Place the DLL in the same folder as your Excel file.

**Alternative**  
- 32-bit → C:\Windows\SysWOW64  
- 64-bit → C:\Windows\System32  

---

## 🔒 License

- A license key is required  
- License verification is performed via a remote server  
- Internet connection is required for activation  
- Activation is performed once per machine (HWID-based)  
- After activation, the license is stored locally  

---

## ⚠️ Disclaimer

This project is under active development and may contain bugs.

The software is provided "as is" without any warranties. Please use it at your own risk.

---

### Notes

- Each license is valid for one hardware device (HWID)  
- Hardware changes may require reactivation  
- License reset is available via support  

---

## 🛡 Stability

- Designed for continuous operation  
- Suitable for production environments  
- Multi-port communication ready  

---

## ⚠️ Notes

- COM port must be correctly configured (e.g., COM5)  
- Baud rate must match your device  
- Device must be connected before execution  
- Internet is required for initial activation  

---

## 📄 License

This software is distributed under a proprietary license.  
Source code is not included.

---

## 👨‍💻 Author

TANI-Lab
