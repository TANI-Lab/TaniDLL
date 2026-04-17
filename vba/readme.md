# TaniDLL VBA Support

This folder contains VBA modules for using TaniDLL easily from Excel.

---

## 📦 Files

* `TaniDLL.bas` → VBA declarations for all API functions

---
## 📘 Documentation

Full API reference is available here:

👉 [TaniDLL API Reference (PDF)](docs/TANI_DLL_specification_en.pdf)

---

## 🚀 Quick Setup

### 1. Import Module

In Excel VBA:

1. Press `Alt + F11`
2. Go to **File → Import File**
3. Select `TaniDLL.bas`

---

### 2. Activate License (Required)

```vba id="x1v3d9"
Sub Test_License()
    Dim ret As Long
    ret = TANI_CheckLicense("YOUR_LICENSE_KEY")
    
    If ret <> 1 Then
        MsgBox "license error", vbCritical
    End If
End Sub
```

⚠️ Run this only once after purchasing or renewing your license.

---

### 3. Test Serial Communication

```vba id="z2gh8k"
Sub TestSerial()

  Dim ret As Long
  Dim buf(255) As Byte

  ret = TANI_OpenEx(0, "COM5", 115200, 8, 0, 1, 0)
  If ret <> 0 Then
    MsgBox "open error", vbCritical
    Exit Sub
  End If

  ret = TANI_SendText(0, "HELLO", vbCr)
  If ret <= 0 Then
    MsgBox "send error", vbCritical
    GoTo EXIT_PROC
  End If

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

## 💡 Tips

* COM port example: `"COM5"`
* Baud rate must match your device
* Always call `TANI_Close` after use

---

## ⚠️ Common Issues

* **"license error"**
  → Check your license key

* **"open error"**
  → Check COM port name and availability

* **No response**
  → Check device connection and baud rate

---

## 🔗 Related

See main README for full documentation.
