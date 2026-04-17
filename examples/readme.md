# TaniDLL Examples

This folder contains simple VBA examples for using TaniDLL.

---

## 📦 Files

* `SerialTest.bas` → Basic send/receive example

---

## 🚀 How to Use

### 1. Import Files

In Excel VBA:

1. Press `Alt + F11`
2. Import the following files:

   * `TaniDLL.bas` (from `/vba`)
   * `SerialTest.bas` (from this folder)

---

### 2. Activate License

```vba id="ex1"
Sub Test_License()
    Dim ret As Long
    ret = TANI_CheckLicense("YOUR_LICENSE_KEY")
    
    If ret <> 1 Then
        MsgBox "license error", vbCritical
    End If
End Sub
```

---

### 3. Run Example

Run:

```vba id="ex2"
Sub TestSerial()
```

---

## 💡 What This Example Does

* Opens COM port (`COM5`)
* Sends text: `"HELLO"`
* Waits for response
* Displays received data

---

## ⚙️ Configuration

You may need to change:

```vba id="ex3"
"COM5"      ' Change to your COM port
115200      ' Match your device baud rate
```

---

## ⚠️ Notes

* Make sure your device is connected
* License activation is required before use
* Always close the port after use

---

## 🔧 Expected Behavior

If everything is correct:

* A message box will show the received data
* No errors will occur

---

## ❌ Common Issues

* No response
  → Check COM port and baud rate

* Timeout
  → Device may not be sending data

* Open error
  → Port may already be in use

---

## 🔗 Related

* `/vba` → VBA module (required)
* Main README → Full setup guide
