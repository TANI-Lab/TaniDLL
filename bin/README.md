# TaniDLL - Binary Files

This folder contains the compiled DLL files for TaniDLL.

---

## 📦 Files

* `TaniDLL32.dll` → For 32-bit applications (Excel 32-bit)
* `TaniDLL64.dll` → For 64-bit applications (Excel 64-bit)

---

## 📥 Installation

### Option 1 (Recommended)

Place the DLL in the same folder as your Excel file.

---

### Option 2 (System-wide)

Copy the DLL to:

* 32-bit → `C:\Windows\SysWOW64`
* 64-bit → `C:\Windows\System32`

---

## ⚠️ Important

* Make sure to use the correct DLL for your Excel version
* Administrator privileges may be required for system folder installation

---

## ❓ How to check Excel bit version

In Excel:

1. Go to **File**
2. Select **Account**
3. Click **About Excel**

You will see either:

* 32-bit
* 64-bit

---

## 🚫 Common Mistakes

* Using 32-bit DLL with 64-bit Excel (or vice versa)
* Placing DLL in the wrong folder
* Forgetting to match VBA declaration with DLL version

---

## 🔗 See also

Please refer to the main README for usage instructions.
