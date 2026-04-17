## 🔌 ESP32 Communication Example

This example demonstrates full communication testing between Excel VBA and ESP32.

---

## 📡 ESP32 Setup (Arduino IDE)

Upload the following code to your ESP32:

```cpp
// ===== TaniDLL Test ESP32 =====

String lineBuffer = "";

void setup() {
  Serial.begin(115200);
}

void loop() {
  while (Serial.available()) {
    char c = Serial.read();

    if (c == '\n') {
      processLine(lineBuffer);
      lineBuffer = "";
    } else {
      lineBuffer += c;
    }
  }
}

void processLine(String cmd) {

  if (cmd == "PING") {
    Serial.println("PONG");
  }

  else if (cmd == "BINTEST") {
    uint8_t data[4] = {1, 2, 3, 4};
    Serial.write(data, 4);
  }

  else if (cmd == "TERMTEST") {
    Serial.write("HELLO");
    Serial.write(0x23);
  }

  else {
    Serial.print("ECHO:");
    Serial.println(cmd);
  }
}
```

---

## 💻 VBA Test Program

Run the following macro:

```vba
Sub TaniDLL_AutoTest()
```

---

## ✅ What This Test Covers

* Text communication (PING → PONG)
* Binary data transfer
* Terminator-based receive
* Timeout handling
* Buffer limit behavior
* Continuous communication (stress test)
* Invalid ID handling

---

## 🎯 Expected Output (Immediate Window)

```
==== AUTO TEST START ====
[PASS]        Open           0 
[PASS]        TextRecv       5 
  Data:       PONG
[PASS]        BinaryRecv     4 
[PASS]        Terminator     5 
[PASS]        Timeout        0 
[PASS]        BufferLimit    10 
[PASS]        CR Check       13 
[PASS]        Continuous     100 
[PASS]        InvalidID     -1 
==== AUTO TEST END ====
```

---

## ⚙️ Configuration

Update these values if needed:

```vba
Const COM_NAME As String = "COM5"
Const BAUD As Long = 115200
```

---

## ⚠️ Notes

* ESP32 must be connected via USB
* Baud rate must match (115200)
* License activation is required before testing
* Internet connection is required for license activation

---

## 💡 Tip

This example acts as a full validation tool for TaniDLL.

If all tests pass, your environment is correctly configured.

