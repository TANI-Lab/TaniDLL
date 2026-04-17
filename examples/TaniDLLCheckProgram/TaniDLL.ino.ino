// ===== TaniDLL テスト用ESP32 =====

String lineBuffer = "";

void setup() {
  Serial.begin(115200);
}

void loop() {
  // --- テキスト受信 ---
  while (Serial.available()) {
    char c = Serial.read();

    // 改行で1行処理
    if (c == '\n') {
      processLine(lineBuffer);
      lineBuffer = "";
    } else {
      lineBuffer += c;
    }
  }
}

// コマンド処理
void processLine(String cmd) {

  if (cmd == "PING") {
    Serial.println("PONG");
  }

  else if (cmd == "BINTEST") {
    // バイナリ送信（例: 0x01 0x02 0x03 0x04）
    uint8_t data[4] = {1, 2, 3, 4};
    Serial.write(data, 4);
  }

  else if (cmd == "TERMTEST") {
    // 0x0Dで終わるデータ
    Serial.write("HELLO");
    Serial.write(0x23);
  }

  else {
    Serial.print("ECHO:");
    Serial.println(cmd);
  }
}