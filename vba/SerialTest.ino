// ===== TaniDLL Test (CR Terminator) =====

String lineBuffer = "";

void setup() {
  Serial.begin(115200);
}

void loop() {
  while (Serial.available()) {
    char c = Serial.read();

    // CR (0x0D)で受信完了
    if (c == '\r') {
      processLine(lineBuffer);
      lineBuffer = "";
    } else {
      lineBuffer += c;
    }
  }
}

// コマンド処理
void processLine(String cmd) {

  // デバッグ用（必要なら）
  // Serial.println("RX:" + cmd);

  if (cmd == "HELLO") {
    Serial.print("WORLD");
    Serial.write('\r');  // ★CRで終了（重要）
  }
  else {
    Serial.print("ECHO:");
    Serial.print(cmd);
    Serial.write('\r');  // ★必ずCRで終わる
  }
}
