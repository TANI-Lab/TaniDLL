void setup() {
  Serial.begin(115200);
  randomSeed(analogRead(0));
}

void loop() {
  static int count = 0;

  // ゆっくり変動する“現実っぽい温度”
  static float temp = 25.0;

  temp += (random(-5, 6) / 10.0);  // -0.5〜+0.5変動

  // 少しだけ制限（現実的レンジ）
  if (temp < 20.0) temp = 20.0;
  if (temp > 30.0) temp = 30.0;

  Serial.print("SENSOR:1,");
  Serial.print("TEMP:");
  Serial.print(temp, 1);
  Serial.print(",COUNT:");
  Serial.println(count++);

  delay(500);
}