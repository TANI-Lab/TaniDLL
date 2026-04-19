void setup() {
  Serial.begin(115200);
  randomSeed(analogRead(0));
}

void loop() {
  static int count = 0;

  static float value = 60.0;

  value += (random(-8, 9) / 10.0);  // -0.8〜+0.8

  if (value < 40.0) value = 40.0;
  if (value > 80.0) value = 80.0;

  Serial.print("SENSOR:2,");
  Serial.print("HUM:");
  Serial.print(value, 1);
  Serial.print(",COUNT:");
  Serial.println(count++);

  delay(500);
}