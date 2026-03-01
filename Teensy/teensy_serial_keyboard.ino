// Teensy (Arduino) sketch: USB = Serial + Keyboard
// Tools: set USB Type to "Serial + Keyboard" in Teensyduino

#include <Arduino.h>
#include <Keyboard.h>

void setup() {
  Serial.begin(115200);
  delay(200);
  Serial.println("TEENSY_ID:VKEY");
}

void pressAndReleaseKey(uint16_t keyCode, int dur) {
  Keyboard.press(keyCode);
  delay(max(0, dur));
  Keyboard.release(keyCode);
}

void handleKeyCommand(String s) {
  // format: KEY:<keySpec>:<duration>
  s.trim();
  if (!s.startsWith("KEY:")) return;
  String body = s.substring(4);
  int p = body.indexOf(':');
  String key = p >= 0 ? body.substring(0, p) : body;
  int dur = p >= 0 ? body.substring(p+1).toInt() : 50;
  key.trim();
  if (key.length() == 1) {
    char c = key.charAt(0);
    Keyboard.press(c);
    delay(max(0, dur));
    Keyboard.release(c);
  } else {
    if (key.equalsIgnoreCase("ENTER")) {
      pressAndReleaseKey(KEY_ENTER, dur);
    } else if (key.equalsIgnoreCase("TAB")) {
      pressAndReleaseKey(KEY_TAB, dur);
    }
    // add more named keys as needed
  }
}

void handleChordCommand(String s) {
  // format: CHORD:A+B+C:duration
  s.trim();
  if (!s.startsWith("CHORD:")) return;
  String body = s.substring(6);
  int p = body.lastIndexOf(':');
  String keys = p >= 0 ? body.substring(0, p) : body;
  int dur = p >= 0 ? body.substring(p+1).toInt() : 100;
  // press all
  int i = 0;
  int len = (int)keys.length();
  while (i < len) {
    int j = keys.indexOf('+', i);
    String k = j < 0 ? keys.substring(i) : keys.substring(i, j);
    k.trim();
    if (k.length() == 1) Keyboard.press(k.charAt(0));
    i = j < 0 ? len : j + 1;
  }
  delay(max(0, dur));
  // release all
  i = 0;
  while (i < len) {
    int j = keys.indexOf('+', i);
    String k = j < 0 ? keys.substring(i) : keys.substring(i, j);
    k.trim();
    if (k.length() == 1) Keyboard.release(k.charAt(0));
    i = j < 0 ? len : j + 1;
  }
}

String rxBuf = "";
void loop() {
  while (Serial.available()) {
    char c = Serial.read();
    if (c == '\n') {
      String line = rxBuf;
      rxBuf = "";
      line.trim();
      if (line.length() == 0) continue;
      if (line.equalsIgnoreCase("PING")) {
        Serial.println("PONG");
      } else if (line.startsWith("KEY:") ) {
        handleKeyCommand(line);
      } else if (line.startsWith("CHORD:")) {
        handleChordCommand(line);
      }
    } else if (c >= 32) {
      rxBuf += c;
    }
  }
}
