/*
  RAPID aux output bench test
  ===========================

  Drives D7, D8, D12 and D13 individually so each can be metered against GND.
  Intended for checking that a 5 V AVR board (Uno, Nano, Mega, Metro 328) can
  stand in for the PCI-DAS6030 AUXPORT lines that carry the vacuum controls.

  Boot state
  ----------
  AVR pins come up high-impedance and stay that way until pinMode() runs. On a
  board with the stock bootloader that window is one to two seconds after every
  reset -- and opening the USB serial port asserts DTR, which resets the board.
  So every managed pin is driven LOW here before anything else, and the PORT
  latch is cleared before the driver is enabled so the pin cannot glitch high
  on the way.

  This matters beyond the bench: on the real system a floating line during that
  window is what decides whether a reset drops the sample. Pull-downs (or
  pull-ups, whichever corresponds to "safe" at the vacuum box) belong on the
  wiring, not in the sketch.

  Protocol
  --------
  115200 8N1, newline terminated. Every command answers exactly one line.

    *IDN?           -> RAPID-AUX-TEST v1
    PIN <n> <0|1>   -> OK <n> <0|1>          n in {7, 8, 12, 13}
    ALL 0           -> OK ALL 0              every managed pin low
    STAT?           -> 7=0 8=0 12=0 13=0

  Anything else answers ERR and changes no pin.

  Note that D13 also drives the board's onboard LED, which is useful here: it
  gives visual confirmation alongside the meter reading.
*/

const uint8_t PINS[] = {7, 8, 12, 13};
const uint8_t N_PINS = sizeof(PINS) / sizeof(PINS[0]);

const unsigned long BAUD = 115200;

char line[24];
uint8_t lineLen = 0;

bool isManaged(long pin) {
  for (uint8_t i = 0; i < N_PINS; i++) {
    if ((long)PINS[i] == pin) return true;
  }
  return false;
}

void allLow() {
  for (uint8_t i = 0; i < N_PINS; i++) digitalWrite(PINS[i], LOW);
}

void reportStatus() {
  for (uint8_t i = 0; i < N_PINS; i++) {
    Serial.print(PINS[i]);
    Serial.print('=');
    Serial.print(digitalRead(PINS[i]) ? '1' : '0');
    if (i + 1 < N_PINS) Serial.print(' ');
  }
  Serial.println();
}

void handleLine(char *s) {
  if (strcmp(s, "*IDN?") == 0) {
    Serial.println(F("RAPID-AUX-TEST v1"));
    return;
  }

  if (strcmp(s, "STAT?") == 0) {
    reportStatus();
    return;
  }

  if (strcmp(s, "ALL 0") == 0) {
    allLow();
    Serial.println(F("OK ALL 0"));
    return;
  }

  if (strncmp(s, "PIN ", 4) == 0) {
    char *pinText = s + 4;
    char *space = strchr(pinText, ' ');
    if (space != NULL) {
      *space = '\0';
      long pin = atol(pinText);
      long value = atol(space + 1);
      if (isManaged(pin) && (value == 0 || value == 1)) {
        digitalWrite((uint8_t)pin, value ? HIGH : LOW);
        Serial.print(F("OK "));
        Serial.print(pin);
        Serial.print(' ');
        Serial.println(value);
        return;
      }
    }
  }

  Serial.println(F("ERR"));
}

void setup() {
  for (uint8_t i = 0; i < N_PINS; i++) {
    digitalWrite(PINS[i], LOW);   // clear the latch first,
    pinMode(PINS[i], OUTPUT);     // then enable the driver -- no glitch
  }

  Serial.begin(BAUD);
  while (!Serial) {
    ;                             // needed on Leonardo/Micro, no-op on Uno
  }
  Serial.println(F("RAPID-AUX-TEST v1 ready"));
}

void loop() {
  while (Serial.available() > 0) {
    char c = (char)Serial.read();

    if (c == '\r') continue;

    if (c == '\n') {
      line[lineLen] = '\0';
      if (lineLen > 0) handleLine(line);
      lineLen = 0;
    } else if (lineLen < sizeof(line) - 1) {
      line[lineLen++] = c;
    } else {
      lineLen = 0;                // overlong input, discard the whole line
      Serial.println(F("ERR too long"));
    }
  }
}
