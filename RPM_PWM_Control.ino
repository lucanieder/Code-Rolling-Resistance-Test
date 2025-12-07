#include <Servo.h>

Servo esc;

// ---- Set up ----
const int hallPin = 2;  //Pin to hallsensor
const int escPin = 5;   //Pin to ESC controller
volatile unsigned long pulseCount = 0;

unsigned long lastRPMTime = 0;
float measuredRPM = 0;

// ---- Calibrated motor values ----
int targetRPM = 100;    //initial target RPM
int escValue  = 1100;
int escValRe  = 1000;
int escValN   = 1100;   //motor in neutral

const int pulsesPerRevolution = 3;

// Status
bool motorStarted = false;
bool motorEnabled = false;
bool manualEscMode = false;    

void hallISR() {
  pulseCount++;   
}

// --------------------------------------------------------
//                SERIELLE COMMANDS
// --------------------------------------------------------
void handleSerialCommands() {
  if (!Serial.available()) return;

  String cmd = Serial.readStringUntil('\n');
  cmd.trim();


  // ---- Start ----
  if (cmd.equalsIgnoreCase("start")) {
    motorEnabled = true;
    motorStarted = false;
    escValue =  escValN;
    Serial.println("Motor wird gestartet…");
    return;
  }

  // ---- Stop ----
  if (cmd.equalsIgnoreCase("stop")) {
    motorEnabled = false;
    manualEscMode = false;
    escValue = escValN;
    esc.writeMicroseconds(1000);
    Serial.println("Motor gestoppt.");
    return;
  }

  // ---- Reset ----
  if (cmd.equalsIgnoreCase("reset")) {
    escValue = escValN;
    motorStarted = false;
    Serial.println("ESC & Werte wurden zurückgesetzt.");
    return;
  }

  // ---- Switch: ESC-direct control mode ----
  if (cmd.equalsIgnoreCase("mode esc")) {
    manualEscMode = true;
    Serial.println("Modus: Direkte ESC-Steuerung aktiviert.");
    return;
  }

  // ---- Switch: RPM-Regulation mode ----
  if (cmd.equalsIgnoreCase("mode rpm")) {
    manualEscMode = false;
    Serial.println("Modus: RPM-Regelung aktiviert.");
    return;
  }

  // ---- Set esc value ----
  if (cmd.startsWith("esc")) {
    int v = cmd.substring(3).toInt();
    if (v >= 1000 && v <= 2000) {
      escValue = v;
      manualEscMode = true;
      Serial.print("Direkter ESC-Wert gesetzt: ");
      Serial.println(escValue);
    } else {
      Serial.println("Fehler: ESC-Wert muss 1000–2000 sein.");
    }
    return;
  }

  // ---- Set final RPM ----
  if (cmd.startsWith("rpm")) {
    int newRPM = cmd.substring(3).toInt();
    if (newRPM > 0) {
      targetRPM = newRPM;
      Serial.print("Neue Ziel-RPM: ");
      Serial.println(targetRPM);
    } else {
      Serial.println("Ungültige RPM Eingabe.");
    }
    return;
  }
}


// --------------------------------------------------------
//                        SETUP
// --------------------------------------------------------
void setup() {
  Serial.begin(9600);

  esc.attach(escPin, 1000, 2000);
  esc.writeMicroseconds(1000);
  delay(3000);

  pinMode(hallPin, INPUT_PULLUP);
  attachInterrupt(digitalPinToInterrupt(hallPin), hallISR, RISING);

  Serial.println("System gestartet...");
  Serial.println("Befehle:");
  Serial.println(" start | stop | reset");
  Serial.println(" rpm XXXX  (Regelmodus)");
  Serial.println(" mode esc | mode rpm");
  Serial.println(" esc XXXX  (ESC-Direktsteuerung)");
}


// --------------------------------------------------------
//                         MAIN LOOP
// --------------------------------------------------------
void loop() {

  handleSerialCommands();

  unsigned long currentTime = millis();

  if (!motorEnabled) {
    escValue = escValN;
    esc.writeMicroseconds(escValue);

    if (currentTime - lastRPMTime >= 500) {
      Serial.print("Motor aus | ESC: ");
      Serial.println(escValue);
      lastRPMTime = currentTime;
    }
    return;
  }

  // ---- Direct ESC-mode ----
  if (manualEscMode) {
    esc.writeMicroseconds(escValue);

    if (currentTime - lastRPMTime >= 500) {
      Serial.print("MANUELL | ESC: ");
      Serial.println(escValue);
      lastRPMTime = currentTime;
    }
    return;
  }


  // ---- RPM-mode ----
  if (currentTime - lastRPMTime >= 200) {

    noInterrupts();
    unsigned long pulses = pulseCount;
    pulseCount = 0;
    interrupts();

    measuredRPM = (pulses)* 10/ pulsesPerRevolution;

    Serial.print("RPM: ");
    Serial.print(measuredRPM);
    Serial.print(" | ESC µs: ");
    Serial.println(escValue);

    // ---- Soft Start ----
    if (!motorStarted) {

      if (measuredRPM < 100 && escValue < 2000) {
        escValue += 10;
      } else if (measuredRPM >= 100) {
        motorStarted = true;
        Serial.println("Motor gestartet — Regelung aktiv.");
      }

      escValue = constrain(escValue, 1000, 2000);
      esc.writeMicroseconds(escValue);
      lastRPMTime = currentTime;
      return;
    }

    // ---- RPM-control ----
    int error = targetRPM - measuredRPM;

    if (error > 1) escValue += 5;
    else if (error < -1) escValue -= 5;

    escValue = constrain(escValue, 1000, 2000);
    esc.writeMicroseconds(escValue);

    lastRPMTime = currentTime;
  }
}
