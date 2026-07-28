/*
 * ================================================================
 *  ESP32-C6 Weather Pod  -  Temperature / Humidity / Battery
 * ================================================================
 *  Board : Waveshare ESP32-C6-Zero  (select "ESP32C6 Dev Module")
 *  Sensor: Adafruit SHT45 over I2C (fixed address 0x44)
 *  Power : Samsung INR18650-35E via protected TP4056 -> board VIN
 *
 *  Behaviour:
 *   - Wakes every 5 minutes from deep sleep, always.
 *   - Before this pod has EVER appeared on the sheet: each wake takes one
 *     reading and makes a single BOUNDED attempt (WiFi timeout ~20s) to
 *     send it as its own one-row upload. Success or failure, it goes back
 *     to sleep either way - a failed attempt is discarded, not queued, so
 *     
 *     the eventual successful send always contains EXACTLY one row. No
 *     infinite loop, no LED - just a plain retry once every 5 minutes
 *     until it gets through.
 *   - After that first success, normal sampling takes over: reads the
 *     SHT45 as a short AVERAGED burst (CPU underclocked), reads battery
 *     voltage via a 2:1 divider, timestamps from the RTC-advanced clock,
 *     and buffers readings in RTC memory (survives deep sleep).
 *   - WiFi is touched ONLY twice in the whole lifecycle: once for that
 *     first registration row, and then once every ~6 hours (72 queued
 *     readings) for the normal batch upload - never anywhere else. On
 *     success the buffer is wiped so the count restarts; on failure the
 *     batch is kept and retried on the next wake.
 *   - RGB LED is turned off to save power.
 *   - A magnet on the reed switch wakes the pod into a BLE "config mode"
 *     where the app sets the name, intervals, WiFi and Sheets URL. Config
 *     lives in flash (NVS) and survives re-flashing and power loss.
 *
 *  Libraries to install (Arduino Library Manager):
 *   - "Adafruit SHT4x Library"        (for the SHT45)
 *   - "Adafruit Unified Sensor"       (dependency, installs alongside)
 *   - "ArduinoJson"  (v7)             (parses the BLE config payload)
 *   WiFi, HTTPClient, BLE and Preferences ship with the ESP32 Arduino
 *   core (use core v3.x for C6; its BLE headers are backed by NimBLE).
 * ================================================================
 */

#include <Arduino.h>
#include <Wire.h>
#include <Adafruit_SHT4x.h>
#include <WiFi.h>
#include <WiFiClientSecure.h>
#include <HTTPClient.h>
#include <time.h>
#include <Preferences.h>
#include <ArduinoJson.h>
#include <BLEDevice.h>
#include <BLEServer.h>
#include <BLEUtils.h>

// ------------------- DEFAULTS (used on the FIRST boot only) -------------------
#define DEFAULT_NAME        "DEWVEE:02"
#define DEFAULT_SSID        "Airtel_The Liana Trust"
#define DEFAULT_PASS        ""   // set via BLE config on first use
#define DEFAULT_URL         "https://script.google.com/macros/s/AKfycbw88DpkYijK98qTzhlSguLAFxFRQI_H3KXreYcRp2sitSr3XWFPRz9QwR88ocdsRsO7/exec"
#define DEFAULT_SAMPLE_MIN  5
#define DEFAULT_UPLOAD_HRS  6

struct Config {
  String   name;
  uint16_t sampleMin;
  uint16_t uploadHrs;
  bool     uploadsEnabled;
  String   ssid;
  String   pass;
  String   url;
} cfg;

Preferences prefs;

// --- BLE config service UUIDs (must match the web app) ---
#define SERVICE_UUID "d5770001-9e0b-4b7a-9c1e-2f3a4b5c6d7e"
#define CONFIG_UUID  "d5770002-9e0b-4b7a-9c1e-2f3a4b5c6d7e"
#define STATUS_UUID  "d5770003-9e0b-4b7a-9c1e-2f3a4b5c6d7e"
const uint32_t CONFIG_MODE_MS = 180000;   // 3 min BLE window
BLECharacteristic* configChar = nullptr;
BLECharacteristic* statusChar = nullptr;
volatile bool configSaved = false;

// --- pins ---
const int SDA_PIN   = 6;
const int SCL_PIN   = 5;
const int VBAT_PIN  = 1;
const int REED_PIN  = 2;    // Terminal A -> GND, Terminal B -> GPIO2
                            // Internal pull-up (~45kΩ), no external resistor needed.

// Reed wake is now ENABLED. Set to 0 to disable if debugging without the switch wired.
#define ENABLE_REED_WAKE 1

// --- sensor / battery ---
const float VBAT_DIVIDER  = 2.0f;
const int   BURST_SAMPLES = 10;

#define uS_PER_MIN 60000000ULL

// ---- RTC-persisted state (survives deep sleep) ----
typedef struct {
  uint32_t epoch;
  float    temp;
  float    hum;
  float    vbat;
  uint8_t  pct;
} Reading;

#define BUFFER_CAPACITY 216
RTC_DATA_ATTR Reading  buffer[BUFFER_CAPACITY];
RTC_DATA_ATTR uint16_t bufCount      = 0;
RTC_DATA_ATTR uint32_t bootCount     = 0;
RTC_DATA_ATTR bool     timeSynced    = false;
RTC_DATA_ATTR uint32_t lastSyncEpoch = 0;
RTC_DATA_ATTR bool     registered    = false;
RTC_DATA_ATTR float    lastVbat      = 0;
RTC_DATA_ATTR uint8_t  lastPct       = 0;

Adafruit_SHT4x sht = Adafruit_SHT4x();

// ---------------- helpers ----------------

void ledOff() {
#ifdef RGB_BUILTIN
  rgbLedWrite(RGB_BUILTIN, 0, 0, 0);
#endif
}

void goToSleep() {
  WiFi.mode(WIFI_OFF);
  Serial.flush();
  esp_sleep_enable_timer_wakeup((uint64_t)cfg.sampleMin * uS_PER_MIN);
#if ENABLE_REED_WAKE
  esp_deep_sleep_enable_gpio_wakeup(1ULL << REED_PIN, ESP_GPIO_WAKEUP_GPIO_LOW);
#endif
  esp_deep_sleep_start();
}

bool readSensor(float &tOut, float &hOut) {
  sensors_event_t humidity, temp;
  float tSum = 0, hSum = 0;
  int   n = 0;
  for (int i = 0; i < BURST_SAMPLES; i++) {
    sht.getEvent(&humidity, &temp);
    float t = temp.temperature;
    float h = humidity.relative_humidity;
    if (!isnan(t) && !isnan(h)) { tSum += t; hSum += h; n++; }
    delay(50);
  }
  if (n == 0) return false;
  tOut = tSum / n;
  hOut = hSum / n;
  return true;
}

// VBAT calibration history:
//   v1: single sample, factor = 4.22/3.672 = 1.149 (13% error)
//   v2: 32-sample averaged, factor over-corrected -> showed 4.700V at full charge
//   v3: recalibrated -> pre-cal reading = 4.700/1.149 = 4.090V, new factor = 4.22/4.090
const float VBAT_CALIBRATION = 4.22f / 4.090f;   // ~1.032
float readBatteryVolts() {
  analogSetPinAttenuation(VBAT_PIN, ADC_11db);
  const int N = 32;
  uint32_t samples[N];
  for (int i = 0; i < N; i++) { samples[i] = analogReadMilliVolts(VBAT_PIN); delay(2); }
  for (int i = 1; i < N; i++) {
    uint32_t key = samples[i]; int j = i - 1;
    while (j >= 0 && samples[j] > key) { samples[j + 1] = samples[j]; j--; }
    samples[j + 1] = key;
  }
  const int trim = N / 4;
  uint32_t sum = 0;
  for (int i = trim; i < N - trim; i++) sum += samples[i];
  uint32_t mv = sum / (N - 2 * trim);
  return (mv / 1000.0f) * VBAT_DIVIDER * VBAT_CALIBRATION;
}

int batteryPercent(float v) {
  if (v >= 4.20f) return 100;
  if (v >= 4.10f) return 90;
  if (v >= 4.00f) return 80;
  if (v >= 3.90f) return 70;
  if (v >= 3.80f) return 60;
  if (v >= 3.75f) return 50;
  if (v >= 3.70f) return 40;
  if (v >= 3.65f) return 30;
  if (v >= 3.60f) return 20;
  if (v >= 3.50f) return 12;
  if (v >= 3.40f) return 6;
  if (v >= 3.30f) return 3;
  return 0;
}

void pushReading(uint32_t epoch, float t, float h, float v, uint8_t pct) {
  if (bufCount >= BUFFER_CAPACITY) {
    memmove(&buffer[0], &buffer[1], (BUFFER_CAPACITY - 1) * sizeof(Reading));
    bufCount = BUFFER_CAPACITY - 1;
  }
  buffer[bufCount++] = { epoch, t, h, v, pct };
}

bool connectWiFi(uint32_t timeoutMs) {
  WiFi.mode(WIFI_STA);
  WiFi.begin(cfg.ssid.c_str(), cfg.pass.c_str());
  uint32_t start = millis();
  while (WiFi.status() != WL_CONNECTED && millis() - start < timeoutMs) delay(100);
  return WiFi.status() == WL_CONNECTED;
}

bool syncNTP(uint32_t timeoutMs) {
  configTime(0, 0, "time.google.com", "pool.ntp.org", "time.cloudflare.com");
  uint32_t start = millis();
  bool resent = false;
  while (time(nullptr) < 1700000000UL && millis() - start < timeoutMs) {
    if (!resent && millis() - start > timeoutMs / 2) {
      configTime(0, 0, "time.google.com", "pool.ntp.org", "time.cloudflare.com");
      resent = true;
    }
    delay(200);
  }
  return time(nullptr) >= 1700000000UL;
}

uint32_t currentEpoch() {
  time_t t = time(nullptr);
  return (t >= 1700000000UL) ? (uint32_t)t : 0;
}

bool uploadBatch() {
  String body;
  body.reserve(64 + bufCount * 70);
  body  = "{\"device\":\""; body += cfg.name; body += "\",\"rows\":[";
  for (uint16_t i = 0; i < bufCount; i++) {
    if (i) body += ",";
    char row[96];
    snprintf(row, sizeof(row),
      "{\"t\":%lu,\"temp\":%.2f,\"hum\":%.1f,\"pct\":%u,\"v\":%.3f}",
      (unsigned long)buffer[i].epoch, buffer[i].temp,
      buffer[i].hum, buffer[i].pct, buffer[i].vbat);
    body += row;
  }
  body += "]}";

  WiFiClientSecure client;
  client.setInsecure();
  HTTPClient https;
  https.setConnectTimeout(15000);
  https.setTimeout(15000);
  if (!https.begin(client, cfg.url.c_str())) return false;
  https.setFollowRedirects(HTTPC_DISABLE_FOLLOW_REDIRECTS);
  https.addHeader("Content-Type", "application/json");
  int code = https.POST(body);
  String resp = https.getString();
  Serial.printf("Upload HTTP %d\n", code);
  https.end();
  bool ok = (code >= 200 && code < 400);
  if (!ok) Serial.println("Upload response: " + resp);
  return ok;
}

// ---------------- config storage + BLE ----------------

void ledConfig() {
#ifdef RGB_BUILTIN
  rgbLedWrite(RGB_BUILTIN, 0, 0, 40);   // dim blue = config mode
#endif
}

void loadConfig() {
  prefs.begin("dewvee", true);
  cfg.name           = prefs.getString("name", DEFAULT_NAME);
  cfg.sampleMin      = prefs.getUShort("sampleMin", DEFAULT_SAMPLE_MIN);
  cfg.uploadHrs      = prefs.getUShort("uploadHrs", DEFAULT_UPLOAD_HRS);
  cfg.uploadsEnabled = prefs.getBool("up", true);
  cfg.ssid           = prefs.getString("ssid", DEFAULT_SSID);
  cfg.pass           = prefs.getString("pass", DEFAULT_PASS);
  cfg.url            = prefs.getString("url", DEFAULT_URL);
  prefs.end();
  if (cfg.sampleMin < 1) cfg.sampleMin = 1;
  if (cfg.uploadHrs < 1) cfg.uploadHrs = 1;
}

void saveConfig() {
  prefs.begin("dewvee", false);
  prefs.putString("name", cfg.name);
  prefs.putUShort("sampleMin", cfg.sampleMin);
  prefs.putUShort("uploadHrs", cfg.uploadHrs);
  prefs.putBool("up", cfg.uploadsEnabled);
  prefs.putString("ssid", cfg.ssid);
  prefs.putString("pass", cfg.pass);
  prefs.putString("url", cfg.url);
  prefs.end();
}

String buildConfigJson() {
  JsonDocument doc;
  doc["name"]      = cfg.name;
  doc["sampleMin"] = cfg.sampleMin;
  doc["uploadHrs"] = cfg.uploadHrs;
  doc["up"]        = cfg.uploadsEnabled ? 1 : 0;
  doc["ssid"]      = cfg.ssid;
  doc["url"]       = cfg.url;
  doc["hasPass"]   = cfg.pass.length() ? 1 : 0;
  String out; serializeJson(doc, out); return out;
}

void applyConfigJson(const String& json) {
  JsonDocument doc;
  DeserializationError e = deserializeJson(doc, json);
  if (e) { if (statusChar) statusChar->setValue("error: bad json"); return; }
  if (!doc["name"].isNull())      cfg.name           = doc["name"].as<String>();
  if (!doc["sampleMin"].isNull()) cfg.sampleMin      = doc["sampleMin"].as<uint16_t>();
  if (!doc["uploadHrs"].isNull()) cfg.uploadHrs      = doc["uploadHrs"].as<uint16_t>();
  if (!doc["up"].isNull())        cfg.uploadsEnabled = doc["up"].as<int>() != 0;
  if (!doc["ssid"].isNull())      cfg.ssid           = doc["ssid"].as<String>();
  if (!doc["url"].isNull())       cfg.url            = doc["url"].as<String>();
  if (!doc["pass"].isNull()) {
    String p = doc["pass"].as<String>();
    if (p.length()) cfg.pass = p;
  }
  if (cfg.sampleMin < 1) cfg.sampleMin = 1;
  if (cfg.uploadHrs < 1) cfg.uploadHrs = 1;
  saveConfig();
  if (configChar) configChar->setValue(buildConfigJson().c_str());
  if (statusChar) statusChar->setValue("saved");
  configSaved = true;
  Serial.println("Config saved over BLE.");
}

class ConfigCallbacks : public BLECharacteristicCallbacks {
  void onWrite(BLECharacteristic* c) override {
    String incoming = c->getValue().c_str();
    applyConfigJson(incoming);
  }
};

void runConfigMode() {
  setCpuFrequencyMhz(160);
  ledConfig();
  Serial.printf("== BLE CONFIG MODE == advertising as \"Dewvee-%s\"\n", cfg.name.c_str());

  BLEDevice::init(("Dewvee-" + cfg.name).c_str());
  BLEServer*  server = BLEDevice::createServer();
  BLEService* svc    = server->createService(SERVICE_UUID);

  configChar = svc->createCharacteristic(CONFIG_UUID,
      BLECharacteristic::PROPERTY_READ | BLECharacteristic::PROPERTY_WRITE);
  configChar->setCallbacks(new ConfigCallbacks());
  configChar->setValue(buildConfigJson().c_str());

  statusChar = svc->createCharacteristic(STATUS_UUID, BLECharacteristic::PROPERTY_READ);
  statusChar->setValue("ready");

  svc->start();
  BLEAdvertising* adv = BLEDevice::getAdvertising();
  adv->addServiceUUID(SERVICE_UUID);
  adv->setScanResponse(true);
  BLEDevice::startAdvertising();

  uint32_t start  = millis();
  uint32_t window = CONFIG_MODE_MS;
  while (millis() - start < window) {
    delay(100);
    if (configSaved) {
      configSaved = false;
      start  = millis();
      window = 30000;   // linger 30 s after a save for follow-up edits
    }
  }
  BLEDevice::deinit(true);
  ledOff();
  Serial.println("Config mode ended.");
}

// ---------------- main ----------------

void setup() {
  setCpuFrequencyMhz(80);
  ledOff();
  Serial.begin(115200);
  delay(50);

  loadConfig();

  // Magnet wake -> BLE config mode, then back to sleep
#if ENABLE_REED_WAKE
  pinMode(REED_PIN, INPUT_PULLUP);
  delay(5);
  bool reedActive = (digitalRead(REED_PIN) == LOW);
  if (esp_sleep_get_wakeup_cause() == ESP_SLEEP_WAKEUP_GPIO || reedActive) {
    // Capture current reading and flush buffered data before config mode.
    // Prevents data loss if the device is reflashed or power-cycled during
    // the config session (both actions wipe RTC_DATA_ATTR memory).
    if (registered && cfg.uploadsEnabled) {
      Wire.begin(SDA_PIN, SCL_PIN);
      float temp = NAN, hum = NAN;
      if (sht.begin(&Wire)) {
        sht.setPrecision(SHT4X_HIGH_PRECISION);
        sht.setHeater(SHT4X_NO_HEATER);
        readSensor(temp, hum);
      }
      float vbat = readBatteryVolts();
      uint8_t pct = batteryPercent(vbat);
      lastVbat = vbat; lastPct = pct;

      if (connectWiFi(15000)) {
        if (syncNTP(10000)) {
          timeSynced    = true;
          lastSyncEpoch = currentEpoch();
        }
        uint32_t epoch = currentEpoch();
        if (!isnan(temp) && !isnan(hum) && epoch != 0)
          pushReading(epoch, temp, hum, vbat, pct);
        if (bufCount > 0) {
          if (uploadBatch()) {
            bufCount = 0;
            Serial.println("Reed wake: buffer flushed before config mode.");
          } else {
            Serial.println("Reed wake: flush failed, buffer kept for next timer wake.");
          }
        }
        WiFi.mode(WIFI_OFF);
        delay(100);
      } else {
        // WiFi unavailable - at least save the reading with the RTC timestamp
        // so it is not lost; it will upload on the next normal timer wake.
        uint32_t epoch = currentEpoch();
        if (!isnan(temp) && !isnan(hum) && epoch != 0)
          pushReading(epoch, temp, hum, vbat, pct);
        Serial.printf("Reed wake: WiFi unavailable, %u readings buffered.\n", bufCount);
      }
    }
    runConfigMode();
    // wait for magnet removal so we don't instantly re-wake
    uint32_t t0 = millis();
    while (digitalRead(REED_PIN) == LOW && millis() - t0 < 10000) delay(100);
    goToSleep();
    return;
  }
#endif

  // --- one-time registration ---
  if (!registered) {
    if (cfg.uploadsEnabled) {
      Wire.begin(SDA_PIN, SCL_PIN);
      float temp = NAN, hum = NAN;
      if (sht.begin(&Wire)) {
        sht.setPrecision(SHT4X_HIGH_PRECISION);
        sht.setHeater(SHT4X_NO_HEATER);
        readSensor(temp, hum);
      } else {
        Serial.println("SHT45 not found - check wiring.");
      }
      float vbat = readBatteryVolts();
      uint8_t pct = batteryPercent(vbat);
      lastVbat = vbat; lastPct = pct;

      if (connectWiFi(20000)) {
        delay(300);
        if (syncNTP(20000)) {
          timeSynced    = true;
          lastSyncEpoch = currentEpoch();
          uint32_t epoch = currentEpoch();
          if (!isnan(temp) && !isnan(hum) && epoch != 0) {
            pushReading(epoch, temp, hum, vbat, pct);
            if (uploadBatch()) {
              bufCount = 0; registered = true;
              Serial.println("Registered - first reading is on the sheet.");
            } else {
              bufCount = 0;
              Serial.println("First upload failed - retrying next wake.");
            }
          }
        } else { Serial.println("NTP sync failed - retrying next wake."); }
      } else { Serial.println("WiFi unavailable - retrying next wake."); }
    } else {
      registered = true;
    }
    bootCount++;
    goToSleep();
    return;
  }

  bootCount++;

  // --- normal wake: sample -> buffer -> upload if due ---
  uint16_t uploadEntries = (uint16_t)(((uint32_t)cfg.uploadHrs * 60UL) / cfg.sampleMin);
  if (uploadEntries < 1) uploadEntries = 1;
  bool uploadDue = cfg.uploadsEnabled && (bufCount + 1 >= uploadEntries);
  bool useWiFi   = uploadDue;

  Wire.begin(SDA_PIN, SCL_PIN);
  float temp = NAN, hum = NAN;
  if (sht.begin(&Wire)) {
    sht.setPrecision(SHT4X_HIGH_PRECISION);
    sht.setHeater(SHT4X_NO_HEATER);
    readSensor(temp, hum);
  } else {
    Serial.println("SHT45 not found - check wiring.");
  }

  float vbat; uint8_t pct;
  if (useWiFi) {
    vbat = readBatteryVolts(); pct = batteryPercent(vbat);
    lastVbat = vbat; lastPct = pct;
  } else {
    vbat = lastVbat; pct = lastPct;
  }

  bool wifiUp = false;
  if (useWiFi) {
    wifiUp = connectWiFi(15000);
    if (wifiUp) {
      uint32_t rtcNow = currentEpoch();
      delay(300);
      if (syncNTP(20000)) {
        timeSynced = true;
        uint32_t trueNow = currentEpoch();
        if (lastSyncEpoch != 0 && rtcNow > lastSyncEpoch && trueNow > lastSyncEpoch) {
          double scale = (double)(trueNow - lastSyncEpoch) /
                         (double)(rtcNow  - lastSyncEpoch);
          if (scale > 0.5 && scale < 2.0) {
            for (uint16_t i = 0; i < bufCount; i++) {
              if (buffer[i].epoch <= lastSyncEpoch) continue;
              double corrected = (double)lastSyncEpoch +
                                 (double)(buffer[i].epoch - lastSyncEpoch) * scale;
              buffer[i].epoch = (uint32_t)(corrected + 0.5);
            }
            Serial.printf("Drift corrected: RTC off by %+ld s (scale %.5f)\n",
                          (long)rtcNow - (long)trueNow, scale);
          }
        }
        lastSyncEpoch = trueNow;
      }
    }
  }

  uint32_t epoch = currentEpoch();
  if (!isnan(temp) && !isnan(hum) && epoch != 0)
    pushReading(epoch, temp, hum, vbat, pct);
  Serial.printf("T=%.2fC RH=%.1f%% Vbat=%.3f (%u%%) epoch=%lu buf=%u\n",
                temp, hum, vbat, pct, (unsigned long)epoch, bufCount);

  if (uploadDue && wifiUp && bufCount > 0) {
    if (uploadBatch()) {
      bufCount = 0; registered = true;
    } else {
      Serial.println("Upload failed - buffer kept, retry next wake.");
    }
  } else if (uploadDue && !wifiUp) {
    Serial.println("WiFi unavailable - buffer kept, retry next wake.");
  }

  goToSleep();
}

void loop() {
  // unused - everything happens in setup() before deep sleep
}
