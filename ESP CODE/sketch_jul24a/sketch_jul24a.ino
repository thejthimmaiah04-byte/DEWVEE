/*
 * ================================================================
 *  ESP32-C6 Weather Pod  -  Temperature / Humidity / Battery
 * ================================================================
 *  Board : Waveshare ESP32-C6-Zero  (select "ESP32C6 Dev Module")
 *  Sensor: Adafruit SHT45 over I2C (fixed address 0x44)
 *  Power : Samsung INR18650-35E via protected TP4056 -> board VIN
 *
 *  Behaviour:
 *   - Wakes every N minutes from deep sleep (default 10).
 *   - Before first registration: attempts one upload per wake; retries
 *     every N minutes until the first row lands on the sheet.
 *   - After registration: reads SHT45 in a 3-sample burst, reads
 *     battery on EVERY wake, timestamps via RTC-advanced clock, and
 *     buffers in RTC memory.  WiFi touched only at upload intervals
 *     (default every 12 h / 72 readings).  Drift-corrected at sync.
 *   - BSSID + channel are cached in RTC RAM after each successful
 *     connect, cutting reconnect time from ~8 s to ~2 s.  Cache is
 *     invalidated on credential change or failed connect.
 *   - Battery protection: below 3.20 V the pod blinks red twice and
 *     sleeps for 60 min to prevent deep-discharge cell damage.
 *     Recovers automatically once voltage rises above 3.50 V.
 *   - Magnet on reed switch wakes into BLE config mode.  Before
 *     entering config mode, any buffered data is flushed so it is
 *     not lost if the device is re-flashed during the session.
 *   - Shared upload key: if cfg.uploadKey is set the pod sends it
 *     in every POST so the Apps Script can authenticate requests.
 *   - Location label: human-readable placement tag ("Server Room")
 *     stored in NVS, included in every upload for the dashboard.
 *
 *  Power budget (3000 mAh, 10 min / 12 h defaults):
 *   Deep sleep  ~3.5 mAh/day  (150 µA board quiescent)
 *   Sampling    ~0.5 mAh/day  (40 MHz, 3-sample burst, 144 wakes)
 *   WiFi        ~2.5 mAh/day  (2 sessions × ~15 s with BSSID cache)
 *   Total       ~6.5 mAh/day  →  ~370 days on 2400 mAh usable
 *
 *  Libraries (Arduino Library Manager):
 *   - Adafruit SHT4x Library + Adafruit Unified Sensor
 *   - ArduinoJson  (v7)
 *   WiFi, HTTPClient, BLE, Preferences ship with ESP32 Arduino core
 *   (use core v3.x for C6; BLE is backed by NimBLE).
 *
 *  Partition scheme: "Huge APP (3MB No OTA / 1MB SPIFFS)"
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

// -------- DEFAULTS (first boot only — configure via BLE) --------
#define DEFAULT_NAME        "DEWVEE:03"
#define DEFAULT_SSID        ""           // set via BLE on first use
#define DEFAULT_PASS        ""           // set via BLE on first use
#define DEFAULT_URL         "https://script.google.com/macros/s/AKfycbw88DpkYijK98qTzhlSguLAFxFRQI_H3KXreYcRp2sitSr3XWFPRz9QwR88ocdsRsO7/exec"
#define DEFAULT_SAMPLE_MIN  10   // 10 min → 144 readings/day; tuned for 6-month battery life
#define DEFAULT_UPLOAD_HRS  12   // 12 h  → 2 WiFi sessions/day

struct Config {
  String   name;
  String   location;       // human-readable placement label
  uint16_t sampleMin;
  uint16_t uploadHrs;
  bool     uploadsEnabled;
  String   ssid;
  String   pass;
  String   url;
  String   uploadKey;      // shared secret for server-side auth
} cfg;

Preferences prefs;

// BLE service UUIDs (must match the web app)
#define SERVICE_UUID "d5770001-9e0b-4b7a-9c1e-2f3a4b5c6d7e"
#define CONFIG_UUID  "d5770002-9e0b-4b7a-9c1e-2f3a4b5c6d7e"
#define STATUS_UUID  "d5770003-9e0b-4b7a-9c1e-2f3a4b5c6d7e"
const uint32_t CONFIG_MODE_MS = 180000;   // 3 min BLE window
BLECharacteristic* configChar = nullptr;
BLECharacteristic* statusChar = nullptr;
volatile bool configSaved     = false;
bool          wifiCredChanged = false;

// Pins
const int SDA_PIN  = 6;
const int SCL_PIN  = 5;
const int VBAT_PIN = 1;
const int REED_PIN = 2;   // Terminal A -> GND, B -> GPIO2, internal pull-up

#define ENABLE_REED_WAKE 1

const float VBAT_DIVIDER  = 2.0f;
const int   BURST_SAMPLES = 3;    // SHT45 high-precision takes 8.2 ms; 3 × 15 ms is enough
#define uS_PER_MIN 60000000ULL

// -------- RTC-persisted state (survives deep sleep) --------
typedef struct {
  uint32_t epoch;
  float    temp;
  float    hum;
  float    vbat;
  uint8_t  pct;
} Reading;

#define BUFFER_CAPACITY 216
RTC_DATA_ATTR Reading  buffer[BUFFER_CAPACITY];
RTC_DATA_ATTR uint16_t bufCount        = 0;
RTC_DATA_ATTR uint32_t bootCount       = 0;
RTC_DATA_ATTR bool     timeSynced      = false;
RTC_DATA_ATTR uint32_t lastSyncEpoch   = 0;
RTC_DATA_ATTR bool     registered      = false;
RTC_DATA_ATTR float    lastVbat        = 0;
RTC_DATA_ATTR uint8_t  lastPct         = 0;
RTC_DATA_ATTR bool     pendingWifiTest = false;  // set before restart after BLE cred change
RTC_DATA_ATTR bool     batteryDead     = false;  // set when Vbat < 3.20V; clears at 3.50V
// BSSID + channel cache: skip AP scan on reconnect, cuts connect time by 3–8 s
RTC_DATA_ATTR uint8_t  cachedBssid[6]  = {0};
RTC_DATA_ATTR uint8_t  cachedChannel   = 0;
RTC_DATA_ATTR bool     bssidCached     = false;

Adafruit_SHT4x sht = Adafruit_SHT4x();

// -------- LED helpers --------
void ledOff() {
#ifdef RGB_BUILTIN
  rgbLedWrite(RGB_BUILTIN, 0, 0, 0);
#endif
}
void ledBlink(uint8_t r, uint8_t g, uint8_t b, int times, int onMs, int offMs) {
#ifdef RGB_BUILTIN
  for (int i = 0; i < times; i++) {
    rgbLedWrite(RGB_BUILTIN, r, g, b);
    delay(onMs);
    rgbLedWrite(RGB_BUILTIN, 0, 0, 0);
    if (i < times - 1) delay(offMs);
  }
#endif
}
void ledConfig() {
#ifdef RGB_BUILTIN
  rgbLedWrite(RGB_BUILTIN, 0, 0, 40);   // dim blue = config mode
#endif
}

// -------- Sleep --------
void goToSleep() {
  WiFi.mode(WIFI_OFF);
  Serial.flush();
  esp_sleep_enable_timer_wakeup((uint64_t)cfg.sampleMin * uS_PER_MIN);
#if ENABLE_REED_WAKE
  esp_deep_sleep_enable_gpio_wakeup(1ULL << REED_PIN, ESP_GPIO_WAKEUP_GPIO_LOW);
#endif
  esp_deep_sleep_start();
}

// -------- Sensor --------
bool readSensor(float &tOut, float &hOut) {
  delay(50);
  sensors_event_t humidity, temp;
  float tSum = 0, hSum = 0;
  int   n = 0;
  for (int i = 0; i < BURST_SAMPLES; i++) {
    sht.getEvent(&humidity, &temp);
    float t = temp.temperature, h = humidity.relative_humidity;
    if (!isnan(t) && !isnan(h)) { tSum += t; hSum += h; n++; }
    if (i < BURST_SAMPLES - 1) delay(15);
  }
  if (n == 0) return false;
  tOut = tSum / n;
  hOut = hSum / n;
  return true;
}

// -------- Battery --------
// Calibration: pre-cal reading = 4.700/1.149 = 4.090V → factor = 4.22/4.090 ≈ 1.032
const float VBAT_CALIBRATION = 4.22f / 4.090f;

float readBatteryVolts() {
  analogSetPinAttenuation(VBAT_PIN, ADC_11db);
  const int N = 16;
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
  if (v >= 4.10f) return  90;
  if (v >= 4.00f) return  80;
  if (v >= 3.90f) return  70;
  if (v >= 3.80f) return  60;
  if (v >= 3.75f) return  50;
  if (v >= 3.70f) return  40;
  if (v >= 3.65f) return  30;
  if (v >= 3.60f) return  20;
  if (v >= 3.50f) return  12;
  if (v >= 3.40f) return   6;
  if (v >= 3.30f) return   3;
  return 0;
}

// -------- Buffer --------
void pushReading(uint32_t epoch, float t, float h, float v, uint8_t pct) {
  if (bufCount >= BUFFER_CAPACITY) {
    memmove(&buffer[0], &buffer[1], (BUFFER_CAPACITY - 1) * sizeof(Reading));
    bufCount = BUFFER_CAPACITY - 1;
  }
  buffer[bufCount++] = { epoch, t, h, v, pct };
}

// -------- WiFi / NTP --------
bool connectWiFi(uint32_t timeoutMs) {
  setCpuFrequencyMhz(80);
  WiFi.mode(WIFI_STA);
  WiFi.setSleep(false);
  if (bssidCached) {
    WiFi.begin(cfg.ssid.c_str(), cfg.pass.c_str(), cachedChannel, cachedBssid);
  } else {
    WiFi.begin(cfg.ssid.c_str(), cfg.pass.c_str());
  }
  uint32_t start = millis();
  while (WiFi.status() != WL_CONNECTED && millis() - start < timeoutMs) delay(100);
  bool ok = (WiFi.status() == WL_CONNECTED);
  if (ok) {
    memcpy(cachedBssid, WiFi.BSSID(), 6);
    cachedChannel = (uint8_t)WiFi.channel();
    bssidCached   = true;
  } else if (bssidCached) {
    bssidCached = false;
    memset(cachedBssid, 0, 6);
    cachedChannel = 0;
  }
  return ok;
}

bool syncNTP(uint32_t timeoutMs) {
  configTime(0, 0, "time.google.com", "pool.ntp.org", "time.cloudflare.com");
  uint32_t start = millis(); bool resent = false;
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

// -------- Upload --------
static String jsonStr(const String& s) {
  String out; out.reserve(s.length() + 2);
  out = "\"";
  for (size_t i = 0; i < s.length(); i++) {
    char c = s[i];
    if      (c == '"')  out += "\\\"";
    else if (c == '\\') out += "\\\\";
    else if (c == '\n') out += "\\n";
    else if (c == '\r') out += "\\r";
    else                out += c;
  }
  out += "\"";
  return out;
}

bool uploadBatch() {
  String body;
  body.reserve(128 + bufCount * 70);
  body  = "{\"device\":"   + jsonStr(cfg.name);
  body += ",\"location\":" + jsonStr(cfg.location);
  body += ",\"sampleMin\":"; body += cfg.sampleMin;
  if (cfg.uploadKey.length()) {
    body += ",\"key\":"; body += jsonStr(cfg.uploadKey);
  }
  body += ",\"rows\":[";
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
  int    code = https.POST(body);
  String resp = https.getString();
  Serial.printf("Upload HTTP %d\n", code);
  https.end();
  bool ok = (code >= 200 && code < 400);
  if (!ok) Serial.println("Upload response: " + resp);
  return ok;
}

// -------- Config storage --------
void loadConfig() {
  prefs.begin("dewvee", true);
  cfg.name           = prefs.getString("name",      DEFAULT_NAME);
  cfg.location       = prefs.getString("loc",       "");
  cfg.sampleMin      = prefs.getUShort("sampleMin", DEFAULT_SAMPLE_MIN);
  cfg.uploadHrs      = prefs.getUShort("uploadHrs", DEFAULT_UPLOAD_HRS);
  cfg.uploadsEnabled = prefs.getBool  ("up",         true);
  cfg.ssid           = prefs.getString("ssid",      DEFAULT_SSID);
  cfg.pass           = prefs.getString("pass",      DEFAULT_PASS);
  cfg.url            = prefs.getString("url",       DEFAULT_URL);
  cfg.uploadKey      = prefs.getString("ukey",      "");
  prefs.end();
  if (cfg.sampleMin < 1) cfg.sampleMin = 1;
  if (cfg.uploadHrs < 1) cfg.uploadHrs = 1;
}

void saveConfig() {
  prefs.begin("dewvee", false);
  prefs.putString("name",      cfg.name);
  prefs.putString("loc",       cfg.location);
  prefs.putUShort("sampleMin", cfg.sampleMin);
  prefs.putUShort("uploadHrs", cfg.uploadHrs);
  prefs.putBool  ("up",        cfg.uploadsEnabled);
  prefs.putString("ssid",      cfg.ssid);
  prefs.putString("pass",      cfg.pass);
  prefs.putString("url",       cfg.url);
  prefs.putString("ukey",      cfg.uploadKey);
  prefs.end();
}

String buildConfigJson() {
  JsonDocument doc;
  doc["name"]      = cfg.name;
  doc["location"]  = cfg.location;
  doc["sampleMin"] = cfg.sampleMin;
  doc["uploadHrs"] = cfg.uploadHrs;
  doc["up"]        = cfg.uploadsEnabled ? 1 : 0;
  doc["ssid"]      = cfg.ssid;
  doc["url"]       = cfg.url;
  doc["hasPass"]   = cfg.pass.length()      ? 1 : 0;
  doc["hasKey"]    = cfg.uploadKey.length() ? 1 : 0;
  String out; serializeJson(doc, out); return out;
}

void applyConfigJson(const String& json) {
  JsonDocument doc;
  DeserializationError e = deserializeJson(doc, json);
  if (e) { if (statusChar) statusChar->setValue("error:badjson"); return; }

  String oldSsid = cfg.ssid, oldPass = cfg.pass;

  if (!doc["name"].isNull())      cfg.name           = doc["name"].as<String>();
  if (!doc["location"].isNull())  cfg.location       = doc["location"].as<String>();
  if (!doc["sampleMin"].isNull()) cfg.sampleMin      = doc["sampleMin"].as<uint16_t>();
  if (!doc["uploadHrs"].isNull()) cfg.uploadHrs      = doc["uploadHrs"].as<uint16_t>();
  if (!doc["up"].isNull())        cfg.uploadsEnabled = doc["up"].as<int>() != 0;
  if (!doc["ssid"].isNull())      cfg.ssid           = doc["ssid"].as<String>();
  if (!doc["url"].isNull())       cfg.url            = doc["url"].as<String>();
  if (!doc["pass"].isNull()) {
    String p = doc["pass"].as<String>(); if (p.length()) cfg.pass = p;
  }
  if (!doc["key"].isNull()) {
    String k = doc["key"].as<String>(); if (k.length()) cfg.uploadKey = k;
  }

  if (cfg.ssid != oldSsid || cfg.pass != oldPass) {
    wifiCredChanged = true;
    bssidCached = false;
    memset(cachedBssid, 0, 6);
    cachedChannel = 0;
  }
  if (cfg.sampleMin < 1) cfg.sampleMin = 1;
  if (cfg.uploadHrs < 1) cfg.uploadHrs = 1;
  saveConfig();
  if (configChar) configChar->setValue(buildConfigJson().c_str());
  if (statusChar) statusChar->setValue("saved");
  configSaved = true;
  Serial.println("Config saved over BLE.");
}

// -------- BLE callbacks --------
class ServerCallbacks : public BLEServerCallbacks {
  void onConnect(BLEServer*) override {
    ledBlink(40, 40, 40, 1, 250, 0);   // 1× white = phone connected
    ledConfig();
    Serial.println("BLE client connected.");
  }
  void onDisconnect(BLEServer*) override {
    BLEDevice::startAdvertising();
    Serial.println("BLE client disconnected - resuming advertising.");
  }
};
class ConfigCallbacks : public BLECharacteristicCallbacks {
  void onWrite(BLECharacteristic* c) override {
    applyConfigJson(c->getValue().c_str());
  }
};

// -------- BLE config mode --------
void runConfigMode() {
  setCpuFrequencyMhz(160);
  ledConfig();
  Serial.printf("== BLE CONFIG MODE == advertising as \"Dewvee-%s\"\n", cfg.name.c_str());

  BLEDevice::init(("Dewvee-" + cfg.name).c_str());
  BLEServer*  server = BLEDevice::createServer();
  server->setCallbacks(new ServerCallbacks());
  BLEService* svc = server->createService(SERVICE_UUID);

  configChar = svc->createCharacteristic(CONFIG_UUID,
      BLECharacteristic::PROPERTY_READ | BLECharacteristic::PROPERTY_WRITE);
  configChar->setCallbacks(new ConfigCallbacks());
  configChar->setValue(buildConfigJson().c_str());

  statusChar = svc->createCharacteristic(STATUS_UUID, BLECharacteristic::PROPERTY_READ);
  statusChar->setValue("ready:180");

  svc->start();
  BLEAdvertising* adv = BLEDevice::getAdvertising();
  adv->addServiceUUID(SERVICE_UUID);
  adv->setScanResponse(true);
  BLEDevice::startAdvertising();

  uint32_t start      = millis();
  uint32_t window     = CONFIG_MODE_MS;
  uint32_t lastStatus = 0;

  while (millis() - start < window) {
    delay(100);

    if (millis() - lastStatus >= 10000) {
      lastStatus = millis();
      uint32_t elapsed = millis() - start;
      uint32_t remaining = (elapsed < window) ? (window - elapsed) / 1000 : 0;
      char buf[32];
      snprintf(buf, sizeof(buf), "ready:%lu", remaining);
      if (statusChar && !configSaved) statusChar->setValue(buf);
    }

    if (configSaved) {
      configSaved = false;
      // If WiFi creds changed we must restart to switch radio — short linger only.
      // Otherwise linger 30 s so the user can make follow-up edits.
      window  = wifiCredChanged ? 5000 : 30000;
      start   = millis();
      lastStatus = 0;
#ifdef RGB_BUILTIN
      rgbLedWrite(RGB_BUILTIN, 0, 40, 0);   // solid green = saved
      delay(1500);
      ledConfig();
#endif
      Serial.println("Config saved - green confirmed.");
      if (statusChar) statusChar->setValue("saved");
    }
  }

  BLEDevice::deinit(true);
  ledOff();
  Serial.println("Config mode ended.");

  // If WiFi creds did NOT change but WiFi is configured, flush buffer now.
  // (When creds DID change we restart below, which handles the upload on next boot.)
  if (!wifiCredChanged && cfg.uploadsEnabled && cfg.ssid.length() > 0 && bufCount > 0) {
    Serial.println("Post-config upload attempt...");
    if (connectWiFi(15000)) {
      if (syncNTP(10000)) { timeSynced = true; lastSyncEpoch = currentEpoch(); }
      if (uploadBatch()) {
        bufCount = 0; registered = true;
        Serial.println("Post-config upload OK.");
      } else {
        Serial.println("Post-config upload failed - buffer kept.");
      }
      WiFi.mode(WIFI_OFF);
    } else {
      Serial.println("Post-config upload: WiFi unavailable, buffer kept.");
    }
  }

  if (wifiCredChanged) {
    wifiCredChanged = false;
    pendingWifiTest = true;
    Serial.println("WiFi credentials changed - restarting to test new network.");
    delay(100);
    esp_restart();
  }
}

// ================================================================
//  MAIN
// ================================================================
void setup() {
  setCpuFrequencyMhz(40);   // 40 MHz for all non-WiFi work saves ~8 mA vs 80 MHz
  ledOff();
  Serial.begin(115200);
  delay(50);
  loadConfig();

  // ---- Post-BLE WiFi credential test ----
  // Always take a fresh reading here so the sheet receives at least one data point
  // the moment new WiFi credentials are confirmed — this is the user's only signal
  // that the pod is online and the config was successful.
  if (pendingWifiTest) {
    pendingWifiTest = false;
    Serial.println("Post-BLE WiFi test starting...");

    Wire.begin(SDA_PIN, SCL_PIN);
    float temp = NAN, hum = NAN;
    if (sht.begin(&Wire)) {
      sht.setPrecision(SHT4X_HIGH_PRECISION);
      sht.setHeater(SHT4X_NO_HEATER);
      readSensor(temp, hum);
    }
    float   vbat = readBatteryVolts();
    uint8_t pct  = batteryPercent(vbat);
    lastVbat = vbat; lastPct = pct;

    if (connectWiFi(20000)) {
      Serial.println("New WiFi OK.");
      if (syncNTP(15000)) { timeSynced = true; lastSyncEpoch = currentEpoch(); }
      uint32_t epoch = currentEpoch();
      if (!isnan(temp) && !isnan(hum) && epoch != 0)
        pushReading(epoch, temp, hum, vbat, pct);
      if (bufCount > 0) {
        if (uploadBatch()) {
          bufCount = 0; registered = true;
          ledBlink(0, 50, 0, 2, 400, 200);   // 2× green = WiFi + upload confirmed
          Serial.println("Post-BLE upload OK - pod confirmed online.");
        } else {
          Serial.println("Upload failed - buffer kept.");
        }
      }
      WiFi.mode(WIFI_OFF);
    } else {
      ledBlink(50, 0, 0, 3, 200, 150);   // 3× red = wrong WiFi credentials
      WiFi.mode(WIFI_OFF);
      Serial.println("New WiFi credentials failed to connect.");
    }
    goToSleep();
    return;
  }

  // ---- Reed switch wake → BLE config mode ----
#if ENABLE_REED_WAKE
  pinMode(REED_PIN, INPUT_PULLUP);
  delay(50);
  bool reedActive = false;
  if (digitalRead(REED_PIN) == LOW) {
    delay(20); if (digitalRead(REED_PIN) == LOW) {
      delay(20); if (digitalRead(REED_PIN) == LOW) reedActive = true;
    }
  }
  if (esp_sleep_get_wakeup_cause() == ESP_SLEEP_WAKEUP_GPIO || reedActive) {
    if (registered && cfg.uploadsEnabled) {
      Wire.begin(SDA_PIN, SCL_PIN);
      float temp = NAN, hum = NAN;
      if (sht.begin(&Wire)) {
        sht.setPrecision(SHT4X_HIGH_PRECISION);
        sht.setHeater(SHT4X_NO_HEATER);
        readSensor(temp, hum);
      }
      float   vbat = readBatteryVolts();
      uint8_t pct  = batteryPercent(vbat);
      lastVbat = vbat; lastPct = pct;

      if (connectWiFi(15000)) {
        if (syncNTP(10000)) { timeSynced = true; lastSyncEpoch = currentEpoch(); }
        uint32_t epoch = currentEpoch();
        if (!isnan(temp) && !isnan(hum) && epoch != 0)
          pushReading(epoch, temp, hum, vbat, pct);
        if (bufCount > 0) {
          if (uploadBatch()) { bufCount = 0; Serial.println("Reed wake: buffer flushed."); }
          else               { Serial.println("Reed wake: flush failed - buffer kept."); }
        }
        WiFi.mode(WIFI_OFF); delay(100);
      } else {
        uint32_t epoch = currentEpoch();
        if (!isnan(temp) && !isnan(hum) && epoch != 0)
          pushReading(epoch, temp, hum, vbat, pct);
        Serial.printf("Reed wake: WiFi unavailable, %u readings buffered.\n", bufCount);
      }
    }
    runConfigMode();
    uint32_t t0 = millis();
    while (digitalRead(REED_PIN) == LOW && millis() - t0 < 10000) delay(100);
    goToSleep();
    return;
  }
#endif

  // ---- Battery protection (all normal / timer wakes) ----
  float   vbat = readBatteryVolts();
  uint8_t pct  = batteryPercent(vbat);
  lastVbat = vbat; lastPct = pct;

  if (vbat < 3.20f) batteryDead = true;
  if (batteryDead) {
    if (vbat >= 3.50f) {
      batteryDead = false;
      Serial.println("Battery recovered - resuming normal operation.");
    } else {
      Serial.printf("Battery critical (%.3fV) - protection sleep 60 min.\n", vbat);
      ledBlink(50, 0, 0, 2, 400, 300);   // 2× red = critical battery
      ledOff(); Serial.flush();
      esp_sleep_enable_timer_wakeup(60ULL * uS_PER_MIN);
#if ENABLE_REED_WAKE
      esp_deep_sleep_enable_gpio_wakeup(1ULL << REED_PIN, ESP_GPIO_WAKEUP_GPIO_LOW);
#endif
      esp_deep_sleep_start();
      return;
    }
  }

  // ---- One-time registration ----
  if (!registered) {
    if (cfg.uploadsEnabled) {
      Wire.begin(SDA_PIN, SCL_PIN);
      float temp = NAN, hum = NAN;
      if (sht.begin(&Wire)) {
        sht.setPrecision(SHT4X_HIGH_PRECISION);
        sht.setHeater(SHT4X_NO_HEATER);
        readSensor(temp, hum);
      } else { Serial.println("SHT45 not found - check wiring."); }

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

  // ---- Normal wake: sample → buffer → upload if due ----
  uint16_t uploadEntries = (uint16_t)(
    ((uint32_t)cfg.uploadHrs * 60UL + cfg.sampleMin - 1) / cfg.sampleMin);
  if (uploadEntries < 1) uploadEntries = 1;
  bool uploadDue = cfg.uploadsEnabled && (bufCount + 1 >= uploadEntries);

  Wire.begin(SDA_PIN, SCL_PIN);
  float temp = NAN, hum = NAN;
  if (sht.begin(&Wire)) {
    sht.setPrecision(SHT4X_HIGH_PRECISION);
    sht.setHeater(SHT4X_NO_HEATER);
    readSensor(temp, hum);
  } else { Serial.println("SHT45 not found - check wiring."); }

  bool wifiUp = false;
  if (uploadDue) {
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
  Serial.printf("T=%.2fC RH=%.1f%% Vbat=%.3fV (%u%%) epoch=%lu buf=%u\n",
                temp, hum, vbat, pct, (unsigned long)epoch, bufCount);

  if (uploadDue && wifiUp && bufCount > 0) {
    if (uploadBatch()) { bufCount = 0; registered = true; }
    else               { Serial.println("Upload failed - buffer kept, retry next wake."); }
  } else if (uploadDue && !wifiUp) {
    Serial.println("WiFi unavailable - buffer kept, retry next wake.");
  }

  goToSleep();
}

void loop() {
  // Everything happens in setup() before deep sleep — loop() never runs.
}
