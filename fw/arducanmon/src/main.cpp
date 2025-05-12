#include <Arduino.h>
#include <mcp_can.h>
#include <SPI.h>

// CAN TX Variables
unsigned long prevTX = 0;
byte data[8] = {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
bool autoSend = false;

// CAN RX Variables
long unsigned int rxId;
unsigned char len;
unsigned char rxBuf[8];

// Serial Buffer
char msgString[128];
char incomingByte;
String inputString = "";
bool stringComplete = false;

// CAN0 INT and CS
#define CAN0_INT 2 // Set INT to pin 2
MCP_CAN CAN0(10);  // Set CS to pin 10

void setup() {
  Serial.begin(115200);
  
  // Initialize MCP2515 running at 8MHz with a baudrate of 125kb/s as used in TP2
  if (CAN0.begin(MCP_ANY, CAN_125KBPS, MCP_8MHZ) == CAN_OK) {
    Serial.println("CAN_INIT_OK\n");
    Serial.print("CAN BaudRate: 125kbps\n");
    Serial.print("MCP2515 Clock: 8MHz\n");
  } else {
    Serial.println("CAN_INIT_FAIL");
  }
  
  // Set to normal mode (not loopback)
  CAN0.setMode(MCP_NORMAL);
  
  pinMode(CAN0_INT, INPUT); // Configuring pin for /INT input
  
  Serial.println("TP2_CAN_MONITOR_READY");
}

// Extrae el ID y datos para enviar un mensaje CAN
void sendCANMessage(String cmd) {
  // Formato: SEND_ID_HEX_BYTE1_BYTE2_...
  // Ejemplo: "SEND_100_52_2D_33_34" envía ID=0x100, data=R-34
  
  int firstSplit = cmd.indexOf('_', 5); // Busca después de "SEND_"
  if (firstSplit == -1) return;
  
  String idStr = cmd.substring(5, firstSplit);
  long id = strtol(idStr.c_str(), NULL, 16); // Convertir ID de hex a long
  
  // Extraer los bytes de datos
  int startIdx = firstSplit + 1;
  int byteCount = 0;
  
  memset(data, 0, sizeof(data)); // Limpiar el array de datos
  
  while (startIdx < cmd.length() && byteCount < 8) {
    int nextSplit = cmd.indexOf('_', startIdx);
    if (nextSplit == -1) nextSplit = cmd.length();
    
    String byteStr = cmd.substring(startIdx, nextSplit);
    data[byteCount] = strtol(byteStr.c_str(), NULL, 16);
    
    byteCount++;
    startIdx = nextSplit + 1;
  }
  
  // Enviar mensaje CAN
  byte sndStat = CAN0.sendMsgBuf(id, 0, byteCount, data);
  
  if (sndStat == CAN_OK) {
    Serial.print("CAN_TX_OK_");
    Serial.print(id, HEX);
    Serial.print("_");
    for (int i = 0; i < byteCount; i++) {
      Serial.print(data[i], HEX);
      if (i < byteCount - 1) Serial.print("_");
    }
    Serial.println();
  } else {
    Serial.println("CAN_TX_FAIL");
  }
}

// Procesa comandos recibidos por serial
void processCommand(String cmd) {
  cmd.trim();
  
  if (cmd.startsWith("SEND_")) {
    sendCANMessage(cmd);
  } 
  else if (cmd == "MODE_NORMAL") {
    CAN0.setMode(MCP_NORMAL);
    Serial.println("MODE_SET_NORMAL");
  }
  else if (cmd == "MODE_LOOPBACK") {
    CAN0.setMode(MCP_LOOPBACK);
    Serial.println("MODE_SET_LOOPBACK");
  }
  else if (cmd == "AUTO_ON") {
    autoSend = true;
    Serial.println("AUTO_SEND_ON");
  }
  else if (cmd == "AUTO_OFF") {
    autoSend = false;
    Serial.println("AUTO_SEND_OFF");
  }
  else if (cmd.startsWith("TP2_ANGLE_")) {
    // Formato: TP2_ANGLE_TYPE_VALUE
    // Ejemplo: TP2_ANGLE_R_-45
    int firstSplit = cmd.indexOf('_', 9);
    if (firstSplit != -1) {
      String angleType = cmd.substring(9, firstSplit);
      String angleValue = cmd.substring(firstSplit + 1);
      
      // Preparar datos para el formato del TP2
      memset(data, 0, sizeof(data)); // Limpiar array
      
      // Primero byte: tipo de ángulo (R, C, O)
      data[0] = angleType.charAt(0);
      
      // Siguientes bytes: valor ASCII del ángulo
      int idx = 1;
      for (int i = 0; i < angleValue.length() && idx < 8; i++) {
        data[idx++] = angleValue.charAt(i);
      }
      
      // Enviar con ID 0x100 (se puede personalizar)
      byte sndStat = CAN0.sendMsgBuf(0x100, 0, idx, data);
      
      if (sndStat == CAN_OK) {
        Serial.println("TP2_ANGLE_SENT_OK");
      } else {
        Serial.println("TP2_ANGLE_SENT_FAIL");
      }
    }
  }
  else {
    Serial.println("UNKNOWN_COMMAND");
  }
}

void loop() {
  // Verificar datos recibidos por serial
  while (Serial.available()) {
    char inChar = (char)Serial.read();
    if (inChar == '\n') {
      stringComplete = true;
    } else {
      inputString += inChar;
    }
  }
  
  // Procesar comando completo
  if (stringComplete) {
    processCommand(inputString);
    inputString = "";
    stringComplete = false;
  }
  
  // Verificar mensajes CAN recibidos
  if (!digitalRead(CAN0_INT)) {
    CAN0.readMsgBuf(&rxId, &len, rxBuf);
    
    // Formato de salida: CAN_RX_ID_LEN_BYTE1_BYTE2_...
    Serial.print("CAN_RX_");
    Serial.print(rxId, HEX);
    Serial.print("_");
    Serial.print(len);
    
    // Imprimir bytes como hexadecimal
    for (byte i = 0; i < len; i++) {
      Serial.print("_");
      Serial.print(rxBuf[i], HEX);
    }
    
    // Añadir interpretación del TP2 si el formato corresponde
    if (len >= 2) {
      char angleType = rxBuf[0];
      if (angleType == 'R' || angleType == 'C' || angleType == 'O') {
        Serial.print("_TP2_");
        Serial.print((char)angleType);
        Serial.print("_");
        
        // Convertir bytes de datos a string
        char angleStr[8] = {0};
        for (int i = 1; i < len && i < 8; i++) {
          angleStr[i-1] = rxBuf[i];
        }
        
        Serial.print(angleStr);
      }
    }
    
    Serial.println();
  }
}