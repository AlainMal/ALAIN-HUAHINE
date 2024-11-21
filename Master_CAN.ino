#include "mcp_can.h"
#include <SPI.h>

MCP_CAN CAN0(10); // CS pin 10 for MCP2515

void setup() {
  Serial.begin(500000);
  if (CAN0.begin(MCP_ANY, CAN_250KBPS, MCP_8MHZ) == CAN_OK) {
  //  Serial.println("CAN Init OK");
  } else {
    Serial.println("CAN Init Failed");
    while (1);
  }
  CAN0.setMode(MCP_NORMAL);
}

void loop() {
  if (CAN0.checkReceive() == CAN_MSGAVAIL) {
    long unsigned int rxId;
    unsigned char len = 0;
    unsigned char rxBuf[8];
    CAN0.readMsgBuf(&rxId, &len, rxBuf);
    //Serial.print("Message ID: "); j'ai modifier pour avoir les trames lisibles
    Serial.print(".");
    Serial.print(rxId, HEX);
    Serial.print(";");
    Serial.print(len, HEX);
    Serial.print(":");
    for (int i = 0; i < len; i++) {
      Serial.print(rxBuf[i], HEX);
      if (i!=len-1) {Serial.print(",");}
    }
    Serial.print("?");
  }
}
