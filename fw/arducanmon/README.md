# arducanmon

CAN<->PC gateway implementation on Arduino UNO using the MCP2515 (SPI<->CAN CNTRLLR) + TJA1050 (XCVR) module.

Connections:

* VCC: 5V
* INT: 2
* CS: 10
* SCK: 13
* MOSI (SI): 12
* MISO (SO): 11

If the board is a CAN terminal node, place the jumper on the 120Ohms terminator.