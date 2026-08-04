#pragma once

#include <stdint.h>

#define MOSHUSHOU_HID_REPORT_ID_KEYBOARD        0x01
#define MOSHUSHOU_HID_REPORT_ID_ABSOLUTE_MOUSE  0x04
#define MOSHUSHOU_HID_REPORT_ID_CONTROL         0x40

#define MOSHUSHOU_HID_CONTROL_REPORT_SIZE       65
#define MOSHUSHOU_HID_CONTROL_PAYLOAD_SIZE      63

#pragma pack(push, 1)

typedef struct _MOSHUSHOU_KEYBOARD_REPORT
{
    uint8_t ReportId;
    uint8_t Modifiers;
    uint8_t Reserved;
    uint8_t Keys[6];
} MOSHUSHOU_KEYBOARD_REPORT;

typedef struct _MOSHUSHOU_ABSOLUTE_MOUSE_REPORT
{
    uint8_t ReportId;
    uint8_t Buttons;
    uint16_t X;
    uint16_t Y;
    int8_t Wheel;
} MOSHUSHOU_ABSOLUTE_MOUSE_REPORT;

typedef struct _MOSHUSHOU_CONTROL_REPORT
{
    uint8_t ReportId;
    uint8_t PayloadLength;
    uint8_t Payload[MOSHUSHOU_HID_CONTROL_PAYLOAD_SIZE];
} MOSHUSHOU_CONTROL_REPORT;

#pragma pack(pop)

typedef char MOSHUSHOU_KEYBOARD_REPORT_SIZE_CHECK[
    sizeof(MOSHUSHOU_KEYBOARD_REPORT) == 9 ? 1 : -1];
typedef char MOSHUSHOU_MOUSE_REPORT_SIZE_CHECK[
    sizeof(MOSHUSHOU_ABSOLUTE_MOUSE_REPORT) == 7 ? 1 : -1];
typedef char MOSHUSHOU_CONTROL_REPORT_SIZE_CHECK[
    sizeof(MOSHUSHOU_CONTROL_REPORT) == MOSHUSHOU_HID_CONTROL_REPORT_SIZE ? 1 : -1];
