#include "Driver.h"

NTSTATUS
MoshushouRequestGetHidXferPacketToRead(
    _In_ WDFREQUEST Request,
    _Out_ HID_XFER_PACKET* Packet
    )
{
    WDFMEMORY inputMemory;
    WDFMEMORY outputMemory;
    size_t inputLength;
    size_t outputLength;
    PVOID inputBuffer;
    PVOID outputBuffer;
    NTSTATUS status;

    status = WdfRequestRetrieveInputMemory(Request, &inputMemory);
    if (!NT_SUCCESS(status))
    {
        return status;
    }

    inputBuffer = WdfMemoryGetBuffer(inputMemory, &inputLength);
    if (inputLength < sizeof(UCHAR))
    {
        return STATUS_INVALID_BUFFER_SIZE;
    }

    status = WdfRequestRetrieveOutputMemory(Request, &outputMemory);
    if (!NT_SUCCESS(status))
    {
        return status;
    }

    outputBuffer = WdfMemoryGetBuffer(outputMemory, &outputLength);
    Packet->reportId = *(PUCHAR)inputBuffer;
    Packet->reportBuffer = (PUCHAR)outputBuffer;
    Packet->reportBufferLen = (ULONG)outputLength;
    return STATUS_SUCCESS;
}

NTSTATUS
MoshushouRequestGetHidXferPacketToWrite(
    _In_ WDFREQUEST Request,
    _Out_ HID_XFER_PACKET* Packet
    )
{
    WDFMEMORY inputMemory;
    WDFMEMORY outputMemory;
    size_t inputLength;
    size_t outputLength;
    PVOID inputBuffer;
    NTSTATUS status;

    status = WdfRequestRetrieveOutputMemory(Request, &outputMemory);
    if (!NT_SUCCESS(status))
    {
        return status;
    }

    (VOID)WdfMemoryGetBuffer(outputMemory, &outputLength);
    status = WdfRequestRetrieveInputMemory(Request, &inputMemory);
    if (!NT_SUCCESS(status))
    {
        return status;
    }

    inputBuffer = WdfMemoryGetBuffer(inputMemory, &inputLength);
    Packet->reportId = (UCHAR)outputLength;
    Packet->reportBuffer = (PUCHAR)inputBuffer;
    Packet->reportBufferLen = (ULONG)inputLength;
    return STATUS_SUCCESS;
}

static NTSTATUS
MoshushouGetStringId(
    _In_ WDFREQUEST Request,
    _Out_ ULONG* StringId
    )
{
    WDFMEMORY inputMemory;
    size_t inputLength;
    PULONG inputValue;
    NTSTATUS status;

    status = WdfRequestRetrieveInputMemory(Request, &inputMemory);
    if (!NT_SUCCESS(status))
    {
        return status;
    }

    inputValue = (PULONG)WdfMemoryGetBuffer(inputMemory, &inputLength);
    if (inputLength < sizeof(ULONG))
    {
        return STATUS_INVALID_BUFFER_SIZE;
    }

    *StringId = *inputValue & 0xFFFF;
    return STATUS_SUCCESS;
}

NTSTATUS
MoshushouGetString(
    _In_ WDFREQUEST Request,
    _In_ BOOLEAN Indexed
    )
{
    ULONG stringId;
    const WCHAR* value;
    size_t bytes;
    NTSTATUS status;

    status = MoshushouGetStringId(Request, &stringId);
    if (!NT_SUCCESS(status))
    {
        return status;
    }

    if (Indexed)
    {
        if (stringId != 1)
        {
            return STATUS_INVALID_PARAMETER;
        }

        value = MOSHUSHOU_PRODUCT_STRING;
        bytes = sizeof(MOSHUSHOU_PRODUCT_STRING);
    }
    else
    {
        switch (stringId)
        {
        case HID_STRING_ID_IMANUFACTURER:
            value = MOSHUSHOU_MANUFACTURER_STRING;
            bytes = sizeof(MOSHUSHOU_MANUFACTURER_STRING);
            break;

        case HID_STRING_ID_IPRODUCT:
            value = MOSHUSHOU_PRODUCT_STRING;
            bytes = sizeof(MOSHUSHOU_PRODUCT_STRING);
            break;

        case HID_STRING_ID_ISERIALNUMBER:
            value = MOSHUSHOU_SERIAL_STRING;
            bytes = sizeof(MOSHUSHOU_SERIAL_STRING);
            break;

        default:
            return STATUS_INVALID_PARAMETER;
        }
    }

    return MoshushouRequestCopyFromBuffer(Request, value, bytes);
}
