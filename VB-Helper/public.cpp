#include "public.h"

qstring charToUUID(uint8* uuid_Addr)
{
	char buff_uuid[40] = { 0 };
	qsnprintf(buff_uuid, sizeof(buff_uuid), "%02X%02X%02X%02X-%02X%02X-%02X%02X-%02X%02X%02X%02X%02X%02X%02X%02X",
		uuid_Addr[3], uuid_Addr[2], uuid_Addr[1], uuid_Addr[0],
		uuid_Addr[5], uuid_Addr[4],
		uuid_Addr[7], uuid_Addr[6],
		uuid_Addr[8], uuid_Addr[9], uuid_Addr[10], uuid_Addr[11], uuid_Addr[12], uuid_Addr[13], uuid_Addr[14], uuid_Addr[15]
	);
	qstring uuid = buff_uuid;
	return uuid;
}


qstring get_uuid(ea_t uuid_Addr)
{
	char buff_uuid[40] = { 0 };
	qsnprintf(buff_uuid, sizeof(buff_uuid), "%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X", get_dword(uuid_Addr), get_word(uuid_Addr + 4), get_word(uuid_Addr + 6),
		get_byte(uuid_Addr + 8),
		get_byte(uuid_Addr + 9),
		get_byte(uuid_Addr + 10),
		get_byte(uuid_Addr + 11),
		get_byte(uuid_Addr + 12),
		get_byte(uuid_Addr + 13),
		get_byte(uuid_Addr + 14),
		get_byte(uuid_Addr + 15)
	);
	qstring uuid = buff_uuid;
	return uuid;
}

qstring get_shortstring(ea_t ea)
{
	if (ea <= 0)
	{
		return "";
	}

	char buffer[255] = { 0 };
	if (get_bytes(buffer, sizeof(buffer), ea, GMB_READALL, NULL) != sizeof(buffer))			//没读取到完整的字节应该算是错误了
	{
		return "";
	}
	qstring ret = buffer;
	return ret;
}