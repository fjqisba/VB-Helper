#pragma once

#define VBHPPFILE "VB.hpp"

#include <typeinf.hpp>
#include <pro.h>

extern til_t* idati;

struct coClassInfo
{
	qstring coClassGuid;
	qstring StructName;
};

struct ObjData
{
	ea_t GuidAddr;
	qstring StructName;
};

qstring get_uuid(ea_t uuidAddr);

qstring get_shortstring(ea_t ea);

qstring charToUUID(uint8* pChar);


