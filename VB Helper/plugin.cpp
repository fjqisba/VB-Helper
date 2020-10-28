#include <ida.hpp>
#include <auto.hpp>
#include <hexrays.hpp>
#include <pro.h>
#include <entry.hpp>
#include <map>


#include <windows.h>

hexdsp_t *hexdsp = NULL;
static bool inited = false;
static std::map<qstring, qstring> map_VBVM;

struct VBHeader
{
	DWORD szVbMagic;

	DWORD lpSubMain;
};


int enumIMPORTS(ea_t ea, const char *name, uval_t ord, void *param)
{
	if(name)
	{
		if(map_VBVM[name]!="")
		{
			apply_cdelc(idati,ea,map_VBVM[name].c_str(),TINFO_DEFINITE);
		}
		else
		{
			msg("%s\n",name);
		}
	}
	
	switch(ord)
	{
		case 561:
			apply_cdecl(idati,ea,"short __stdcall rtcIsNumeric(VARIANT* iVar);",TINFO_DEFINITE);
			break;
		case 573:
			apply_cdelc(idati,ea,"VARIANT* __stdcall rtcHexVarFromVar(VARIANT*,VARIANT*);",TINFO_DEFINITE);
			break;
		case 617:
			apply_cdecl(idati,ea,"VARIANT* __stdcall rtcLeftCharVar(VARIANT*,VARIANT*,int);",TINFO_DEFINITE);
			break;
		case 632:
			apply_cdecl(idati,ea,"VARIANT* __stdcall rtcMidCharVar(VARIANT* outStr,VARIANT* Src,int offset,VARIANT* subLen);",TINFO_DEFINITE);
			break;
		default:
			break;
	}
	return 1;
}

void FixVBImports()
{
	if (map_VBVM.size() == 0)
	{
		map_VBVM["__vbaAryMove"]="SAFEARRAY* __stdcall _vbaAryMove(SAFEARRAY **, SAFEARRAY **);";
		map_VBVM["__vbaExitProc"] = "void __stdcall _vbaExitProc();";
		map_VBVM["__vbaEnd"] = "void __usercall __noreturn _vbaEnd(int a1@<esi>, int a2@<edi>);";
		
		map_VBVM["__vbaErrorOverflow"] = "void __stdcall __noreturn _vbaErrorOverflow();";
		
		map_VBVM["__vbaFreeObj"] = "void __thiscall _vbaFreeObj(LPVOID);";
		map_VBVM["__vbaFreeStr"] = "void __fastcall _vbaFreeStr(BSTR *);";
		map_VBVM["__vbaFreeVar"] = "void __fastcall _vbaFreeVar(VARIANT*);";
		
		//To do...
		
		map_VBVM["__vbaNew"] = "LPVOID __stdcall _vbaNew(LPVOID ppv);";
		map_VBVM["__vbaObjSet"] = "LPVOID __stdcall _vbaObjSet(LPVOID a1, LPVOID a2);";
		map_VBVM["__vbaStrCat"] = "BSTR __stdcall _vbaStrCat(wchar_t *, wchar_t *);";

		map_VBVM["__vbaUI1I2"] = "unsigned __int8 __fastcall _vbaUI1I2(unsigned __int16);";
		map_VBVM["__vbaVarAdd"] = "VARIANT* __stdcall _vbaVarAdd(VARIANT *Dest, VARIANT *V1, VARIANT *V2);";
		map_VBVM["__vbaVarMove"] = "VARIANT *__fastcall _vbaVarMove(VARIANT *Dst, VARIANT *Src);";
	}

	int num = get_import_module_qty();
	for (int index = 0; index < num; index++)
	{
		qstring Module_Name;
		if (!get_import_module_name(&Module_Name, index))//获取模块名称失败
		{
			continue;
		}

		if (Module_Name == "MSVBVM60")
		{
			enum_import_names(index, enumIMPORTS);
		}
	}

}

void DoMyVB()
{
	ea_t PEentry = get_entry(get_entry_ordinal(0));

	uchar FirstCommand = get_byte(PEentry);

	uint32 aVBHeaderAddr;
	if (FirstCommand == 0x68)
	{
		aVBHeaderAddr = get_dword(PEentry + 1);		////EXE program
		msg("aVBHeader is at  %.8X\n", aVBHeaderAddr);
	}
	else
	{
		
		return;
	}

	VBHeader vbHead;
	get_bytes(&vbHead, sizeof(vbHead), aVBHeaderAddr);

	if (vbHead.szVbMagic != 0x21354256)		//VB5
	{
		msg("VB Signiture not found \n");
		return;
	}

	if (vbHead.lpSubMain)
	{
		set_name(vbHead.lpSubMain, "SubMain");
	}

	uint16 formcount = get_word(aVBHeaderAddr + 0x44);
	for (unsigned int n = 0; n < formcount; n++)
	{

	}
}

//--------------------------------------------------------------------------
int idaapi init(void)
{
	if (!init_hexrays_plugin())
		return PLUGIN_SKIP; // no decompiler
	const char *hxver = get_hexrays_version();
	msg("Hex-rays version %s has been detected, %s ready to use\n", hxver, PLUGIN.wanted_name);
	inited = true;
	return PLUGIN_KEEP;
}

//--------------------------------------------------------------------------
void idaapi term(void)
{
	if (inited)
		term_hexrays_plugin();
}

//--------------------------------------------------------------------------
bool idaapi run(size_t)
{
	if (auto_is_ok() || ask_yn(0, "The autoanalysis has not finished yet.\nDo you want to continue?") >= 1)
	{
		FixVBImports();
		DoMyVB();
	}

	return true;
}

//--------------------------------------------------------------------------
static char comment[] = "It's a tool for VB programs";


plugin_t PLUGIN =
{
	IDP_INTERFACE_VERSION,
	0,                    // plugin flags
	init,                 // initialize
	term,                 // terminate. this pointer may be NULL.
	run,                  // invoke plugin
	comment,              // long comment about the plugin
	"fjqisba@sohu.com",   // multiline help about the plugin
	"VB Helper",		  // the preferred short name of the plugin
	"Alt + F1"           // the preferred hotkey to run the plugin
};
