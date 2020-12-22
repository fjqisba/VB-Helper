#include "ImportsFix.h"
#include "public.h"

std::map<qstring, qstring> map_VBVM;

int idaapi enumVBImports(ea_t ea, const char* name, uval_t ord, void* param)
{
	qstring FuncName;
	if (!name)
	{
		get_ea_name(&FuncName, ea);
	}
	else
	{
		FuncName = name;
	}

	if (map_VBVM[FuncName] != "")
	{
		apply_cdecl(idati, ea, map_VBVM[FuncName].c_str(), TINFO_DEFINITE);
	}

	//msg("ord:%d\n", ord);
	return 1;
}

void FixVBImports()
{
	if (map_VBVM.size() == 0)
	{
		map_VBVM["__vbaAryDestruct"] = "void __stdcall _vbaAryDestruct(int, SAFEARRAY **);";
		map_VBVM["__vbaAryMove"] = "SAFEARRAY* __stdcall _vbaAryMove(SAFEARRAY **, SAFEARRAY **);";
		map_VBVM["__vbaExitProc"] = "void __stdcall _vbaExitProc();";
		map_VBVM["__vbaEnd"] = "void __usercall __noreturn _vbaEnd(int @<esi>, int @<edi>);";

		map_VBVM["__vbaErrorOverflow"] = "void __stdcall __noreturn _vbaErrorOverflow();";

		map_VBVM["__vbaForEachCollObj"] = "void __stdcall _vbaForEachCollObj(const GUID *, LPVOID, LPVOID, LPVOID);";
		map_VBVM["__vbaFreeObj"] = "void __thiscall _vbaFreeObj(LPVOID);";
		map_VBVM["__vbaFreeStr"] = "void __fastcall _vbaFreeStr(BSTR *);";
		map_VBVM["__vbaFreeVar"] = "void __fastcall _vbaFreeVar(VARIANT *);";
		map_VBVM["__vbaFreeVarList"] = "void _vbaFreeVarList(int, ...);";

		map_VBVM["__vbaHresultCheckObj"] = "int __stdcall _vbaHresultCheckObj(int, void *,const GUID *, unsigned int);";

		map_VBVM["__vbaI2I4"] = "short __fastcall _vbaI2I4(int);";
		map_VBVM["__vbaI2Str"] = "SHORT __stdcall _vbaI2Str(BSTR);";
		map_VBVM["__vbaI2Var"] = "short __stdcall _vbaI2Var(VARIANT *);";
		map_VBVM["__vbaI4Var"] = "int __stdcall _vbaI4Var(VARIANT *);";

		map_VBVM["__vbaLateIdCall"] = "void _vbaLateIdCall(LPVOID, ...);";
		map_VBVM["__vbaLbound"] = "int __stdcall _vbaLbound(int , SAFEARRAY *);";
		map_VBVM["__vbaLenBstr"] = "int __stdcall _vbaLenBstr(BSTR );";
		map_VBVM["__vbaLenVar"] = "VARIANT* __stdcall _vbaLenVar(VARIANT *out_Len, VARIANT *);";

		map_VBVM["__vbaNew"] = "LPVOID __stdcall _vbaNew(LPVOID );";
		map_VBVM["__vbaNew2"] = "LPVOID __stdcall _vbaNew(LPVOID , LPVOID );";
		map_VBVM["__vbaObjSet"] = "LPVOID __stdcall _vbaObjSet(LPVOID , LPVOID );";

		map_VBVM["__vbaOnError"] = "void __stdcall _vbaOnError(int);";

		map_VBVM["__vbaRedim"] = "SAFEARRAY* __cdecl _vbaRedim(int , int , SAFEARRAY **ppsaOut, int , int , int , int);";

		map_VBVM["__vbaStrCat"] = "BSTR __stdcall _vbaStrCat(BSTR , BSTR );";
		map_VBVM["__vbaStrCmp"] = "int __stdcall _vbaStrCmp(BSTR , BSTR );";
		map_VBVM["__vbaStrCopy"] = "BSTR __fastcall _vbaStrCopy(BSTR *, BSTR );";

		map_VBVM["__vbaStrI2"] = "BSTR __stdcall _vbaStrI2(SHORT);";
		map_VBVM["__vbaStrI4"] = "BSTR __stdcall _vbaStrI4(int);";

		map_VBVM["__vbaStrMove"] = "BSTR __fastcall _vbaStrMove(BSTR *, BSTR);";

		map_VBVM["__vbaStrVarMove"] = "BSTR __stdcall _vbaStrVarMove(VARIANT *);";

		map_VBVM["__vbaUbound"] = "int __stdcall _vbaUbound(int , SAFEARRAY *);";
		map_VBVM["__vbaUI1I2"] = "unsigned char __fastcall _vbaUI1I2(short );";
		map_VBVM["__vbaVarAdd"] = "VARIANT* __stdcall _vbaVarAdd(VARIANT *, VARIANT *, VARIANT *);";
		map_VBVM["__vbaVarAnd"] = "VARIANT* __stdcall _vbaVarAnd(VARIANT *, VARIANT *, VARIANT *);";

		map_VBVM["__vbaVarCat"] = "VARIANT *__stdcall _vbaVarCat(VARIANT *, VARIANT *, VARIANT *);";

		map_VBVM["__vbaVarMove"] = "VARIANT *__fastcall _vbaVarMove(VARIANT *, VARIANT *);";
		map_VBVM["__vbaVarTstEq"] = "int __stdcall _vbaVarTstEq(VARIANT *, VARIANT *);";
		map_VBVM["__vbaVarTstGt"] = "int __stdcall _vbaVarTstGt(VARIANT *, VARIANT *);";
		map_VBVM["__vbaVarTstNe"] = "int __stdcall _vbaVarTstNe(VARIANT *, VARIANT *);";
		map_VBVM["__vbaVarVargNofree"] = "VARIANT* __fastcall _vbaVarVargNofree(VARIANT *, VARIANT *);";

		//rtc函数
		map_VBVM["rtcIsNumeric"] = "short __stdcall rtcIsNumeric(VARIANT *iVar);";
		map_VBVM["rtcHexVarFromVar"] = "VARIANT *__stdcall rtcHexVarFromVar(VARIANT *, VARIANT *);";
		map_VBVM["rtcLeftCharVar"] = "VARIANT *__stdcall rtcLeftCharVar(VARIANT *, VARIANT *, int);";
		map_VBVM["rtcMidCharVar"] = "VARIANT* __stdcall rtcMidCharVar(VARIANT *oStr, VARIANT *iStr, int offset, VARIANT *SubLen);";
	}

	uint num = get_import_module_qty();
	for (uint index = 0; index < num; index++)
	{
		qstring Module_Name;
		if (!get_import_module_name(&Module_Name, index))		//获取模块名称失败
		{
			continue;
		}

		if (Module_Name == "MSVBVM60")
		{
			enum_import_names(index, enumVBImports);
		}
	}
}