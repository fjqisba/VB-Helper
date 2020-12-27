#include "ComParser.h"
#include <diskio.hpp>
#include <objbase.h>
#include <atlbase.h>
#include <atlcom.h>
#include <comdef.h>
#include <set>
#include "public.h"

struct mid_INTERFACE
{
	qstring m_BaseClassName;			//继承的接口类名称
	qstring m_INTERFACEGUID;			//接口类的GUID
	qstring m_INTERFACEName;			//接口类的名称
	qstring m_COCLASSGUID;				//所属COCLASS的guid
	qstring m_COCLASSName;				//所属COCLASS的名称
	qvector<qstring> mVec_VTableFunc;	//成员函数名称数组
};

struct mid_COMINFO
{
	qstring m_LibraryGUID;								//libraryGUID
	qstring m_LibraryName;								//library
	qstring m_LibraryHelp;								//helpstring
	qvector<mid_INTERFACE> m_Interface;					//所有的COM接口

	//辅助成员
	std::set<qstring> mSet_Interface;					//用于记录接口是否被解析过
};

qstring StringFromGUID(REFGUID guid)
{
	qstring ret;
	ret.sprnt("%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X", guid.Data1, guid.Data2, guid.Data3,
	guid.Data4[0], guid.Data4[1], guid.Data4[2], guid.Data4[3], guid.Data4[4], guid.Data4[5], guid.Data4[6], guid.Data4[7]);

	return ret;
}

inline qstring GetStringFromBSTR(_bstr_t& bstr)
{
	qstring ret;
	utf16_utf8(&ret, bstr.GetBSTR());
	return ret;
}

qstring GetBaseClassName(ITypeInfo* pti)
{
	qstring ret;
	TYPEATTR* pTypeAttr = NULL;
	ITypeInfo* pBaseTypeInfo = NULL;

	try
	{
		if (FAILED(pti->GetTypeAttr(&pTypeAttr)))
		{
			throw NULL;
		}
		for (unsigned int n = 0; n < pTypeAttr->cImplTypes; n++)
		{
			HREFTYPE href = NULL;
			if (FAILED(pti->GetRefTypeOfImplType(n, &href)))
			{
				throw NULL;
			}

			if (FAILED(pti->GetRefTypeInfo(href, &pBaseTypeInfo)))
			{
				throw NULL;
			}

			_bstr_t bstrName, bstrDoc, bstrHelp;
			if (FAILED(pBaseTypeInfo->GetDocumentation(MEMBERID_NIL, bstrName.GetAddress(), bstrDoc.GetAddress(), nullptr, bstrHelp.GetAddress())))
			{
				throw NULL;
			}
			ret = GetStringFromBSTR(bstrName);
		}
	}
	catch (...)
	{
		if (pTypeAttr)
		{
			pti->ReleaseTypeAttr(pTypeAttr);
		}
		if (pBaseTypeInfo)
		{
			pBaseTypeInfo->Release();
		}
	}

	if (pTypeAttr)
	{
		pti->ReleaseTypeAttr(pTypeAttr);
	}
	if (pBaseTypeInfo)
	{
		pBaseTypeInfo->Release();
	}
	return ret;
}

//反编译接口信息,参数bCOCLASS表示接口是否来源于COCLASS
bool DeCompileInterfaceInfo(ITypeInfo* pti, mid_INTERFACE& eInterface)
{
	TYPEATTR* pattr = NULL;
	if (FAILED(pti->GetTypeAttr(&pattr)))
	{
		return false;
	}

	_bstr_t bstrName;
	pti->GetDocumentation(MEMBERID_NIL, bstrName.GetAddress(), nullptr, nullptr, nullptr);
	
	eInterface.m_INTERFACEName = GetStringFromBSTR(bstrName);
	eInterface.m_INTERFACEGUID = StringFromGUID(pattr->guid);

	eInterface.m_BaseClassName = GetBaseClassName(pti);
	if (eInterface.m_BaseClassName.empty())
	{
		pti->ReleaseTypeAttr(pattr);
		return false;
	}

	if (pattr->wTypeFlags & TYPEFLAG_FDUAL)
	{
		HREFTYPE iRefType = NULL;
		if (FAILED(pti->GetRefTypeOfImplType(-1, &iRefType)))
		{
			//说明是单纯的dispinterface接口,不是我们所需要的
			return true;
		}
		pti->ReleaseTypeAttr(pattr);
		ITypeInfo* tmpType = NULL;
		pti->GetRefTypeInfo(iRefType, &tmpType);
		pti->Release();
		pti = tmpType;
		pti->GetTypeAttr(&pattr);
	}

	eInterface.m_BaseClassName = GetBaseClassName(pti);
	if (eInterface.m_BaseClassName.empty())
	{
		pti->ReleaseTypeAttr(pattr);
		return true;
	}

	if (pattr->cFuncs && pattr->typekind == TKIND_INTERFACE)
	{
		unsigned int VtableCount = pattr->cbSizeVft / pattr->cbSizeInstance;
		unsigned int* pVtableIndex = new unsigned int[VtableCount];
		memset(pVtableIndex, -1, sizeof(unsigned int) * VtableCount);

		unsigned int StartIndex = 3;
		for (unsigned int n = 0; n < pattr->cImplTypes; ++n)
		{
			HREFTYPE baseType;
			if (FAILED(pti->GetRefTypeOfImplType(n, &baseType)))
			{
				return false;
			}
			ITypeInfo* pBaseInfo = NULL;
			if (FAILED(pti->GetRefTypeInfo(baseType, &pBaseInfo)))
			{
				return false;
			}
			TYPEATTR* pAttr_Base = NULL;
			pBaseInfo->GetTypeAttr(&pAttr_Base);

			StartIndex = pAttr_Base->cbSizeVft / pAttr_Base->cbSizeInstance;

			pBaseInfo->ReleaseTypeAttr(pAttr_Base);
			pBaseInfo->Release();
		}

		for (UINT iFunc = 0; iFunc < pattr->cFuncs; ++iFunc)
		{
			FUNCDESC* pFuncDesc = NULL;
			pti->GetFuncDesc(iFunc, &pFuncDesc);

			unsigned int index = pFuncDesc->oVft / sizeof(int*);
			if (index >= VtableCount)
			{
				return false;
			}

			pti->ReleaseFuncDesc(pFuncDesc);
			pVtableIndex[index] = iFunc;
		}

		for (unsigned int iIndex = StartIndex; iIndex < VtableCount; ++iIndex)
		{
			qstring FuncName;
			if (pVtableIndex[iIndex] == -1)
			{
				FuncName.sprnt("Missing%d", iIndex);
				eInterface.mVec_VTableFunc.push_back(FuncName);
				continue;
			}
			FUNCDESC* pRealFunc = NULL;
			pti->GetFuncDesc(pVtableIndex[iIndex], &pRealFunc);
			if (pRealFunc->invkind & INVOKE_PROPERTYGET)
			{
				FuncName += "Get_";
			}
			if (pRealFunc->invkind & INVOKE_PROPERTYPUT)
			{
				FuncName += "Set_";
			}
			if (pRealFunc->invkind & INVOKE_PROPERTYPUTREF)
			{
				FuncName += "Ref_";
			}

			_bstr_t  name;
			pti->GetDocumentation(pRealFunc->memid, name.GetAddress(), nullptr, nullptr, nullptr);
			FuncName += GetStringFromBSTR(name);
			eInterface.mVec_VTableFunc.push_back(FuncName);

			pti->ReleaseFuncDesc(pRealFunc);
		}
		delete[] pVtableIndex;
	}

	pti->ReleaseTypeAttr(pattr);
	return true;
}

bool DeCompileCoClass(ITypeInfo* pti, mid_COMINFO& COMInfo)
{
	TYPEATTR* pattr = NULL;
	ITypeInfo* ptiImpl = NULL;

	if (FAILED(pti->GetTypeAttr(&pattr)))
	{
		return false;
	}

	qstring coClassGuid = StringFromGUID(pattr->guid);

	_bstr_t name;
	if (FAILED(pti->GetDocumentation(MEMBERID_NIL, name.GetAddress(), nullptr, nullptr, nullptr)))
	{
		return false;
	}

	qstring coClassName = GetStringFromBSTR(name);;

	for (UINT n = 0; n < pattr->cImplTypes; n++)
	{
		HREFTYPE href = NULL;
		int impltype = NULL;
		if (FAILED(pti->GetImplTypeFlags(n, &impltype)))
		{
			return false;
		}

		if (FAILED(pti->GetRefTypeOfImplType(n, &href)))
		{
			return false;
		}

		if (FAILED(pti->GetRefTypeInfo(href, &ptiImpl)))
		{
			return false;
		}
		
		mid_INTERFACE eInterface;
		eInterface.m_COCLASSName = coClassName;
		eInterface.m_COCLASSGUID = coClassGuid;

		if (DeCompileInterfaceInfo(ptiImpl, eInterface))
		{
			COMInfo.m_Interface.push_back(eInterface);
			COMInfo.mSet_Interface.insert(eInterface.m_INTERFACEName);
		}

		ptiImpl->Release();
		ptiImpl = NULL;
	}

	pti->ReleaseTypeAttr(pattr);
	return true;
}

bool DeCompileTypedef(ITypeInfo* pti, mid_COMINFO& COMInfo)
{
	TYPEATTR* pattr = NULL;

	_bstr_t		bstrName;


	mid_INTERFACE eInterface;
	try
	{
		if (FAILED(pti->GetTypeAttr(&pattr)))
		{
			throw NULL;
		}
		if (IsEqualGUID(pattr->guid, GUID_NULL))
		{
			throw NULL;
		}

		eInterface.m_INTERFACEGUID = StringFromGUID(pattr->guid);

		pti->GetDocumentation(MEMBERID_NIL, bstrName.GetAddress(), nullptr, nullptr, nullptr);
		eInterface.m_INTERFACEName = bstrName;
		if (pattr->typekind != TKIND_ALIAS)
		{
			//还没遇到过
			throw NULL;
		}

		TYPEDESC* ptdesc = &pattr->tdescAlias;
		if (ptdesc->vt != VT_USERDEFINED)
		{
			//还没遇到过
			throw NULL;
		}

		ITypeInfo* ptiRefType = NULL;
		if (FAILED(pti->GetRefTypeInfo(ptdesc->hreftype, &ptiRefType)))
		{
			throw NULL;
		}
		ptiRefType->GetDocumentation(MEMBERID_NIL, bstrName.GetAddress(), nullptr, nullptr, nullptr);

		qstring TypeDefInterfaceName = GetStringFromBSTR(bstrName);
		ptiRefType->Release();

		bool bFind = false;
		for (unsigned int n = 0; n < COMInfo.m_Interface.size(); ++n)
		{
			if (COMInfo.m_Interface[n].m_INTERFACEName == TypeDefInterfaceName)
			{
				eInterface.m_BaseClassName = COMInfo.m_Interface[n].m_BaseClassName;
				eInterface.m_COCLASSName = COMInfo.m_Interface[n].m_COCLASSName;
				eInterface.m_COCLASSGUID = COMInfo.m_Interface[n].m_COCLASSGUID;
				eInterface.mVec_VTableFunc = COMInfo.m_Interface[n].mVec_VTableFunc;
				COMInfo.m_Interface.push_back(eInterface);
				bFind = true;
				break;
			}
		}

		if (!bFind)
		{
			throw NULL;
		}

		pti->ReleaseTypeAttr(pattr);
	}
	catch (...)
	{
		if (pattr)
		{
			pti->ReleaseTypeAttr(pattr);
		}
	}

	return true;
}

//解析COM组件信息，得到COMInfo
bool ParseComInfo(qstring DllPath, mid_COMINFO& COMInfo)
{
	qwstring wDllPath;
	utf8_utf16(&wDllPath, DllPath.c_str());

	ITypeLib* ptlb = NULL;
	TLIBATTR* attr_Library = NULL;

	bool ret = true;
	try
	{
		if (FAILED(LoadTypeLib(wDllPath.c_str(), &ptlb)))
		{
			throw NULL;
		}

		if (FAILED(ptlb->GetLibAttr(&attr_Library)))
		{
			throw NULL;
		}

		COMInfo.m_LibraryGUID = StringFromGUID(attr_Library->guid);
		ptlb->ReleaseTLibAttr(attr_Library);

		//解析Library
		_bstr_t bstrName, bstrDoc, bstrHelp;
		if (FAILED(ptlb->GetDocumentation(MEMBERID_NIL, bstrName.GetAddress(), bstrDoc.GetAddress(), nullptr, bstrHelp.GetAddress())))
		{
			throw NULL;
		}

		COMInfo.m_LibraryName = GetStringFromBSTR(bstrName);
		if (bstrDoc.length())
		{
			COMInfo.m_LibraryHelp = GetStringFromBSTR(bstrDoc);
		}

		//先解析所有的COCLASS接口
		for (unsigned int n = 0; n < ptlb->GetTypeInfoCount(); ++n)
		{
			TYPEKIND typeKind{};
			ptlb->GetTypeInfoType(n, &typeKind);

			ITypeInfo* pTypeInfo = NULL;
			ptlb->GetTypeInfo(n, &pTypeInfo);

			if (typeKind == TKIND_COCLASS)
			{
				ret = DeCompileCoClass(pTypeInfo, COMInfo);
			}
			pTypeInfo->Release();
		}

		//解析独立的Interface接口
		for (unsigned int n = 0; n < ptlb->GetTypeInfoCount(); ++n)
		{
			ITypeInfo* pTypeInfo = NULL;
			ptlb->GetTypeInfo(n, &pTypeInfo);

			TYPEATTR* TypeAttr = NULL;
			pTypeInfo->GetTypeAttr(&TypeAttr);
			if (TypeAttr->typekind == TKIND_INTERFACE || TypeAttr->typekind == TKIND_DISPATCH)
			{
				_bstr_t bstrName;
				pTypeInfo->GetDocumentation(MEMBERID_NIL, bstrName.GetAddress(), nullptr, nullptr, nullptr);

				qstring InterfaceName = GetStringFromBSTR(bstrName);
				if (!COMInfo.mSet_Interface.count(InterfaceName))
				{
					mid_INTERFACE eInterface;
					eInterface.m_COCLASSGUID = "00000000-0000-0000-0000-000000000000";
					if (DeCompileInterfaceInfo(pTypeInfo, eInterface))
					{
						COMInfo.m_Interface.push_back(eInterface);
						COMInfo.mSet_Interface.insert(InterfaceName);
					}
				}
			}
			pTypeInfo->ReleaseTypeAttr(TypeAttr);
			pTypeInfo->Release();
		}

		//处理TypeDef
		for (unsigned int n = 0; n < ptlb->GetTypeInfoCount(); ++n)
		{
			TYPEKIND typeKind{};
			ptlb->GetTypeInfoType(n, &typeKind);

			ITypeInfo* pTypeInfo = NULL;
			ptlb->GetTypeInfo(n, &pTypeInfo);

			if (typeKind == TKIND_ALIAS)
			{
				ret = DeCompileTypedef(pTypeInfo, COMInfo);
			}
			pTypeInfo->Release();
		}
	}
	catch (...)
	{
		if (ptlb)
		{
			ptlb->Release();
			ptlb = NULL;
		}
		ret = false;
	}

	return ret;
}

void GenerateComHpp(const qvector<qstring>& vec_comPath)
{
	if (FAILED(CoInitialize(nullptr)))
	{
		return;
	}
	
	qstring qstr_hppFile = idadir(PLG_SUBDIR);
	qstr_hppFile += "\\COM\\";
	qstr_hppFile += VBHPPFILE;
	
	FILE* fpOut = qfopen(qstr_hppFile.c_str(), "w");

	qstring 用户提示 = "//请勿随意修改此文件内容!\n";
	qfwrite(fpOut, 用户提示.c_str(), 用户提示.length());

	//用来检测是否存在相同的结构体名称
	std::map<qstring, bool> map_Name;

	//用来将原始的接口名映射到新的名称
	std::map<qstring, qstring> map_ChangeName;
	map_ChangeName["IDispatch"] = "IDispatchVtbl";
	map_ChangeName["IUnknown"] = "IUnknownVtbl";

	for (unsigned int nDllIndex = 0; nDllIndex < vec_comPath.size(); ++nDllIndex)
	{
		mid_COMINFO info;
		if (!ParseComInfo(vec_comPath[nDllIndex], info))
		{
			continue;
		}
		
		int nPos = vec_comPath[nDllIndex].rfind('\\');
		qstring DllName = vec_comPath[nDllIndex].substr(nPos + 1);
		int xPos = DllName.rfind('.');
		DllName = DllName.substr(0, xPos);

		//已经解析好的结构体名称
		std::set<qstring> set_CheckOK;
		//支持的基类
		std::set<qstring> set_SupportBaseClass;
		set_SupportBaseClass.insert("IDispatch");
		set_SupportBaseClass.insert("IUnknown");

		bool bCheckAgain = false;

		do
		{
			bCheckAgain = false;
			for (unsigned int n = 0; n < info.m_Interface.size(); ++n)
			{
				mid_INTERFACE& eInterface = info.m_Interface.at(n);

				//遇到没有解析过的结构体
				if (!set_CheckOK.count(eInterface.m_COCLASSName + eInterface.m_INTERFACEName))
				{
					//遇到不支持的基类
					if (!set_SupportBaseClass.count(eInterface.m_BaseClassName))
					{
						bCheckAgain = true;
						continue;
					}

					if (map_ChangeName[eInterface.m_BaseClassName].empty())
					{
						bCheckAgain = true;
						continue;
					}

					qstring coclassGuid;
					coclassGuid.sprnt("//%s\n", eInterface.m_COCLASSGUID.c_str());
					qfwrite(fpOut, coclassGuid.c_str(), coclassGuid.length());

					qstring InterfaceGuid;
					InterfaceGuid.sprnt("//%s\n", eInterface.m_INTERFACEGUID.c_str());
					qfwrite(fpOut, InterfaceGuid.c_str(), InterfaceGuid.length());

					//结构体虚表名称 = 模块名称 + COCLASS名称 + 接口名称
					qstring StructName = DllName + "_" + eInterface.m_COCLASSName + eInterface.m_INTERFACEName;
					if (!map_Name[StructName])
					{
						map_Name[StructName] = true;
					}
					else
					{
						msg("[GenerateComHpp]:Found the same Struct,ignored.\n");
					}

					qstring VTableName = StructName + "_vtbl";
					map_ChangeName[eInterface.m_INTERFACEName] = VTableName;

					qstring StructVTLine;
					StructVTLine.sprnt("struct %s : %s\n{\n", VTableName.c_str(), map_ChangeName[eInterface.m_BaseClassName].c_str());
					qfwrite(fpOut, StructVTLine.c_str(), StructVTLine.length());

					for (unsigned int iFuncIndex = 0; iFuncIndex < eInterface.mVec_VTableFunc.size(); ++iFuncIndex)
					{
						qstring eFuncLine;
						eFuncLine.sprnt("\tint %s;\n", eInterface.mVec_VTableFunc[iFuncIndex].c_str());
						qfwrite(fpOut, eFuncLine.c_str(), eFuncLine.length());
					}

					qstring EndLine = "};\n";
					qfwrite(fpOut, EndLine.c_str(), EndLine.length());

					qstring StructLine;
					StructLine.sprnt("struct %s\n{\n\t%s* lpvt;\n};\n", StructName.c_str(), VTableName.c_str());
					qfwrite(fpOut, StructLine.c_str(), StructLine.length());

					set_CheckOK.insert(eInterface.m_COCLASSName + eInterface.m_INTERFACEName);
					set_SupportBaseClass.insert(eInterface.m_INTERFACEName);
					bCheckAgain = true;
				}
			}
		} while (bCheckAgain);
	}

	qfclose(fpOut);
	CoUninitialize();
}