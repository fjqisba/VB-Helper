#include "ComManager.h"

#include <hexrays.hpp>
#include <diskio.hpp>


int idaapi EnumerateTypeLib(const char* file, void* ud)
{
	qvector<qstring>* pRecvData = (qvector<qstring>*)ud;
	pRecvData->push_back(file);
	return 0;
}

void InitComManager()
{
	//To do...判断COM目录是否存在

	qstring qstr_idaDir = idadir(PLG_SUBDIR);
	qstring qstr_ComDir = qstr_idaDir + "\\COM";
	qstring qstr_VBHpp = qstr_ComDir + "\\";
	qstr_VBHpp += VBHPPFILE;

	//To do...后续可以增加判断是否需要重新生成VBhpp头文件的功能
	bool needUpdateHpp = false;

	qvector<qstring> vec_ComFile;
	enumerate_files(0, 0, qstr_ComDir.c_str(), "*.*", EnumerateTypeLib, &vec_ComFile);

	
	GenerateComHpp(vec_ComFile);
	
	
	return;
}