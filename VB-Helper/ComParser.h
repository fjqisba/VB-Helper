#ifndef ComParser_h__
#define ComParser_h__

#include <pro.h>
#include <diskio.hpp>

//根据COM组件生成hpp头文件
void GenerateComHpp(const qvector<qstring>& vec_comPath);

#endif // ComParser_h__
