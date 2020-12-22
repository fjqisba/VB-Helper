#pragma once
#include <ida.hpp>
#include <bytes.hpp>
#include <kernwin.hpp>
#include <vector>
#include "public.h"

#include "middleVB.h"

//——————————————————————————————————————————
#pragma pack(push,1)

//结构体大小为0x18
struct DefaultEvent
{
	uint32 dwNull;				//为0,未知
	uint32 EventInfo;			//事件所属对象信息
	uint32 lpObjectInfo;		//指向当前对象的备份指针
	uint32 QueryInterface;		//QueryInterface
	uint32 AddRef;				//AddRef
	uint32 Release;				//Release
};

//结构体大小为0x28
struct ControlInfo
{
	uint16 szNone;				//未知
	uint16 UserEventCount;		//用户事件个数
	uint16 EventSupcount;		//支持的事件处理的个数
	uint16 bWEventsOffset;		//用来拷贝事件的内存结构体的偏移
	uint32 lpGuid;				//指向控件GUID的指针
	uint32 dwIndex;				//当前控件的索引ID
	uint32 dwNull1;				//无用
	uint32 dwNull2;				//无用
	uint32 aEventTable;			//执行事件处理表的指针
	uint32 lpIdeData;			//无用
	uint32 lpszName;			//控件的名称
	uint32 dwIndexCopy;			//当前控件的第二索引ID
};

//结构体大小至少为0x18
struct EventTable
{
	uint32 dwNull;				//为0
	uint32 aControl;			//指向控件信息...
	uint32 aObjectInfo;			//指向对象信息
	uint32 aQueryInterface;		//QueryInterface函数
	uint32 aAddRef;				//AddRef函数
	uint32 aRelease;			//Release函数
	//uint32 aEventPointer;		//事件,实际大小为aEventPointer[aControl.EventCount - 1]
};

//结构体大小为0x34
struct ParentObjInfo
{
	uint16 flag;				//标志值
	uint16 offset;				//某个位置
	//To do...
};

//结构体大小为0x40
struct PrivateObjectInfo
{
	uint32 lpHeapLink;			//值为0,无用
	uint32 lpObjectInfo;		//指向当前私有对象的对象指针,ObjectInfo*
	uint32 dwReserved;			//值为-1,无用
	uint32 szNone1;				//未知
	uint16 wClassCount;			//类变量个数
	uint16 oSize;				//某个重要值
	uint32 szNone3;				//未知
	uint32 lpObjectList;		//****************,指向ParentObjInfo数组,成员大小和ObjectDescription.MethodCount相同
	uint32 IDEData2;			//无用
	uint32 ClassVarList;		//指向类变量内存空间
	uint32 szNone4;				//未知
	uint32 szNone5;				//未知
	uint32 IDEData3[3];			//无用
	uint32 dwObjectType;		//所描述对象的类型
	uint32 dwIdentifier;		//结构体的模板版本
};

//结构体大小为0x38
struct ObjectInfo
{
	uint16 wRefCount;			//值为1,未知
	uint16 ObjectIndex;			//当前对象在工程对象数组中的下标
	uint32 ObjectTable;			//指向ObjectTable的备份指针
	uint32 lpIDEData1;			//值为0,无用
	uint32 lpPrivateObject;		//****************指向私有对象的指针,PrivateObjectInfo*
	uint32 dwReserved;			//值为-1,无用
	uint32 dwNull;				//值为0,无用
	uint32 lpObject;			//指向对象描述信息的备份指针,ObjectDescription*
	uint32 lpProjectData;		//指向内存中的数据
	uint16 szNone1;				//未知
	uint16 szNone2;				//为0,未知
	uint32 szNone3;				//未知数
	uint16 wConstants;			//在常量池中的常量个数
	uint16 wMaxConstants;		//申请的常量个数
	uint32 lpIDEData2;			//无用
	uint32 lpIDEData3;			//无用
	uint32 lpOptionObject;		//如果存在额外信息,则该成员下一个接着是OptionalObjectInfo
};

//结构体大小为0x40
struct OptionalObjectInfo
{
	uint32 lguidObjectGUICount;		//guid的个数
	uint32 aObjectCLSID;			//指向CLSID的指针
	uint32 dwNull;					//未知
	uint32 aguidObjectGUITable;		//指向guidObjectGUI的指针的指针
	uint32 lObjectDefaultIIDCount;	//DefaultIID的个数
	uint32 aObjectEventsIIDTable;	//指向EventsIID的指针的指针
	uint32 lObjectEventsIIDCount;	//EventsIID的个数
	uint32 aObjectDefaultIIDTable;	//指向DefaultIID的指针的指针
	uint32 lControlCount;			//控件个数
	uint32 aControlArray;			//指向控件数组的指针
	int16 iMethodLinkCount;			//链接的函数个数
	uint16 iPCodeCount;				//使用的Opcodes个数
	uint16 oInitializeEvent;		//MethodLinkTable中初始化事件的偏移
	uint16 oTerminateEvent;			//MethodLinkTable中终结事件的偏移
	uint32 MethodLinkTable;			//指向函数链接的指针的指针
	uint32 aBasicClassObject;		//指向一个内存结构
	uint32 szNone;					//无用
	uint32 lFlag;					//未知
};

//结构体大小为0x30
struct ObjectDescription
{
	uint32 ObjectInfo;		//指向ObjectInfo的指针,ObjectInfo*
	uint32 dwReserved;		//值为-1,未知
	uint32 PublicBytes;		//指向Public变量大小整数的指针
	uint32 StaticBytes;		//指向Static变量大小整数的指针
	uint32 ModulePublic;	//指向Public变量数据内存
	uint32 ModuleStatic;	//指向Static变量数据内存
	uint32 NTSObjectName;	//****************对象的名称
	int32 MethodCount;		//****************对象所有的函数个数
	int32 MethodNameTable;	//****************所有函数的名称指针
	uint32 StaticVars;		//ModuleStatic中静态变量的偏移
	uint32 ObjectType;		//****************描述对象的bitmask
	uint32 dwNull;			//无用
};


//结构体大小为0x54
struct ObjectTable
{
	uint32 lpHeapLink;			//永远为0
	uint32 lpExecProj;			//VB Exec COM对象的内存指针
	uint32 lpProjectInfo2;		//第二工程对象信息,ProjectInfo2*
	uint32 dwReserved;			//永远为-1
	uint32 dwNull;				//未被使用
	uint32 lpProjectObject;		//在内存中的工程数据指针
	uint8  uuidObject[0x10];	//对象表的GUID
	uint16 fCompileState;		//编译时候的内部标记
	uint16 dwTotalObjects;		//****************工程中存在的对象总数
	uint16 dwCompiledObjects;	//编译对象个数,和上面的相同
	uint16 dwObjectsInUse;		//使用对象个数,通常和上面的相同
	uint32 lpObjectArray;		//****************指向对象描述信息数组的指针,ObjectDescription*
	uint32 fIdeFlag;			//无用
	uint32 lpIdeData;			//无用
	uint32 lpIdeData2;			//无用
	uint32 lpszProjectName;		//工程名称的指针
	uint32 dwLcid;				//工程的LocaleID 
	uint32 dwLcid2;				//工程的Second LocalID
	uint32 lpIdeData3;			//无用
	uint32 dwIdentifier;		//当前结构体的版本,2
};

//结构体大小为0x23C
struct ProjectInfo
{
	uint32 dwVersion;			//版本号，为0x1F4
	uint32 lpObjectTable;		//****************对象表的指针ObjectTable*
	uint32 dwNull;				//无用
	uint32 lpCodeStart;			//VB用户代码起始地址,未被使用
	uint32 lpCodeEnd;			//VB用户代码结束地址,未被使用
	uint32 dwDataSize;			//VB对象结构体大小,未被使用
	uint32 lpThreadSpace;		//线程对象的二级指针
	uint32 lpVbaSeh;			//VBA异常处理函数指针
	uint32 lpNativeCode;		//数据段指针
	uint8  szNone[0x210];		//空白,Internal Project Identifier
	uint32 lpExternalTable;		//****************API导入表的指针
	uint32 dwExternalCount;		//****************导入API的个数
};


struct ProjectInfo2
{
	uint32 lpHeapLink;				//无用,为0
	uint32 lpObjectTable;			//指向ObjectTable的备份指针
	uint32 dwReserved1;				//无用,为-1
	uint32 szNone1;					//无用
	uint32 lpObjectList;			//指向对象描述信息的指针,不清楚...
	uint32 szNone2;					//无用
	uint32 szProjectDescription;	//指向工程描述的指针
	uint32 szProjectHelpFile;		//指向工程帮助文件的指针
	uint32 dwReserved2;				//无用,为-1
	uint32 dwHelpContextId;			//工程设置中的帮助ID
};

struct SControl
{
	uint8 szNone1;
	uint8 szNone2;
	qstring ControlName;
	uint8 szNone3;
	uint8 ControlID;
};

//结构体大小为0x50
struct GUITable
{
	uint32 StructSize;					//当前结构体的大小,为0x50
	uint8  guidObjectGUI[16];			//当前GUI对象的GUID
	uint8  guidFreezeEventsIID[16];		//Freeze事件的GUID
	uint32 ObjectId;					//在工程中的对象ID
	uint32 ControlFlags;				//定义控件如何使用的flag
	uint32 fOLEMisc;					//OLEMisc flag
	uint8  guidObjectCLSID[16];			//CLSID
	uint32 lGUIObjectInfoSize;			//****************GUI对象信息的大小
	uint32 dwNull;						//未知
	uint32 GUIObjectInfo;				//****************GUI对象信息,GUIObjectInfo*,里面也许包含着控件的名称位置等信息
	uint32 Offset;						//从VBHeader.lpGuiTable到这个成员的偏移
};

//结构体大小为0x44
struct RegistrationInfo
{
	uint32 bNextObject;				//COM接口信息的偏移
	uint32 bObjectName;				//对象名称的偏移
	uint32 bObjectDescription;		//对象描述的偏移
	uint32 dwInstancing;			//实例模式
	uint32 dwObjectId;				//当前对象在工程中的ID
	uint8 uuidObject[16];			//对象的CLSID
	uint32 fIsInterface;			//指示下一个CLSID是否有效
	uint32 bUuidObjectIFace;		//对象的接口的CLSID的偏移
	uint32 bUuidEventsIFace;		//对象事件的CLSID的偏移
	uint32 fHasEvents;				//指明上面的CLSID是否有效
	uint32 dwMiscStatus;			//OLEMIC Flags
	uint8 fClassType;				//Class Type
	uint8 fObjectType;				//对象类型
	uint16 wToolboxBitmap32;		//工具箱中的控件图ID
	uint16 wDefaultIcon;			//控件窗口最小图标
	uint16 fIsDesigner;				//指明是否为设计器
	uint32 bDesignerData;			//设计的数据
};


//大小为0x2A
struct COMRegData
{
	uint32 bRegInfo;				//COM接口信息的偏移,RegistrationInfo*
	uint32 bSZProjectName;			//Project或者TypeLib的偏移
	uint32 bSZHelpDirectory;		//帮助目录的偏移
	uint32 bSZProjectDescription;	//工程描述的偏移
	uint8  uuidProjectCLSID[16];	//Project或者TypeLib的CLSID
	uint32 dwTlbLcid;				//LCID of Type Library
	uint16 szNone;					//未知
	uint16 wTlbVerMajor;			//TypeLib主版本号
	uint16 wTlbVerMinor;			//TypeLib次版本号
};


struct ExternalControlInfo
{
	uint32 NextControlOffset;		//下一个控件的偏移
	uint32 GuidOffset1;				//未知Guid的偏移
	uint32 szNone1;					//未知
	uint32 szNone2;					//未知
	uint32 NoneOffset1;				//未知偏移
	uint32 NoneOffset2;				//未知偏移
	uint32 NoneOffset3;				//未知偏移
	uint32 szNone3;					//未知
	uint32 szNone4;					//未知
	uint32 szNone5;					//未知
	uint32 LibNameOffset;			//导入的库名称偏移
	uint32 ControlOffset;			//控件全称偏移
	uint32 ControlNameOffset;		//控件简称偏移
};

//结构体大小为0x68
struct VBHeader
{
	uint32 szVbMagic;			//VB标记
	uint16 wRuntimeBuild;		//运行库版本?
	uint16 szLangDLL;			//语言扩展DLL
	uint32 szNone1;				//未知
	uint32 szNone2;				//未知
	uint32 szNone3;				//未知
	uint16 szSecLangDLL;		//第二语言扩展DLL
	uint32 szNone4;				//未知
	uint32 szNone5;				//未知
	uint32 szNone6;				//未知
	uint16 wRuntimeRevision;	//****************内部运行库的版本
	uint32 dwLCID;				//语言DLL的LCID
	uint32 dwSecLCID;			//第二语言DLL的LCID
	uint32 lpSubMain;			//****************Main函数入口
	uint32 lpProjectInfo;		//****************工程信息的指针,ProjectInfo*
	uint32 fMdlIntCtls;
	uint32 fMDlIntCtls2;
	uint32 dwThreadFlags;
	uint32 dwThreadCount;
	uint16 wFormCount;				//****************窗口的个数
	uint16 wExternalCount;			//额外控件的个数
	uint32 dwThunkCount;			//
	uint32 lpGuiTable;				//****************窗口信息的指针,GUITable*
	uint32 lpExternalTable;			//指向额外控件信息的指针,ExternalControlInfo*
	uint32 lpComRegisterData;		//注册Com组件信息,COMRegData*
	uint32 bSZProjectDescription;	//工程描述信息的偏移
	uint32 bSZProjectExeName;		//EXE名称的偏移
	uint32 bSZProjectHelpFile;		//工程帮助文件的偏移
	uint32 bSZProjectName;			//****************工程的名称的偏移
};

#pragma pack(pop)


class VBDecompilerEngine
{
public:
	VBDecompilerEngine();
	~VBDecompilerEngine();
private:
	mid_VBHeader VBHEAD;
	void GetSControl(uint32 sControlAddr, SControl& result);

	qstring GetControlTypeName(uint8 ControlID);
	bool HandleNativeCode(uint32 vbHeadAddr);

	bool hasControl(uint32 ObjectType);
	bool hasOptionalObjectInfo(uint32 ObjType);
	qstring GetObjectTypeName(uint32 ObjType);

	bool LoadMethodLink(OptionalObjectInfo& info, qvector<mid_VTable>& oVec_MethodLink);

	//将原始VB信息提炼为中间件VB信息,分为PCode和NativeCode两种
	bool TranslateNVBInfo(uint32 NativeHeadAddr);
	bool TranslatePVBInfo(uint32 PCodeHeadAddr);

	bool VBDE_ComRegData(uint32 lpComRegDataAddr);
	bool VBDE_ExternalControl(uint32 lpExControlAddr, uint32 cCount);
public:
	bool DoDecompile(ea_t PEEntry);
	void MakeFunction();
	void CreateVTable();
	void SetSubMain();
	void SetEventFuncName(std::map<qstring, qvector<coClassInfo>>& oMap);
	void AddClassGuid(std::map<qstring, qvector<coClassInfo>>& oMap);

	//获取用户代码段起始地址
	ea_t GetUserCodeStartAddr();
	//获取用户代码段结束地址
	ea_t GetUserCodeEndAddr();
};

