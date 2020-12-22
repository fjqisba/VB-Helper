#pragma once
#include <pro.h>

struct mid_ProjectInfo2
{
	uint32 Todo;
};


struct mid_ControlInfo
{
	qstring m_ControlName;			//控件的名称
	qstring m_ControlEventGuid;		//控件对应事件的GUID
	qvector<uint32> mVec_UseEvent;	//控件所有事件的地址
};

enum midEnum_EventType
{
	e_NoneEvent,
	e_UserEvent,
	e_PutMemEvent,
	e_GetMemEvent,
	e_SetMemEvent,
	e_PutMemStr,
	e_GetMemStr,
	e_GetMem2,
	e_PutMem2,
	e_PutMem4,
	e_GetMem4,
	e_PutMemObj,
	e_GetMemObj,
	e_SetMemObj,
	e_GetMemNewObj,
	e_PutMemNewObj,
	e_SetMemNewObj
};

struct mid_VTable
{
	midEnum_EventType m_EventType;
	uint32 m_Offset;
	uint32 m_MethodAddr;
};

struct mid_OptionalObjectInfo
{
	bool bHasControl;		//该Object是否包含控件
	qvector<mid_ControlInfo> mVec_ControlInfo;

	qvector<mid_VTable> mVec_MethodLink;
};


struct mid_ObjectMess
{
	qstring m_ObjectName;	//对象的名称
	qstring m_ObjectTypeName;
	uint32 m_ObjectType;
	qlist<qstring> mVec_MethodName;

	uint16 m_ClassSize;		//申请的类空间大小
	qstring m_ObjectIID;	//调用函数时候的GUID

	uint16 m_iPCodeCount;	//对象的opcode函数个数

	bool bHasClasses;		//该Object是否包含类

	bool bHasOptionalObj;
	mid_OptionalObjectInfo m_OptionObjInfo;
};


struct mid_WindowsApi
{
	qstring LibraryName;
	qstring FunctionName;
};

struct mid_ComApi
{
	qstring uuid_coclass;
	uint32 GlobalVarAddr;
};

struct mid_ApiInfo
{
	qvector<mid_WindowsApi> mVec_WinApi;
	qvector<mid_ComApi> mVec_ComApi;
};

struct mid_ProjectInfo
{
	qstring ProjectName;			//工程名称
	qstring ProjectDescription;		//工程描述

	mid_ProjectInfo2 m_ProjectInfo2;
	uint32 lpCodeStart;				//VB用户代码起始地址
	uint32 lpCodeEnd;				//VB用户代码结束地址
	qvector<uint32> mVec_FuncTable;	//所有的函数列表
	mid_ApiInfo m_ApiInfo;			//工程引入的Api
};

struct mid_GUITable
{
	uint32 ToDo;
};

struct mid_ExternalControl
{
	qstring LibName;
	qstring IDEName;
	qstring ControlName;
};

struct mid_COMRegData
{
	qstring RegComName;
	qstring RegComUUID;
};

struct mid_VBHeader
{
	uint32 m_lpSubMain;									//Main函数入口

	mid_ProjectInfo m_ProjectInfo;						//工程对象信息
	qvector<mid_ObjectMess> mVec_ObjectTable;			//对象信息
	qvector<mid_GUITable> mVec_GuiTable;				//窗体对象信息
	qvector<mid_ExternalControl> mVec_ExternalControl;	//额外COM控件信息
	qvector<mid_COMRegData> mVec_ComRegData;			//注册COM组件信息
};