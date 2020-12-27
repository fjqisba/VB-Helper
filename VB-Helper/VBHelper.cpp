#include <hexrays.hpp>
#include <auto.hpp>
#include <diskio.hpp>
#include <entry.hpp>
#include "ImportsFix.h"
#include "ComManager.h"
#include "public.h"
#include "VBDecompiler.h"

// Hex-Rays API pointer
hexdsp_t* hexdsp = nullptr;

//--------------------------------------------------------------------------
struct plugin_ctx_t : public plugmod_t
{
	bool inited = true;

	~plugin_ctx_t()
	{
		if (inited)
			term_hexrays_plugin();
	}
	virtual bool idaapi run(size_t) override;
};

//--------------------------------------------------------------------------
static plugmod_t* idaapi init()
{
	qstring qstr_ProcName = inf.procname;
	if (qstr_ProcName != "metapc" || inf.filetype != f_PE)
	{
		return NULL;
	}
	if (!init_hexrays_plugin())
		return nullptr; // no decompiler
	const char* hxver = get_hexrays_version();
	msg("[VB-Helper] plugin 1.0 loaded,Author: fjqisba\n");
	return new plugin_ctx_t;
}

//--------------------------------------------------------------------------
bool idaapi plugin_ctx_t::run(size_t)
{
	//确保IDA已完成初始化分析
	if (!auto_is_ok() && ask_yn(0, "The autoanalysis has not finished yet.\nDo you want to continue?") < 1)
	{
		return true;
	}

	//修复VB导入表
	FixVBImports();

	InitComManager();

	//初始化VB反编译引擎
	VBDecompilerEngine vbEngine;
	ea_t PEentry = get_entry(get_entry_ordinal(0));
	if (!vbEngine.DoDecompile(PEentry))
	{
		msg("[VBHelper]:Init VB Decompiler Engine failed\n");
		return false;
	}
	
	//识别出扫描得到的函数
	vbEngine.MakeFunction();

	//设置主函数的名称
	vbEngine.SetSubMain();

	show_wait_box("Create Object VTable...");
	//建立类的虚表
	vbEngine.CreateVTable();
	hide_wait_box();

	//加载VB.hpp
	show_wait_box("Load COM Structure...");
	if (!vbEngine.Load_VBHpp())
	{
		hide_wait_box();
		info("Load VB.hpp error");
		return false;
	}
	hide_wait_box();


	//解析程序自身的Object,将类的GUID添加到map中
	vbEngine.AddClassGuid();

	//设置控件函数的事件名称
	vbEngine.SetEventFuncName();

	show_wait_box("Flash VB Imports...");
	auto_wait();
	hide_wait_box();

	
	vbEngine.FlashComInterface();
	return true;
}

//--------------------------------------------------------------------------
static char comment[] = "It's a tool used to help Analysis VB Application";

//--------------------------------------------------------------------------
//
//      PLUGIN DESCRIPTION BLOCK
//
//--------------------------------------------------------------------------
plugin_t PLUGIN =
{
  IDP_INTERFACE_VERSION,
  PLUGIN_MULTI,         // The plugin can work with multiple idbs in parallel
  init,                 // initialize
  nullptr,
  nullptr,
  comment,              // long comment about the plugin
  nullptr,              // multiline help about the plugin
  "VBHelper",           // the preferred short name of the plugin
  nullptr,              // the preferred hotkey to run the plugin
};
