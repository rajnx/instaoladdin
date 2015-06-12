// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include "instaoladdin.h"
#include "instaoladdin_i.c"

class CInstaOLAddinLib : public CAtlDllModuleT<CInstaOLAddinLib>
{
public:
	DECLARE_LIBID(LIBID_InstaOLAddinLib)
};

CInstaOLAddinLib  _AtlDllModule;

// DLL Entry Point
BOOL APIENTRY DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{

	return _AtlDllModule.DllMain(dwReason, lpReserved);
}

STDAPI DllCanUnloadNow(void)
{
	return _AtlDllModule.DllCanUnloadNow();
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return _AtlDllModule.DllGetClassObject(rclsid, riid, ppv);
}

STDAPI DllRegisterServer(void)
{
	return _AtlDllModule.DllRegisterServer();
}

STDAPI DllUnregisterServer(void)
{
	return _AtlDllModule.DllUnregisterServer();
}

OBJECT_ENTRY_AUTO(__uuidof(InstaOLAddin), CInstaOLAddin)