// instaoladdin.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
#include "instaoladdin.h"
#include "utils.h"

CInstaOLAddin::CInstaOLAddin()
{
}


CInstaOLAddin::~CInstaOLAddin()
{
}

//
//_IDTExtensibility2
//
STDMETHODIMP CInstaOLAddin::OnConnection(IDispatch * Application, enum ext_ConnectMode ConnectMode, IDispatch * AddInInst,
										SAFEARRAY * * custom)
{
	MessageBox(NULL, L"CInstaOLAddin", L"OnConnection", MB_OK);
	return S_OK;
}

STDMETHODIMP CInstaOLAddin::OnDisconnection(enum ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
{
	return S_OK;
}

STDMETHODIMP CInstaOLAddin::OnAddInsUpdate(SAFEARRAY * * custom)
{
	return S_OK;
}

STDMETHODIMP CInstaOLAddin::OnStartupComplete(SAFEARRAY * * custom)
{
	return S_OK;
}

STDMETHODIMP CInstaOLAddin::OnBeginShutdown(SAFEARRAY * * custom)
{
	return S_OK;
}

//
//_FormRegionStartup
//
STDMETHODIMP CInstaOLAddin::GetFormRegionStorage(BSTR FormRegionName, LPDISPATCH Item, long LCID, OlFormRegionMode FormRegionMode,
												OlFormRegionSize FormRegionSize, VARIANT * Storage)
{
	return S_OK;
}

STDMETHODIMP CInstaOLAddin::BeforeFormRegionShow(_FormRegion * FormRegion)
{
	return S_OK;
}

STDMETHODIMP CInstaOLAddin::GetFormRegionManifest(BSTR FormRegionName, long LCID, VARIANT * Manifest)
{
	return S_OK;
}

STDMETHODIMP CInstaOLAddin::GetFormRegionIcon(BSTR FormRegionName, long LCID, OlFormRegionIcon Icon, VARIANT * Result)
{
	return S_OK;
}

//
//IRibbonExtensibility
//
STDMETHODIMP CInstaOLAddin::GetCustomUI(BSTR RibbonID, BSTR * RibbonXml)
{
	if (!RibbonXml)
	{
		return E_POINTER;
	}

	*RibbonXml = GetXMLResource(IDR_RIBBON_XML);

	return S_OK;
}

//
//IRibbonCallback
//
STDMETHODIMP CInstaOLAddin::ButtonClicked(IDispatch* ribbonControl)
{
	return S_OK;
}

//
//IDispatch::Invoke
//
STDMETHODIMP CInstaOLAddin::Invoke(DISPID dispidMember, const IID &riid, LCID lcid, WORD wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexceptinfo, UINT *puArgErr)
{
	return DISP_E_MEMBERNOTFOUND;
}
