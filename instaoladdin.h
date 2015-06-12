#pragma once

#include "stdafx.h"
#include "oltlbs\MSADDNDR.tlh"
#include "oltlbs\\MSO.tlh"
#include "oltlbs\\FM20.tlh"
#include "oltlbs\\MSOUTL.tlh"
#include "instaoladdin_h.h"

using namespace AddinDesign;
using namespace Office;
using namespace Forms;
using namespace Outlook;

class ATL_NO_VTABLE CInstaOLAddin :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CInstaOLAddin, &CLSID_InstaOLAddin>,
	public IDispatchImpl<IInstaOLAddinConnect, &IID_IInstaOLAddinConnect, &LIBID_InstaOLAddinLib, /* wMajor = */ 1, /* wMajor = */0>,
	public IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &__uuidof(__AddInDesignerObjects), /* wMajor = */ 1>,
	public IDispatchImpl<_FormRegionStartup, &__uuidof(_FormRegionStartup), &__uuidof(__Outlook), /* wMajor = */ 9, /* wMinor = */ 4>,
	public IDispatchImpl<IRibbonExtensibility, &__uuidof(IRibbonExtensibility), &__uuidof(__Office), /* wMajor = */ 2, /* wMinor = */ 5>,
	public IDispatchImpl<IRibbonCallback, &__uuidof(IRibbonCallback)>
{
public:
	CInstaOLAddin();
	virtual ~CInstaOLAddin();

	DECLARE_REGISTRY_RESOURCEID(IDR_INSTAOLADDIN)

	BEGIN_COM_MAP(CInstaOLAddin)
		COM_INTERFACE_ENTRY2(IDispatch, IRibbonExtensibility)
		COM_INTERFACE_ENTRY(IInstaOLAddinConnect)
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
		COM_INTERFACE_ENTRY(IRibbonExtensibility)
		COM_INTERFACE_ENTRY(_FormRegionStartup)
		COM_INTERFACE_ENTRY(IRibbonCallback)
	END_COM_MAP()

	STDMETHOD(Invoke)(DISPID dispidMember, const IID &riid, LCID lcid, WORD wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexceptinfo, UINT *puArgErr);

	//_IDTExtensibility2
	STDMETHOD(OnConnection)(IDispatch * Application, enum ext_ConnectMode ConnectMode, IDispatch * AddInInst, SAFEARRAY * * custom);
	STDMETHOD(OnDisconnection)(enum ext_DisconnectMode RemoveMode, SAFEARRAY * * custom);
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom);
	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom);
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom);

	//_FormRegionStartup
	STDMETHOD(GetFormRegionStorage)(BSTR FormRegionName, LPDISPATCH Item, long LCID, OlFormRegionMode FormRegionMode, OlFormRegionSize FormRegionSize, VARIANT * Storage);
	STDMETHOD(BeforeFormRegionShow)(_FormRegion * FormRegion);
	STDMETHOD(GetFormRegionManifest)(BSTR FormRegionName, long LCID, VARIANT * Manifest);
	STDMETHOD(GetFormRegionIcon)(BSTR FormRegionName, long LCID, OlFormRegionIcon Icon, VARIANT * Result);

	//IRibbonExtensibility
	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml);

	//IRibbonCallback
	STDMETHOD(ButtonClicked)(IDispatch* ribbonControl);
};

