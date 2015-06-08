#pragma once

#include "stdafx.h"
#include "oltlbs\MSADDNDR.tlh"
#include "instaoladdin_h.h"

using namespace AddinDesign;

class ATL_NO_VTABLE CInstaOLAddin :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CInstaOLAddin, &CLSID_InstaOLAddin>,
    public IDispatchImpl<IInstaOLAddinConnect, &IID_IInstaOLAddinConnect, &LIBID_InstaOLAddinLib, 1, 0>,
    public IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &__uuidof(__AddInDesignerObjects), 1, 0>
{
public:
    CInstaOLAddin();
    virtual ~CInstaOLAddin();

    DECLARE_REGISTRY_RESOURCEID(IDR_INSTAOLADDIN)

    BEGIN_COM_MAP(CInstaOLAddin)
        COM_INTERFACE_ENTRY(IInstaOLAddinConnect)
        COM_INTERFACE_ENTRY(_IDTExtensibility2)
    END_COM_MAP()
};
