#pragma once

HRESULT HrGetResource(int nId, LPCTSTR lpType, LPVOID* ppvResourceData, DWORD* pdwSizeInBytes);
BSTR GetXMLResource(int nId);