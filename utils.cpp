#include "stdafx.h"
#include "utils.h"

HRESULT HrGetResource(int nId, LPCTSTR lpType, LPVOID* ppvResourceData, DWORD* pdwSizeInBytes)
{
	if (!ppvResourceData || !pdwSizeInBytes)
	{
		return E_POINTER;
	}

	HMODULE hModule = _AtlBaseModule.GetModuleInstance();
	if (!hModule)
	{
		return E_UNEXPECTED;
	}

	HRSRC hRsrc = FindResource(hModule, MAKEINTRESOURCE(nId), lpType);
	if (!hRsrc)
	{
		return HRESULT_FROM_WIN32(GetLastError());
	}

	HGLOBAL hGlobal = LoadResource(hModule, hRsrc);
	if (!hGlobal)
	{
		return HRESULT_FROM_WIN32(GetLastError());
	}

	*pdwSizeInBytes = SizeofResource(hModule, hRsrc);
	*ppvResourceData = LockResource(hGlobal);
	
	return S_OK;
}

BSTR GetXMLResource(int nId)
{
	LPVOID pResourceData = NULL;
	DWORD dwSizeInBytes = 0;

	HRESULT hr = HrGetResource(nId, TEXT("XML"), &pResourceData, &dwSizeInBytes);
	if (FAILED(hr))
	{
		return NULL;
	}

	// Assumes that the data is not stored in Unicode.
	CComBSTR cbstr(dwSizeInBytes, reinterpret_cast<LPCSTR>(pResourceData));
	return cbstr.Detach();
}