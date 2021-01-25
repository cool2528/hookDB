// hookDB.cpp : 定义 DLL 应用程序的导出函数。
//

#include "stdafx.h"
#include <Ole2.h>
#include <string>
#include <comutil.h>
#include <atlconv.h>
#include "plog/Log.h"
#include "plog/Initializers/RollingFileInitializer.h"
#include "mhook.h"
#include <mutex>
#include "Conver.hpp"
#include <atlstr.h>
#include <ATLComTime.h>
#include <vector>
#include <fstream>
#include "fmt/format.h"
#include "nlohmann/json.hpp"
std::mutex g_mutex;
using json = nlohmann::json;
#import "E:\\VS2015\\Test\\hookDB\\msado15.dll"  no_namespace rename("EOF", "adoEOF") rename("BOF", "FirstOfFile")
typedef HRESULT(__stdcall *PROC_CREATE_FUN)(REFCLSID  rclsid,
	LPUNKNOWN pUnkOuter,
	DWORD     dwClsContext,
	REFIID    riid,
	LPVOID    *ppv);
HMODULE oleHandle = GetModuleHandle("ole32.dll");

PROC_CREATE_FUN CoCreateInstance_OLD = NULL;

typedef HRESULT(__stdcall *FUNC_OPEN)(IUnknown* This, BSTR ConnectionString, BSTR UserID, BSTR Password, long Options);
typedef HRESULT(__stdcall *FUNC_EXCUTE)(IUnknown* This, BSTR CommandText, VARIANT * RecordsAffected, long Options, struct _Recordset * * ppiRset);
std::string varToStr(VARIANT var);
unsigned long writeFile(const std::string& stfFilePath, std::vector<char>& data)
{
	unsigned long dwLength = 0;
	if (stfFilePath.empty())
	{
		return dwLength;
	}
	std::fstream out(stfFilePath, std::ios_base::out | std::ios_base::binary | std::ios_base::trunc);
	if (out.is_open()) {
		out.write(data.data(), data.size());
		dwLength = 1;
		out.close();
	}
	return dwLength;
}
void parseRecordset(_Recordset * pRecord)
{

	try
	{
		std::vector<std::string> column_name;
		//LOG_ERROR_(1) << pRecord->Fields->GetCount();
		for (int i = 0; i < pRecord->Fields->GetCount(); ++i)
		{
			auto item = pRecord->Fields->GetItem(_variant_t((long)i));
			auto nameptr = _com_util::ConvertBSTRToString(item->Name);
			std::string name(nameptr);
			delete[] nameptr;
			item.Release();
			//LOG_ERROR_(1) << name;
			column_name.push_back(name);

		}
		json result, arraylist;
		while (!(pRecord->adoEOF))
		{
			json item;

			for (auto v : column_name)
			{
				auto it = pRecord->GetCollect(v.c_str());
				std::string name = v;
				name = Coding_Conver::Conver::AnsiToUtf8(name.c_str());
				if (it.vt == VT_BSTR && it.bstrVal)
				{
					std::string value;
					auto tmpval = _com_util::ConvertBSTRToString(it.bstrVal);
					value.assign(tmpval);
					delete[] tmpval;
					value = Coding_Conver::Conver::AnsiToUtf8(value.c_str());
					//cout << name << "	"<< value << "	";
					item[name] = value.c_str();
				}
				else if (it.vt == VT_I4)
				{
					//cout << name<< "	" << it.iVal << "	";
					item[name] = it.iVal;
				}
				else if (it.vt == VT_DATE)
				{
					COleDateTime old(it.date);
					CString str = old.Format(_T("%Y-%m-%d %H:%M:%S"));
					std::string value = str.GetBuffer();
					str.ReleaseBuffer();
					item[name] = value.c_str();

				}
				else if (it.vt == VT_I2)
				{
					item[name] = it.iVal;
				}
				else if (it.vt == VT_R4)
				{
					item[name] = it.fltVal;
					//item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_R4").c_str();
				}
				else if (it.vt == VT_R8)
				{
					item[name] = it.dblVal;
					//item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_R8").c_str();
				}
				else if (it.vt == VT_CY)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8(std::to_string((unsigned long long)it.cyVal.int64).c_str()).c_str();
					//item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_CY").c_str();
				}
				else if (it.vt == VT_DISPATCH)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("IDispatch").c_str();
				}
				else if (it.vt == VT_ERROR)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8(std::to_string((LONG)it.scode).c_str()).c_str();
				}
				else if (it.vt == VT_BOOL)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8(std::to_string((short)it.boolVal).c_str()).c_str();
				}
				else if (it.vt == VT_VARIANT)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_VARIANT").c_str();
				}
				else if (it.vt == VT_UNKNOWN)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_UNKNOWN").c_str();
				}
				else if (it.vt == VT_DECIMAL)
				{
					
					//LOG_ERROR_(1) << it.decVal.Lo64;
					item[name] = it.decVal.Lo64;
					//item[name] = Coding_Conver::Conver::AnsiToUtf8(std::to_string((unsigned long long)it.pdecVal->Lo64).c_str()).c_str();
				}
				else if (it.vt == VT_I1)
				{
					std::string value;
					if (it.pcVal)
					{
						LOG_ERROR_(1) << (DWORD)it.pcVal;
						value = it.pcVal;
					}
					value = Coding_Conver::Conver::AnsiToUtf8(value.c_str());
					item[name] = value.c_str();
				}
				else if (it.vt == VT_UI1)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_UI1").c_str();
				}
				else if (it.vt == VT_UI2)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_UI2").c_str();
				}
				else if (it.vt == VT_UI4)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_UI4").c_str();
				}
				else if (it.vt == VT_I8)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_I8").c_str();
				}
				else if (it.vt == VT_UI8)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_UI8").c_str();
				}
				else if (it.vt == VT_INT)
				{
					item[name] = it.intVal;
				}
				else if (it.vt == VT_UINT)
				{
					item[name] = it.uintVal;
				}
				else if (it.vt == VT_VOID)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_VOID").c_str();
				}
				else if (it.vt == VT_HRESULT)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_HRESULT").c_str();
				}
				else if (it.vt == VT_PTR)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_PTR").c_str();
				}
				else if (it.vt == VT_SAFEARRAY)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_SAFEARRAY").c_str();
				}
				else if (it.vt == VT_CARRAY)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_CARRAY").c_str();
				}
				else if (it.vt == VT_USERDEFINED)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_USERDEFINED").c_str();
				}
				else if (it.vt == VT_LPSTR)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_LPSTR").c_str();
				}
				else if (it.vt == VT_LPWSTR)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_LPWSTR").c_str();
				}
				else if (it.vt == VT_RECORD)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_RECORD").c_str();
				}
				else if (it.vt == VT_INT_PTR)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_INT_PTR").c_str();
				}
				else if (it.vt == VT_UINT_PTR)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_UINT_PTR").c_str();
				}
				else if (it.vt == VT_FILETIME)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_FILETIME").c_str();
				}
				else if (it.vt == VT_STREAM)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_STREAM").c_str();
				}
				else if (it.vt == VT_ARRAY)
				{
					item[name] = Coding_Conver::Conver::AnsiToUtf8("VT_ARRAY").c_str();
				}
				else
				{
					std::string val = varToStr(it);
					item[name] = Coding_Conver::Conver::AnsiToUtf8(val.c_str()).c_str();
				}
				it.Clear();
			}
			arraylist.push_back(item);
			pRecord->MoveNext();
		}
		result["list"] = arraylist;
		if (!arraylist.empty())
		{
			std::string bResult = result.dump();
			std::string dataPath = "C:\\sqldata\\";
			::CreateDirectory(dataPath.c_str(), NULL);
			dataPath += "Connection_Execute_" + std::to_string(time(NULL));
			dataPath += ".json";
			std::vector<char> data;
			data.resize(bResult.length());
			data.assign(bResult.begin(), bResult.end());
			writeFile(dataPath, data);
		}
		if (!(pRecord->adoEOF))
		{
			pRecord->MoveLast();
		}
		if (!(pRecord->FirstOfFile))
		{
			pRecord->MoveFirst();
		}
	}
	catch (_com_error& e)
	{
		auto szError = e.ErrorMessage();
		if (szError)
		{
			LOG_ERROR_(1) << szError;
			delete[] szError;
		}
		return;
	}

}
typedef HRESULT(__stdcall* FUNC_PUTSTR)(IUnknown* This, BSTR pbstr);
class hookConnectIon 
{
public:
	static bool hookMethods(DWORD dwAddress) {
		bool bResult = false;
		if (!dwAddress) {

			LOG_INFO << "虚表地址是空的" << dwAddress;
			return bResult;
		}
		m_vttableAddress = dwAddress;
		DWORD dwlen = 21;
		DWORD m_dwOldProtect = NULL;
		auto m_Bret = ::VirtualProtectEx(::GetCurrentProcess(), (LPVOID)dwAddress, sizeof(DWORD) * dwlen, PAGE_EXECUTE_READWRITE, &m_dwOldProtect);
		if (!m_Bret)
		{
			return bResult;
		}
		if (!m_putStr)
		{
			m_putStr = (FUNC_PUTSTR)((LPDWORD)dwAddress)[9];
			((LPDWORD)dwAddress)[9] = (DWORD)&hookConnectIon::put_ConnectionString;
		}
		if (!m_excute_old)
		{
			m_excute_old = (FUNC_EXCUTE)((LPDWORD)dwAddress)[16];
			((LPDWORD)dwAddress)[16] = (DWORD)&hookConnectIon::raw_Execute;
		}
		if (!m_open_old)
		{
			m_open_old = (FUNC_OPEN)((LPDWORD)dwAddress)[20];
			((LPDWORD)dwAddress)[20] = (DWORD)&hookConnectIon::raw_Open;
		}
		m_Bret = ::VirtualProtectEx(::GetCurrentProcess(), (LPVOID)dwAddress, sizeof(DWORD) * dwlen, m_dwOldProtect, NULL);	//改回原来的属性
		bResult = true;
		return bResult;
	}
	static HRESULT __stdcall put_ConnectionString(IUnknown* This, BSTR ConnectionString) {
		std::string buf;
		if (ConnectionString)
		{
			auto strcon = _com_util::ConvertBSTRToString(ConnectionString);
			if (strcon)
			{
				buf = "SQL语句:";
				buf+= strcon;
				delete[] strcon;
			}
		}
		//LOG_INFO<< "Connection put_ConnectionString :" << buf;
		if (!buf.empty())
		{
			LOG_INFO << buf.c_str();
		}
		//LOG_INFO << "Connection put_ConnectionString ";
		return m_putStr(This,ConnectionString);
	}
	static HRESULT __stdcall raw_Open(IUnknown* This,BSTR ConnectionString,BSTR UserID,BSTR Password,long Options) {
		std::string buf;
		if (ConnectionString)
		{
			auto strcon = _com_util::ConvertBSTRToString(ConnectionString);
			if (strcon)
			{
				buf = "登录语句:";
				buf += strcon;
				delete[] strcon;
			}
		}
		if (UserID)
		{
			auto strcon = _com_util::ConvertBSTRToString(UserID);
			if (strcon)
			{
				buf += "\t";
				buf += "用户名:";
				buf += strcon;
				delete[] strcon;
			}
		}
		if (Password)
		{
			auto strcon = _com_util::ConvertBSTRToString(Password);
			if (strcon)
			{
				buf += "\t";
				buf += "密码:";
				buf += strcon;
				delete[] strcon;
			}
		}
		if (!buf.empty())
		{
			LOG_INFO << buf.c_str();
		}
		//LOG_INFO << "Connection Open";
		return m_open_old(This, ConnectionString, UserID, Password, Options);
	}
	static HRESULT __stdcall raw_Execute(IUnknown* This, BSTR CommandText, VARIANT * RecordsAffected, long Options, struct _Recordset ** ppiRset) {
		std::string buf;
		if (CommandText) {
			auto strcon = _com_util::ConvertBSTRToString(CommandText);
			if (strcon)
			{
				buf = "SQL 语句:";
				buf += strcon;
				delete[] strcon;
			}
		}
		//LOG_INFO << buf.c_str() << "\nConnection Execute RecordsAffected vt type" << RecordsAffected->vt;
		if (!buf.empty())
		{
			LOG_INFO << buf.c_str();
		}
		HRESULT hr = m_excute_old(This, CommandText, RecordsAffected, Options, ppiRset);
		try
		{
			if (SUCCEEDED(hr))
			{
				if (*ppiRset)
				{
					parseRecordset(*ppiRset);
				}
			}
		}
		catch (_com_error& e)
		{
			auto szError = e.ErrorMessage();
			if (szError)
			{
				LOG_ERROR_(1) << szError;
				delete[] szError;
			}
			(*ppiRset)->MoveFirst();
			return hr;
		}
		return hr;
	}
	static bool unhookMethods() {
		bool bResult = false;
		DWORD m_dwOldProtect = NULL;
		DWORD dwlen = 21;
		auto m_Bret = ::VirtualProtectEx(::GetCurrentProcess(), (LPVOID)m_vttableAddress, sizeof(DWORD) * dwlen, PAGE_EXECUTE_READWRITE, &m_dwOldProtect);
		if (!m_Bret)
		{
			return bResult;
		}
		((LPDWORD)m_vttableAddress)[9] = (DWORD)m_putStr;
		((LPDWORD)m_vttableAddress)[16] = (DWORD)m_excute_old;
		((LPDWORD)m_vttableAddress)[20] = (DWORD)m_open_old;
		m_Bret = ::VirtualProtectEx(::GetCurrentProcess(), (LPVOID)m_vttableAddress, sizeof(DWORD) * dwlen, m_dwOldProtect, NULL);	//改回原来的属性
		bResult = true;
		return bResult;
	}
	static DWORD m_vttableAddress;
	static FUNC_OPEN m_open_old;
	static FUNC_PUTSTR m_putStr;
	static FUNC_EXCUTE m_excute_old;

};
DWORD hookConnectIon::m_vttableAddress = NULL;
FUNC_OPEN hookConnectIon::m_open_old = nullptr;
FUNC_EXCUTE hookConnectIon::m_excute_old = nullptr;
FUNC_PUTSTR hookConnectIon::m_putStr = nullptr;
std::string printArgType(DataTypeEnum ntype) {
	std::string strType;
	switch (ntype)
	{
	case adArray:
	{
		strType = "array";
		break;
	}
	case adBigInt:
	{
		strType = "Short";
		break;
	}
	case adBinary:
	{
		strType = "byte";
		break;
	}
	case adBoolean:
	{
		strType = "Boolean";
		break;
	}
	case adBSTR:
	{
		strType = "wchar_t*";
		break;
	}
	case adChapter:
	{
		strType = "numerical";
		break;
	}
	case adChar:
	{
		strType = "char*";
		break;
	}
	case adCurrency:
	{
		strType = "numerical";
		break;
	}
	case adDate:
	{
		strType = "DBTYPE_DBDATE";
		break;
	}
	case adDBDate:
	{
		strType = "date";
		break;
	}
	case adDBTime:
	{
		strType = "DBTYPE_DBTIME";
		break;
	}
	case adDBTimeStamp:
	{
		strType = "DBTYPE_DBTIMESTAMP";
		break;
	}
	case adDecimal:
	{
		strType = "DBTYPE_DECIMAL";
		break;
	}
	case adDouble:
	{
		strType = "double";
		break;
	}
	case adEmpty:
	{
		strType = "DBTYPE_EMPTY";
		break;
	}
	case adError:
	{
		strType = "DBTYPE_ERROR";
		break;
	}
	case adFileTime:
	{
		strType = "DBTYPE_FILETIME";
		break;
	}
	case adGUID:
	{
		strType = "DBTYPE_GUID";
		break;
	}
	case adIDispatch:
	{
		strType = "DBTYPE_IDISPATCH";
		break;
	}
	case adInteger:
	{
		strType = "DBTYPE_I4";
		break;
	}
	case adIUnknown:
	{
		strType = "DBTYPE_IUNKNOWN";
		break;
	}
	case adLongVarBinary:
	{
		strType = "long byte";
		break;
	}
	case adLongVarChar:
	{
		strType = "long char*";
		break;
	}
	case adLongVarWChar:
	{
		strType = "long wchar_t*";
		break;
	}
	case adNumeric:
	{
		strType = "DBTYPE_NUMERIC";
		break;
	}
	case adPropVariant:
	{
		strType = "DBTYPE_PROP_VARIANT";
		break;
	}
	case adSingle:
	{
		strType = "DBTYPE_R4";
		break;
	}
	case adSmallInt:
	{
		strType = "DBTYPE_I2";
		break;
	}
	case adTinyInt:
	{
		strType = "DBTYPE_I1";
		break;
	}
	case adUnsignedBigInt:
	{
		strType = "DBTYPE_UI8";
		break;
	}
	case adUnsignedInt:
	{
		strType = "DBTYPE_UI4";
		break;
	}
	case adUnsignedSmallInt:
	{
		strType = "DBTYPE_UI2";
		break;
	}
	case adUnsignedTinyInt:
	{
		strType = "DBTYPE_UI1";
		break;
	}
	case adUserDefined:
	{
		strType = "DBTYPE_UDT";
		break;
	}
	case adVarBinary:
	{
		strType = "binary";
		break;
	}
	case adVarChar:
	{
		strType = "string";
		break;
	}
	case adVariant:
	{
		strType = "DBTYPE_VARIANT";
		break;
	}
	case adVarNumeric:
	{
		strType = "numeric";
		break;
	}
	case adVarWChar:
	{
		strType = "Unicode string";
		break;
	}
	case adWChar:
	{
		strType = "DBTYPE_WSTR";
		break;
	}
	default:
		strType = std::to_string((UINT)ntype);
		break;
	}
	return strType;
}
std::string varToStr(VARIANT var) {
	std::string result;
	VARIANT desc;
	if (SUCCEEDED(VariantChangeType(&desc, &var, VARIANT_NOUSEROVERRIDE | VARIANT_LOCALBOOL, VT_BSTR)))
	{
		if (desc.vt == VT_BSTR && desc.bstrVal)
		{
			std::string value;
			auto tmpval = _com_util::ConvertBSTRToString(desc.bstrVal);
			value.assign(tmpval);
			delete[] tmpval;
			//value = Coding_Conver::Conver::AnsiToUtf8(value.c_str());
			
			result = value;
		}
	}
	return result;
}
typedef HRESULT(__stdcall * FUNRECO_OPEN)(IUnknown* This, VARIANT Source, VARIANT ActiveConnection, enum CursorTypeEnum CursorType, enum LockTypeEnum LockType, long Options);
//raw_NextRecordset
typedef HRESULT(__stdcall* FUNC_NETRECORDSET)(IUnknown* This, VARIANT * RecordsAffected, struct _Recordset * * ppiRs);
class hookRecordset {
public:
	static bool hookMethods(DWORD dwAddress) {
		bool bResult = false;
		if (!dwAddress) {

			LOG_INFO << "虚表地址是空的" << dwAddress;
			return bResult;
		}
		m_vttableAddress = dwAddress;
		DWORD dwlen = 62;
		DWORD m_dwOldProtect = NULL;
		auto m_Bret = ::VirtualProtectEx(::GetCurrentProcess(), (LPVOID)dwAddress, sizeof(DWORD) * dwlen, PAGE_EXECUTE_READWRITE, &m_dwOldProtect);
		if (!m_Bret)
		{
			return bResult;
		}
		if (!m_open_old)
		{
			m_open_old = (FUNRECO_OPEN)((LPDWORD)dwAddress)[40];
			((LPDWORD)dwAddress)[40] = (DWORD)&hookRecordset::raw_Open;
		}
		if (!m_NetRecordset)
		{
			m_NetRecordset = (FUNC_NETRECORDSET)((LPDWORD)dwAddress)[61];
			((LPDWORD)dwAddress)[61] = (DWORD)&hookRecordset::raw_NextRecordset;
		}
		m_Bret = ::VirtualProtectEx(::GetCurrentProcess(), (LPVOID)dwAddress, sizeof(DWORD) * dwlen, m_dwOldProtect, NULL);	//改回原来的属性
		bResult = true;
		return bResult;
	}
	static HRESULT __stdcall raw_Open(IUnknown* This, VARIANT Source, VARIANT ActiveConnection, enum CursorTypeEnum CursorType, enum LockTypeEnum LockType, long Options) {
		std::string buf;
		if (Source.vt != VT_EMPTY && Source.vt != VT_NULL)
		{
			if (Source.vt == VT_BSTR && Source.bstrVal) {
				auto strbuf = _com_util::ConvertBSTRToString(Source.bstrVal);
				if (strbuf)
				{
					buf = "SQL 语句:";
					buf += strbuf;
					delete[] strbuf;
				}
			}
			
		}
		if (ActiveConnection.vt != VT_EMPTY && ActiveConnection.vt != VT_NULL)
		{
			if (ActiveConnection.vt == VT_BSTR && ActiveConnection.bstrVal) {
				auto strbuf = _com_util::ConvertBSTRToString(ActiveConnection.bstrVal);
				LOG_INFO<< strbuf;
				if (strbuf)
				{
					buf += strbuf;
					delete[] strbuf;
				}
			}
		}
		if (!buf.empty())
		{
			LOG_INFO << buf.c_str();
		}
		//LOG_INFO << "open";
		return m_open_old(This, Source, ActiveConnection, CursorType, LockType, Options);
	}
	static HRESULT __stdcall raw_NextRecordset(IUnknown* This, VARIANT * RecordsAffected, struct _Recordset * * ppiRs)
	{
		HRESULT hr = m_NetRecordset(This, RecordsAffected, ppiRs);
		LOG_ERROR_(1) << "raw_NextRecordset RecordsAffected: " << RecordsAffected->lVal;
		try
		{
			if (SUCCEEDED(hr))
			{
				if (*ppiRs)
				{
					parseRecordset(*ppiRs);

				}
			}
		}
		catch (_com_error& e)
		{
			auto szError = e.ErrorMessage();
			if (szError)
			{
				LOG_ERROR_(1) << szError;
				delete[] szError;
			}
			(*ppiRs)->MoveFirst();
			return hr;
		}
		return hr;
	}
	static bool unhookMethods() {
		bool bResult = false;
		DWORD m_dwOldProtect = NULL;
		DWORD dwlen =62;
		auto m_Bret = ::VirtualProtectEx(::GetCurrentProcess(), (LPVOID)m_vttableAddress, sizeof(DWORD) * dwlen, PAGE_EXECUTE_READWRITE, &m_dwOldProtect);
		if (!m_Bret)
		{
			return bResult;
		}
		((LPDWORD)m_vttableAddress)[40] = (DWORD)m_open_old;
		((LPDWORD)m_vttableAddress)[61] = (DWORD)m_NetRecordset;
		m_Bret = ::VirtualProtectEx(::GetCurrentProcess(), (LPVOID)m_vttableAddress, sizeof(DWORD) * dwlen, m_dwOldProtect, NULL);	//改回原来的属性
		bResult = true;
		return bResult;
	}
	static DWORD m_vttableAddress;
	static FUNRECO_OPEN m_open_old;
	static FUNC_NETRECORDSET m_NetRecordset;
};
FUNRECO_OPEN hookRecordset::m_open_old = nullptr;
DWORD hookRecordset::m_vttableAddress = NULL;
FUNC_NETRECORDSET hookRecordset::m_NetRecordset = nullptr;
typedef HRESULT(__stdcall *FUNCCOM_OPEN)(IUnknown* This, VARIANT * RecordsAffected, VARIANT * Parameters, long Options, struct _Recordset * * ppiRs);
typedef HRESULT(__stdcall *FUNCCOM_PUT)(IUnknown* This, BSTR pbstr);

class hookCommand {
public:
	static bool hookMethods(DWORD dwAddress) {
		bool bResult = false;
		if (!dwAddress) {

			LOG_INFO << "虚表地址是空的" << dwAddress;
			return bResult;
		}
		m_vttableAddress = dwAddress;
		DWORD dwlen = 20;
		DWORD m_dwOldProtect = NULL;
		auto m_Bret = ::VirtualProtectEx(::GetCurrentProcess(), (LPVOID)dwAddress, sizeof(DWORD) * dwlen, PAGE_EXECUTE_READWRITE, &m_dwOldProtect);
		if (!m_Bret)
		{
			return bResult;
		}
		if (!m_open_old)
		{
			m_open_old = (FUNCCOM_OPEN)((LPDWORD)dwAddress)[17];
			((LPDWORD)dwAddress)[17] = (DWORD)&hookCommand::raw_Execute;
		}

		if (!m_put)
		{
			m_put = (FUNCCOM_PUT)((LPDWORD)dwAddress)[12];
			((LPDWORD)dwAddress)[12] = (DWORD)&hookCommand::put_CommandText;
		}
		m_Bret = ::VirtualProtectEx(::GetCurrentProcess(), (LPVOID)dwAddress, sizeof(DWORD) * dwlen, m_dwOldProtect, NULL);	//改回原来的属性
		bResult = true;
		return bResult;
	}
	static HRESULT __stdcall put_CommandText(IUnknown* This,BSTR pbstr) {
		std::string buf;
		if (pbstr)
		{
			auto strbuf = _com_util::ConvertBSTRToString(pbstr);
			if (strbuf)
			{
				buf = "SQL 语句:";
				buf+= strbuf;
				delete[] strbuf;
				
			}
		}
		if (!buf.empty())
		{
			LOG_INFO << buf.c_str();
		}
		return m_put(This, pbstr);
	}
	static HRESULT __stdcall raw_Execute(IUnknown* This, VARIANT * RecordsAffected, VARIANT * Parameters, long Options, struct _Recordset** ppiRs) {
		std::string buf;
		if (Parameters) {
			if (Parameters->vt != VT_EMPTY && Parameters->vt != VT_NULL)
			{
				if (Parameters->vt == VT_BSTR && Parameters->bstrVal) {
					auto strbuf = _com_util::ConvertBSTRToString(Parameters->bstrVal);
					if (strbuf)
					{
						buf = "SQL 语句:";
						buf += strbuf;
						delete[] strbuf;
					}
				}
				//LOG_INFO<< "Parameters vt 类型不是 BSTR 类型如下:" << Parameters->vt;
			}
		}
		if (RecordsAffected) 
		{
			if (RecordsAffected->vt != VT_EMPTY && RecordsAffected->vt != VT_NULL)
			{
				if (RecordsAffected->vt == VT_BSTR && RecordsAffected->bstrVal) {
					auto strbuf = _com_util::ConvertBSTRToString(RecordsAffected->bstrVal);
					if (strbuf)
					{
						buf += strbuf;
						delete[] strbuf;
					}
				}
				//LOG_INFO << "RecordsAffected vt 类型不是 BSTR 类型如下:" << RecordsAffected->vt;
			}
		}
		if (!buf.empty())
		{
			LOG_INFO << buf.c_str();
		}
		if (This)
		{
			_Command* pcmd = (_Command*)This;
			if (pcmd->Parameters)
			{
				long dwArgLen = pcmd->Parameters->GetCount();
				LOG_INFO << "当前SQL语句带有\t" << dwArgLen << "\t条参数 参数值如下:";
				for (long index = 0;index < dwArgLen;++index)
				{
					_Parameter* item = nullptr;
					VARIANT varIndex;
					varIndex.vt = VT_I4;
					varIndex.lVal = index;
					 HRESULT pHr = pcmd->Parameters->get_Item(varIndex, &item);
					 if (SUCCEEDED(pHr) && item)
					 {
						 _bstr_t name = item->GetName();
						 auto szStr = _com_util::ConvertBSTRToString(name.GetBSTR());
						 std::string strName;
						 if (szStr)
						 {
							 strName.assign(szStr);
							 delete[] szStr;
						 }
						 std::string val = varToStr(item->GetValue());
						 auto nType = item->GetType();
						 LOG_INFO << "参数名字:\t" << strName.c_str() << "\t参数类型:\t" << printArgType(nType).c_str() << "\t参数内容:\t" << val;
						
					 }
				}
			}
		}
		HRESULT hr =  m_open_old(This,RecordsAffected,Parameters,Options,ppiRs);
		try
		{
			if (SUCCEEDED(hr))
			{
				if (*ppiRs)
				{
					parseRecordset(*ppiRs);

				}
			}
		}
		catch (_com_error& e)
		{
			auto szError = e.ErrorMessage();
			if (szError)
			{
				LOG_ERROR_(1) << szError;
				delete[] szError;
			}
			(*ppiRs)->MoveFirst();
			return hr;
		}
		return hr;
	}
	static bool unhookMethods() {
		bool bResult = false;
		DWORD m_dwOldProtect = NULL;
		DWORD dwlen = 20;
		auto m_Bret = ::VirtualProtectEx(::GetCurrentProcess(), (LPVOID)m_vttableAddress, sizeof(DWORD) * dwlen, PAGE_EXECUTE_READWRITE, &m_dwOldProtect);
		if (!m_Bret)
		{
			return bResult;
		}
		((LPDWORD)m_vttableAddress)[12] = (DWORD)m_put;
		((LPDWORD)m_vttableAddress)[17] = (DWORD)m_open_old;
		m_Bret = ::VirtualProtectEx(::GetCurrentProcess(), (LPVOID)m_vttableAddress, sizeof(DWORD) * dwlen, m_dwOldProtect, NULL);	//改回原来的属性
		bResult = true;
		return bResult;
	}
	static DWORD m_vttableAddress;
	static FUNCCOM_OPEN m_open_old;
	static FUNCCOM_PUT m_put;
};
DWORD hookCommand::m_vttableAddress = NULL;
FUNCCOM_OPEN hookCommand::m_open_old = nullptr;
FUNCCOM_PUT hookCommand::m_put = nullptr;
static bool g_connect_hook_status = false;
static bool g_recordest_hook_status = false;
static bool g_command_hook_status = false;
std::string GetClsidToProgId(const CLSID clsid)
{
	std::string strProgId;
	LPOLESTR str = nullptr;
	char * debugstr = nullptr;
	ProgIDFromCLSID(clsid, &str);
	debugstr = _com_util::ConvertBSTRToString(str);
	if (debugstr)
	{
		strProgId.assign(debugstr);
		delete[] debugstr;
	}
	CoTaskMemFree(str);
	str = nullptr;
	return strProgId;
}
void printClsid(CLSID clsid,char* title) {
#ifdef _DEBUG
	LPOLESTR str = nullptr;
	char * debugstr = nullptr;
	StringFromCLSID(clsid, &str);
	debugstr = _com_util::ConvertBSTRToString(str);
	if (debugstr)
	{
		LOG_INFO << title << "clsid " << debugstr;
		delete[] debugstr;
	}
	CoTaskMemFree(str);
	str = nullptr;
	ProgIDFromCLSID(clsid, &str);
	debugstr = _com_util::ConvertBSTRToString(str);
	if (debugstr)
	{
		LOG_INFO << title << "ProgID " << debugstr;
		delete[] debugstr;
	}
	CoTaskMemFree(str);
	str = nullptr;
#endif
}
_RecordsetPtr lpRecordset = nullptr;
_CommandPtr lpconmand = nullptr;
_ConnectionPtr lpConnection = nullptr;
HRESULT __stdcall CoCreateInstance_New(
	REFCLSID  rclsid,
	LPUNKNOWN pUnkOuter,
	DWORD     dwClsContext,
	REFIID    riid,
	LPVOID    *ppv
) 
{
	printClsid(rclsid, "rclsid");
	//LOG_INFO << GetClsidToProgId(rclsid);
	CLSID clsid;
	LPVOID pdata = nullptr;
	DWORD dwOldFlag = NULL;
	HRESULT hr1 = CLSIDFromProgID(OLESTR("ADODB.Connection"), &clsid);
	//clsid = __uuidof(Connection);
	if (!g_connect_hook_status)
	{
		hr1 = CoCreateInstance_OLD(clsid, NULL, CLSCTX_ALL, __uuidof(IUnknown), &pdata);
		if (SUCCEEDED(hr1))
		{
			if (::VirtualProtectEx(::GetCurrentProcess(), pdata, sizeof(LPVOID), PAGE_EXECUTE_READWRITE, &dwOldFlag))
			{
				LPDWORD vtrtab = (LPDWORD)*((PDWORD)pdata);
				LOG_INFO << "Connection this addr:--->" << pdata;
				LOG_INFO << "Connection vtr table addr:--->" << vtrtab;
				if (hookConnectIon::hookMethods((DWORD)vtrtab)) {
					lpConnection = reinterpret_cast<_Connection*>(pdata);
					//lpconmand->Release();
					LOG_INFO << "Connection 成功";
				}
				g_connect_hook_status = true;
			}
			::VirtualProtectEx(::GetCurrentProcess(), pdata, sizeof(LPVOID), dwOldFlag, NULL);
		}
		else
		{
			LOG_INFO << hr1;
		}
	}
	
	hr1 = CLSIDFromProgID(OLESTR("ADODB.recordset"), &clsid);
	//clsid = __uuidof(Recordset);
	if (!g_recordest_hook_status)
	{
		hr1 = CoCreateInstance_OLD(clsid, NULL, CLSCTX_ALL, __uuidof(IUnknown), &pdata);
		if (SUCCEEDED(hr1))
		{
			if (::VirtualProtectEx(::GetCurrentProcess(), pdata, sizeof(LPVOID), PAGE_EXECUTE_READWRITE, &dwOldFlag))
			{
				LPDWORD vtrtab = (LPDWORD)*((PDWORD)pdata);
				LOG_INFO << "recordset this addr:--->" << pdata;
				LOG_INFO << "recordset vtr table addr:--->" << vtrtab;
				if (hookRecordset::hookMethods((DWORD)vtrtab))
				{
					lpRecordset = reinterpret_cast<_Recordset*>(pdata);
					//lpconmand->Release();
					LOG_INFO << "recordset hook 成功";
				}
				g_recordest_hook_status = true;
			}
			::VirtualProtectEx(::GetCurrentProcess(), pdata, sizeof(LPVOID), dwOldFlag, NULL);
		}
		else
		{
			LOG_INFO << hr1;
		}
	}

	hr1 = CLSIDFromProgID(OLESTR("ADODB.command"), &clsid);
	//clsid = __uuidof(Command);
	if (!g_command_hook_status)
	{
		hr1 = CoCreateInstance_OLD(clsid, NULL, CLSCTX_ALL, __uuidof(IUnknown), &pdata);
		if (SUCCEEDED(hr1))
		{
			if (::VirtualProtectEx(::GetCurrentProcess(), pdata, sizeof(LPVOID), PAGE_EXECUTE_READWRITE, &dwOldFlag))
			{
				LPDWORD vtrtab = (LPDWORD)*((PDWORD)pdata);
				LOG_INFO << "command this addr:--->" << pdata;
				LOG_INFO << "command vtr table addr:--->" << vtrtab;
				if (hookCommand::hookMethods((DWORD)vtrtab))
				{
					lpconmand = reinterpret_cast<_Command*>(pdata);
					//lpconmand->Release();
					LOG_INFO << "command hook 成功";
				}
				g_command_hook_status = true;
				LOG_INFO << "★★★★★★★★★★★★★★★★★★★★以下是获取的数据语句★★★★★★★★★★★★★★★★★★★★★★★★";
			}
			::VirtualProtectEx(::GetCurrentProcess(), pdata, sizeof(LPVOID), dwOldFlag, NULL);
		}
		else
		{

			LOG_INFO << hr1;
		}
	}
	HRESULT hr = CoCreateInstance_OLD(rclsid, pUnkOuter, dwClsContext, riid, ppv);
	return hr;
}

BOOL APIENTRY DllMain(HMODULE hModule,
	DWORD  ul_reason_for_call,
	LPVOID lpReserved
)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
	{
// 		CHAR szPath[MAX_PATH] = { 0 };
// 		GetModuleFileNameA(NULL, szPath, MAX_PATH);
// 		google::InitGoogleLogging(szPath);
// 		google::SetLogDestination(google::GLOG_INFO, "C:\\");
// 		//设置特定严重级别的日志的输出目录和前缀。第一个参数为日志级别，第二个参数表示输出目录及日志文件名前缀
// 		google::SetStderrLogging(google::GLOG_INFO);  //大于指定级别的日志都输出到标准输出
// 		google::SetLogFilenameExtension("hookdb_sqlinfo.");  //在日志文件名中级别后添加一个扩展名。适用于所有严重级别
// 		FLAGS_colorlogtostderr = true;  //设置记录到标准输出的颜色消息（如果终端支持）
// 		FLAGS_max_log_size = 10;  //设置最大日志文件大小（以MB为单位）
// 		FLAGS_stop_logging_if_full_disk = true;  //设置是否在磁盘已满时避免日志记录到磁盘
		plog::init(plog::info, "c:\\sql.log", 0xA00000, 10240);
		plog::init<1>(plog::error, "c:\\error.log", 0xA00000, 10240);
		if (oleHandle == NULL) {
			oleHandle = LoadLibrary("ole32.dll");
		}
		if (oleHandle)
		{
			CoCreateInstance_OLD = (PROC_CREATE_FUN)GetProcAddress(oleHandle, "CoCreateInstance");
			LOG_INFO << (DWORD)CoCreateInstance_OLD;
			if (Mhook_SetHook((PVOID*)&CoCreateInstance_OLD, CoCreateInstance_New)) {
				LOG_INFO << "注入....";
			}
//			init();
		}
	}
	break;
	case DLL_PROCESS_DETACH:
	{
		hookCommand::unhookMethods();
		hookRecordset::unhookMethods();
		hookConnectIon::unhookMethods();
		Mhook_Unhook((PVOID*)&CoCreateInstance_OLD);
//		google::ShutdownGoogleLogging();	//关闭
	}
	break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
		break;
	}
	return TRUE;
}
