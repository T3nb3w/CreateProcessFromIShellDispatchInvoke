#include <windows.h>
#include <stdio.h>
#include <initguid.h>
#include <stdint.h>

DEFINE_GUID(clsid, 0x13709620, 0xc279, 0x11ce, 0xa4, 0x9e, 0x44, 0x45, 0x53, 0x54, 0x00, 0x00);

 int main(int argc,  char **argv){
    
    LPOLESTR clsidstr = NULL;
    StringFromCLSID(&clsid, &clsidstr);
    printf("Our targeted CLSID is %ls\n", clsidstr);
    HRESULT hr;
    hr = CoInitialize(NULL);

    FARPROC DllGetClassObject = GetProcAddress(LoadLibrary("shell32.dll"), "DllGetClassObject");
    printf("DllGetClassObject is at 0x%p\n\n", DllGetClassObject);
    
    IClassFactory *icf = NULL;
    hr = DllGetClassObject(&clsid, &IID_IClassFactory, (void **)&icf);

    if(hr != S_OK) {
            printf("DllGetClassObject failed to do something. Error %d HRESULT 0x%08x\n", GetLastError(), (unsigned int)hr);
            CoUninitialize(); 
            ExitProcess(0);          
    }

   IDispatch *id = NULL;
   hr = icf->lpVtbl->CreateInstance(icf, NULL, &IID_IDispatch, (void **)&id);
    if(hr != S_OK) {
            printf("CreateInstance failed to do something. Error %d HRESULT 0x%08x\n", GetLastError(), (unsigned int)hr);
            CoUninitialize();
            ExitProcess(0);          
    }

   WCHAR *member = L"ShellExecute";
   DISPID dispid = 0;

    hr = id->lpVtbl->GetIDsOfNames(id, &IID_NULL, &member, 1, LOCALE_USER_DEFAULT, &dispid);
    if(hr != S_OK) {
            printf("GetIDsOfNames failed to do something. Error %d HRESULT 0x%08x\n", GetLastError(), (unsigned int)hr);
            CoUninitialize();
            ExitProcess(0);          
    }

    VARIANT args = { VT_EMPTY };
    args.vt = VT_BSTR;
    
    WCHAR buffer[512];
    ZeroMemory(buffer, sizeof(buffer));
    char* param = argv[1];
    MultiByteToWideChar(CP_ACP, MB_ERR_INVALID_CHARS, param, strlen(param), buffer, sizeof(buffer)-1);
    args.bstrVal = SysAllocString(buffer);

    DISPPARAMS dp = {&args, NULL, 1, 0};

    VARIANT output = { VT_EMPTY };
    hr = id->lpVtbl->Invoke(id, dispid, &IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &output, NULL, NULL);
     if(hr != S_OK) {
            printf("Invoke failed to do something. Error %d HRESULT 0x%08x\n", GetLastError(), (unsigned int)hr);
            CoUninitialize();
            ExitProcess(0);          
    }


    id->lpVtbl->Release(id);
    icf->lpVtbl->Release(icf);
    SysFreeString(args.bstrVal);
    CoUninitialize();
    return 0;
 }
