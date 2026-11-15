#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <winstring.h>
#include <intrin.h>
#include <map>
#include <string>
#include <string_view>
#include <cstdio>
#include "BoundaryExpression.h"

static HINSTANCE ownModule;
static BYTE* officeBase;
static void** importSlot;
static void* savedImport;
static SRWLOCK cacheLock = SRWLOCK_INIT;
static std::map<std::wstring, std::wstring> cache;
static std::wstring logPath;
static unsigned long rewritten = 0;
static unsigned long maximumRewritten = 0;
static unsigned long minimumRewritten = 0;
static HWND testWindow;
static std::wstring targetWorkbookName=L"触控板滚动测试.xlsx";
static bool verboseLog = false;
static void** canvasSlot;
using SetCanvasSizeFn = void(*)(void*,double,double);
static SetCanvasSizeFn originalCanvasSize;
#include "NativePerformance.h"
#ifdef CODEX_NATIVE_LAB
static LONG WINAPI RecordLabException(EXCEPTION_POINTERS* exception) {
    if(exception->ExceptionRecord->ExceptionCode==0xe0000002 || exception->ExceptionRecord->ExceptionCode==0xc0000005 || exception->ExceptionRecord->ExceptionCode==0xc0000409) {
        FILE* file=nullptr;_wfopen_s(&file,logPath.c_str(),L"a");if(file){fprintf(file,"{\"event\":\"lab-exception\",\"stack\":[");void* frames[40];auto count=CaptureStackBackTrace(0,40,frames,nullptr);for(unsigned i=0;i<count;++i){HMODULE module=nullptr;GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS|GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,reinterpret_cast<LPCWSTR>(frames[i]),&module);wchar_t path[MAX_PATH]{};GetModuleFileNameW(module,path,MAX_PATH);const auto name=wcsrchr(path,L'\\');fprintf(file,"%s\"%ls+%llx\"",i?",":"",name?name+1:path,reinterpret_cast<uintptr_t>(frames[i])-reinterpret_cast<uintptr_t>(module));}fprintf(file,"]}\n");fclose(file);}
    }
    return EXCEPTION_CONTINUE_SEARCH;
}
#endif
#include "FrozenPaneBoundary.h"
#ifdef CODEX_NATIVE_LAB
#include "NativeCanvasProbe.h"
#endif

struct CanvasFrame { bool found=false; RECT viewport{}; RECT content{}; int left=0,top=0,pane=-1,type=-1; double zoom=0; };
static CanvasFrame GetCanvasFrame() {
    NATIVE_PERF(2);
    CanvasFrame result;
    CONTEXT context{};RtlCaptureContext(&context);
    const auto excel=reinterpret_cast<uintptr_t>(GetModuleHandleW(nullptr));
    for(int frame=0;frame<12 && context.Rip;++frame) {
        if(context.Rip-excel==0x96d10) {
            result.found=true;
            result.viewport=*reinterpret_cast<RECT*>(context.Rsp+0x30);
            auto content=reinterpret_cast<BYTE*>(context.Rdi);
            result.content={*reinterpret_cast<int*>(content),*reinterpret_cast<int*>(content+0x10),*reinterpret_cast<int*>(content+0x20),*reinterpret_cast<int*>(content+0x30)};
            result.type=*reinterpret_cast<int*>(content+0x80);
            result.pane=*reinterpret_cast<int*>(content+0x90);
            memcpy(&result.zoom,&context.Xmm6,sizeof(double));
            break;
        }
        if(context.Rip-excel>=0x46f620 && context.Rip-excel<0x46f9e7) {
            result.found=true;
            result.viewport=*reinterpret_cast<RECT*>(context.Rbp-0x29);
            result.left=*reinterpret_cast<int*>(context.Rsp+0x28);
            result.top=*reinterpret_cast<int*>(context.Rsp+0x2c);
            result.pane=static_cast<int>(context.Rdi);
            memcpy(&result.zoom,&context.Xmm7,sizeof(double));
            break;
        }
        DWORD64 base=0;auto function=RtlLookupFunctionEntry(context.Rip,&base,nullptr);
        if(!function) {context.Rip=*reinterpret_cast<DWORD64*>(context.Rsp);context.Rsp+=8;continue;}
        PVOID handler=nullptr;DWORD64 establisher;
        RtlVirtualUnwind(UNW_FLAG_NHANDLER,base,context.Rip,function,&context,&handler,&establisher,nullptr);
    }
    return result;
}
#include "DynamicCanvas.h"

static void TraceCanvasSize(void* layer,double width,double height) {
    auto table=*reinterpret_cast<void***>(layer);
    void* base=reinterpret_cast<void*(*)(void*)>(table[24])(layer);
    auto baseTable=*reinterpret_cast<void***>(base);
    auto tag=reinterpret_cast<const void*(*)(void*)>(baseTable[37])(base);
    const auto info=GetCanvasFrame();
    const auto caller=reinterpret_cast<uintptr_t>(_ReturnAddress())-reinterpret_cast<uintptr_t>(GetModuleHandleW(nullptr));
    AcquireSRWLockExclusive(&cacheLock);
    FILE* file=nullptr;
    _wfopen_s(&file,logPath.c_str(),L"a");
    if(file) {
        fprintf(file,"{\"event\":\"canvas-size\",\"pid\":%lu,\"tick\":%llu,\"layer\":\"%p\",\"tag\":\"%p\",\"width\":%.3f,\"height\":%.3f,\"caller\":%llu,\"found\":%s,\"pane\":%d,\"type\":%d,\"zoom\":%.3f,\"viewport\":[%ld,%ld,%ld,%ld],\"content\":[%ld,%ld,%ld,%ld],\"contentOrigin\":[%d,%d]}\n",GetCurrentProcessId(),GetTickCount64(),layer,tag,width,height,caller,info.found?"true":"false",info.pane,info.type,info.zoom,info.viewport.left,info.viewport.top,info.viewport.right,info.viewport.bottom,info.content.left,info.content.top,info.content.right,info.content.bottom,info.left,info.top);
        fclose(file);
    }
    ReleaseSRWLockExclusive(&cacheLock);
    originalCanvasSize(layer,width,height);
}

static void Log(const char* event, uintptr_t caller = 0, const wchar_t* expression = nullptr) {
    if(!verboseLog)return;
    FILE* file = nullptr;
    _wfopen_s(&file, logPath.c_str(), L"a");
    if (!file) return;
    fprintf(file, "{\"event\":\"%s\",\"pid\":%lu,\"tick\":%llu,\"callerRva\":%llu,\"rewritten\":%lu",
        event, GetCurrentProcessId(), GetTickCount64(), static_cast<unsigned long long>(caller), rewritten);
    if (expression) fprintf(file, ",\"expression\":\"%ls\"", expression);
    fprintf(file, "}\n");
    fclose(file);
}

static HRESULT WINAPI BoundedStringReference(PCWSTR text, UINT32 length, HSTRING_HEADER* header, HSTRING* result) {
    NATIVE_PERF(0);
    if (text && length && length < 2048) {
        const std::wstring_view candidate(text, length);
        if (candidate.find(L"Tracker.Position.") != std::wstring_view::npos ||
            (dynamicBounds && candidate.find(L"freeFloatingPadding") != std::wstring_view::npos)) {
            void* frames[10];
            const USHORT depth = CaptureStackBackTrace(0, 10, frames, nullptr);
            uintptr_t caller = 0;
            bool maximum=false;
            bool minimum=false;
            for (USHORT i = 0; i < depth; ++i) {
                const auto rva = reinterpret_cast<uintptr_t>(frames[i]) - reinterpret_cast<uintptr_t>(officeBase);
                if ((rva >= 0x0010b908 && rva < 0x0010c15e) || (rva >= 0x0010ce90 && rva < 0x0010d67b)) {
                    caller = rva;
                    break;
                }
                if(dynamicBounds && rva>=0x109a30 && rva<0x10a880){caller=rva;maximum=true;break;}
                if(dynamicBounds && rva>=0x10a880 && rva<0x10b75b){caller=rva;minimum=true;break;}
            }
            if (!caller) return WindowsCreateStringReference(text, length, header, result);
            if((minimum||maximum)&&candidate.find(L"this.target.position.x")!=std::wstring_view::npos&&!BindFrozenBackend())return WindowsCreateStringReference(text,length,header,result);
#ifdef CODEX_NATIVE_ADDIN
            if(!minimum&&!maximum&&!BindFrozenBackend())return WindowsCreateStringReference(text,length,header,result);
#endif
#ifdef CODEX_NATIVE_LAB
            if(!maximum&&!minimum)CaptureNativeProbe();
#endif
            AcquireSRWLockExclusive(&cacheLock);
            auto [entry, inserted] = cache.emplace(std::wstring(candidate), std::wstring());
            if (inserted) entry->second = minimum?FrozenMinimum(entry->first):(maximum?ContentMaximum(entry->first):ClampTrackerPositions(entry->first));
            if(maximum){++maximumRewritten;if(verboseLog||maximumRewritten<=8)Log("hard-content-maximum",caller,entry->second.c_str());}
            else if(minimum){++minimumRewritten;if(verboseLog||minimumRewritten<=8)Log("frozen-pane-minimum",caller,entry->second.c_str());}
            else {++rewritten;if(verboseLog||rewritten<=8)Log("bounded-native-expression",caller,entry->second.c_str());}
            const wchar_t* replacement = entry->second.c_str();
            const UINT32 count = static_cast<UINT32>(entry->second.size());
            ReleaseSRWLockExclusive(&cacheLock);
            return WindowsCreateStringReference(replacement, count, header, result);
        }
    }
    return WindowsCreateStringReference(text, length, header, result);
}

static BOOL CALLBACK FindTestWindow(HWND hwnd, LPARAM) {
    DWORD pid = 0;
    GetWindowThreadProcessId(hwnd, &pid);
    wchar_t title[256], className[32];
    GetWindowTextW(hwnd, title, 256);
    GetClassNameW(hwnd, className, 32);
    if (pid == GetCurrentProcessId() && wcscmp(className, L"XLMAIN") == 0 &&
        wcsstr(title, targetWorkbookName.c_str())) {
        testWindow = hwnd;
        return FALSE;
    }
    return TRUE;
}

extern "C" __declspec(dllexport) int WINAPI xlAutoOpen() {
    if (importSlot) return 1;
    wchar_t targetName[MAX_PATH];
    const auto targetLength=GetEnvironmentVariableW(L"CODEX_EXCEL_BOUNDARY_WORKBOOK",targetName,MAX_PATH);
    if(targetLength>0&&targetLength<MAX_PATH)targetWorkbookName=targetName;
    EnumWindows(FindTestWindow, 0);
#ifdef CODEX_NATIVE_ADDIN
    wchar_t bootstrap[2];
    const bool addinBootstrap=GetEnvironmentVariableW(L"CODEX_EXCEL_ADDIN_BOOTSTRAP",bootstrap,2)==1&&bootstrap[0]==L'1';
    if(!testWindow&&!addinBootstrap)return 0;
#else
    if (!testWindow) return 0;
#endif
    officeBase = reinterpret_cast<BYTE*>(GetModuleHandleW(L"mso40uiwin32client.dll"));
    if (!officeBase) return 0;
    const auto dos = reinterpret_cast<IMAGE_DOS_HEADER*>(officeBase);
    const auto nt = reinterpret_cast<IMAGE_NT_HEADERS64*>(officeBase + dos->e_lfanew);
    // Offsets correspond to the supported Office build; validate before patching its imports.
    if (nt->FileHeader.TimeDateStamp != 0x6af2712e || nt->OptionalHeader.SizeOfImage != 0x138a000) return 0;
    wchar_t path[MAX_PATH];
    GetModuleFileNameW(ownModule, path, MAX_PATH);
    logPath = path;
    logPath = logPath.substr(0, logPath.find_last_of(L"\\/")) + L"\\boundary-native-events.jsonl";
    wchar_t logFolder[MAX_PATH];
    if(GetEnvironmentVariableW(L"CODEX_EXCEL_BOUNDARY_LOG_DIR",logFolder,MAX_PATH)>0)
        logPath=std::wstring(logFolder)+L"\\boundary-native-events.jsonl";
    wchar_t verbose[2];
    verboseLog = GetEnvironmentVariableW(L"CODEX_EXCEL_BOUNDARY_VERBOSE", verbose, 2) == 1 && verbose[0] == L'1';
    HMODULE retained;
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_PIN,
        reinterpret_cast<LPCWSTR>(xlAutoOpen), &retained);
    auto slot = reinterpret_cast<void**>(officeBase + 0x01347e10);
    DWORD protection;
    if (!VirtualProtect(slot, sizeof(void*), PAGE_READWRITE, &protection)) return 0;
    savedImport = InterlockedExchangePointer(slot, reinterpret_cast<void*>(BoundedStringReference));
    importSlot = slot;
    DWORD ignored;
    VirtualProtect(slot, sizeof(void*), protection, &ignored);
    wchar_t extent[2];
    dynamicBounds=GetEnvironmentVariableW(L"CODEX_EXCEL_DYNAMIC_BOUNDARY",extent,2)==1 && extent[0]==L'1';
    if(dynamicBounds) {
        if(!InstallDynamicBounds()) {
            RemoveDynamicBounds();
            if(VirtualProtect(importSlot,sizeof(void*),PAGE_READWRITE,&protection)) {
                InterlockedExchangePointer(importSlot,savedImport);VirtualProtect(importSlot,sizeof(void*),protection,&ignored);
            }
            importSlot=nullptr;return 0;
        }
    } else if(GetEnvironmentVariableW(L"CODEX_EXCEL_EXTENT_LOG",extent,2)==1 && extent[0]==L'1') {
        canvasSlot=reinterpret_cast<void**>(officeBase+0x00ac8e18);
        if(VirtualProtect(canvasSlot,sizeof(void*),PAGE_READWRITE,&protection)) {
            originalCanvasSize=reinterpret_cast<SetCanvasSizeFn>(InterlockedExchangePointer(canvasSlot,reinterpret_cast<void*>(TraceCanvasSize)));
            VirtualProtect(canvasSlot,sizeof(void*),protection,&ignored);
        }
    }
    Log("installed-scoped-boundary");
#ifdef CODEX_NATIVE_LAB
    AddVectoredExceptionHandler(1,RecordLabException);
#endif
    return 1;
}

extern "C" __declspec(dllexport) int WINAPI xlAutoClose() {
    RemoveDynamicBounds();
    if(canvasSlot && originalCanvasSize) {
        DWORD protection,ignored;
        if(VirtualProtect(canvasSlot,sizeof(void*),PAGE_READWRITE,&protection)) {
            InterlockedExchangePointer(canvasSlot,reinterpret_cast<void*>(originalCanvasSize));
            VirtualProtect(canvasSlot,sizeof(void*),protection,&ignored);
        }
        canvasSlot=nullptr;
    }
    if (importSlot) {
        DWORD protection;
        if (VirtualProtect(importSlot, sizeof(void*), PAGE_READWRITE, &protection)) {
            InterlockedExchangePointer(importSlot, savedImport);
            DWORD ignored;
            VirtualProtect(importSlot, sizeof(void*), protection, &ignored);
            importSlot = nullptr;
        }
        Log("restored-import");
    }
    return 1;
}
extern "C" __declspec(dllexport) int WINAPI xlAutoRemove() { return xlAutoClose(); }
BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) {
        ownModule = instance;
        DisableThreadLibraryCalls(instance);
    }
    return TRUE;
}
