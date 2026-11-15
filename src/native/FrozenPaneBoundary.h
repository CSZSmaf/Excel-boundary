#pragma once
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.UI.Composition.h>
#include <winrt/Windows.UI.Composition.Interactions.h>
#include <mutex>
#include <cmath>

struct FrozenBinding {
    float x=0,y=0;
    bool canvasBound=false;
    bool observedPosition=false;
    float lastX=0,lastY=0,lastScale=1;
    winrt::Windows::UI::Composition::Interactions::InteractionTracker tracker{nullptr};
    winrt::Windows::UI::Composition::CompositionPropertySet properties{nullptr};
    winrt::Windows::UI::Composition::Visual content{nullptr};
    winrt::Windows::UI::Composition::CompositionClip originalClip{nullptr};
    winrt::Windows::UI::Composition::InsetClip canvasClip{nullptr};
};
static std::mutex frozenGate;
static std::map<unsigned,FrozenBinding> frozenBindings;
static ULONGLONG lastNativeMotion;
static void AttachCanvasClip(FrozenBinding& binding) {
    if(!binding.canvasBound||!binding.content||binding.canvasClip)return;
    binding.originalClip=binding.content.Clip();
    binding.canvasClip=binding.content.Compositor().CreateInsetClip();
    binding.content.Clip(binding.canvasClip);
    FILE* log=nullptr;if(verboseLog)_wfopen_s(&log,logPath.c_str(),L"a");
    if(log){fprintf(log,"{\"event\":\"content-drawing-clip\",\"pid\":%lu,\"clipped\":true}\n",GetCurrentProcessId());fclose(log);}
}
static void RestoreCanvasClip(FrozenBinding& binding) {
    if(binding.content&&binding.canvasClip&&binding.content.Clip()==binding.canvasClip)binding.content.Clip(binding.originalClip);
    binding.canvasClip=nullptr;binding.originalClip=nullptr;
}
static BYTE* CurrentScrollingBackend() {
    CONTEXT context{};RtlCaptureContext(&context);
    for(int frame=0;frame<12&&context.Rip;++frame) {
        const auto rva=context.Rip-reinterpret_cast<uintptr_t>(officeBase);
        if((rva>=0x109a30&&rva<0x10a1b8)||(rva>=0x10a880&&rva<0x10b0ce))return reinterpret_cast<BYTE*>(context.R15);
        if((rva>=0x10b908&&rva<0x10c15e)||(rva>=0x10ce90&&rva<0x10d67b))return reinterpret_cast<BYTE*>(context.R14);
        DWORD64 base=0;auto function=RtlLookupFunctionEntry(context.Rip,&base,nullptr);
        if(!function){context.Rip=*reinterpret_cast<DWORD64*>(context.Rsp);context.Rsp+=8;continue;}
        PVOID handler=nullptr;DWORD64 establisher;
        RtlVirtualUnwind(UNW_FLAG_NHANDLER,base,context.Rip,function,&context,&handler,&establisher,nullptr);
    }
    return nullptr;
}
static bool BindFrozenBackend() {
    NATIVE_PERF(3);
    BYTE* backend=CurrentScrollingBackend();if(!backend)return false;
    const auto id=*reinterpret_cast<unsigned*>(backend+0x80);
    void* native=*reinterpret_cast<void**>(backend+0x480);if(!native)return false;
    try {
        std::lock_guard<std::mutex> lock(frozenGate);
        auto& binding=frozenBindings[id];
        if(!binding.properties||winrt::get_abi(binding.properties)!=native) {
            winrt::Windows::Foundation::IInspectable object{nullptr};winrt::copy_from_abi(object,native);
            binding.properties=object.as<winrt::Windows::UI::Composition::CompositionPropertySet>();
            binding.properties.InsertScalar(L"CodexFrozenX",binding.x);
            binding.properties.InsertScalar(L"CodexFrozenY",binding.y);
#ifdef CODEX_NATIVE_ADDIN
            binding.properties.InsertScalar(L"CodexEnabled",binding.canvasBound?1.0f:0.0f);
#endif
        }
#ifdef CODEX_NATIVE_ADDIN
        void* trackerNative=*reinterpret_cast<void**>(backend+0x400);
        if(trackerNative&&(!binding.tracker||winrt::get_abi(binding.tracker)!=trackerNative)) {
            winrt::Windows::Foundation::IInspectable object{nullptr};winrt::copy_from_abi(object,trackerNative);
            binding.tracker=object.as<winrt::Windows::UI::Composition::Interactions::InteractionTracker>();
            binding.tracker.Properties().InsertScalar(L"CodexEnabled",binding.canvasBound?1.0f:0.0f);
        }
#endif
        void* contentNative=*reinterpret_cast<void**>(backend+0x3d8);
        if(contentNative&&(!binding.content||winrt::get_abi(binding.content)!=contentNative)) {
            RestoreCanvasClip(binding);
            winrt::Windows::Foundation::IInspectable object{nullptr};winrt::copy_from_abi(object,contentNative);
            binding.content=object.as<winrt::Windows::UI::Composition::Visual>();
        }
        AttachCanvasClip(binding);
#ifdef CODEX_NATIVE_LAB
        float padding=0;binding.properties.TryGetScalar(L"freeFloatingPadding",padding);
        FILE* log=nullptr;_wfopen_s(&log,logPath.c_str(),L"a");if(log){fprintf(log,"{\"event\":\"frozen-binding\",\"id\":%u,\"x\":%.3f,\"y\":%.3f,\"padding\":%.3f}\n",id,binding.x,binding.y,padding);fclose(log);}
#endif
        return true;
    } catch(winrt::hresult_error const&){return false;}
}
static void SetFrozenBoundary(unsigned id,float x,float y,bool enabled=true) {
    std::lock_guard<std::mutex> lock(frozenGate);auto& binding=frozenBindings[id];
#ifdef CODEX_NATIVE_ADDIN
    if(binding.canvasBound!=enabled) {
        if(binding.properties)binding.properties.InsertScalar(L"CodexEnabled",enabled?1.0f:0.0f);
        if(binding.tracker)binding.tracker.Properties().InsertScalar(L"CodexEnabled",enabled?1.0f:0.0f);
    }
#endif
    binding.canvasBound=enabled;
    if(!enabled)binding.observedPosition=false;
    if(enabled)AttachCanvasClip(binding);else RestoreCanvasClip(binding);
    if(binding.x==x&&binding.y==y)return;
    binding.x=x;binding.y=y;
    if(binding.properties){binding.properties.InsertScalar(L"CodexFrozenX",x);binding.properties.InsertScalar(L"CodexFrozenY",y);}
}
static bool NativeScrollingBusy() {
    std::lock_guard<std::mutex> lock(frozenGate);
    const auto now=GetTickCount64();bool any=false;
    for(auto& entry:frozenBindings) {
        auto& binding=entry.second;
        if(!binding.canvasBound||!binding.tracker)continue;
        any=true;const auto position=binding.tracker.Position();const float scale=binding.tracker.Scale();
        if(binding.observedPosition&&(std::abs(position.x-binding.lastX)>0.01f||std::abs(position.y-binding.lastY)>0.01f||scale!=binding.lastScale))lastNativeMotion=now;
        binding.observedPosition=true;binding.lastX=position.x;binding.lastY=position.y;binding.lastScale=scale;
    }
    return any&&now-lastNativeMotion<500;
}
static unsigned EnabledNativeBindings() {
    std::lock_guard<std::mutex> lock(frozenGate);unsigned count=0;
    for(auto& entry:frozenBindings)if(entry.second.canvasBound&&entry.second.tracker)++count;
    return count;
}
static void ForgetFrozenBoundary(unsigned id){std::lock_guard<std::mutex> lock(frozenGate);auto binding=frozenBindings.find(id);if(binding!=frozenBindings.end()){RestoreCanvasClip(binding->second);frozenBindings.erase(binding);}}
static void ClearFrozenBoundaries(){std::lock_guard<std::mutex> lock(frozenGate);for(auto& entry:frozenBindings)RestoreCanvasClip(entry.second);frozenBindings.clear();}
