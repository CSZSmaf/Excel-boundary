#pragma once
#include <vector>
#include <cmath>
#include <algorithm>

struct ExtentHeader {unsigned version,columns,rows,revision,dpi,frozenColumns,frozenRows,reserved;};
struct CanvasState {
    int type=-1,pane=-1;
    unsigned id=0;
    double zoom=1,originalWidth=0,originalHeight=0,effectiveWidth=-1,effectiveHeight=-1;
};
static std::map<void*,CanvasState> canvases;
static std::vector<double> columnPoints,rowPoints;
static unsigned extentRevision;
static unsigned extentDpi=96;
static unsigned frozenColumns,frozenRows;
struct ExtentPixels {double width,height,frozenWidth,frozenHeight;};
static std::map<double,ExtentPixels> pixelCache;
static unsigned long long canvasCalls,canvasFastCalls;
static HWND messageWindow;
static bool dynamicBounds;
using DeleteFn=void*(*)(void*,unsigned);
static DeleteFn originalDelete,originalLayerDelete;
struct ExtraPatch {void** slot;void* original;};
static std::vector<ExtraPatch> extraPatches;

static double PixelExtent(const std::vector<double>& points,double zoom,size_t count=SIZE_MAX) {
    const double scale=extentDpi/72.0*zoom;
    double pixels=0;
    for(size_t i=0;i<std::min(count,points.size());++i)pixels+=std::floor(points[i]*scale+0.5);
    return pixels/zoom;
}
static void ApplyCanvas(void* layer,CanvasState& state) {
    NATIVE_PERF(4);
    double width=state.originalWidth,height=state.originalHeight;
    if(!columnPoints.empty()&&!rowPoints.empty()) {
        auto [cached,inserted]=pixelCache.emplace(state.zoom,ExtentPixels{});
        if(inserted)cached->second={PixelExtent(columnPoints,state.zoom),PixelExtent(rowPoints,state.zoom),PixelExtent(columnPoints,state.zoom,frozenColumns),PixelExtent(rowPoints,state.zoom,frozenRows)};
        const auto& extent=cached->second;
        if(state.type==0x37||state.type==0x2e)width=extent.width+(state.type==0x2e?1/state.zoom:0);
        if(state.type==0x37||state.type==0x25)height=extent.height+(state.type==0x25?0.5/state.zoom:0);
        SetFrozenBoundary(state.id,static_cast<float>((state.pane>=2)?extent.frozenWidth:0),static_cast<float>((state.pane==1||state.pane==3)?extent.frozenHeight:0));
    }
#ifdef CODEX_NATIVE_ADDIN
    else SetFrozenBoundary(state.id,0,0,false);
#endif
    if(width!=state.effectiveWidth||height!=state.effectiveHeight) {
        state.effectiveWidth=width;state.effectiveHeight=height;
        FILE* file=nullptr;if(verboseLog)_wfopen_s(&file,logPath.c_str(),L"a");
        if(file){fprintf(file,"{\"event\":\"content-boundary\",\"pid\":%lu,\"tick\":%llu,\"layer\":\"%p\",\"type\":%d,\"pane\":%d,\"zoom\":%.4f,\"width\":%.4f,\"height\":%.4f,\"columns\":%zu,\"rows\":%zu,\"revision\":%u}\n",GetCurrentProcessId(),GetTickCount64(),layer,state.type,state.pane,state.zoom,width,height,columnPoints.size(),rowPoints.size(),extentRevision);fclose(file);}
    }
    originalCanvasSize(layer,width,height);
}
static void BoundedCanvasSize(void* layer,double width,double height) {
    NATIVE_PERF(1);
    ++canvasCalls;
    auto found=canvases.find(layer);
    if(found!=canvases.end()&&found->second.originalWidth==width&&found->second.originalHeight==height) {
        ++canvasFastCalls;originalCanvasSize(layer,found->second.effectiveWidth,found->second.effectiveHeight);return;
    }
    const auto info=GetCanvasFrame();
    if(info.type==0x37||info.type==0x2e||info.type==0x25) {
        auto& state=canvases[layer];state.type=info.type;state.pane=info.pane;state.zoom=info.zoom;state.id=*reinterpret_cast<unsigned*>(static_cast<BYTE*>(layer)+0xfc);
        state.originalWidth=width;state.originalHeight=height;ApplyCanvas(layer,state);
    } else if(found!=canvases.end()) {
        found->second.originalWidth=width;found->second.originalHeight=height;ApplyCanvas(layer,found->second);
    } else originalCanvasSize(layer,width,height);
}
static void ForgetCanvas(void* layer){
    // Backend expressions also exist for non-worksheet UI layers. Release their bindings too.
    const auto id=*reinterpret_cast<unsigned*>(static_cast<BYTE*>(layer)+0xfc);
    ForgetFrozenBoundary(id);canvases.erase(layer);
#ifdef CODEX_NATIVE_LAB
    ForgetNativeProbe(id);
#endif
}
static void* DeleteScroller(void* layer,unsigned flags){ForgetCanvas(layer);return originalDelete(layer,flags);}
static void* DeleteBaseLayer(void* base,unsigned flags){ForgetCanvas(static_cast<BYTE*>(base)-0xe0);return originalLayerDelete(base,flags);}

static LRESULT CALLBACK BoundaryWindow(HWND hwnd,UINT message,WPARAM wp,LPARAM lp) {
#ifdef CODEX_NATIVE_ADDIN
    if(message==WM_APP+87)return NativeScrollingBusy();
    if(message==WM_APP+88)return EnabledNativeBindings();
#endif
#ifdef CODEX_NATIVE_PROFILE
    if(message==WM_APP+85){WriteNativePerformance(static_cast<unsigned>(wp));return TRUE;}
#endif
#ifdef CODEX_NATIVE_LAB
    if(message>=WM_APP+81&&message<=WM_APP+84)return NativeProbeMessage(message,wp);
    if(message==WM_APP+86)return NativeProbeMessage(message,wp);
#endif
    if(message==WM_COPYDATA) {
        const auto copy=reinterpret_cast<COPYDATASTRUCT*>(lp);
        if(copy->dwData!=0x43425831||copy->cbData<sizeof(ExtentHeader))return FALSE;
        const auto header=static_cast<const ExtentHeader*>(copy->lpData);
        const size_t count=static_cast<size_t>(header->columns)+header->rows;
        if(header->version!=12||header->columns>16384||header->rows>1048576||copy->cbData!=sizeof(ExtentHeader)+count*sizeof(double))return FALSE;
        const auto points=reinterpret_cast<const double*>(header+1);
        columnPoints.assign(points,points+header->columns);rowPoints.assign(points+header->columns,points+count);
        extentRevision=header->revision;extentDpi=header->dpi;
        frozenColumns=header->frozenColumns;frozenRows=header->frozenRows;pixelCache.clear();
        for(auto& entry:canvases)ApplyCanvas(entry.first,entry.second);
        return TRUE;
    }
    return DefWindowProcW(hwnd,message,wp,lp);
}
static bool ExtraPatchSlot(uintptr_t rva,void* replacement,void** original) {
    auto slot=reinterpret_cast<void**>(officeBase+rva);DWORD protection,ignored;
    if(!VirtualProtect(slot,sizeof(void*),PAGE_READWRITE,&protection))return false;
    auto previous=InterlockedExchangePointer(slot,replacement);VirtualProtect(slot,sizeof(void*),protection,&ignored);
    extraPatches.push_back({slot,previous});*original=previous;return true;
}
static bool InstallDynamicBounds() {
    auto base=reinterpret_cast<BYTE*>(GetModuleHandleW(nullptr));
    const auto dos=reinterpret_cast<IMAGE_DOS_HEADER*>(base);
    const auto nt=reinterpret_cast<IMAGE_NT_HEADERS64*>(base+dos->e_lfanew);
    if(nt->FileHeader.TimeDateStamp!=0x6af278b3||nt->OptionalHeader.SizeOfImage!=0x4ca7000)return false;
    if(!ExtraPatchSlot(0xac8e18,reinterpret_cast<void*>(BoundedCanvasSize),reinterpret_cast<void**>(&originalCanvasSize))||
       !ExtraPatchSlot(0xac8ee8,reinterpret_cast<void*>(DeleteScroller),reinterpret_cast<void**>(&originalDelete))||
       !ExtraPatchSlot(0xadd848,reinterpret_cast<void*>(DeleteBaseLayer),reinterpret_cast<void**>(&originalLayerDelete)))return false;
    WNDCLASSW cls{};cls.lpfnWndProc=BoundaryWindow;cls.hInstance=ownModule;cls.lpszClassName=L"CodexExcelDynamicBoundary10";RegisterClassW(&cls);
    const auto name=L"CodexExcelBoundary_"+std::to_wstring(GetCurrentProcessId());
    messageWindow=CreateWindowExW(0,cls.lpszClassName,name.c_str(),0,0,0,0,0,HWND_MESSAGE,nullptr,ownModule,nullptr);
    return messageWindow!=nullptr;
}
static void RemoveDynamicBounds() {
    if(messageWindow){DestroyWindow(messageWindow);messageWindow=nullptr;}
    for(auto it=extraPatches.rbegin();it!=extraPatches.rend();++it){DWORD protection,ignored;if(VirtualProtect(it->slot,sizeof(void*),PAGE_READWRITE,&protection)){InterlockedExchangePointer(it->slot,it->original);VirtualProtect(it->slot,sizeof(void*),protection,&ignored);}}
    extraPatches.clear();canvases.clear();pixelCache.clear();ClearFrozenBoundaries();
}
