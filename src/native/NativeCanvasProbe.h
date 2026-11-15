#pragma once
#include <winrt/Windows.Foundation.h>
#include <winrt/Windows.UI.Composition.h>
#include <winrt/Windows.UI.Composition.Interactions.h>
#include <mutex>
#include <memory>
#include <vector>
using namespace winrt;
using namespace Windows::UI::Composition;
using namespace Windows::UI::Composition::Interactions;
using Windows::Foundation::Numerics::float3;
struct BoundsObserver:implements<BoundsObserver,IInteractionTrackerOwner> {
    std::mutex gate;float3 value{0,0,0};
    void CustomAnimationStateEntered(InteractionTracker const&,InteractionTrackerCustomAnimationStateEnteredArgs const&){}
    void IdleStateEntered(InteractionTracker const&,InteractionTrackerIdleStateEnteredArgs const&){}
    void InertiaStateEntered(InteractionTracker const&,InteractionTrackerInertiaStateEnteredArgs const&){}
    void InteractingStateEntered(InteractionTracker const&,InteractionTrackerInteractingStateEnteredArgs const&){}
    void RequestIgnored(InteractionTracker const&,InteractionTrackerRequestIgnoredArgs const&){}
    void ValuesChanged(InteractionTracker const&,InteractionTrackerValuesChangedArgs const& args){std::lock_guard<std::mutex> lock(gate);value=args.Position();}
    float3 Value(){std::lock_guard<std::mutex> lock(gate);return value;}
};
struct NativeProbe {
    unsigned id,ownerId;
    InteractionTracker source{nullptr},bounds{nullptr},minimum{nullptr},rendered{nullptr};
    Visual content{nullptr},viewport{nullptr},output{nullptr};
    CompositionPropertySet properties{nullptr};
    com_ptr<BoundsObserver> observer;
    com_ptr<BoundsObserver> minimumObserver,renderedObserver;
};
static std::mutex probeGate;
static std::map<void*,std::shared_ptr<NativeProbe>> nativeProbes;
static void ForgetNativeProbe(unsigned ownerId) {
    std::lock_guard<std::mutex> lock(probeGate);
    for(auto it=nativeProbes.begin();it!=nativeProbes.end();) {
        if(it->second->ownerId==ownerId)it=nativeProbes.erase(it);else ++it;
    }
}
template<class T> static T QueryNative(void* object) {
    Windows::Foundation::IInspectable inspectable{nullptr};copy_from_abi(inspectable,object);return inspectable.as<T>();
}
static void CaptureNativeProbe() {
    CONTEXT context{};RtlCaptureContext(&context);BYTE* backend=nullptr;
    for(int frame=0;frame<12&&context.Rip;++frame) {
        const auto rva=context.Rip-reinterpret_cast<uintptr_t>(officeBase);
        if((rva>=0x10b908&&rva<0x10c15e)||(rva>=0x10ce90&&rva<0x10d67b)){backend=reinterpret_cast<BYTE*>(context.R14);break;}
        DWORD64 base=0;auto function=RtlLookupFunctionEntry(context.Rip,&base,nullptr);
        if(!function){context.Rip=*reinterpret_cast<DWORD64*>(context.Rsp);context.Rsp+=8;continue;}
        PVOID handler=nullptr;DWORD64 establisher;RtlVirtualUnwind(UNW_FLAG_NHANDLER,base,context.Rip,function,&context,&handler,&establisher,nullptr);
    }
    if(!backend)return;
    auto identity=*reinterpret_cast<void**>(backend+0x400);
    std::lock_guard<std::mutex> lock(probeGate);
    if(nativeProbes.count(identity))return;
    try {
        FILE* trace=nullptr;_wfopen_s(&trace,logPath.c_str(),L"a");if(trace){fprintf(trace,"{\"event\":\"probe-start\",\"backend\":\"%p\",\"tracker\":\"%p\"}\n",backend,identity);fclose(trace);}
        auto probe=std::make_shared<NativeProbe>();
        probe->id=static_cast<unsigned>(reinterpret_cast<uintptr_t>(identity));
        probe->ownerId=*reinterpret_cast<unsigned*>(backend+0x80);
        probe->source=QueryNative<InteractionTracker>(identity);
        _wfopen_s(&trace,logPath.c_str(),L"a");if(trace){fprintf(trace,"{\"event\":\"probe-source\"}\n");fclose(trace);}
        probe->content=QueryNative<Visual>(*reinterpret_cast<void**>(backend+0x3d8));
        const auto composition=*reinterpret_cast<void**>(backend+0x1a0);
        probe->viewport=QueryNative<Visual>(composition);
        probe->properties=QueryNative<CompositionPropertySet>(*reinterpret_cast<void**>(backend+0x480));
        _wfopen_s(&trace,logPath.c_str(),L"a");if(trace){fprintf(trace,"{\"event\":\"probe-references\"}\n");fclose(trace);}
        probe->observer=make_self<BoundsObserver>();
        auto compositor=probe->source.Compositor();
        probe->bounds=InteractionTracker::CreateWithOwner(compositor,*probe->observer);
        _wfopen_s(&trace,logPath.c_str(),L"a");if(trace){fprintf(trace,"{\"event\":\"probe-created\"}\n");fclose(trace);}
        probe->bounds.MinPosition({-100000000,-100000000,0});probe->bounds.MaxPosition({100000000,100000000,0});
        auto expression=compositor.CreateExpressionAnimation(L"Vector3(Tracker.MaxPosition.X,Tracker.MaxPosition.Y,0)");
        expression.SetReferenceParameter(L"Tracker",probe->source);probe->bounds.TryUpdatePositionWithAnimation(expression);
        probe->minimumObserver=make_self<BoundsObserver>();probe->minimum=InteractionTracker::CreateWithOwner(compositor,*probe->minimumObserver);
        probe->minimum.MinPosition({-100000000,-100000000,0});probe->minimum.MaxPosition({100000000,100000000,0});
        auto minExpression=compositor.CreateExpressionAnimation(L"Vector3(Tracker.MinPosition.X,Tracker.MinPosition.Y,0)");
        minExpression.SetReferenceParameter(L"Tracker",probe->source);probe->minimum.TryUpdatePositionWithAnimation(minExpression);
        probe->renderedObserver=make_self<BoundsObserver>();probe->rendered=InteractionTracker::CreateWithOwner(compositor,*probe->renderedObserver);
        probe->rendered.MinPosition({-100000000,-100000000,0});probe->rendered.MaxPosition({100000000,100000000,0});
        auto offsetExpression=compositor.CreateExpressionAnimation(L"Vector3(Output.Offset.X,Output.Offset.Y,0)");
        probe->output=QueryNative<Visual>(*reinterpret_cast<void**>(backend+0x3e0));
        offsetExpression.SetReferenceParameter(L"Output",probe->output);probe->rendered.TryUpdatePositionWithAnimation(offsetExpression);
        nativeProbes.emplace(identity,probe);
    } catch(hresult_error const& error) {
        FILE* file=nullptr;_wfopen_s(&file,logPath.c_str(),L"a");if(file){fprintf(file,"{\"event\":\"probe-error\",\"code\":%d}\n",static_cast<int>(error.code()));fclose(file);}
    }
}
static LRESULT NativeProbeMessage(UINT message,WPARAM sample) {
    std::vector<std::shared_ptr<NativeProbe>> copy;
    {std::lock_guard<std::mutex> lock(probeGate);for(auto& entry:nativeProbes)copy.push_back(entry.second);}
    if(message==WM_APP+81) {
        for(auto& probe:copy)probe->properties.InsertScalar(L"isProgrammaticFreeFloating",0);
        return copy.size();
    }
    if(message==WM_APP+82) {
        FILE* file=nullptr;_wfopen_s(&file,logPath.c_str(),L"a");if(!file)return 0;
        for(auto& probe:copy) {
            const auto content=probe->content.Size(),viewport=probe->viewport.Size();const auto bound=probe->observer->Value();
            const auto scale=probe->source.Scale();const auto minimum=probe->minimumObserver->Value(),rendered=probe->renderedObserver->Value();
            fprintf(file,"{\"event\":\"native-bound-probe\",\"pid\":%lu,\"sample\":%llu,\"id\":%u,\"content\":[%.4f,%.4f],\"viewport\":[%.4f,%.4f],\"scale\":%.4f,\"maximum\":[%.4f,%.4f],\"minimum\":[%.4f,%.4f],\"renderedOffset\":[%.4f,%.4f]}\n",GetCurrentProcessId(),static_cast<unsigned long long>(sample),probe->id,content.x,content.y,viewport.x,viewport.y,scale,bound.x,bound.y,minimum.x,minimum.y,rendered.x,rendered.y);
            const auto raw=probe->source.Position();float floating=0,enabled=0;probe->properties.TryGetScalar(L"isProgrammaticFreeFloating",floating);probe->properties.TryGetScalar(L"CodexEnabled",enabled);
            fprintf(file,"{\"event\":\"source-position-probe\",\"pid\":%lu,\"sample\":%llu,\"id\":%u,\"position\":[%.4f,%.4f],\"floating\":%.0f,\"enabled\":%.0f}\n",GetCurrentProcessId(),static_cast<unsigned long long>(sample),probe->id,raw.x,raw.y,floating,enabled);
            auto clip=probe->content.Clip().try_as<InsetClip>();
            if(clip)fprintf(file,"{\"event\":\"drawing-clip-probe\",\"pid\":%lu,\"sample\":%llu,\"id\":%u,\"content\":[%.4f,%.4f],\"insets\":[%.4f,%.4f,%.4f,%.4f]}\n",GetCurrentProcessId(),static_cast<unsigned long long>(sample),probe->id,content.x,content.y,clip.LeftInset(),clip.TopInset(),clip.RightInset(),clip.BottomInset());
            if(sample==1) {
                for(auto visual:{probe->content,probe->output,probe->viewport}) {
                    const auto offset=visual.Offset(),vscale=visual.Scale();const auto size=visual.Size();
                    auto container=visual.try_as<ContainerVisual>();
                    fprintf(file,"{\"event\":\"visual-tree-probe\",\"id\":%u,\"role\":\"%s\",\"size\":[%.3f,%.3f],\"offset\":[%.3f,%.3f],\"scale\":[%.3f,%.3f],\"clip\":%s,\"children\":%u}\n",probe->id,visual==probe->content?"content":visual==probe->output?"output":"viewport",size.x,size.y,offset.x,offset.y,vscale.x,vscale.y,visual.Clip()?"true":"false",container?container.Children().Count():0);
                }
            }
        }
        fclose(file);return copy.size();
    }
    if(message==WM_APP+83) {
        for(auto& probe:copy) {
            const auto content=probe->content.Size(),viewport=probe->viewport.Size();const auto minimum=probe->minimumObserver->Value();
            if(content.x>8000&&content.x<8500&&content.y>8000&&viewport.x!=2000&&(minimum.x>0||minimum.y>0))
                probe->source.TryUpdatePosition({minimum.x-100,minimum.y-100,0},InteractionTrackerClampingOption::Disabled);
        }
        return copy.size();
    }
    if(message==WM_APP+84) {
        for(auto& probe:copy) {
            const auto content=probe->content.Size(),viewport=probe->viewport.Size();const auto maximum=probe->observer->Value();
            if(content.x>8000&&content.x<8500&&content.y>8000&&viewport.x!=2000)
                probe->source.TryUpdatePosition({maximum.x+100,maximum.y+100,0},InteractionTrackerClampingOption::Disabled);
        }
        return copy.size();
    }
    if(message==WM_APP+86) {
        for(auto& probe:copy) {
            const auto content=probe->content.Size(),viewport=probe->viewport.Size();const auto maximum=probe->observer->Value();
            float enabled=0;probe->properties.TryGetScalar(L"CodexEnabled",enabled);
            if(enabled>0&&viewport.x>500&&viewport.y>400&&content.x>viewport.x&&content.y>viewport.y) {
                if(sample==0)probe->source.TryUpdatePosition({maximum.x*0.3f,maximum.y*0.4f,0},InteractionTrackerClampingOption::Disabled);
                if(sample==1)probe->source.TryUpdatePositionWithAdditionalVelocity({150,100,0});
                if(sample==2)probe->source.TryUpdatePosition({maximum.x+500,maximum.y+500,0},InteractionTrackerClampingOption::Disabled);
                if(sample==3)probe->source.TryUpdatePositionWithAdditionalVelocity({-150,-100,0});
                if(sample==4)probe->properties.InsertScalar(L"isProgrammaticFreeFloating",0);
                if(sample==5)probe->source.TryUpdatePosition({-100,-100,0},InteractionTrackerClampingOption::Disabled);
            }
        }
        return copy.size();
    }
    return 0;
}
