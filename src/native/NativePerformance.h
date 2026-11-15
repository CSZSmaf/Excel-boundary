#pragma once
#ifdef CODEX_NATIVE_PROFILE
#include <atomic>
struct PerfCounter {std::atomic<unsigned long long> count{0},ticks{0},maximum{0};};
static PerfCounter nativePerf[5];
struct PerfScope {
    int index;LARGE_INTEGER start;
    PerfScope(int value):index(value){QueryPerformanceCounter(&start);}
    ~PerfScope(){LARGE_INTEGER end;QueryPerformanceCounter(&end);auto elapsed=static_cast<unsigned long long>(end.QuadPart-start.QuadPart);auto& counter=nativePerf[index];++counter.count;counter.ticks+=elapsed;auto old=counter.maximum.load();while(old<elapsed&&!counter.maximum.compare_exchange_weak(old,elapsed)){};}
};
#define NATIVE_PERF(index) PerfScope measuredScope(index)
static void WriteNativePerformance(unsigned phase) {
    if(!phase){for(auto& counter:nativePerf){counter.count=0;counter.ticks=0;counter.maximum=0;}return;}
    const auto fileName=logPath.substr(0,logPath.find_last_of(L"\\/"))+L"\\native-profile-"+std::to_wstring(phase)+L".json";
    LARGE_INTEGER frequency;QueryPerformanceFrequency(&frequency);
    const char* names[]={"stringReference","canvasSize","canvasUnwind","backendBind","applyCanvas"};
    FILE* file=nullptr;_wfopen_s(&file,fileName.c_str(),L"w");if(!file)return;
    fprintf(file,"{");for(int i=0;i<5;++i){const auto& counter=nativePerf[i];fprintf(file,"%s\"%s\":{\"calls\":%llu,\"totalMs\":%.6f,\"maxMs\":%.6f}",i?",":"",names[i],counter.count.load(),counter.ticks.load()*1000.0/frequency.QuadPart,counter.maximum.load()*1000.0/frequency.QuadPart);}fprintf(file,"}");fclose(file);
}
#else
#define NATIVE_PERF(index)
#endif
