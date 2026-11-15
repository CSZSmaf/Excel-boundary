#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <string>
#include <cstdio>
int wmain(){
    const std::wstring desktopName=L"ExcelBoundaryLab_"+std::to_wstring(GetCurrentProcessId());
    HDESK desktop=CreateDesktopW(desktopName.c_str(),nullptr,nullptr,0,GENERIC_ALL,nullptr);
    if(!desktop)return 1;
    wchar_t file[MAX_PATH];GetModuleFileNameW(nullptr,file,MAX_PATH);
    const auto folder=std::wstring(file).substr(0,std::wstring(file).find_last_of(L'\\'));
    std::wstring command=L"\""+folder+L"\\ExcelIntegration.exe\"";
    std::wstring fullDesktop=L"WinSta0\\"+desktopName;
    STARTUPINFOW startup{sizeof(startup)};startup.lpDesktop=fullDesktop.data();startup.dwFlags=STARTF_USESHOWWINDOW;startup.wShowWindow=SW_HIDE;
    PROCESS_INFORMATION process{};
    if(!CreateProcessW(nullptr,command.data(),nullptr,nullptr,FALSE,CREATE_NO_WINDOW,nullptr,folder.c_str(),&startup,&process)){CloseDesktop(desktop);return 2;}
    printf("Excel test started on a private desktop; no desktop switch or input injection.\n");fflush(stdout);
    const auto wait=WaitForSingleObject(process.hProcess,55000);
    DWORD code=3;
    if(wait==WAIT_OBJECT_0)GetExitCodeProcess(process.hProcess,&code);
    else fprintf(stderr,"Integration test exceeded its 55-second deadline.\n");
    CloseHandle(process.hThread);CloseHandle(process.hProcess);CloseDesktop(desktop);
    return static_cast<int>(code);
}
