#include "../src/native/BoundaryExpression.h"
#include <iostream>
#include <stdexcept>
static void check(bool value,const char* name){if(!value)throw std::runtime_error(name);}
int main(){
    for(const wchar_t axis:{L'X',L'Y'}){
        const auto suffix=std::wstring(1,axis);
        const auto offset=ClampTrackerPositions(L"Tracker.IsPositionRoundingSuggested ? Round(-Tracker.Position."+suffix+L") : (-Tracker.Position."+suffix+L")");
        check(offset.find(L"Tracker.CodexEnabled > 0")!=std::wstring::npos,"scoped offset");
        check(offset.find(L"Clamp(Tracker.Position."+suffix+L", Tracker.MinPosition."+suffix+L", Tracker.MaxPosition."+suffix+L")")!=std::wstring::npos,"native offset bounds");
        check(offset.find(L"IsPositionRoundingSuggested ? Round(")!=std::wstring::npos,"native rounding policy");
        const auto lower=std::wstring(1,axis==L'X'?L'x':L'y');
        const auto p=L"this.target.position."+lower;
        const auto original=p+L" > (Content.Size."+suffix+L" * this.target.Scale) ? "+p+L" + ScrollingLayer.freeFloatingPadding : (Content.Size."+suffix+L" * this.target.Scale) + ScrollingLayer.freeFloatingPadding";
        const auto maximum=ContentMaximum(original);
        check(maximum.find(L"ScrollingLayer.CodexEnabled > 0")!=std::wstring::npos,"scoped maximum");
        check(maximum.find(original)!=std::wstring::npos,"inactive maximum preserved");
        check((maximum.find(L"Viewport.Size.Y * 0.5")!=std::wstring::npos)==(axis==L'Y'),"centered lower boundary");
        const auto minimum=FrozenMinimum(L"("+p+L" > 0 ? -ScrollingLayer.freeFloatingPadding : "+p+L" - ScrollingLayer.freeFloatingPadding) * ScrollingLayer.isProgrammaticFreeFloating");
        check(minimum.find(L"ScrollingLayer.CodexEnabled * ScrollingLayer.CodexFrozen"+suffix)!=std::wstring::npos,"frozen-pane minimum");
    }
    check(ClampTrackerPositions(L"Header.Width")==L"Header.Width","unrelated expressions unchanged");
    std::cout<<"Native expression tests passed.\n";
}
