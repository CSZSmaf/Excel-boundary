#pragma once
#include <string>

inline std::wstring ClampTrackerPositions(std::wstring expression) {
    for (const wchar_t axis : {L'X', L'Y'}) {
        const std::wstring member = std::wstring(L"Tracker.Position.") + axis;
        std::wstring clamped = L"Clamp(" + member + L", Tracker.MinPosition." + axis +
            L", Tracker.MaxPosition." + axis + L")";
#ifdef CODEX_NATIVE_ADDIN
        clamped=L"(Tracker.CodexEnabled > 0 ? "+clamped+L" : "+member+L")";
#endif
        size_t offset = 0;
        while ((offset = expression.find(member, offset)) != std::wstring::npos) {
            expression.replace(offset, member.size(), clamped);
            offset += clamped.size();
        }
    }
    return expression;
}

inline std::wstring ContentMaximum(std::wstring expression) {
    for(const wchar_t axis:{L'X',L'Y'}) {
        const wchar_t lower=axis==L'X'?L'x':L'y';
        const std::wstring position=std::wstring(L"this.target.position.")+lower;
        const std::wstring size=std::wstring(L"(Content.Size.")+axis+L" * this.target.Scale)";
        const std::wstring original=position+L" > "+size+L" ? "+position+L" + ScrollingLayer.freeFloatingPadding : "+size+L" + ScrollingLayer.freeFloatingPadding";
        const std::wstring viewportFactor=axis==L'Y'?L" * 0.5":L"";
        std::wstring replacement=L"Max(ScrollingLayer.CodexFrozen"+std::wstring(1,axis)+L" * this.target.Scale, Content.Size."+axis+L" * this.target.Scale - Viewport.Size."+axis+viewportFactor+L")";
#ifdef CODEX_NATIVE_ADDIN
        replacement=L"(ScrollingLayer.CodexEnabled > 0 ? "+replacement+L" : ("+original+L"))";
#endif
        size_t offset=0;
        while((offset=expression.find(original,offset))!=std::wstring::npos){expression.replace(offset,original.size(),replacement);offset+=replacement.size();}
    }
    return expression;
}

inline std::wstring FrozenMinimum(std::wstring expression) {
    for(const wchar_t axis:{L'X',L'Y'}) {
        const std::wstring position=std::wstring(L"this.target.position.")+(axis==L'X'?L'x':L'y');
        const std::wstring original=L"("+position+L" > 0 ? -ScrollingLayer.freeFloatingPadding : "+position+L" - ScrollingLayer.freeFloatingPadding) * ScrollingLayer.isProgrammaticFreeFloating";
        const std::wstring replacement=L"("+original+L" + (1 - ScrollingLayer.isProgrammaticFreeFloating) * "
#ifdef CODEX_NATIVE_ADDIN
            L"ScrollingLayer.CodexEnabled * "
#endif
            L"ScrollingLayer.CodexFrozen"+axis+L" * this.target.Scale)";
        size_t offset=0;
        while((offset=expression.find(original,offset))!=std::wstring::npos){expression.replace(offset,original.size(),replacement);offset+=replacement.size();}
    }
    return expression;
}
