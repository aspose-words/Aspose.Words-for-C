#pragma once

#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Text/DropCapPosition.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExParagraphFormat : public ApiExampleBase
{
    typedef ExParagraphFormat ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void AsianTypographyProperties();
    void DropCap(Aspose::Words::DropCapPosition dropCapPosition);
    void LineSpacing();
    void ParagraphSpacingAuto(bool autoSpacing);
    void ParagraphSpacingSameStyle(bool noSpaceBetweenParagraphsOfSameStyle);
    void ParagraphOutlineLevel();
    void PageBreakBefore(bool pageBreakBefore);
    void WidowControl(bool widowControl);
    void LinesToDrop();
    void SuppressHyphens(bool suppressAutoHyphens);
    void ParagraphSpacingAndIndents();
    void ParagraphBaselineAlignment();
    void MirrorIndents();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


