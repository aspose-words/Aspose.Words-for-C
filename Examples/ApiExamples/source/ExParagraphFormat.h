#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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


