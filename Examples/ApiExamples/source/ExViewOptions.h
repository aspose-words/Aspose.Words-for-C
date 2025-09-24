#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <Aspose.Words.Cpp/Model/Settings/ZoomType.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Settings;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExViewOptions : public ApiExampleBase
{
    typedef ExViewOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void SetZoomPercentage();
    void SetZoomType(Aspose::Words::Settings::ZoomType zoomType);
    void DisplayBackgroundShape(bool displayBackgroundShape);
    void DisplayPageBoundaries(bool doNotDisplayPageBoundaries);
    void FormsDesign(bool useFormsDesign);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


