#pragma once

#include <gtest/gtest.h>
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


