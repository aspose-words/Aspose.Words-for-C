#pragma once

#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Loading/TxtTrailingSpacesOptions.h>
#include <Aspose.Words.Cpp/Model/Loading/TxtLeadingSpacesOptions.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Loading;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExTxtLoadOptions : public ApiExampleBase
{
    typedef ExTxtLoadOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void DetectNumberingWithWhitespaces(bool detectNumberingWithWhitespaces);
    void TrailSpaces(Aspose::Words::Loading::TxtLeadingSpacesOptions txtLeadingSpacesOptions, Aspose::Words::Loading::TxtTrailingSpacesOptions txtTrailingSpacesOptions);
    void DetectDocumentDirection();
    void AutoNumberingDetection();
    void DetectHyperlinks();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


