#pragma once

#include <gtest/gtest.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExXpsSaveOptions : public ApiExampleBase
{
    typedef ExXpsSaveOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void OutlineLevels();
    void BookFold(bool renderTextAsBookFold);
    void ExportExactPages();
    void XpsDigitalSignature();
    void CompressionLevelXps();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


