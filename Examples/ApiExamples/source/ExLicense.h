#pragma once

#include <gtest/gtest.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExLicense : public ApiExampleBase
{
    typedef ExLicense ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void LicenseFromFileNoPath();
    void LicenseFromStream();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


