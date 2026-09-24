#pragma once

#include <gtest/gtest.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExCleanupOptions : public ApiExampleBase
{
    typedef ExCleanupOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void RemoveUnusedResources();
    void RemoveDuplicateStyles();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


