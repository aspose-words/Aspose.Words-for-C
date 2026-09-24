#pragma once

#include <gtest/gtest.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExTabStop : public ApiExampleBase
{
    typedef ExTabStop ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void AddTabStops();
    void TabStopCollection();
    void RemoveByIndex();
    void GetPositionByIndex();
    void GetIndexByPosition();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


