#pragma once

#include <gtest/gtest.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExRtfSaveOptions : public ApiExampleBase
{
    typedef ExRtfSaveOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void ExportImages(bool exportImagesForOldReaders);
    void SaveImagesAsWmf(bool saveImagesAsWmf);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


