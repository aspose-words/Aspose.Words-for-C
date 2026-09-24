#pragma once

#include <gtest/gtest.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExDocSaveOptions : public ApiExampleBase
{
    typedef ExDocSaveOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void SaveAsDoc();
    void TempFolder();
    void PictureBullets();
    void UpdateLastPrintedProperty(bool isUpdateLastPrintedProperty);
    void UpdateCreatedTimeProperty(bool isUpdateCreatedTimeProperty);
    void AlwaysCompressMetafiles(bool compressAllMetafiles);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


