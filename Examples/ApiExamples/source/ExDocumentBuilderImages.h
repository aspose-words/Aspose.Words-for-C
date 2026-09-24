#pragma once

#include <gtest/gtest.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExDocumentBuilderImages : public ApiExampleBase
{
    typedef ExDocumentBuilderImages ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void InsertImageFromStream();
    void InsertImageFromFilename();
    void InsertSvgImage();
    void InsertImageFromImageObject();
    void InsertImageFromByteArray();
    void InsertGif();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


