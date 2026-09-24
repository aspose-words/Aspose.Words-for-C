#pragma once

#include <gtest/gtest.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExPlainTextDocument : public ApiExampleBase
{
    typedef ExPlainTextDocument ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void Load();
    void LoadFromStream();
    void LoadEncrypted();
    void LoadEncryptedUsingStream();
    void BuiltInProperties();
    void CustomDocumentProperties();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


