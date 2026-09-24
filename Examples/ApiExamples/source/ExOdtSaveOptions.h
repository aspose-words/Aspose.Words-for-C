#pragma once

#include <gtest/gtest.h>
#include <Aspose.Words.Cpp/Model/Document/SaveFormat.h>

#include "ApiExampleBase.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExOdtSaveOptions : public ApiExampleBase
{
    typedef ExOdtSaveOptions ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    void Odt11Schema(bool exportToOdt11Specs);
    void Encrypt(Aspose::Words::SaveFormat saveFormat);
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


