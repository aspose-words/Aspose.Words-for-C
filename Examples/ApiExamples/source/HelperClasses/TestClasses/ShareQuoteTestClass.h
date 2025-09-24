#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.

#include <system/string.h>
#include <cstdint>

#include "HelperClasses/Common.h"

namespace Aspose
{
namespace Words
{
namespace ApiExamples
{
namespace TestData
{
class Common;
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose

namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace HelperClasses {

namespace TestClasses {

class ShareQuoteTestClass : public System::Object
{
    typedef ShareQuoteTestClass ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
    friend class TestData::Common;
    
public:

    int32_t Date;
    int32_t Volume;
    double Open;
    double High;
    double Low;
    double Close;
    
    ShareQuoteTestClass(int32_t date, int32_t volume, double open, double high, double low, double close);
    
    System::String Color();
    
};

} // namespace TestClasses
} // namespace HelperClasses
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


