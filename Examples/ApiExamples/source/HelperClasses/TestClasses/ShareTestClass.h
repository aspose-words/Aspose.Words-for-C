#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.

#include <system/string.h>

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

class ShareTestClass : public System::Object
{
    typedef ShareTestClass ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
    friend class TestData::Common;
    
public:

    System::String Sector;
    System::String Industry;
    System::String Ticker;
    double Weight;
    double Delta;
    
    ShareTestClass(System::String sector, System::String industry, System::String ticker, double weight, double delta);
    
    System::String Title();
    System::String Color();
    System::String IndustryColor();
    
};

} // namespace TestClasses
} // namespace HelperClasses
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


