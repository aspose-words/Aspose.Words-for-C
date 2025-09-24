#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.

#include <system/string.h>
#include <drawing/color.h>
#include <cstdint>

#include "HelperClasses/TestClasses/ColorItemTestClass.h"


using namespace Aspose::Words::ApiExamples::TestData::TestClasses;

namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

namespace TestBuilders {

class ColorItemTestBuilder : public System::Object
{
    typedef ColorItemTestBuilder ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    System::String Name;
    System::Drawing::Color Color;
    int32_t ColorCode;
    double Value1;
    double Value2;
    double Value3;
    
    ColorItemTestBuilder();
    
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ColorItemTestBuilder> WithColor(System::String name, System::Drawing::Color color);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ColorItemTestBuilder> WithColorCode(System::String name, int32_t colorCode);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ColorItemTestBuilder> WithColorAndValues(System::String name, System::Drawing::Color color, double value1, double value2, double value3);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ColorItemTestBuilder> WithColorCodeAndValues(System::String name, int32_t colorCode, double value1, double value2, double value3);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ColorItemTestClass> Build();
    
protected:

    virtual ~ColorItemTestBuilder();
    
};

} // namespace TestBuilders
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


