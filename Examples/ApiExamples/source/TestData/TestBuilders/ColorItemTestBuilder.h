#pragma once

#include <cstdint>
#include <drawing/color.h>
#include <system/string.h>

#include "TestData/TestClasses/ColorItemTestClass.h"

using namespace ApiExamples::TestData::TestClasses;

namespace ApiExamples { namespace TestData { namespace TestBuilders {

class ColorItemTestBuilder : public System::Object
{
public:
    System::String Name;
    System::Drawing::Color Color;
    int32_t ColorCode;
    double Value1;
    double Value2;
    double Value3;

    ColorItemTestBuilder();

    System::SharedPtr<ColorItemTestBuilder> WithColor(System::String name, System::Drawing::Color color);
    System::SharedPtr<ColorItemTestBuilder> WithColorCode(System::String name, int32_t colorCode);
    System::SharedPtr<ColorItemTestBuilder> WithColorAndValues(System::String name, System::Drawing::Color color, double value1, double value2, double value3);
    System::SharedPtr<ColorItemTestBuilder> WithColorCodeAndValues(System::String name, int32_t colorCode, double value1, double value2, double value3);
    System::SharedPtr<ColorItemTestClass> Build();

protected:
    virtual ~ColorItemTestBuilder();
};

}}} // namespace ApiExamples::TestData::TestBuilders
