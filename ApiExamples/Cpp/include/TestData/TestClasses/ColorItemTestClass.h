#pragma once

#include <system/string.h>
#include <system/object.h>
#include <drawing/color.h>
#include <cstdint>

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

class ColorItemTestClass : public System::Object
{
public:

    System::String Name;
    System::Drawing::Color Color;
    int32_t ColorCode;
    double Value1;
    double Value2;
    double Value3;

    ColorItemTestClass(System::String name, System::Drawing::Color color, int32_t colorCode, double value1, double value2, double value3);

};

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples

