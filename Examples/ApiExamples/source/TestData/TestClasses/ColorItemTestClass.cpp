#include "TestData/TestClasses/ColorItemTestClass.h"

namespace ApiExamples { namespace TestData { namespace TestClasses {

ColorItemTestClass::ColorItemTestClass(System::String name, System::Drawing::Color color, int32_t colorCode, double value1, double value2, double value3)
    : ColorCode(0), Value1(0), Value2(0), Value3(0)
{
    Name = name;
    Color = color;
    ColorCode = colorCode;
    Value1 = value1;
    Value2 = value2;
    Value3 = value3;
}

}}} // namespace ApiExamples::TestData::TestClasses
