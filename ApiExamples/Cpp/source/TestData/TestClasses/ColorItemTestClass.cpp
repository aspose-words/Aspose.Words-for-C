#include "TestData/TestClasses/ColorItemTestClass.h"

#include <system/string.h>
#include <drawing/color.h>
#include <cstdint>

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

RTTI_INFO_IMPL_HASH(1008333637u, ::ApiExamples::TestData::TestClasses::ColorItemTestClass, ThisTypeBaseTypesInfo);

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

System::Object::shared_members_type ApiExamples::TestData::TestClasses::ColorItemTestClass::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::TestData::TestClasses::ColorItemTestClass::Color", this->Color);

    return result;
}

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
