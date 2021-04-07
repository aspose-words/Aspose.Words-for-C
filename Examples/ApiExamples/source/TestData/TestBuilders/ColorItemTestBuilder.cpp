#include "TestData/TestBuilders/ColorItemTestBuilder.h"

using namespace ApiExamples::TestData::TestClasses;
namespace ApiExamples { namespace TestData { namespace TestBuilders {

ColorItemTestBuilder::ColorItemTestBuilder() : ColorCode(0), Value1(0), Value2(0), Value3(0)
{
    Name = u"DefaultName";
    Color = System::Drawing::Color::get_Black();
    ColorCode = System::Drawing::Color::get_Black().ToArgb();
    Value1 = 1.0;
    Value2 = 1.0;
    Value3 = 1.0;
}

System::SharedPtr<ColorItemTestBuilder> ColorItemTestBuilder::WithColor(System::String name, System::Drawing::Color color)
{
    Name = name;
    Color = color;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<ColorItemTestBuilder> ColorItemTestBuilder::WithColorCode(System::String name, int32_t colorCode)
{
    Name = name;
    ColorCode = colorCode;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<ColorItemTestBuilder> ColorItemTestBuilder::WithColorAndValues(System::String name, System::Drawing::Color color, double value1,
                                                                                 double value2, double value3)
{
    Name = name;
    Color = color;
    Value1 = value1;
    Value2 = value2;
    Value3 = value3;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<ColorItemTestBuilder> ColorItemTestBuilder::WithColorCodeAndValues(System::String name, int32_t colorCode, double value1, double value2,
                                                                                     double value3)
{
    Name = name;
    ColorCode = colorCode;
    Value1 = value1;
    Value2 = value2;
    Value3 = value3;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<ColorItemTestClass> ColorItemTestBuilder::Build()
{
    return System::MakeObject<ColorItemTestClass>(Name, Color, ColorCode, Value1, Value2, Value3);
}

ColorItemTestBuilder::~ColorItemTestBuilder()
{
}

}}} // namespace ApiExamples::TestData::TestBuilders
