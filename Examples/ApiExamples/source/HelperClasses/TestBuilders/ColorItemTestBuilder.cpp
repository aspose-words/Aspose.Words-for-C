// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#include "HelperClasses/TestBuilders/ColorItemTestBuilder.h"


using namespace Aspose::Words::ApiExamples::TestData::TestClasses;
namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

namespace TestBuilders {

RTTI_INFO_IMPL_HASH(3647658262u, ::Aspose::Words::ApiExamples::TestData::TestBuilders::ColorItemTestBuilder, ThisTypeBaseTypesInfo);

ColorItemTestBuilder::ColorItemTestBuilder() : ColorCode(0), Value1(0), Value2(0), Value3(0)
{
    Name = u"DefaultName";
    Color = System::Drawing::Color::get_Black();
    ColorCode = System::Drawing::Color::get_Black().ToArgb();
    Value1 = 1.0;
    Value2 = 1.0;
    Value3 = 1.0;
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ColorItemTestBuilder> ColorItemTestBuilder::WithColor(System::String name, System::Drawing::Color color)
{
    Name = name;
    Color = color;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ColorItemTestBuilder> ColorItemTestBuilder::WithColorCode(System::String name, int32_t colorCode)
{
    Name = name;
    ColorCode = colorCode;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ColorItemTestBuilder> ColorItemTestBuilder::WithColorAndValues(System::String name, System::Drawing::Color color, double value1, double value2, double value3)
{
    Name = name;
    Color = color;
    Value1 = value1;
    Value2 = value2;
    Value3 = value3;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestBuilders::ColorItemTestBuilder> ColorItemTestBuilder::WithColorCodeAndValues(System::String name, int32_t colorCode, double value1, double value2, double value3)
{
    Name = name;
    ColorCode = colorCode;
    Value1 = value1;
    Value2 = value2;
    Value3 = value3;
    return System::MakeSharedPtr(this);
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ColorItemTestClass> ColorItemTestBuilder::Build()
{
    return System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ColorItemTestClass>(Name, Color, ColorCode, Value1, Value2, Value3);
}

ColorItemTestBuilder::~ColorItemTestBuilder()
{
}

} // namespace TestBuilders
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
