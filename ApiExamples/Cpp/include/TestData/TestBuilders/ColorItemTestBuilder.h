#pragma once

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <drawing/color.h>
#include <cstdint>

namespace ApiExamples { namespace TestData { namespace TestClasses { class ColorItemTestClass; } } }

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
    
    System::SharedPtr<ColorItemTestBuilder> WithColor(System::String name, System::Drawing::Color color);
    System::SharedPtr<ColorItemTestBuilder> WithColorCode(System::String name, int32_t colorCode);
    System::SharedPtr<ColorItemTestBuilder> WithColorAndValues(System::String name, System::Drawing::Color color, double value1, double value2, double value3);
    System::SharedPtr<ColorItemTestBuilder> WithColorCodeAndValues(System::String name, int32_t colorCode, double value1, double value2, double value3);
    System::SharedPtr<ApiExamples::TestData::TestClasses::ColorItemTestClass> Build();
    
protected:

    virtual ~ColorItemTestBuilder();
    
    System::Object::shared_members_type GetSharedMembers() override;
    
};

} // namespace TestBuilders
} // namespace TestData
} // namespace ApiExamples


