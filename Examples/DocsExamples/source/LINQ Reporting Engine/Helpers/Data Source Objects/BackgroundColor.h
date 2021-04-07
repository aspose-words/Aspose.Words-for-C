#pragma once

#include <cstdint>
#include <drawing/color.h>
#include <system/nullable.h>
#include <system/string.h>

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {

//ExStart:Color
class BackgroundColor : public System::Object
{
public:
    String get_Name();
    void set_Name(String value);
    System::Drawing::Color get_Color();
    void set_Color(System::Drawing::Color value);
    System::Nullable<int> get_ColorCode();
    void set_ColorCode(System::Nullable<int> value);
    System::Nullable<double> get_Value1();
    void set_Value1(System::Nullable<double> value);
    System::Nullable<double> get_Value2();
    void set_Value2(System::Nullable<double> value);
    System::Nullable<double> get_Value3();
    void set_Value3(System::Nullable<double> value);

private:
    String pr_Name;
    System::Drawing::Color pr_Color;
    System::Nullable<int> pr_ColorCode;
    System::Nullable<double> pr_Value1;
    System::Nullable<double> pr_Value2;
    System::Nullable<double> pr_Value3;
};

//ExEnd:Color
}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects
