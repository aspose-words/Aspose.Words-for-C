#include "LINQ Reporting Engine/Helpers/Data Source Objects/BackgroundColor.h"

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {

System::String BackgroundColor::get_Name()
{
    return pr_Name;
}

void BackgroundColor::set_Name(System::String value)
{
    pr_Name = value;
}

System::Drawing::Color BackgroundColor::get_Color()
{
    return pr_Color;
}

void BackgroundColor::set_Color(System::Drawing::Color value)
{
    pr_Color = value;
}

System::Nullable<int32_t> BackgroundColor::get_ColorCode()
{
    return pr_ColorCode;
}

void BackgroundColor::set_ColorCode(System::Nullable<int32_t> value)
{
    pr_ColorCode = value;
}

System::Nullable<double> BackgroundColor::get_Value1()
{
    return pr_Value1;
}

void BackgroundColor::set_Value1(System::Nullable<double> value)
{
    pr_Value1 = value;
}

System::Nullable<double> BackgroundColor::get_Value2()
{
    return pr_Value2;
}

void BackgroundColor::set_Value2(System::Nullable<double> value)
{
    pr_Value2 = value;
}

System::Nullable<double> BackgroundColor::get_Value3()
{
    return pr_Value3;
}

void BackgroundColor::set_Value3(System::Nullable<double> value)
{
    pr_Value3 = value;
}

}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects
