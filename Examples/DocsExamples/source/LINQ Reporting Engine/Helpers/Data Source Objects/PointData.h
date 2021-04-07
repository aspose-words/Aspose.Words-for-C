#pragma once

#include <cstdint>
#include <system/string.h>

using System::MakeObject;
using System::SharedPtr;
using System::String;

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {

//ExStart:PointDataClass
class PointData : public System::Object
{
public:
    String get_Time();
    void set_Time(String value);
    int get_Flow();
    void set_Flow(int value);
    int get_Rainfall();
    void set_Rainfall(int value);

    PointData();

private:
    String pr_Time;
    int pr_Flow;
    int pr_Rainfall;
};

//ExEnd:PointDataClass
}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects
