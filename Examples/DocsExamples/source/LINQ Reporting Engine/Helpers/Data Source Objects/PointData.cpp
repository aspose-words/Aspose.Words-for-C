#include "LINQ Reporting Engine/Helpers/Data Source Objects/PointData.h"

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {

System::String PointData::get_Time()
{
    return pr_Time;
}

void PointData::set_Time(System::String value)
{
    pr_Time = value;
}

int32_t PointData::get_Flow()
{
    return pr_Flow;
}

void PointData::set_Flow(int32_t value)
{
    pr_Flow = value;
}

int32_t PointData::get_Rainfall()
{
    return pr_Rainfall;
}

void PointData::set_Rainfall(int32_t value)
{
    pr_Rainfall = value;
}

PointData::PointData() : pr_Flow(0), pr_Rainfall(0)
{
}

}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects
