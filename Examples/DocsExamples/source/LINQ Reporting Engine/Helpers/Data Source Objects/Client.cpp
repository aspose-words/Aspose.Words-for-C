#include "LINQ Reporting Engine/Helpers/Data Source Objects/Client.h"

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {

System::String Client::get_Name()
{
    return pr_Name;
}

void Client::set_Name(System::String value)
{
    pr_Name = value;
}

System::String Client::get_Country()
{
    return pr_Country;
}

void Client::set_Country(System::String value)
{
    pr_Country = value;
}

System::String Client::get_LocalAddress()
{
    return pr_LocalAddress;
}

void Client::set_LocalAddress(System::String value)
{
    pr_LocalAddress = value;
}

}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects
