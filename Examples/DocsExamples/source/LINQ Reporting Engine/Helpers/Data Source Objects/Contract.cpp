#include "LINQ Reporting Engine/Helpers/Data Source Objects/Contract.h"

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {

System::SharedPtr<DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects::Manager> Contract::get_Manager()
{
    return pr_Manager;
}

void Contract::set_Manager(System::SharedPtr<DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects::Manager> value)
{
    pr_Manager = value;
}

System::SharedPtr<DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects::Client> Contract::get_Client()
{
    return pr_Client;
}

void Contract::set_Client(System::SharedPtr<DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects::Client> value)
{
    pr_Client = value;
}

float Contract::get_Price()
{
    return pr_Price;
}

void Contract::set_Price(float value)
{
    pr_Price = value;
}

System::DateTime Contract::get_Date()
{
    return pr_Date;
}

void Contract::set_Date(System::DateTime value)
{
    pr_Date = value;
}

Contract::Contract() : pr_Price(0)
{
}

}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects
