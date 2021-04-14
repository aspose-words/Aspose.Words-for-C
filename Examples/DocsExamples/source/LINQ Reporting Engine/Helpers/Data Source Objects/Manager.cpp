#include "LINQ Reporting Engine/Helpers/Data Source Objects/Manager.h"

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {

System::String Manager::get_Name()
{
    return pr_Name;
}

void Manager::set_Name(System::String value)
{
    pr_Name = value;
}

int32_t Manager::get_Age()
{
    return pr_Age;
}

void Manager::set_Age(int32_t value)
{
    pr_Age = value;
}

System::ArrayPtr<uint8_t> Manager::get_Photo()
{
    return pr_Photo;
}

void Manager::set_Photo(System::ArrayPtr<uint8_t> value)
{
    pr_Photo = value;
}

System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Contract>>> Manager::get_Contracts()
{
    return pr_Contracts;
}

void Manager::set_Contracts(System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Contract>>> value)
{
    pr_Contracts = value;
}

Manager::Manager() : pr_Age(0)
{
}

}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects
