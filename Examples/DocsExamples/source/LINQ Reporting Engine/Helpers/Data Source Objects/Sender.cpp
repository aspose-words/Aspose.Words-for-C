#include "LINQ Reporting Engine/Helpers/Data Source Objects/Sender.h"

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {

System::String Sender::get_Name()
{
    return pr_Name;
}

void Sender::set_Name(System::String value)
{
    pr_Name = value;
}

System::String Sender::get_Message()
{
    return pr_Message;
}

void Sender::set_Message(System::String value)
{
    pr_Message = value;
}

}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects
