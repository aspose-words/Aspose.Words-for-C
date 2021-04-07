#pragma once

#include <system/string.h>

using System::MakeObject;
using System::SharedPtr;
using System::String;

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {

//ExStart:Sender
class Sender : public System::Object
{
public:
    String get_Name();
    void set_Name(String value);
    String get_Message();
    void set_Message(String value);

private:
    String pr_Name;
    String pr_Message;
};

//ExEnd:Sender
}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects
