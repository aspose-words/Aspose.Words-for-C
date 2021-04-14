#pragma once

#include <system/string.h>

using System::MakeObject;
using System::SharedPtr;
using System::String;

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {

//ExStart:Client
class Client : public System::Object
{
public:
    String get_Name();
    void set_Name(String value);
    String get_Country();
    void set_Country(String value);
    String get_LocalAddress();
    void set_LocalAddress(String value);

private:
    String pr_Name;
    String pr_Country;
    String pr_LocalAddress;
};

//ExEnd:Client
}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects
