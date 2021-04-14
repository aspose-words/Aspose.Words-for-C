#pragma once

#include <system/date_time.h>
#include <system/object.h>

#include "LINQ Reporting Engine/Helpers/Data Source Objects/Client.h"
#include "LINQ Reporting Engine/Helpers/Data Source Objects/Manager.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {
class Manager;
}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {

//ExStart:Contract
class Contract : public System::Object
{
public:
    SharedPtr<DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects::Manager> get_Manager();
    void set_Manager(SharedPtr<DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects::Manager> value);
    SharedPtr<DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects::Client> get_Client();
    void set_Client(SharedPtr<DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects::Client> value);
    float get_Price();
    void set_Price(float value);
    System::DateTime get_Date();
    void set_Date(System::DateTime value);

    Contract();

private:
    SharedPtr<DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects::Manager> pr_Manager;
    SharedPtr<DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects::Client> pr_Client;
    float pr_Price;
    System::DateTime pr_Date;
};

//ExEnd:Contract
}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects
