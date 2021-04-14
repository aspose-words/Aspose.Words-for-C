#pragma once

#include <cstdint>
#include <system/array.h>
#include <system/collections/ienumerable.h>

#include "LINQ Reporting Engine/Helpers/Data Source Objects/Contract.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {
class Contract;
}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects

namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers { namespace Data_Source_Objects {

//ExStart:Manager
class Manager : public System::Object
{
public:
    String get_Name();
    void set_Name(String value);
    int get_Age();
    void set_Age(int value);
    ArrayPtr<uint8_t> get_Photo();
    void set_Photo(ArrayPtr<uint8_t> value);
    SharedPtr<System::Collections::Generic::IEnumerable<SharedPtr<Contract>>> get_Contracts();
    void set_Contracts(SharedPtr<System::Collections::Generic::IEnumerable<SharedPtr<Contract>>> value);

    Manager();

private:
    String pr_Name;
    int pr_Age;
    ArrayPtr<uint8_t> pr_Photo;
    SharedPtr<System::Collections::Generic::IEnumerable<SharedPtr<Contract>>> pr_Contracts;
};

//ExEnd:Manager
}}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects
