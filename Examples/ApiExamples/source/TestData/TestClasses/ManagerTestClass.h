#pragma once

#include <cstdint>
#include <system/collections/ienumerable.h>
#include <system/string.h>

#include "TestData/TestClasses/ContractTestClass.h"

namespace ApiExamples { namespace TestData { namespace TestClasses {
class ContractTestClass;
}}} // namespace ApiExamples::TestData::TestClasses

namespace ApiExamples { namespace TestData { namespace TestClasses {

class ManagerTestClass : public System::Object
{
public:
    System::String get_Name();
    void set_Name(System::String value);
    int32_t get_Age();
    void set_Age(int32_t value);
    System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<ContractTestClass>>> get_Contracts();
    void set_Contracts(System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<ContractTestClass>>> value);

    ManagerTestClass();

private:
    System::String pr_Name;
    int32_t pr_Age;
    System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<ContractTestClass>>> pr_Contracts;
};

}}} // namespace ApiExamples::TestData::TestClasses
