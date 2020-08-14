#include "TestData/TestClasses/ContractTestClass.h"

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

RTTI_INFO_IMPL_HASH(1659307295u, ::ApiExamples::TestData::TestClasses::ContractTestClass, ThisTypeBaseTypesInfo);

System::SharedPtr<ManagerTestClass> ContractTestClass::get_Manager()
{
    return pr_Manager;
}

void ContractTestClass::set_Manager(System::SharedPtr<ManagerTestClass> value)
{
    pr_Manager = value;
}

System::SharedPtr<ClientTestClass> ContractTestClass::get_Client()
{
    return pr_Client;
}

void ContractTestClass::set_Client(System::SharedPtr<ClientTestClass> value)
{
    pr_Client = value;
}

float ContractTestClass::get_Price()
{
    return pr_Price;
}

void ContractTestClass::set_Price(float value)
{
    pr_Price = value;
}

System::DateTime ContractTestClass::get_Date()
{
    return pr_Date;
}

void ContractTestClass::set_Date(System::DateTime value)
{
    pr_Date = value;
}

ContractTestClass::ContractTestClass() : pr_Price(0)
{
}

System::Object::shared_members_type ApiExamples::TestData::TestClasses::ContractTestClass::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::TestData::TestClasses::ContractTestClass::pr_Manager", this->pr_Manager);
    result.Add("ApiExamples::TestData::TestClasses::ContractTestClass::pr_Client", this->pr_Client);
    result.Add("ApiExamples::TestData::TestClasses::ContractTestClass::pr_Date", this->pr_Date);

    return result;
}

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
