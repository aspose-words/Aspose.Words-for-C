#include "TestData/TestClasses/ContractTestClass.h"

namespace ApiExamples { namespace TestData { namespace TestClasses {

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

}}} // namespace ApiExamples::TestData::TestClasses
