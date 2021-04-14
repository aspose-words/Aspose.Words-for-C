#include "TestData/TestClasses/ClientTestClass.h"

namespace ApiExamples { namespace TestData { namespace TestClasses {

System::String ClientTestClass::get_Name()
{
    return pr_Name;
}

void ClientTestClass::set_Name(System::String value)
{
    pr_Name = value;
}

System::String ClientTestClass::get_Country()
{
    return pr_Country;
}

void ClientTestClass::set_Country(System::String value)
{
    pr_Country = value;
}

System::String ClientTestClass::get_LocalAddress()
{
    return pr_LocalAddress;
}

void ClientTestClass::set_LocalAddress(System::String value)
{
    pr_LocalAddress = value;
}

}}} // namespace ApiExamples::TestData::TestClasses
