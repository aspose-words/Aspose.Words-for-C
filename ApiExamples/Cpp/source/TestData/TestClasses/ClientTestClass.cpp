#include "TestData/TestClasses/ClientTestClass.h"

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

RTTI_INFO_IMPL_HASH(2046037350u, ::ApiExamples::TestData::TestClasses::ClientTestClass, ThisTypeBaseTypesInfo);

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

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
