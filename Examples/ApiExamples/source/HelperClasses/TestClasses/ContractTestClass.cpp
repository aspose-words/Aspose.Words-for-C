// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#include "HelperClasses/TestClasses/ContractTestClass.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

RTTI_INFO_IMPL_HASH(3752613209u, ::Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass, ThisTypeBaseTypesInfo);

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass> ContractTestClass::get_Manager() const
{
    return pr_Manager;
}

void ContractTestClass::set_Manager(System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass> value)
{
    pr_Manager = value;
}

System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass> ContractTestClass::get_Client() const
{
    return pr_Client;
}

void ContractTestClass::set_Client(System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass> value)
{
    pr_Client = value;
}

float ContractTestClass::get_Price() const
{
    return pr_Price;
}

void ContractTestClass::set_Price(float value)
{
    pr_Price = value;
}

System::DateTime ContractTestClass::get_Date() const
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

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
