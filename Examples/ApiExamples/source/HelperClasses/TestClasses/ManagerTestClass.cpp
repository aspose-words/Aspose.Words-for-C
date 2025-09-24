// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#include "HelperClasses/TestClasses/ManagerTestClass.h"

namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

RTTI_INFO_IMPL_HASH(3617414548u, ::Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass, ThisTypeBaseTypesInfo);

System::String ManagerTestClass::get_Name() const
{
    return pr_Name;
}

void ManagerTestClass::set_Name(System::String value)
{
    pr_Name = value;
}

int32_t ManagerTestClass::get_Age() const
{
    return pr_Age;
}

void ManagerTestClass::set_Age(int32_t value)
{
    pr_Age = value;
}

System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>>> ManagerTestClass::get_Contracts() const
{
    return pr_Contracts;
}

void ManagerTestClass::set_Contracts(System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>>> value)
{
    pr_Contracts = value;
}

ManagerTestClass::ManagerTestClass() : pr_Age(0)
{
}

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
