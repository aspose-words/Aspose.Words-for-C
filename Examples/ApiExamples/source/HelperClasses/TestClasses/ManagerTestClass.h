#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.

#include <system/string.h>
#include <system/collections/ienumerable.h>
#include <cstdint>

#include "HelperClasses/TestClasses/ContractTestClass.h"

namespace Aspose
{
namespace Words
{
namespace ApiExamples
{
namespace TestData
{
namespace TestClasses
{
class ContractTestClass;
} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose

namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

class ManagerTestClass : public System::Object
{
    typedef ManagerTestClass ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    System::String get_Name() const;
    void set_Name(System::String value);
    int32_t get_Age() const;
    void set_Age(int32_t value);
    System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>>> get_Contracts() const;
    void set_Contracts(System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>>> value);
    
    ManagerTestClass();
    
private:

    System::String pr_Name;
    int32_t pr_Age;
    System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>>> pr_Contracts;
    
};

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


