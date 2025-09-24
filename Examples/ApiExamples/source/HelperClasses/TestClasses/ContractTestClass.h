#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.

#include <system/object.h>
#include <system/date_time.h>

#include "HelperClasses/TestClasses/ManagerTestClass.h"
#include "HelperClasses/TestClasses/ClientTestClass.h"

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
class ManagerTestClass;
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

class ContractTestClass : public System::Object
{
    typedef ContractTestClass ThisType;
    typedef System::Object BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass> get_Manager() const;
    void set_Manager(System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass> value);
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass> get_Client() const;
    void set_Client(System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass> value);
    float get_Price() const;
    void set_Price(float value);
    System::DateTime get_Date() const;
    void set_Date(System::DateTime value);
    
    ContractTestClass();
    
private:

    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass> pr_Manager;
    System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass> pr_Client;
    float pr_Price;
    System::DateTime pr_Date;
    
};

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


