#pragma once

#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/date_time.h>

namespace ApiExamples { namespace TestData { namespace TestClasses { class ManagerTestClass; } } }
namespace ApiExamples { namespace TestData { namespace TestClasses { class ClientTestClass; } } }

namespace ApiExamples {

namespace TestData {

namespace TestClasses {

class ContractTestClass : public System::Object
{
public:

    System::SharedPtr<ManagerTestClass> get_Manager();
    void set_Manager(System::SharedPtr<ManagerTestClass> value);
    System::SharedPtr<ClientTestClass> get_Client();
    void set_Client(System::SharedPtr<ClientTestClass> value);
    float get_Price();
    void set_Price(float value);
    System::DateTime get_Date();
    void set_Date(System::DateTime value);

    ContractTestClass();

private:

    System::SharedPtr<ManagerTestClass> pr_Manager;
    System::SharedPtr<ClientTestClass> pr_Client;
    float pr_Price;
    System::DateTime pr_Date;

};

} // namespace TestClasses
} // namespace TestData
} // namespace ApiExamples

