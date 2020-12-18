#pragma once

#include <system/shared_ptr.h>
#include <system/collections/ienumerable.h>

namespace ApiExamples { namespace TestData { namespace TestClasses { class ManagerTestClass; } } }
namespace ApiExamples { namespace TestData { namespace TestClasses { class ClientTestClass; } } }
namespace ApiExamples { namespace TestData { namespace TestClasses { class ContractTestClass; } } }

using namespace ApiExamples::TestData::TestClasses;

namespace ApiExamples {

namespace TestData {

class Common
{
public:

    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<TestClasses::ManagerTestClass>>> GetManagers();
    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<TestClasses::ManagerTestClass>>> GetEmptyManagers();
    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<TestClasses::ClientTestClass>>> GetClients();
    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<TestClasses::ContractTestClass>>> GetContracts();

};

} // namespace TestData
} // namespace ApiExamples

