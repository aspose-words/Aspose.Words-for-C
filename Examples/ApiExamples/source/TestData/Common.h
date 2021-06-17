#pragma once

#include <system/collections/ienumerable.h>
#include <system/shared_ptr.h>

#include "TestData/TestClasses/ClientTestClass.h"
#include "TestData/TestClasses/ContractTestClass.h"
#include "TestData/TestClasses/ManagerTestClass.h"

namespace ApiExamples { namespace TestData { namespace TestClasses {
class ManagerTestClass;
}}} // namespace ApiExamples::TestData::TestClasses
namespace ApiExamples { namespace TestData { namespace TestClasses {
class ContractTestClass;
}}} // namespace ApiExamples::TestData::TestClasses

using namespace ApiExamples::TestData::TestClasses;

namespace ApiExamples { namespace TestData {

class Common
{
public:
    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<TestClasses::ManagerTestClass>>> GetManagers();
    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<TestClasses::ManagerTestClass>>> GetEmptyManagers();
    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<TestClasses::ClientTestClass>>> GetClients();
    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<TestClasses::ContractTestClass>>> GetContracts();

public:
    Common() = delete;
};

}} // namespace ApiExamples::TestData
