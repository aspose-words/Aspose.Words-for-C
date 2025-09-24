#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.

#include <system/collections/ienumerable.h>
#include <system/array.h>

#include "HelperClasses/TestClasses/ShareTestClass.h"
#include "HelperClasses/TestClasses/ShareQuoteTestClass.h"
#include "HelperClasses/TestClasses/ManagerTestClass.h"
#include "HelperClasses/TestClasses/ContractTestClass.h"
#include "HelperClasses/TestClasses/ClientTestClass.h"

namespace Aspose
{
namespace Words
{
namespace ApiExamples
{
namespace HelperClasses
{
namespace TestClasses
{
class ShareQuoteTestClass;
class ShareTestClass;
} // namespace TestClasses
} // namespace HelperClasses
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


using namespace Aspose::Words::ApiExamples::HelperClasses::TestClasses;
using namespace Aspose::Words::ApiExamples::TestData::TestClasses;

namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

class Common
{
    typedef Common ThisType;
    
public:

    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass>>> GetManagers();
    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass>>> GetEmptyManagers();
    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass>>> GetClients();
    static System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>>> GetContracts();
    static System::ArrayPtr<System::SharedPtr<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>> GetShares();
    static System::ArrayPtr<System::SharedPtr<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareQuoteTestClass>> GetShareQuotes();
    
public:
    Common() = delete;
};

} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


