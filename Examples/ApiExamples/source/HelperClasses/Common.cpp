// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
#include "HelperClasses/Common.h"

#include <system/linq/enumerable.h>
#include <system/enumerator_adapter.h>
#include <system/date_time.h>
#include <system/collections/list.h>
#include <cstdint>


using namespace Aspose::Words::ApiExamples::HelperClasses::TestClasses;
using namespace Aspose::Words::ApiExamples::TestData::TestClasses;
namespace Aspose {

namespace Words {

namespace ApiExamples {

namespace TestData {

System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass>>> Common::GetManagers()
{
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass>>> result = System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass>>>();
    auto manager = System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass>();
    manager->set_Name(u"John Smith");
    manager->set_Age(36);
    auto initValue = System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>();
    initValue->set_Client(System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass>());
    initValue->get_Client()->set_Name(u"A Company");
    initValue->get_Client()->set_Country(u"Australia");
    initValue->get_Client()->set_LocalAddress(u"219-241 Cleveland St STRAWBERRY HILLS  NSW  1427");
    initValue->set_Manager(manager);
    initValue->set_Price(1200000.0f);
    initValue->set_Date(System::DateTime(2017, 1, 1));
    auto initValue2 = System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>();
    initValue2->set_Client(System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass>());
    initValue2->get_Client()->set_Name(u"B Ltd.");
    initValue2->get_Client()->set_Country(u"Brazil");
    initValue2->get_Client()->set_LocalAddress(u"Avenida João Jorge, 112, ap. 31 Vila Industrial Campinas - SP 13035-680");
    initValue2->set_Manager(manager);
    initValue2->set_Price(750000.0f);
    initValue2->set_Date(System::DateTime(2017, 4, 1));
    auto initValue3 = System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>();
    initValue3->set_Client(System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass>());
    initValue3->get_Client()->set_Name(u"C & D");
    initValue3->get_Client()->set_Country(u"Canada");
    initValue3->get_Client()->set_LocalAddress(u"101-3485 RUE DE LA MONTAGNE MONTRÉAL (QUÉBEC) H3G 2A6");
    initValue3->set_Manager(manager);
    initValue3->set_Price(350000.0f);
    initValue3->set_Date(System::DateTime(2017, 7, 1));
    
    manager->set_Contracts(System::MakeArray<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>>({initValue, initValue2, initValue3}));
    
    result->Add(manager);
    manager = System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass>();
    manager->set_Name(u"Tony Anderson");
    manager->set_Age(37);
    auto initValue4 = System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>();
    initValue4->set_Client(System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass>());
    initValue4->get_Client()->set_Name(u"E Corp.");
    initValue4->get_Client()->set_LocalAddress(u"445 Mount Eden Road Mount Eden Auckland 1024");
    initValue4->set_Manager(manager);
    initValue4->set_Price(650000.0f);
    initValue4->set_Date(System::DateTime(2017, 2, 1));
    auto initValue5 = System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>();
    initValue5->set_Client(System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass>());
    initValue5->get_Client()->set_Name(u"F & Partners");
    initValue5->get_Client()->set_LocalAddress(u"20 Greens Road Tuahiwi Kaiapoi 7691 ");
    initValue5->set_Manager(manager);
    initValue5->set_Price(550000.0f);
    initValue5->set_Date(System::DateTime(2017, 8, 1));
    
    manager->set_Contracts(System::MakeArray<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>>({initValue4, initValue5}));
    
    result->Add(manager);
    manager = System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass>();
    manager->set_Name(u"July James");
    manager->set_Age(38);
    auto initValue6 = System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>();
    initValue6->set_Client(System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass>());
    initValue6->get_Client()->set_Name(u"G & Co.");
    initValue6->get_Client()->set_Country(u"Greece");
    initValue6->get_Client()->set_LocalAddress(u"Karkisias 6 GR-111 42  ATHINA GRÉCE");
    initValue6->set_Manager(manager);
    initValue6->set_Price(350000.0f);
    initValue6->set_Date(System::DateTime(2017, 2, 1));
    auto initValue7 = System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>();
    initValue7->set_Client(System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass>());
    initValue7->get_Client()->set_Name(u"H Group");
    initValue7->get_Client()->set_Country(u"Hungary");
    initValue7->get_Client()->set_LocalAddress(u"Budapest Fiktív utca 82., IV. em./28.2806");
    initValue7->set_Manager(manager);
    initValue7->set_Price(250000.0f);
    initValue7->set_Date(System::DateTime(2017, 5, 1));
    auto initValue8 = System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>();
    initValue8->set_Client(System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass>());
    initValue8->get_Client()->set_Name(u"I & Sons");
    initValue8->get_Client()->set_LocalAddress(u"43 Vogel Street Roslyn Palmerston North 4414");
    initValue8->set_Manager(manager);
    initValue8->set_Price(100000.0f);
    initValue8->set_Date(System::DateTime(2017, 7, 1));
    auto initValue9 = System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>();
    initValue9->set_Client(System::MakeObject<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass>());
    initValue9->get_Client()->set_Name(u"J Ent.");
    initValue9->get_Client()->set_Country(u"Japan");
    initValue9->get_Client()->set_LocalAddress(u"Hakusan 4-Chōme 3-2 Bunkyō-ku, TŌKYŌ 112-0001 Japan");
    initValue9->set_Manager(manager);
    initValue9->set_Price(100000.0f);
    initValue9->set_Date(System::DateTime(2017, 8, 1));
    
    manager->set_Contracts(System::MakeArray<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>>({initValue6, initValue7, initValue8, initValue9}));
    
    result->Add(manager);
    return result;
}

System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass>>> Common::GetEmptyManagers()
{
    return System::Linq::Enumerable::Empty<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ManagerTestClass> >();
}

System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass>>> Common::GetClients()
{
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass>>> result = System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ClientTestClass>>>();
    for (auto&& manager : System::IterateOver(GetManagers()))
    {
        for (auto&& contract : System::IterateOver(manager->get_Contracts()))
        {
            result->Add(contract->get_Client());
        }
    }
    
    return result;
}

System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>>> Common::GetContracts()
{
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>>> result = System::MakeObject<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::ApiExamples::TestData::TestClasses::ContractTestClass>>>();
    for (auto&& manager : System::IterateOver(GetManagers()))
    {
        for (auto&& contract : System::IterateOver(manager->get_Contracts()))
        {
            result->Add(contract);
        }
    }
    
    return result;
}

System::ArrayPtr<System::SharedPtr<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>> Common::GetShares()
{
    return System::MakeArray<System::SharedPtr<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>>({System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>(u"Technology", u"Consumer Electronics", u"AAPL", 6.602835, -0.0054), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>(u"Technology", u"Software - Infrastructure", u"MSFT", 5.832072, -0.005), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>(u"Technology", u"Software - Infrastructure", u"ADBE", 0.562561, -0.0274), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>(u"Technology", u"Semiconductors", u"NVDA", 1.335994, -0.0074), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>(u"Technology", u"Semiconductors", u"QCOM", 0.462198, 0.0248), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>(u"Communication Services", u"Internet Content & Information", u"GOOG", 3.771651, 0.011), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>(u"Communication Services", u"Entertainment", u"DIS", 0.575768, 0.0102), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>(u"Communication Services", u"Entertainment", u"WBD", 0.116579, -0.0165), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>(u"Consumer Cyclical", u"Internet Retail", u"AMZN", 3.011482, 0.044), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>(u"Consumer Cyclical", u"Auto Manufactures", u"TSLA", 1.816734, -0.0018), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>(u"Consumer Cyclical", u"Auto Manufactures", u"GM", 0.160205, 0.0026), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareTestClass>(u"Financial", u"Credit Services", u"V", 1.1, 0.005)});
}

System::ArrayPtr<System::SharedPtr<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareQuoteTestClass>> Common::GetShareQuotes()
{
    return System::MakeArray<System::SharedPtr<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareQuoteTestClass>>({System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareQuoteTestClass>(45131, 15232450, 171.32, 172.50, 170.69, 171.98), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareQuoteTestClass>(45132, 13962990, 172.20, 172.70, 171.40, 171.86), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareQuoteTestClass>(45133, 14902060, 171.86, 171.93, 170.31, 171.35), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareQuoteTestClass>(45134, 16962540, 171.64, 173.10, 171.35, 172.00), System::MakeObject<Aspose::Words::ApiExamples::HelperClasses::TestClasses::ShareQuoteTestClass>(45135, 15588280, 171.98, 172.40, 170.00, 171.44)});
}

} // namespace TestData
} // namespace ApiExamples
} // namespace Words
} // namespace Aspose
