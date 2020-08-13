#include "TestData/Common.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/object.h>
#include <system/linq/enumerable.h>
#include <system/enumerator_adapter.h>
#include <system/date_time.h>
#include <system/collections/list.h>
#include <system/collections/ienumerable.h>
#include <system/array.h>
#include <cstdint>

#include "TestData/TestClasses/ManagerTestClass.h"
#include "TestData/TestClasses/ContractTestClass.h"
#include "TestData/TestClasses/ClientTestClass.h"

using namespace ApiExamples::TestData::TestClasses;
namespace ApiExamples {

namespace TestData {

System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<TestClasses::ManagerTestClass>>> Common::GetManagers()
{
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<TestClasses::ManagerTestClass>>> managers = System::MakeObject<System::Collections::Generic::List<System::SharedPtr<TestClasses::ManagerTestClass>>>();

    {
        auto manager = System::MakeObject<TestClasses::ManagerTestClass>();
        manager->set_Name(u"John Smith");
        manager->set_Age(36);
        manager->set_Contracts(System::MakeArray<System::SharedPtr<ApiExamples::TestData::TestClasses::ContractTestClass>>({[&]{ auto tmp_0 = System::MakeObject<TestClasses::ContractTestClass>(); tmp_0->set_Client([&]{ auto tmp_1 = System::MakeObject<TestClasses::ClientTestClass>(); tmp_1->set_Name(u"A Company"); tmp_1->set_Country(u"Australia"); tmp_1->set_LocalAddress(u"219-241 Cleveland St STRAWBERRY HILLS  NSW  1427"); return tmp_1; }()); tmp_0->set_Manager(manager); tmp_0->set_Price(1200000); tmp_0->set_Date(System::DateTime(2017, 1, 1)); return tmp_0; }(), [&]{ auto tmp_2 = System::MakeObject<TestClasses::ContractTestClass>(); tmp_2->set_Client([&]{ auto tmp_3 = System::MakeObject<TestClasses::ClientTestClass>(); tmp_3->set_Name(u"B Ltd."); tmp_3->set_Country(u"Brazil"); tmp_3->set_LocalAddress(u"Avenida João Jorge, 112, ap. 31 Vila Industrial Campinas - SP 13035-680"); return tmp_3; }()); tmp_2->set_Manager(manager); tmp_2->set_Price(750000); tmp_2->set_Date(System::DateTime(2017, 4, 1)); return tmp_2; }(), [&]{ auto tmp_4 = System::MakeObject<TestClasses::ContractTestClass>(); tmp_4->set_Client([&]{ auto tmp_5 = System::MakeObject<TestClasses::ClientTestClass>(); tmp_5->set_Name(u"C & D"); tmp_5->set_Country(u"Canada"); tmp_5->set_LocalAddress(u"101-3485 RUE DE LA MONTAGNE MONTRÉAL (QUÉBEC) H3G 2A6"); return tmp_5; }()); tmp_4->set_Manager(manager); tmp_4->set_Price(350000); tmp_4->set_Date(System::DateTime(2017, 7, 1)); return tmp_4; }()}));
        managers->Add(manager);
    }

    {
        auto manager = System::MakeObject<TestClasses::ManagerTestClass>();
        manager->set_Name(u"Tony Anderson");
        manager->set_Age(37);
        manager->set_Contracts(System::MakeArray<System::SharedPtr<ApiExamples::TestData::TestClasses::ContractTestClass>>({[&]{ auto tmp_6 = System::MakeObject<TestClasses::ContractTestClass>(); tmp_6->set_Client([&]{ auto tmp_7 = System::MakeObject<TestClasses::ClientTestClass>(); tmp_7->set_Name(u"E Corp."); tmp_7->set_LocalAddress(u"445 Mount Eden Road Mount Eden Auckland 1024"); return tmp_7; }()); tmp_6->set_Manager(manager); tmp_6->set_Price(650000); tmp_6->set_Date(System::DateTime(2017, 2, 1)); return tmp_6; }(), [&]{ auto tmp_8 = System::MakeObject<TestClasses::ContractTestClass>(); tmp_8->set_Client([&]{ auto tmp_9 = System::MakeObject<TestClasses::ClientTestClass>(); tmp_9->set_Name(u"F & Partners"); tmp_9->set_LocalAddress(u"20 Greens Road Tuahiwi Kaiapoi 7691 "); return tmp_9; }()); tmp_8->set_Manager(manager); tmp_8->set_Price(550000); tmp_8->set_Date(System::DateTime(2017, 8, 1)); return tmp_8; }()}));
        managers->Add(manager);
    }

    {
        auto manager = System::MakeObject<TestClasses::ManagerTestClass>();
        manager->set_Name(u"July James");
        manager->set_Age(38);
        manager->set_Contracts(System::MakeArray<System::SharedPtr<ApiExamples::TestData::TestClasses::ContractTestClass>>({[&]{ auto tmp_10 = System::MakeObject<TestClasses::ContractTestClass>(); tmp_10->set_Client([&]{ auto tmp_11 = System::MakeObject<TestClasses::ClientTestClass>(); tmp_11->set_Name(u"G & Co."); tmp_11->set_Country(u"Greece"); tmp_11->set_LocalAddress(u"Karkisias 6 GR-111 42  ATHINA GRÉCE"); return tmp_11; }()); tmp_10->set_Manager(manager); tmp_10->set_Price(350000); tmp_10->set_Date(System::DateTime(2017, 2, 1)); return tmp_10; }(), [&]{ auto tmp_12 = System::MakeObject<TestClasses::ContractTestClass>(); tmp_12->set_Client([&]{ auto tmp_13 = System::MakeObject<TestClasses::ClientTestClass>(); tmp_13->set_Name(u"H Group"); tmp_13->set_Country(u"Hungary"); tmp_13->set_LocalAddress(u"Budapest Fiktív utca 82., IV. em./28.2806"); return tmp_13; }()); tmp_12->set_Manager(manager); tmp_12->set_Price(250000); tmp_12->set_Date(System::DateTime(2017, 5, 1)); return tmp_12; }(), [&]{ auto tmp_14 = System::MakeObject<TestClasses::ContractTestClass>(); tmp_14->set_Client([&]{ auto tmp_15 = System::MakeObject<TestClasses::ClientTestClass>(); tmp_15->set_Name(u"I & Sons"); tmp_15->set_LocalAddress(u"43 Vogel Street Roslyn Palmerston North 4414"); return tmp_15; }()); tmp_14->set_Manager(manager); tmp_14->set_Price(100000); tmp_14->set_Date(System::DateTime(2017, 7, 1)); return tmp_14; }(), [&]{ auto tmp_16 = System::MakeObject<TestClasses::ContractTestClass>(); tmp_16->set_Client([&]{ auto tmp_17 = System::MakeObject<TestClasses::ClientTestClass>(); tmp_17->set_Name(u"J Ent."); tmp_17->set_Country(u"Japan"); tmp_17->set_LocalAddress(u"Hakusan 4-Chōme 3-2 Bunkyō-ku, TŌKYŌ 112-0001 Japan"); return tmp_17; }()); tmp_16->set_Manager(manager); tmp_16->set_Price(100000); tmp_16->set_Date(System::DateTime(2017, 8, 1)); return tmp_16; }()}));
        managers->Add(manager);
    }

    return managers;
}

System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<TestClasses::ManagerTestClass>>> Common::GetEmptyManagers()
{
    return System::Linq::Enumerable::Empty<System::SharedPtr<TestClasses::ManagerTestClass> >();
}

System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<TestClasses::ClientTestClass>>> Common::GetClients()
{
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<TestClasses::ClientTestClass>>> clients = System::MakeObject<System::Collections::Generic::List<System::SharedPtr<TestClasses::ClientTestClass>>>();
    for (auto manager : System::IterateOver(GetManagers()))
    {
        for (auto contract : System::IterateOver(manager->get_Contracts()))
        {
            clients->Add(contract->get_Client());
        }
    }
    return clients;
}

System::SharedPtr<System::Collections::Generic::IEnumerable<System::SharedPtr<TestClasses::ContractTestClass>>> Common::GetContracts()
{
    System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<TestClasses::ContractTestClass>>> contracts = System::MakeObject<System::Collections::Generic::List<System::SharedPtr<TestClasses::ContractTestClass>>>();
    for (auto manager : System::IterateOver(GetManagers()))
    {
        contracts->AddRange(manager->get_Contracts());
    }
    return contracts;
}

} // namespace TestData
} // namespace ApiExamples
