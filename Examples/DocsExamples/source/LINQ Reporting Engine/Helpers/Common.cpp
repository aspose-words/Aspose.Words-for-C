#include "Common.h"

#include <cstdint>
#include <system/collections/ienumerator.h>
#include <system/collections/list.h>
#include <system/date_time.h>
#include <system/enumerator_adapter.h>
#include <system/io/file.h>

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace DocsExamples::LINQ_Reporting_Engine::Helpers::Data_Source_Objects;
namespace DocsExamples { namespace LINQ_Reporting_Engine { namespace Helpers {

SharedPtr<Data_Source_Objects::Manager> Common::GetManager()
{
    //ExStart:GetManager
    SharedPtr<System::Collections::Generic::IEnumerator<SharedPtr<Data_Source_Objects::Manager>>> managers = GetManagers()->GetEnumerator();
    managers->MoveNext();

    return managers->get_Current();
    //ExEnd:GetManager
}

SharedPtr<System::Collections::Generic::IEnumerable<SharedPtr<Data_Source_Objects::Client>>> Common::GetClients()
{
    //ExStart:GetClients
    auto clients = MakeObject<System::Collections::Generic::List<SharedPtr<Data_Source_Objects::Client>>>();
    for (const auto& manager : System::IterateOver(GetManagers()))
    {
        for (const auto& contract : System::IterateOver(manager->get_Contracts()))
        {
            clients->Add(contract->get_Client());
        }
    }
    return clients;
    //ExEnd:GetClients
}

SharedPtr<System::Collections::Generic::IEnumerable<SharedPtr<Data_Source_Objects::Manager>>> Common::GetManagers()
{
    //ExStart:GetManagers
    auto managers = MakeObject<System::Collections::Generic::List<SharedPtr<Data_Source_Objects::Manager>>>();

    {
        auto manager = MakeObject<Data_Source_Objects::Manager>();
        manager->set_Name(u"John Smith");
        manager->set_Age(36);
        manager->set_Photo(Photo());
        auto contracts = MakeObject<System::Collections::Generic::List<SharedPtr<Data_Source_Objects::Contract>>>();

        {
            auto client = MakeObject<Data_Source_Objects::Client>();
            client->set_Name(u"A Company");
            client->set_Country(u"Australia");
            client->set_LocalAddress(u"219-241 Cleveland St STRAWBERRY HILLS  NSW  1427");
            auto contract = MakeObject<Data_Source_Objects::Contract>();
            contract->set_Client(client);
            contract->set_Manager(manager);
            contract->set_Price(1200000);
            contract->set_Date(System::DateTime(2015, 1, 1));
            contracts->Add(contract);
        }

        {
            auto client = MakeObject<Data_Source_Objects::Client>();
            client->set_Name(u"B Ltd.");
            client->set_Country(u"Brazil");
            client->set_LocalAddress(u"Avenida João Jorge, 112, ap. 31 Vila Industrial Campinas - SP 13035-680");
            auto contract = MakeObject<Data_Source_Objects::Contract>();
            contract->set_Client(client);
            contract->set_Manager(manager);
            contract->set_Price(750000);
            contract->set_Date(System::DateTime(2015, 4, 1));
            contracts->Add(contract);
        }

        {
            auto client = MakeObject<Data_Source_Objects::Client>();
            client->set_Name(u"C & D");
            client->set_Country(u"Canada");
            client->set_LocalAddress(u"101-3485 RUE DE LA MONTAGNE MONTRÉAL (QUÉBEC) H3G 2A6");
            auto contract = MakeObject<Data_Source_Objects::Contract>();
            contract->set_Client(client);
            contract->set_Manager(manager);
            contract->set_Price(350000);
            contract->set_Date(System::DateTime(2015, 7, 1));
            contracts->Add(contract);
        }
        manager->set_Contracts(contracts);
        managers->Add(manager);
    }

    {
        auto manager = MakeObject<Data_Source_Objects::Manager>();
        manager->set_Name(u"Tony Anderson");
        manager->set_Age(37);
        manager->set_Photo(Photo());
        auto contracts = MakeObject<System::Collections::Generic::List<SharedPtr<Data_Source_Objects::Contract>>>();

        {
            auto client = MakeObject<Data_Source_Objects::Client>();
            client->set_Name(u"E Corp.");
            client->set_LocalAddress(u"445 Mount Eden Road Mount Eden Auckland 1024");
            auto contract = MakeObject<Data_Source_Objects::Contract>();
            contract->set_Client(client);
            contract->set_Manager(manager);
            contract->set_Price(650000);
            contract->set_Date(System::DateTime(2015, 2, 1));
            contracts->Add(contract);
        }

        {
            auto client = MakeObject<Data_Source_Objects::Client>();
            client->set_Name(u"F & Partners");
            client->set_LocalAddress(u"20 Greens Road Tuahiwi Kaiapoi 7691 ");
            auto contract = MakeObject<Data_Source_Objects::Contract>();
            contract->set_Client(client);
            contract->set_Manager(manager);
            contract->set_Price(550000);
            contract->set_Date(System::DateTime(2015, 8, 1));
            contracts->Add(contract);
        }
        manager->set_Contracts(contracts);
        managers->Add(manager);
    }

    {
        auto manager = MakeObject<Data_Source_Objects::Manager>();
        manager->set_Name(u"July James");
        manager->set_Age(38);
        manager->set_Photo(Photo());
        auto contracts = MakeObject<System::Collections::Generic::List<SharedPtr<Data_Source_Objects::Contract>>>();

        {
            auto client = MakeObject<Data_Source_Objects::Client>();
            client->set_Name(u"G & Co.");
            client->set_Country(u"Greece");
            client->set_LocalAddress(u"Karkisias 6 GR-111 42  ATHINA GRÉCE");
            auto contract = MakeObject<Data_Source_Objects::Contract>();
            contract->set_Client(client);
            contract->set_Manager(manager);
            contract->set_Price(350000);
            contract->set_Date(System::DateTime(2015, 2, 1));
            contracts->Add(contract);
        }

        {
            auto client = MakeObject<Data_Source_Objects::Client>();
            client->set_Name(u"H Group");
            client->set_Country(u"Hungary");
            client->set_LocalAddress(u"Budapest Fiktív utca 82., IV. em./28.2806");
            auto contract = MakeObject<Data_Source_Objects::Contract>();
            contract->set_Client(client);
            contract->set_Manager(manager);
            contract->set_Price(250000);
            contract->set_Date(System::DateTime(2015, 5, 1));
            contracts->Add(contract);
        }

        {
            auto client = MakeObject<Data_Source_Objects::Client>();
            client->set_Name(u"I & Sons");
            client->set_LocalAddress(u"43 Vogel Street Roslyn Palmerston North 4414");
            auto contract = MakeObject<Data_Source_Objects::Contract>();
            contract->set_Client(client);
            contract->set_Manager(manager);
            contract->set_Price(100000);
            contract->set_Date(System::DateTime(2015, 7, 1));
            contracts->Add(contract);
        }

        {
            auto client = MakeObject<Data_Source_Objects::Client>();
            client->set_Name(u"J Ent.");
            client->set_Country(u"Japan");
            client->set_LocalAddress(u"Hakusan 4-Chōme 3-2 Bunkyō-ku, TŌKYŌ 112-0001 Japan");
            auto contract = MakeObject<Data_Source_Objects::Contract>();
            contract->set_Client(client);
            contract->set_Manager(manager);
            contract->set_Price(100000);
            contract->set_Date(System::DateTime(2015, 8, 1));
            contracts->Add(contract);
        }
        manager->set_Contracts(contracts);
        managers->Add(manager);
    }
    return managers;
    //ExEnd:GetManagers
}

ArrayPtr<uint8_t> Common::Photo()
{
    //ExStart:Photo
    // Load the photo and read all bytes
    ArrayPtr<uint8_t> logo = System::IO::File::ReadAllBytes(ImagesDir + u"Logo.jpg");

    return logo;
    //ExEnd:Photo
}

SharedPtr<System::Collections::Generic::IEnumerable<SharedPtr<Data_Source_Objects::Contract>>> Common::GetContracts()
{
    //ExStart:GetContracts
    auto contracts = MakeObject<System::Collections::Generic::List<SharedPtr<Data_Source_Objects::Contract>>>();
    for (const auto& manager : System::IterateOver(GetManagers()))
    {
        for (const auto& contract : System::IterateOver(manager->get_Contracts()))
        {
            contracts->Add(contract);
        }
    }
    return contracts;
    //ExEnd:GetContracts
}

}}} // namespace DocsExamples::LINQ_Reporting_Engine::Helpers
