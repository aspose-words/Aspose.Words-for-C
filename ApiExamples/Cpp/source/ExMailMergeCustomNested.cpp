// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
#include "ExMailMergeCustomNested.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/scope_guard.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/enumerator_adapter.h>
#include <system/convert.h>
#include <system/collections/list.h>
#include <system/array.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IMailMergeDataSource.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>

#include "TestUtil.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;
namespace ApiExamples {

RTTI_INFO_IMPL_HASH(3659109035u, ::ApiExamples::ExMailMergeCustomNested::OrderMailMergeDataSource, ThisTypeBaseTypesInfo);

String ExMailMergeCustomNested::OrderMailMergeDataSource::get_TableName()
{
    return u"Order";
}

bool ExMailMergeCustomNested::OrderMailMergeDataSource::get_IsEof()
{
    return mRecordIndex >= mOrders->get_Count();
}

ExMailMergeCustomNested::OrderMailMergeDataSource::OrderMailMergeDataSource(SharedPtr<ExMailMergeCustomNested::OrderList> orders)
     : mRecordIndex(0)
{
    mOrders = orders;

    // When the data source is initialized, it must be positioned before the first record
    mRecordIndex = -1;
}

bool ExMailMergeCustomNested::OrderMailMergeDataSource::GetValue(String fieldName, SharedPtr<System::Object>& fieldValue)
{
    const String& switch_value_2 = fieldName;
    if (switch_value_2 == u"Name")
    {
        fieldValue = System::ObjectExt::Box<String>(mOrders->idx_get(mRecordIndex)->get_Name());
        return true;
    }
    else if (switch_value_2 == u"Quantity")
    {
        fieldValue = System::ObjectExt::Box<int>(mOrders->idx_get(mRecordIndex)->get_Quantity());
        return true;
    }
    else if (true)
    {
        fieldValue.reset();
        return false;
    }
}

bool ExMailMergeCustomNested::OrderMailMergeDataSource::MoveNext()
{
    if (!get_IsEof())
    {
        mRecordIndex++;
    }

    return !get_IsEof();
}

SharedPtr<Aspose::Words::MailMerging::IMailMergeDataSource> ExMailMergeCustomNested::OrderMailMergeDataSource::GetChildDataSource(String tableName)
{
    return nullptr;
}

System::Object::shared_members_type ApiExamples::ExMailMergeCustomNested::OrderMailMergeDataSource::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::ExMailMergeCustomNested::OrderMailMergeDataSource::mOrders", this->mOrders);

    return result;
}

RTTI_INFO_IMPL_HASH(3693202837u, ::ApiExamples::ExMailMergeCustomNested::CustomerMailMergeDataSource, ThisTypeBaseTypesInfo);

String ExMailMergeCustomNested::CustomerMailMergeDataSource::get_TableName()
{
    return u"Customer";
}

bool ExMailMergeCustomNested::CustomerMailMergeDataSource::get_IsEof()
{
    return mRecordIndex >= mCustomers->get_Count();
}

ExMailMergeCustomNested::CustomerMailMergeDataSource::CustomerMailMergeDataSource(SharedPtr<ExMailMergeCustomNested::CustomerList> customers)
     : mRecordIndex(0)
{
    mCustomers = customers;

    // When the data source is initialized, it must be positioned before the first record
    mRecordIndex = -1;
}

bool ExMailMergeCustomNested::CustomerMailMergeDataSource::GetValue(String fieldName, SharedPtr<System::Object>& fieldValue)
{
    const String& switch_value_0 = fieldName;
    if (switch_value_0 == u"FullName")
    {
        fieldValue = System::ObjectExt::Box<String>(mCustomers->idx_get(mRecordIndex)->get_FullName());
        return true;
    }
    else if (switch_value_0 == u"Address")
    {
        fieldValue = System::ObjectExt::Box<String>(mCustomers->idx_get(mRecordIndex)->get_Address());
        return true;
    }
    else if (switch_value_0 == u"Order")
    {
        fieldValue = mCustomers->idx_get(mRecordIndex)->get_Orders();
        return true;
    }
    else if (true)
    {
        fieldValue.reset();
        return false;
    }
}

bool ExMailMergeCustomNested::CustomerMailMergeDataSource::MoveNext()
{
    if (!get_IsEof())
    {
        mRecordIndex++;
    }

    return !get_IsEof();
}

SharedPtr<Aspose::Words::MailMerging::IMailMergeDataSource> ExMailMergeCustomNested::CustomerMailMergeDataSource::GetChildDataSource(String tableName)
{
    const String& switch_value_1 = tableName;
    if (switch_value_1 == u"Order")
    {
        return MakeObject<ExMailMergeCustomNested::OrderMailMergeDataSource>(mCustomers->idx_get(mRecordIndex)->get_Orders());
    }
    else if (true)
    {
        return nullptr;
    }
}

System::Object::shared_members_type ApiExamples::ExMailMergeCustomNested::CustomerMailMergeDataSource::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::ExMailMergeCustomNested::CustomerMailMergeDataSource::mCustomers", this->mCustomers);

    return result;
}

RTTI_INFO_IMPL_HASH(3451081561u, ::ApiExamples::ExMailMergeCustomNested::OrderList, ThisTypeBaseTypesInfo);

SharedPtr<ExMailMergeCustomNested::Order> ExMailMergeCustomNested::OrderList::idx_get(int index) const
{
    return System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Order>>::idx_get(index);
}

void ExMailMergeCustomNested::OrderList::idx_set(int index, SharedPtr<ExMailMergeCustomNested::Order> value)
{
    System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Order>>::idx_set(index, value);
}

ExMailMergeCustomNested::OrderList::~OrderList()
{
}

RTTI_INFO_IMPL_HASH(2600516859u, ::ApiExamples::ExMailMergeCustomNested::Order, ThisTypeBaseTypesInfo);

String ExMailMergeCustomNested::Order::get_Name()
{
    return pr_Name;
}

void ExMailMergeCustomNested::Order::set_Name(String value)
{
    pr_Name = value;
}

int ExMailMergeCustomNested::Order::get_Quantity()
{
    return pr_Quantity;
}

void ExMailMergeCustomNested::Order::set_Quantity(int value)
{
    pr_Quantity = value;
}

ExMailMergeCustomNested::Order::Order(String oName, int oQuantity) : pr_Quantity(0)
{
    //Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    System::Details::ThisProtector __local_self_ref(this);
    //---------------------------------------------------------Self Reference

    set_Name(oName);
    set_Quantity(oQuantity);
}

RTTI_INFO_IMPL_HASH(687680239u, ::ApiExamples::ExMailMergeCustomNested::CustomerList, ThisTypeBaseTypesInfo);

SharedPtr<ExMailMergeCustomNested::Customer> ExMailMergeCustomNested::CustomerList::idx_get(int index) const
{
    return System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Customer>>::idx_get(index);
}

void ExMailMergeCustomNested::CustomerList::idx_set(int index, SharedPtr<ExMailMergeCustomNested::Customer> value)
{
    System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Customer>>::idx_set(index, value);
}

ExMailMergeCustomNested::CustomerList::~CustomerList()
{
}

RTTI_INFO_IMPL_HASH(337111697u, ::ApiExamples::ExMailMergeCustomNested::Customer, ThisTypeBaseTypesInfo);

String ExMailMergeCustomNested::Customer::get_FullName()
{
    return pr_FullName;
}

void ExMailMergeCustomNested::Customer::set_FullName(String value)
{
    pr_FullName = value;
}

String ExMailMergeCustomNested::Customer::get_Address()
{
    return pr_Address;
}

void ExMailMergeCustomNested::Customer::set_Address(String value)
{
    pr_Address = value;
}

SharedPtr<ExMailMergeCustomNested::OrderList> ExMailMergeCustomNested::Customer::get_Orders()
{
    return pr_Orders;
}

void ExMailMergeCustomNested::Customer::set_Orders(SharedPtr<ExMailMergeCustomNested::OrderList> value)
{
    pr_Orders = value;
}

ExMailMergeCustomNested::Customer::Customer(String aFullName, String anAddress)
{
    //Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    System::Details::ThisProtector __local_self_ref(this);
    //---------------------------------------------------------Self Reference

    set_FullName(aFullName);
    set_Address(anAddress);
    set_Orders(MakeObject<ExMailMergeCustomNested::OrderList>());
}

System::Object::shared_members_type ApiExamples::ExMailMergeCustomNested::Customer::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::ExMailMergeCustomNested::Customer::pr_Orders", this->pr_Orders);

    return result;
}

void ExMailMergeCustomNested::TestCustomDataSource(SharedPtr<ExMailMergeCustomNested::CustomerList> customers, SharedPtr<Aspose::Words::Document> doc)
{
    SharedPtr<System::Collections::Generic::List<ArrayPtr<String>>> mailMergeData = MakeObject<System::Collections::Generic::List<ArrayPtr<String>>>();

    for (auto customer : System::IterateOver(customers))
    {
        for (auto order : System::IterateOver(customer->get_Orders()))
        {
            mailMergeData->Add(MakeArray<String>({order->get_Name(), System::Convert::ToString(order->get_Quantity())}));
        }
        mailMergeData->Add(MakeArray<String>({customer->get_FullName(), customer->get_Address()}));
    }

    TestUtil::MailMergeMatchesArray(mailMergeData->ToArray(), doc, false);
}

namespace gtest_test
{

class ExMailMergeCustomNested : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExMailMergeCustomNested> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExMailMergeCustomNested>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExMailMergeCustomNested> ExMailMergeCustomNested::s_instance;

} // namespace gtest_test

void ExMailMergeCustomNested::CustomDataSource()
{
    // Create a destination document for the mail merge
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->InsertField(u" MERGEFIELD TableStart:Customer");

    builder->Write(u"Full name:\t");
    builder->InsertField(u" MERGEFIELD FullName ");
    builder->Write(u"\nAddress:\t");
    builder->InsertField(u" MERGEFIELD Address ");
    builder->Write(u"\nOrders:\n");

    builder->InsertField(u" MERGEFIELD TableStart:Order");

    builder->Write(u"\tItem name:\t");
    builder->InsertField(u" MERGEFIELD Name ");
    builder->Write(u"\n\tQuantity:\t");
    builder->InsertField(u" MERGEFIELD Quantity ");
    builder->InsertParagraph();

    builder->InsertField(u" MERGEFIELD TableEnd:Order");

    builder->InsertField(u" MERGEFIELD TableEnd:Customer");

    // Create some data that we will use in the mail merge
    auto customers = MakeObject<ExMailMergeCustomNested::CustomerList>();
    customers->Add(MakeObject<ExMailMergeCustomNested::Customer>(u"Thomas Hardy", u"120 Hanover Sq., London"));
    customers->Add(MakeObject<ExMailMergeCustomNested::Customer>(u"Paolo Accorti", u"Via Monte Bianco 34, Torino"));

    // Create some data for nesting in the mail merge
    customers->idx_get(0)->get_Orders()->Add(MakeObject<ExMailMergeCustomNested::Order>(u"Rugby World Cup Cap", 2));
    customers->idx_get(0)->get_Orders()->Add(MakeObject<ExMailMergeCustomNested::Order>(u"Rugby World Cup Ball", 1));
    customers->idx_get(1)->get_Orders()->Add(MakeObject<ExMailMergeCustomNested::Order>(u"Rugby World Cup Guide", 1));

    // To be able to mail merge from your own data source, it must be wrapped
    // into an object that implements the IMailMergeDataSource interface
    auto customersDataSource = MakeObject<ExMailMergeCustomNested::CustomerMailMergeDataSource>(customers);

    // Now you can pass your data source into Aspose.Words
    doc->get_MailMerge()->ExecuteWithRegions(customersDataSource);

    doc->Save(ArtifactsDir + u"NestedMailMergeCustom.CustomDataSource.docx");
    TestCustomDataSource(customers, MakeObject<Document>(ArtifactsDir + u"NestedMailMergeCustom.CustomDataSource.docx"));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExMailMergeCustomNested, CustomDataSource)
{
    s_instance->CustomDataSource();
}

} // namespace gtest_test

} // namespace ApiExamples
