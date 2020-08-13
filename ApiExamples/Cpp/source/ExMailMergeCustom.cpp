// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
// CPPDEFECT: System.Data is not supported
#include "ExMailMergeCustom.h"

#include <system/string.h>
#include <system/shared_ptr.h>
#include <system/scope_guard.h>
#include <system/object_ext.h>
#include <system/object.h>
#include <system/exceptions.h>
#include <system/collections/list.h>
#include <system/collections/dictionary.h>
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

RTTI_INFO_IMPL_HASH(1943753940u, ::ApiExamples::ExMailMergeCustom::EmployeeListMailMergeSource, ThisTypeBaseTypesInfo);

bool ExMailMergeCustom::EmployeeListMailMergeSource::get_IsEof()
{
    return mRecordIndex >= mEmployees->get_Count();
}

String ExMailMergeCustom::EmployeeListMailMergeSource::get_TableName()
{
    return u"Employees";
}

ExMailMergeCustom::EmployeeListMailMergeSource::EmployeeListMailMergeSource(SharedPtr<ExMailMergeCustom::EmployeeList> employees)
     : mRecordIndex(0)
{
    mEmployees = employees;
    mRecordIndex = -1;
}

bool ExMailMergeCustom::EmployeeListMailMergeSource::MoveNext()
{
    if (!get_IsEof())
    {
        mRecordIndex++;
    }

    return !get_IsEof();
}

void ExMailMergeCustom::EmployeeListMailMergeSource::Reset()
{
    mRecordIndex = -1;
}

bool ExMailMergeCustom::EmployeeListMailMergeSource::GetValue(String fieldName, SharedPtr<System::Object>& fieldValue)
{
    const String& switch_value_1 = fieldName;
    if (switch_value_1 == u"FullName")
    {
        fieldValue = System::ObjectExt::Box<String>(mEmployees->idx_get(mRecordIndex)->get_FullName());
        return true;
    }
    else if (switch_value_1 == u"Department")
    {
        fieldValue = System::ObjectExt::Box<String>(mEmployees->idx_get(mRecordIndex)->get_Department());
        return true;
    }
    else if (true)
    {
        fieldValue.reset();
        return false;
    }
}

SharedPtr<Aspose::Words::MailMerging::IMailMergeDataSource> ExMailMergeCustom::EmployeeListMailMergeSource::GetChildDataSource(String tableName)
{
    throw System::NotImplementedException();
}

System::Object::shared_members_type ApiExamples::ExMailMergeCustom::EmployeeListMailMergeSource::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::ExMailMergeCustom::EmployeeListMailMergeSource::mEmployees", this->mEmployees);

    return result;
}

RTTI_INFO_IMPL_HASH(2311686147u, ::ApiExamples::ExMailMergeCustom::DataSourceRoot, ThisTypeBaseTypesInfo);

SharedPtr<Aspose::Words::MailMerging::IMailMergeDataSource> ExMailMergeCustom::DataSourceRoot::GetDataSource(String tableName)
{
    SharedPtr<ExMailMergeCustom::EmployeeListMailMergeSource> source = mSources->idx_get(tableName);
    source->Reset();
    return mSources->idx_get(tableName);
}

void ExMailMergeCustom::DataSourceRoot::RegisterSource(String sourceName, SharedPtr<ExMailMergeCustom::EmployeeListMailMergeSource> source)
{
    mSources->Add(sourceName, source);
}

ExMailMergeCustom::DataSourceRoot::DataSourceRoot()
     : mSources(MakeObject<System::Collections::Generic::Dictionary<String, SharedPtr<ExMailMergeCustom::EmployeeListMailMergeSource>>>())
{
}

System::Object::shared_members_type ApiExamples::ExMailMergeCustom::DataSourceRoot::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::ExMailMergeCustom::DataSourceRoot::mSources", this->mSources);

    return result;
}

RTTI_INFO_IMPL_HASH(890233096u, ::ApiExamples::ExMailMergeCustom::EmployeeList, ThisTypeBaseTypesInfo);

SharedPtr<ExMailMergeCustom::Employee> ExMailMergeCustom::EmployeeList::idx_get(int index) const
{
    return System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Employee>>::idx_get(index);
}

void ExMailMergeCustom::EmployeeList::idx_set(int index, SharedPtr<ExMailMergeCustom::Employee> value)
{
    System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Employee>>::idx_set(index, value);
}

ExMailMergeCustom::EmployeeList::~EmployeeList()
{
}

RTTI_INFO_IMPL_HASH(1290033578u, ::ApiExamples::ExMailMergeCustom::Employee, ThisTypeBaseTypesInfo);

String ExMailMergeCustom::Employee::get_FullName()
{
    return pr_FullName;
}

String ExMailMergeCustom::Employee::get_Department()
{
    return pr_Department;
}

ExMailMergeCustom::Employee::Employee(String aFullName, String aDepartment)
{
    //Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    System::Details::ThisProtector __local_self_ref(this);
    //---------------------------------------------------------Self Reference

    pr_FullName = aFullName;
    pr_Department = aDepartment;
}

RTTI_INFO_IMPL_HASH(2441965292u, ::ApiExamples::ExMailMergeCustom::CustomerMailMergeDataSource, ThisTypeBaseTypesInfo);

String ExMailMergeCustom::CustomerMailMergeDataSource::get_TableName()
{
    return u"Customer";
}

bool ExMailMergeCustom::CustomerMailMergeDataSource::get_IsEof()
{
    return mRecordIndex >= mCustomers->get_Count();
}

ExMailMergeCustom::CustomerMailMergeDataSource::CustomerMailMergeDataSource(SharedPtr<ExMailMergeCustom::CustomerList> customers)
     : mRecordIndex(0)
{
    mCustomers = customers;

    // When the data source is initialized, it must be positioned before the first record.
    mRecordIndex = -1;
}

bool ExMailMergeCustom::CustomerMailMergeDataSource::GetValue(String fieldName, SharedPtr<System::Object>& fieldValue)
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
    else if (true)
    {
        fieldValue.reset();
        return false;
    }
}

bool ExMailMergeCustom::CustomerMailMergeDataSource::MoveNext()
{
    if (!get_IsEof())
    {
        mRecordIndex++;
    }

    return !get_IsEof();
}

SharedPtr<Aspose::Words::MailMerging::IMailMergeDataSource> ExMailMergeCustom::CustomerMailMergeDataSource::GetChildDataSource(String tableName)
{
    return nullptr;
}

System::Object::shared_members_type ApiExamples::ExMailMergeCustom::CustomerMailMergeDataSource::GetSharedMembers()
{
    auto result = System::Object::GetSharedMembers();

    result.Add("ApiExamples::ExMailMergeCustom::CustomerMailMergeDataSource::mCustomers", this->mCustomers);

    return result;
}

RTTI_INFO_IMPL_HASH(1701774168u, ::ApiExamples::ExMailMergeCustom::CustomerList, ThisTypeBaseTypesInfo);

SharedPtr<ExMailMergeCustom::Customer> ExMailMergeCustom::CustomerList::idx_get(int index) const
{
    return System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Customer>>::idx_get(index);
}

void ExMailMergeCustom::CustomerList::idx_set(int index, SharedPtr<ExMailMergeCustom::Customer> value)
{
    System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Customer>>::idx_set(index, value);
}

ExMailMergeCustom::CustomerList::~CustomerList()
{
}

RTTI_INFO_IMPL_HASH(337103866u, ::ApiExamples::ExMailMergeCustom::Customer, ThisTypeBaseTypesInfo);

String ExMailMergeCustom::Customer::get_FullName()
{
    return pr_FullName;
}

void ExMailMergeCustom::Customer::set_FullName(String value)
{
    pr_FullName = value;
}

String ExMailMergeCustom::Customer::get_Address()
{
    return pr_Address;
}

void ExMailMergeCustom::Customer::set_Address(String value)
{
    pr_Address = value;
}

ExMailMergeCustom::Customer::Customer(String aFullName, String anAddress)
{
    //Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    System::Details::ThisProtector __local_self_ref(this);
    //---------------------------------------------------------Self Reference

    set_FullName(aFullName);
    set_Address(anAddress);
}

void ExMailMergeCustom::TestCustomDataSource(SharedPtr<ExMailMergeCustom::CustomerList> customerList, SharedPtr<Aspose::Words::Document> doc)
{
    ArrayPtr<ArrayPtr<String>> mergeData = MakeArray<ArrayPtr<String>>(customerList->get_Count());

    for (int i = 0; i < customerList->get_Count(); i++)
    {
        mergeData[i] = MakeArray<String>({customerList->idx_get(i)->get_FullName(), customerList->idx_get(i)->get_Address()});
    }

    TestUtil::MailMergeMatchesArray(mergeData, doc, true);
}

SharedPtr<Aspose::Words::Document> ExMailMergeCustom::CreateSourceDocumentWithMailMergeRegions(ArrayPtr<String> regions)
{
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);

    for (String s : regions)
    {
        builder->Writeln(String(u"\n") + s + u" branch: ");
        builder->InsertField(String(u" MERGEFIELD TableStart:") + s);
        builder->InsertField(u" MERGEFIELD FullName");
        builder->Write(u", ");
        builder->InsertField(u" MERGEFIELD Department");
        builder->InsertField(String(u" MERGEFIELD TableEnd:") + s);
    }

    return doc;
}

namespace gtest_test
{

class ExMailMergeCustom : public ::testing::Test
{
protected:
    static SharedPtr<::ApiExamples::ExMailMergeCustom> s_instance;

    void SetUp() override
    {
        s_instance->SetUp();
    };

    static void SetUpTestCase()
    {
        s_instance = MakeObject<::ApiExamples::ExMailMergeCustom>();
        s_instance->OneTimeSetUp();
    };

    static void TearDownTestCase()
    {
        s_instance->OneTimeTearDown();
        s_instance = nullptr;
    };

};

SharedPtr<::ApiExamples::ExMailMergeCustom> ExMailMergeCustom::s_instance;

} // namespace gtest_test

void ExMailMergeCustom::CustomDataSource()
{
    // Create a destination document for the mail merge
    auto doc = MakeObject<Document>();
    auto builder = MakeObject<DocumentBuilder>(doc);
    builder->InsertField(u" MERGEFIELD FullName ");
    builder->InsertParagraph();
    builder->InsertField(u" MERGEFIELD Address ");

    // Create some data that we will use in the mail merge
    auto customers = MakeObject<ExMailMergeCustom::CustomerList>();
    customers->Add(MakeObject<ExMailMergeCustom::Customer>(u"Thomas Hardy", u"120 Hanover Sq., London"));
    customers->Add(MakeObject<ExMailMergeCustom::Customer>(u"Paolo Accorti", u"Via Monte Bianco 34, Torino"));

    // To be able to mail merge from your own data source, it must be wrapped
    // into an object that implements the IMailMergeDataSource interface
    auto dataSource = MakeObject<ExMailMergeCustom::CustomerMailMergeDataSource>(customers);

    // Now you can pass your data source into Aspose.Words
    doc->get_MailMerge()->Execute(dataSource);

    doc->Save(ArtifactsDir + u"MailMergeCustom.CustomDataSource.docx");
    TestCustomDataSource(customers, MakeObject<Document>(ArtifactsDir + u"MailMergeCustom.CustomDataSource.docx"));
    //ExSkip
}

namespace gtest_test
{

TEST_F(ExMailMergeCustom, CustomDataSource)
{
    s_instance->CustomDataSource();
}

} // namespace gtest_test

void ExMailMergeCustom::CustomDataSourceRoot()
{
    // Create a document with two mail merge regions named "Washington" and "Seattle"
    ArrayPtr<String> mailMergeRegions = MakeArray<String>({u"Vancouver", u"Seattle"});
    SharedPtr<Document> doc = CreateSourceDocumentWithMailMergeRegions(mailMergeRegions);

    // Create two data sources
    auto employeesWashingtonBranch = MakeObject<ExMailMergeCustom::EmployeeList>();
    employeesWashingtonBranch->Add(MakeObject<ExMailMergeCustom::Employee>(u"John Doe", u"Sales"));
    employeesWashingtonBranch->Add(MakeObject<ExMailMergeCustom::Employee>(u"Jane Doe", u"Management"));

    auto employeesSeattleBranch = MakeObject<ExMailMergeCustom::EmployeeList>();
    employeesSeattleBranch->Add(MakeObject<ExMailMergeCustom::Employee>(u"John Cardholder", u"Management"));
    employeesSeattleBranch->Add(MakeObject<ExMailMergeCustom::Employee>(u"Joe Bloggs", u"Sales"));

    // Register our data sources by name in a data source root
    auto sourceRoot = MakeObject<ExMailMergeCustom::DataSourceRoot>();
    sourceRoot->RegisterSource(mailMergeRegions[0], MakeObject<ExMailMergeCustom::EmployeeListMailMergeSource>(employeesWashingtonBranch));
    sourceRoot->RegisterSource(mailMergeRegions[1], MakeObject<ExMailMergeCustom::EmployeeListMailMergeSource>(employeesSeattleBranch));

    // Since we have consecutive mail merge regions, we would normally have to perform two mail merges
    // However, one mail merge source data root call every relevant data source and merge automatically
    doc->get_MailMerge()->ExecuteWithRegions(sourceRoot);

    doc->Save(ArtifactsDir + u"MailMergeCustom.CustomDataSourceRoot.docx");
}

namespace gtest_test
{

TEST_F(ExMailMergeCustom, CustomDataSourceRoot)
{
    s_instance->CustomDataSourceRoot();
}

} // namespace gtest_test

} // namespace ApiExamples
