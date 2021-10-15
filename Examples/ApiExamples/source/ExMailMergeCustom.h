#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/MailMerging/IMailMergeDataSource.h>
#include <Aspose.Words.Cpp/MailMerging/IMailMergeDataSourceRoot.h>
#include <Aspose.Words.Cpp/MailMerging/MailMerge.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <system/array.h>
#include <system/collections/dictionary.h>
#include <system/collections/list.h>
#include <system/exceptions.h>
#include <system/object_ext.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;

namespace ApiExamples {

class ExMailMergeCustom : public ApiExampleBase
{
public:
    class EmployeeListMailMergeSource;

public:
    //ExStart
    //ExFor:IMailMergeDataSource
    //ExFor:IMailMergeDataSource.TableName
    //ExFor:IMailMergeDataSource.MoveNext
    //ExFor:IMailMergeDataSource.GetValue
    //ExFor:IMailMergeDataSource.GetChildDataSource
    //ExFor:MailMerge.Execute(IMailMergeDataSource)
    //ExSummary:Shows how to execute a mail merge with a data source in the form of a custom object.
    void CustomDataSource()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        builder->InsertField(u" MERGEFIELD FullName ");
        builder->InsertParagraph();
        builder->InsertField(u" MERGEFIELD Address ");

        SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Customer>>> customers =
            MakeObject<System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Customer>>>();
        customers->Add(MakeObject<ExMailMergeCustom::Customer>(u"Thomas Hardy", u"120 Hanover Sq., London"));
        customers->Add(MakeObject<ExMailMergeCustom::Customer>(u"Paolo Accorti", u"Via Monte Bianco 34, Torino"));

        // To use a custom object as a data source, it must implement the IMailMergeDataSource interface.
        auto dataSource = MakeObject<ExMailMergeCustom::CustomerMailMergeDataSource>(customers);

        doc->get_MailMerge()->Execute(dataSource);

        doc->Save(ArtifactsDir + u"MailMergeCustom.CustomDataSource.docx");
        TestCustomDataSource(customers, MakeObject<Document>(ArtifactsDir + u"MailMergeCustom.CustomDataSource.docx"));
        //ExSkip
    }

    /// <summary>
    /// An example of a "data entity" class in your application.
    /// </summary>
    class Customer : public System::Object
    {
    public:
        String get_FullName()
        {
            return pr_FullName;
        }

        void set_FullName(String value)
        {
            pr_FullName = value;
        }

        String get_Address()
        {
            return pr_Address;
        }

        void set_Address(String value)
        {
            pr_Address = value;
        }

        Customer(String aFullName, String anAddress)
        {
            set_FullName(aFullName);
            set_Address(anAddress);
        }

    private:
        String pr_FullName;
        String pr_Address;
    };

    /// <summary>
    /// A custom mail merge data source that you implement to allow Aspose.Words
    /// to mail merge data from your Customer objects into Microsoft Word documents.
    /// </summary>
    class CustomerMailMergeDataSource : public IMailMergeDataSource
    {
    public:
        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        String get_TableName() override
        {
            return u"Customer";
        }

        CustomerMailMergeDataSource(SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Customer>>> customers) : mRecordIndex(0)
        {
            mCustomers = customers;

            // When we initialize the data source, its position must be before the first record.
            mRecordIndex = -1;
        }

        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        bool GetValue(String fieldName, SharedPtr<System::Object>& fieldValue) override
        {
            if (fieldName == u"FullName")
            {
                fieldValue = System::ObjectExt::Box<String>(mCustomers->idx_get(mRecordIndex)->get_FullName());
                return true;
            }
            else if (fieldName == u"Address")
            {
                fieldValue = System::ObjectExt::Box<String>(mCustomers->idx_get(mRecordIndex)->get_Address());
                return true;
            }
            else
            {
                fieldValue.reset();
                return false;
            }
        }

        /// <summary>
        /// A standard implementation for moving to a next record in a collection.
        /// </summary>
        bool MoveNext() override
        {
            if (!get_IsEof())
            {
                mRecordIndex++;
            }

            return !get_IsEof();
        }

        SharedPtr<IMailMergeDataSource> GetChildDataSource(String tableName) override
        {
            return nullptr;
        }

    private:
        bool get_IsEof()
        {
            return mRecordIndex >= mCustomers->get_Count();
        }

        SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Customer>>> mCustomers;
        int mRecordIndex;
    };
    //ExEnd

    void TestCustomDataSource(SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Customer>>> customerList, SharedPtr<Document> doc)
    {
        ArrayPtr<ArrayPtr<String>> mergeData = MakeArray<ArrayPtr<String>>(customerList->get_Count());

        for (int i = 0; i < customerList->get_Count(); i++)
        {
            mergeData[i] = MakeArray<String>({customerList->idx_get(i)->get_FullName(), customerList->idx_get(i)->get_Address()});
        }

        TestUtil::MailMergeMatchesArray(mergeData, doc, true);
    }

    //ExStart
    //ExFor:IMailMergeDataSourceRoot
    //ExFor:IMailMergeDataSourceRoot.GetDataSource(String)
    //ExFor:MailMerge.ExecuteWithRegions(IMailMergeDataSourceRoot)
    //ExSummary:Performs mail merge from a custom data source with master-detail data.
    void CustomDataSourceRoot()
    {
        // Create a document with two mail merge regions named "Washington" and "Seattle".
        ArrayPtr<String> mailMergeRegions = MakeArray<String>({u"Vancouver", u"Seattle"});
        SharedPtr<Document> doc = CreateSourceDocumentWithMailMergeRegions(mailMergeRegions);

        // Create two data sources for the mail merge.
        SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Employee>>> employeesWashingtonBranch =
            MakeObject<System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Employee>>>();
        employeesWashingtonBranch->Add(MakeObject<ExMailMergeCustom::Employee>(u"John Doe", u"Sales"));
        employeesWashingtonBranch->Add(MakeObject<ExMailMergeCustom::Employee>(u"Jane Doe", u"Management"));

        SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Employee>>> employeesSeattleBranch =
            MakeObject<System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Employee>>>();
        employeesSeattleBranch->Add(MakeObject<ExMailMergeCustom::Employee>(u"John Cardholder", u"Management"));
        employeesSeattleBranch->Add(MakeObject<ExMailMergeCustom::Employee>(u"Joe Bloggs", u"Sales"));

        // Register our data sources by name in a data source root.
        //  If we are about to use this data source root in a mail merge with regions,
        // each source's registered name must match the name of an existing mail merge region in the mail merge source document.
        auto sourceRoot = MakeObject<ExMailMergeCustom::DataSourceRoot>();
        sourceRoot->RegisterSource(mailMergeRegions[0], MakeObject<ExMailMergeCustom::EmployeeListMailMergeSource>(employeesWashingtonBranch));
        sourceRoot->RegisterSource(mailMergeRegions[1], MakeObject<ExMailMergeCustom::EmployeeListMailMergeSource>(employeesSeattleBranch));

        // Since we have consecutive mail merge regions, we would normally have to perform two mail merges.
        // However, one mail merge source with a data root can fill in multiple regions
        // if the root contains tables with corresponding names/column names.
        doc->get_MailMerge()->ExecuteWithRegions(sourceRoot);

        doc->Save(ArtifactsDir + u"MailMergeCustom.CustomDataSourceRoot.docx");
    }

    /// <summary>
    /// Create a document that contains consecutive mail merge regions, with names designated by the input array,
    /// for a data table of employees.
    /// </summary>
    static SharedPtr<Document> CreateSourceDocumentWithMailMergeRegions(ArrayPtr<String> regions)
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

    /// <summary>
    /// An example of a "data entity" class in your application.
    /// </summary>
    class Employee : public System::Object
    {
    public:
        String get_FullName()
        {
            return pr_FullName;
        }

        String get_Department()
        {
            return pr_Department;
        }

        Employee(String aFullName, String aDepartment)
        {
            pr_FullName = aFullName;
            pr_Department = aDepartment;
        }

    private:
        String pr_FullName;
        String pr_Department;
    };

    /// <summary>
    /// Data source root that can be passed directly into a mail merge which can register and contain many child data sources.
    /// These sources must all implement IMailMergeDataSource, and are registered and differentiated by a name
    /// which corresponds to a mail merge region that will read the respective data.
    /// </summary>
    class DataSourceRoot : public IMailMergeDataSourceRoot
    {
    public:
        SharedPtr<IMailMergeDataSource> GetDataSource(String tableName) override
        {
            SharedPtr<ExMailMergeCustom::EmployeeListMailMergeSource> source = mSources->idx_get(tableName);
            source->Reset();
            return mSources->idx_get(tableName);
        }

        void RegisterSource(String sourceName, SharedPtr<ExMailMergeCustom::EmployeeListMailMergeSource> source)
        {
            mSources->Add(sourceName, source);
        }

        DataSourceRoot() : mSources(MakeObject<System::Collections::Generic::Dictionary<String, SharedPtr<ExMailMergeCustom::EmployeeListMailMergeSource>>>())
        {
        }

    private:
        SharedPtr<System::Collections::Generic::Dictionary<String, SharedPtr<ExMailMergeCustom::EmployeeListMailMergeSource>>> mSources;
    };

    /// <summary>
    /// Custom mail merge data source.
    /// </summary>
    class EmployeeListMailMergeSource : public IMailMergeDataSource
    {
    public:
        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        String get_TableName() override
        {
            return u"Employees";
        }

        EmployeeListMailMergeSource(SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Employee>>> employees) : mRecordIndex(0)
        {
            mEmployees = employees;
            mRecordIndex = -1;
        }

        /// <summary>
        /// A standard implementation for moving to a next record in a collection.
        /// </summary>
        bool MoveNext() override
        {
            if (!get_IsEof())
            {
                mRecordIndex++;
            }

            return !get_IsEof();
        }

        void Reset()
        {
            mRecordIndex = -1;
        }

        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        bool GetValue(String fieldName, SharedPtr<System::Object>& fieldValue) override
        {
            if (fieldName == u"FullName")
            {
                fieldValue = System::ObjectExt::Box<String>(mEmployees->idx_get(mRecordIndex)->get_FullName());
                return true;
            }
            else if (fieldName == u"Department")
            {
                fieldValue = System::ObjectExt::Box<String>(mEmployees->idx_get(mRecordIndex)->get_Department());
                return true;
            }
            else
            {
                fieldValue.reset();
                return false;
            }
        }

        /// <summary>
        /// Child data sources are for nested mail merges.
        /// </summary>
        SharedPtr<IMailMergeDataSource> GetChildDataSource(String tableName) override
        {
            throw System::NotImplementedException();
        }

    private:
        bool get_IsEof()
        {
            return mRecordIndex >= mEmployees->get_Count();
        }

        SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustom::Employee>>> mEmployees;
        int mRecordIndex;
    };
    //ExEnd
};

} // namespace ApiExamples
