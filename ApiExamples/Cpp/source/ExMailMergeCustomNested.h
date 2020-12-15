#pragma once
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <cstdint>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/Model/Document/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Model/Fields/Field.h>
#include <Aspose.Words.Cpp/Model/MailMerge/IMailMergeDataSource.h>
#include <Aspose.Words.Cpp/Model/MailMerge/MailMerge.h>
#include <Aspose.Words.Cpp/Model/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Model/Text/Paragraph.h>
#include <system/array.h>
#include <system/collections/list.h>
#include <system/convert.h>
#include <system/details/pointer_collection_helpers.h>
#include <system/enumerator_adapter.h>
#include <system/object_ext.h>
#include <system/scope_guard.h>

#include "ApiExampleBase.h"
#include "TestUtil.h"

using System::String;
using System::SharedPtr;
using System::ArrayPtr;
using System::MakeObject;
using System::MakeArray;

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;

namespace ApiExamples {

class ExMailMergeCustomNested : public ApiExampleBase
{
public:
    class OrderList;
    class OrderMailMergeDataSource;

public:
    //ExStart
    //ExFor:MailMerge.ExecuteWithRegions(IMailMergeDataSource)
    //ExSummary:Performs mail merge with regions from a custom data source.
    void CustomDataSource()
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

        SharedPtr<ExMailMergeCustomNested::OrderList> get_Orders()
        {
            return pr_Orders;
        }

        void set_Orders(SharedPtr<ExMailMergeCustomNested::OrderList> value)
        {
            pr_Orders = value;
        }

        Customer(String aFullName, String anAddress)
        {
            // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            System::Details::ThisProtector __local_self_ref(this);
            //---------------------------------------------------------Self Reference

            set_FullName(aFullName);
            set_Address(anAddress);
            set_Orders(MakeObject<ExMailMergeCustomNested::OrderList>());
        }

    private:
        String pr_FullName;
        String pr_Address;
        SharedPtr<ExMailMergeCustomNested::OrderList> pr_Orders;
    };

    /// <summary>
    /// An example of a typed collection that contains your "data" objects.
    /// </summary>
    class CustomerList : public System::Collections::Generic::List<SharedPtr<ApiExamples::ExMailMergeCustomNested::Customer>>
    {
    public:
        void SetTemplateWeakPtr(unsigned int argument) override;

    protected:
        virtual ~CustomerList()
        {
        }
    };

    /// <summary>
    /// An example of a child "data entity" class in your application.
    /// </summary>
    class Order : public System::Object
    {
    public:
        String get_Name()
        {
            return pr_Name;
        }

        void set_Name(String value)
        {
            pr_Name = value;
        }

        int get_Quantity()
        {
            return pr_Quantity;
        }

        void set_Quantity(int value)
        {
            pr_Quantity = value;
        }

        Order(String oName, int oQuantity) : pr_Quantity(0)
        {
            // Self Reference+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            System::Details::ThisProtector __local_self_ref(this);
            //---------------------------------------------------------Self Reference

            set_Name(oName);
            set_Quantity(oQuantity);
        }

    private:
        String pr_Name;
        int pr_Quantity;
    };

    /// <summary>
    /// An example of a typed collection that contains your "data" objects.
    /// </summary>
    class OrderList : public System::Collections::Generic::List<SharedPtr<ApiExamples::ExMailMergeCustomNested::Order>>
    {
    public:
        void SetTemplateWeakPtr(unsigned int argument) override;

    protected:
        virtual ~OrderList()
        {
        }
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

        CustomerMailMergeDataSource(SharedPtr<ExMailMergeCustomNested::CustomerList> customers) : mRecordIndex(0)
        {
            mCustomers = customers;

            // When the data source is initialized, it must be positioned before the first record
            mRecordIndex = -1;
        }

        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        bool GetValue(String fieldName, SharedPtr<System::Object>& fieldValue) override
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

    private:
        bool get_IsEof()
        {
            return mRecordIndex >= mCustomers->get_Count();
        }

        SharedPtr<ExMailMergeCustomNested::CustomerList> mCustomers;
        int mRecordIndex;
    };

    class OrderMailMergeDataSource : public IMailMergeDataSource
    {
    public:
        /// <summary>
        /// The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
        /// </summary>
        String get_TableName() override
        {
            return u"Order";
        }

        OrderMailMergeDataSource(SharedPtr<ExMailMergeCustomNested::OrderList> orders) : mRecordIndex(0)
        {
            mOrders = orders;

            // When the data source is initialized, it must be positioned before the first record
            mRecordIndex = -1;
        }

        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        bool GetValue(String fieldName, SharedPtr<System::Object>& fieldValue) override
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

        /// <summary>
        /// Return null because we do not have any child elements for this sort of object.
        /// </summary>
        SharedPtr<IMailMergeDataSource> GetChildDataSource(String tableName) override
        {
            return nullptr;
        }

    private:
        bool get_IsEof()
        {
            return mRecordIndex >= mOrders->get_Count();
        }

        SharedPtr<ExMailMergeCustomNested::OrderList> mOrders;
        int mRecordIndex;
    };
    //ExEnd

protected:
    void TestCustomDataSource(SharedPtr<ExMailMergeCustomNested::CustomerList> customers, SharedPtr<Document> doc)
    {
        SharedPtr<System::Collections::Generic::List<ArrayPtr<String>>> mailMergeData =
            MakeObject<System::Collections::Generic::List<ArrayPtr<String>>>();

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
};

} // namespace ApiExamples
