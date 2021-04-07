#pragma once
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
#include <system/enumerator_adapter.h>
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

class ExMailMergeCustomNested : public ApiExampleBase
{
public:
    class Order;
    class OrderMailMergeDataSource;

public:
    //ExStart
    //ExFor:MailMerge.ExecuteWithRegions(IMailMergeDataSource)
    //ExSummary:Shows how to use mail merge regions to execute a nested mail merge.
    void CustomDataSource()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Normally, MERGEFIELDs contain the name of a column of a mail merge data source.
        // Instead, we can use "TableStart:" and "TableEnd:" prefixes to begin/end a mail merge region.
        // Each region will belong to a table with a name that matches the string immediately after the prefix's colon.
        builder->InsertField(u" MERGEFIELD TableStart:Customers");

        // These MERGEFIELDs are inside the mail merge region of the "Customers" table.
        // When we execute the mail merge, this field will receive data from rows in a data source named "Customers".
        builder->Write(u"Full name:\t");
        builder->InsertField(u" MERGEFIELD FullName ");
        builder->Write(u"\nAddress:\t");
        builder->InsertField(u" MERGEFIELD Address ");
        builder->Write(u"\nOrders:\n");

        // Create a second mail merge region inside the outer region for a data source named "Orders".
        // The "Orders" data entries have a many-to-one relationship with the "Customers" data source.
        builder->InsertField(u" MERGEFIELD TableStart:Orders");

        builder->Write(u"\tItem name:\t");
        builder->InsertField(u" MERGEFIELD Name ");
        builder->Write(u"\n\tQuantity:\t");
        builder->InsertField(u" MERGEFIELD Quantity ");
        builder->InsertParagraph();

        builder->InsertField(u" MERGEFIELD TableEnd:Orders");
        builder->InsertField(u" MERGEFIELD TableEnd:Customers");

        // Create related data with names that match those of our mail merge regions.
        SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Customer>>> customers =
            MakeObject<System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Customer>>>();
        customers->Add(MakeObject<ExMailMergeCustomNested::Customer>(u"Thomas Hardy", u"120 Hanover Sq., London"));
        customers->Add(MakeObject<ExMailMergeCustomNested::Customer>(u"Paolo Accorti", u"Via Monte Bianco 34, Torino"));

        customers->idx_get(0)->get_Orders()->Add(MakeObject<ExMailMergeCustomNested::Order>(u"Rugby World Cup Cap", 2));
        customers->idx_get(0)->get_Orders()->Add(MakeObject<ExMailMergeCustomNested::Order>(u"Rugby World Cup Ball", 1));
        customers->idx_get(1)->get_Orders()->Add(MakeObject<ExMailMergeCustomNested::Order>(u"Rugby World Cup Guide", 1));

        // To mail merge from your data source, we must wrap it into an object that implements the IMailMergeDataSource interface.
        auto customersDataSource = MakeObject<ExMailMergeCustomNested::CustomerMailMergeDataSource>(customers);

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

        SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Order>>> get_Orders()
        {
            return pr_Orders;
        }

        void set_Orders(SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Order>>> value)
        {
            pr_Orders = value;
        }

        Customer(String aFullName, String anAddress)
        {
            set_FullName(aFullName);
            set_Address(anAddress);
            set_Orders(MakeObject<System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Order>>>());
        }

    private:
        String pr_FullName;
        String pr_Address;
        SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Order>>> pr_Orders;
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
            set_Name(oName);
            set_Quantity(oQuantity);
        }

    private:
        String pr_Name;
        int pr_Quantity;
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
            return u"Customers";
        }

        CustomerMailMergeDataSource(SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Customer>>> customers) : mRecordIndex(0)
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
            else if (fieldName == u"Order")
            {
                fieldValue = mCustomers->idx_get(mRecordIndex)->get_Orders();
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
            if (tableName == u"Orders")
            {
                return MakeObject<ExMailMergeCustomNested::OrderMailMergeDataSource>(mCustomers->idx_get(mRecordIndex)->get_Orders());
            }
            else
            {
                return nullptr;
            }
        }

    private:
        bool get_IsEof()
        {
            return mRecordIndex >= mCustomers->get_Count();
        }

        SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Customer>>> mCustomers;
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
            return u"Orders";
        }

        OrderMailMergeDataSource(SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Order>>> orders) : mRecordIndex(0)
        {
            mOrders = orders;

            // When we initialize the data source, its position must be before the first record.
            mRecordIndex = -1;
        }

        /// <summary>
        /// Aspose.Words calls this method to get a value for every data field.
        /// </summary>
        bool GetValue(String fieldName, SharedPtr<System::Object>& fieldValue) override
        {
            if (fieldName == u"Name")
            {
                fieldValue = System::ObjectExt::Box<String>(mOrders->idx_get(mRecordIndex)->get_Name());
                return true;
            }
            else if (fieldName == u"Quantity")
            {
                fieldValue = System::ObjectExt::Box<int>(mOrders->idx_get(mRecordIndex)->get_Quantity());
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

        SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Order>>> mOrders;
        int mRecordIndex;
    };
    //ExEnd

    void TestCustomDataSource(SharedPtr<System::Collections::Generic::List<SharedPtr<ExMailMergeCustomNested::Customer>>> customers, SharedPtr<Document> doc)
    {
        SharedPtr<System::Collections::Generic::List<ArrayPtr<String>>> mailMergeData = MakeObject<System::Collections::Generic::List<ArrayPtr<String>>>();

        for (auto customer : customers)
        {
            for (auto order : customer->get_Orders())
            {
                mailMergeData->Add(MakeArray<String>({order->get_Name(), System::Convert::ToString(order->get_Quantity())}));
            }
            mailMergeData->Add(MakeArray<String>({customer->get_FullName(), customer->get_Address()}));
        }

        TestUtil::MailMergeMatchesArray(mailMergeData->ToArray(), doc, false);
    }
};

} // namespace ApiExamples
