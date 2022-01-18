#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/MailMerging/IMailMergeDataSource.h>
#include <Aspose.Words.Cpp/MailMerging/MailMerge.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <system/collections/list.h>
#include <system/object_ext.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;

namespace DocsExamples { namespace Mail_Merge_and_Reporting { namespace Custom_examples {

class NestedMailMergeCustom : public DocsExamplesBase
{
public:
    class Order;
    class OrderMailMergeDataSource;

public:
    void CustomMailMerge()
    {
        //ExStart:NestedMailMergeCustom
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

        SharedPtr<System::Collections::Generic::List<SharedPtr<NestedMailMergeCustom::Customer>>> customers =
            MakeObject<System::Collections::Generic::List<SharedPtr<NestedMailMergeCustom::Customer>>>();
        customers->Add(MakeObject<NestedMailMergeCustom::Customer>(u"Thomas Hardy", u"120 Hanover Sq., London"));
        customers->Add(MakeObject<NestedMailMergeCustom::Customer>(u"Paolo Accorti", u"Via Monte Bianco 34, Torino"));

        customers->idx_get(0)->get_Orders()->Add(MakeObject<NestedMailMergeCustom::Order>(u"Rugby World Cup Cap", 2));
        customers->idx_get(0)->get_Orders()->Add(MakeObject<NestedMailMergeCustom::Order>(u"Rugby World Cup Ball", 1));
        customers->idx_get(1)->get_Orders()->Add(MakeObject<NestedMailMergeCustom::Order>(u"Rugby World Cup Guide", 1));

        // To be able to mail merge from your data source,
        // it must be wrapped into an object that implements the IMailMergeDataSource interface.
        auto customersDataSource = MakeObject<NestedMailMergeCustom::CustomerMailMergeDataSource>(customers);

        doc->get_MailMerge()->ExecuteWithRegions(customersDataSource);

        doc->Save(ArtifactsDir + u"NestedMailMergeCustom.CustomMailMerge.docx");
        //ExEnd:NestedMailMergeCustom
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

        SharedPtr<System::Collections::Generic::List<SharedPtr<NestedMailMergeCustom::Order>>> get_Orders()
        {
            return pr_Orders;
        }

        void set_Orders(SharedPtr<System::Collections::Generic::List<SharedPtr<NestedMailMergeCustom::Order>>> value)
        {
            pr_Orders = value;
        }

        Customer(String aFullName, String anAddress)
        {
            set_FullName(aFullName);
            set_Address(anAddress);
            set_Orders(MakeObject<System::Collections::Generic::List<SharedPtr<NestedMailMergeCustom::Order>>>());
        }

    private:
        String pr_FullName;
        String pr_Address;
        SharedPtr<System::Collections::Generic::List<SharedPtr<NestedMailMergeCustom::Order>>> pr_Orders;
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
            return u"Customer";
        }

        CustomerMailMergeDataSource(SharedPtr<System::Collections::Generic::List<SharedPtr<NestedMailMergeCustom::Customer>>> customers) : mRecordIndex(0)
        {
            mCustomers = customers;

            // When the data source is initialized, it must be positioned before the first record.
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
            if (tableName == u"Order")
            {
                return MakeObject<NestedMailMergeCustom::OrderMailMergeDataSource>(mCustomers->idx_get(mRecordIndex)->get_Orders());
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

        SharedPtr<System::Collections::Generic::List<SharedPtr<NestedMailMergeCustom::Customer>>> mCustomers;
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

        OrderMailMergeDataSource(SharedPtr<System::Collections::Generic::List<SharedPtr<NestedMailMergeCustom::Order>>> orders) : mRecordIndex(0)
        {
            mOrders = orders;

            // When the data source is initialized, it must be positioned before the first record.
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

        SharedPtr<IMailMergeDataSource> GetChildDataSource(String tableName) override
        {
            ASPOSE_UNUSED(tableName);
            // Return null because we haven't any child elements for this sort of object.
            return nullptr;
        }

    private:
        bool get_IsEof()
        {
            return mRecordIndex >= mOrders->get_Count();
        }

        SharedPtr<System::Collections::Generic::List<SharedPtr<NestedMailMergeCustom::Order>>> mOrders;
        int mRecordIndex;
    };
};

}}} // namespace DocsExamples::Mail_Merge_and_Reporting::Custom_examples
