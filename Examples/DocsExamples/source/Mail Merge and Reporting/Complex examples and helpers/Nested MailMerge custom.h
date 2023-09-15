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
#include <SQLiteCpp.h>

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
        //ExStart:NestedMailMerge
        //GistId:3435df005db9907ec9ba3a6b777ae6fb
        void NestedMailMerge()
        {            
            auto doc = MakeObject<Document>();
            auto builder = MakeObject<DocumentBuilder>(doc);
            builder->InsertField(u" MERGEFIELD TableStart:Customers");

            builder->Write(u"Full name:\t");
            builder->InsertField(u" MERGEFIELD FullName ");
            builder->Write(u"\nAddress:\t");
            builder->InsertField(u" MERGEFIELD Address ");
            builder->Write(u"\nOrders:\n");

            builder->InsertField(u" MERGEFIELD TableStart:Orders");

            builder->Write(u"\tItem name:\t");
            builder->InsertField(u" MERGEFIELD ItemName ");
            builder->Write(u"\n\tQuantity:\t");
            builder->InsertField(u" MERGEFIELD Quantity ");
            builder->InsertParagraph();

            builder->InsertField(u" MERGEFIELD TableEnd:Orders");

            builder->InsertField(u" MERGEFIELD TableEnd:Customers");            

            SQLite::Database database((std::string)(DatabaseDir + u"customers.db3"));

            // To be able to mail merge from your data source,
            // it must be wrapped into an object that implements the IMailMergeDataSource interface.
            auto customersDataSource = MakeObject<NestedMailMergeCustom::CustomersDataSource>(database);            

            doc->get_MailMerge()->ExecuteWithRegions(customersDataSource);

            doc->Save(ArtifactsDir + u"NestedMailMergeCustom.CustomMailMerge.docx");            
        }

        class OrdersDataSource : public MailMerging::IMailMergeDataSource
        {
        public:
            OrdersDataSource(SQLite::Database& database, int32_t id)
                : mQuery(database, "SELECT ItemName, Quantity FROM Orders WHERE CustomerId = ?")
            {
                mQuery.bind(1, id);
            }

            String get_TableName() override
            {
                return u"Orders";
            }

            bool GetValue(String fieldName, SharedPtr<Object>& fieldValue) override
            {

                if (fieldName == u"ItemName")
                {
                    fieldValue = System::ObjectExt::Box<String>(String::FromUtf8(mQuery.getColumn(0).getString()));
                    return true;
                }

                if (fieldName == u"Quantity")
                {
                    fieldValue = System::ObjectExt::Box<int32_t>(mQuery.getColumn(1).getInt());
                    return true;
                }

                fieldValue.reset();
                return false;
            }

            bool MoveNext() override
            {
                return mQuery.executeStep();
            }

            SharedPtr<IMailMergeDataSource> GetChildDataSource(String tableName) override
            {
                return nullptr;
            }

        private:
            SQLite::Statement mQuery;
        };

        class CustomersDataSource : public MailMerging::IMailMergeDataSource
        {
        public:
            CustomersDataSource(SQLite::Database& database)
                : mDatabase{ database }
                , mQuery{ mDatabase, "SELECT * FROM Customers" }
            {
            }

            String get_TableName() override
            {
                return u"Customers";
            }

            bool GetValue(String fieldName, SharedPtr<Object>& fieldValue) override
            {
                if (fieldName == u"FullName")
                {
                    fieldValue = System::ObjectExt::Box<String>(String::FromUtf8(mQuery.getColumn(1).getString()));
                    return true;
                }

                if (fieldName == u"Address")
                {
                    fieldValue = System::ObjectExt::Box<String>(String::FromUtf8(mQuery.getColumn(2).getString()));
                    return true;
                }

                fieldValue.reset();
                return false;
            }

            bool MoveNext() override
            {
                return mQuery.executeStep();
            }

            SharedPtr<IMailMergeDataSource> GetChildDataSource(String tableName) override
            {
                if (tableName == u"Orders")
                {
                    return MakeObject<OrdersDataSource>(mDatabase, mQuery.getColumn(0).getInt());
                }
                return nullptr;
            }
        private:
            SQLite::Database& mDatabase;
            SQLite::Statement mQuery;
        };
        //ExEnd:NestedMailMerge
    };

}}} // namespace DocsExamples::Mail_Merge_and_Reporting::Custom_examples
