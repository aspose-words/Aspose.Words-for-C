#pragma once

#include <cstdint>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/MailMerging/IMailMergeDataSource.h>
#include <Aspose.Words.Cpp/MailMerging/MailMerge.h>
#include <Aspose.Words.Cpp/MailMerging/MailMergeRegionInfo.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <system/array.h>
#include <system/collections/ilist.h>
#include <system/enumerator_adapter.h>
#include <system/object_ext.h>
#include <system/test_tools/compare.h>
#include <system/test_tools/test_tools.h>
#include <testing/test_predicates.h>
#include <SQLiteCpp.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::MailMerging;

namespace DocsExamples { namespace Mail_Merge_and_Reporting {

class BaseOperations : public DocsExamplesBase
{
public:
    void SimpleMailMerge()
    {
        //ExStart:SimpleMailMerge
        //GistId:3435df005db9907ec9ba3a6b777ae6fb
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        
        builder->InsertField(u" MERGEFIELD CustomerName ");
        builder->InsertParagraph();
        builder->InsertField(u" MERGEFIELD Item ");
        builder->InsertParagraph();
        builder->InsertField(u" MERGEFIELD Quantity ");

        auto boxHelper = [](const char16_t* value) { return System::ObjectExt::Box<String>(value); };

        // Fill the fields in the document with user data.
        doc->get_MailMerge()->Execute(
            MakeArray<String>({u"CustomerName", u"Item", u"Quantity"}),
            MakeArray<SharedPtr<System::Object>>(
                { boxHelper(u"John Doe"), boxHelper(u"Hawaiian"), boxHelper(u"2")}));

        doc->Save(ArtifactsDir + u"BaseOperations.SimpleMailMerge.docx");
        //ExEnd:SimpleMailMerge
    }

    void UseIfElseMustache()
    {
        //ExStart:UseOfifelseMustacheSyntax
        auto doc = MakeObject<Document>(MyDir + u"Mail merge destinations - Mustache syntax.docx");

        doc->get_MailMerge()->set_UseNonMergeFields(true);
        doc->get_MailMerge()->Execute(MakeArray<String>({u"GENDER"}), MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(u"MALE")}));

        doc->Save(ArtifactsDir + u"BaseOperations.IfElseMustache.docx");
        //ExEnd:UseOfifelseMustacheSyntax
    }

    void GetRegionsByName()
    {
        //ExStart:GetRegionsByName
        auto doc = MakeObject<Document>(MyDir + u"Mail merge regions.docx");

        SharedPtr<System::Collections::Generic::IList<SharedPtr<MailMergeRegionInfo>>> regions = doc->get_MailMerge()->GetRegionsByName(u"Region1");
        ASSERT_EQ(1, doc->get_MailMerge()->GetRegionsByName(u"Region1")->get_Count());
        for (const auto& region : System::IterateOver(regions))
        {
            ASSERT_EQ(u"Region1", region->get_Name());
        }

        regions = doc->get_MailMerge()->GetRegionsByName(u"Region2");
        ASSERT_EQ(1, doc->get_MailMerge()->GetRegionsByName(u"Region2")->get_Count());
        for (const auto& region : System::IterateOver(regions))
        {
            ASSERT_EQ(u"Region2", region->get_Name());
        }

        regions = doc->get_MailMerge()->GetRegionsByName(u"NestedRegion1");
        ASSERT_EQ(2, doc->get_MailMerge()->GetRegionsByName(u"NestedRegion1")->get_Count());
        for (const auto& region : System::IterateOver(regions))
        {
            ASSERT_EQ(u"NestedRegion1", region->get_Name());
        }
        //ExEnd:GetRegionsByName
    }

    //ExStart:MultipleDocumentsMailMerge
    //GistId:3435df005db9907ec9ba3a6b777ae6fb
    void MultipleDocumentsMailMerge()
    {
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u" MERGEFIELD FullName ");
        builder->InsertParagraph();
        builder->InsertField(u" MERGEFIELD Address ");

        // Fill the fields in the document with user data.
        SQLite::Database database{ (std::string)(DatabaseDir + u"customers.db3") };
        SQLite::Statement query{ database, "SELECT * FROM Customers" };
        auto dataSource = MakeObject<CustomersRowMailMergeDataSource>(query);

        int32_t counter = 1;
        while (query.executeStep())
        {
            dataSource->Reset();
            doc->get_MailMerge()->Execute(dataSource);
            doc->Save(ArtifactsDir + u"BaseOperations.MultipleDocumentsMailMerge_" + counter++ + u".docx");
        }
    }

    class CustomersRowMailMergeDataSource : public MailMerging::IMailMergeDataSource
    {
    public:
        CustomersRowMailMergeDataSource(SQLite::Statement& query) : mQuery(query) {}

        String get_TableName() override
        {
            return u"Customers";
        }

        bool GetValue(String fieldName, SharedPtr<Object>& fieldValue) override
        {
            auto boxHelper = [](const std::string& value) { return System::ObjectExt::Box<String>(String::FromUtf8(value)); };

            if (fieldName == u"FullName")
            {
                fieldValue = boxHelper(mQuery.getColumn(1).getString());
                return true;
            }

            if (fieldName == u"Address")
            {
                fieldValue = boxHelper(mQuery.getColumn(2).getString());
                return true;
            }

            fieldValue.reset();
            return false;
        }

        bool MoveNext() override
        {
            if (!isInitialized)
            {
                isInitialized = true;
                return true;
            }
            return false;
        }

        SharedPtr<IMailMergeDataSource> GetChildDataSource(String tableName) override
        {
            return nullptr;
        }

        void Reset()
        {
            isInitialized = false;
        }

    private:
        SQLite::Statement& mQuery;
        bool isInitialized = false;
    };
    //ExEnd:MultipleDocumentsMailMerge
};

}} // namespace DocsExamples::Mail_Merge_and_Reporting
