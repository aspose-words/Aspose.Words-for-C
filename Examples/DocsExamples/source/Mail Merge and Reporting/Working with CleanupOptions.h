#pragma once

#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldMergeField.h>
#include <Aspose.Words.Cpp/MailMerging/IMailMergeDataSource.h>
#include <Aspose.Words.Cpp/MailMerging/MailMerge.h>
#include <Aspose.Words.Cpp/MailMerging/MailMergeCleanupOptions.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <system/array.h>
#include <system/enum_helpers.h>
#include <system/exceptions.h>
#include <system/object_ext.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;
using namespace Aspose::Words::MailMerging;

namespace DocsExamples { namespace Mail_Merge_and_Reporting {

class WorkingWithCleanupOptions : public DocsExamplesBase
{

public:
    void CleanupParagraphsWithPunctuationMarks()
    {
        //ExStart:CleanupParagraphsWithPunctuationMarks
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto mergeFieldOption1 = System::ExplicitCast<FieldMergeField>(builder->InsertField(u"MERGEFIELD", u"Option_1"));
        mergeFieldOption1->set_FieldName(u"Option_1");

        // Here is the complete list of cleanable punctuation marks: ! , . : ; ? ¡ ¿.
        builder->Write(u" ?  ");

        auto mergeFieldOption2 = System::ExplicitCast<FieldMergeField>(builder->InsertField(u"MERGEFIELD", u"Option_2"));
        mergeFieldOption2->set_FieldName(u"Option_2");

        doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveEmptyParagraphs);
        // The option's default value is true, which means that the behavior was changed to mimic MS Word.
        // If you rely on the old behavior can revert it by setting the option to false.
        doc->get_MailMerge()->set_CleanupParagraphsWithPunctuationMarks(true);

        doc->get_MailMerge()->Execute(MakeArray<String>({u"Option_1", u"Option_2"}), MakeArray<SharedPtr<System::Object>>({nullptr, nullptr}));

        doc->Save(ArtifactsDir + u"WorkingWithCleanupOptions.CleanupParagraphsWithPunctuationMarks.docx");
        //ExEnd:CleanupParagraphsWithPunctuationMarks
    }

    void RemoveEmptyParagraphs()
    {
        //ExStart:RemoveEmptyParagraphs
        //GistId:f39874821cb317d245a769c9ce346fea
        auto doc = MakeObject<Document>(MyDir + u"Table with fields.docx");

        doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveEmptyParagraphs);

        doc->get_MailMerge()->Execute(
            MakeArray<String>({u"FullName", u"Company", u"Address", u"Address2", u"City"}),
            MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(u"James Bond"), System::ObjectExt::Box<String>(u"MI5 Headquarters"),
                                                  System::ObjectExt::Box<String>(u"Milbank"), System::ObjectExt::Box<String>(u""),
                                                  System::ObjectExt::Box<String>(u"London")}));

        doc->Save(ArtifactsDir + u"WorkingWithCleanupOptions.RemoveEmptyParagraphs.docx");
        //ExEnd:RemoveEmptyParagraphs
    }

    void RemoveUnusedFields()
    {
        //ExStart:RemoveUnusedFields
        //GistId:f39874821cb317d245a769c9ce346fea
        auto doc = MakeObject<Document>(MyDir + u"Table with fields.docx");

        doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveUnusedFields);

        doc->get_MailMerge()->Execute(
            MakeArray<String>({u"FullName", u"Company", u"Address", u"Address2", u"City"}),
            MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(u"James Bond"), System::ObjectExt::Box<String>(u"MI5 Headquarters"),
                                                  System::ObjectExt::Box<String>(u"Milbank"), System::ObjectExt::Box<String>(u""),
                                                  System::ObjectExt::Box<String>(u"London")}));

        doc->Save(ArtifactsDir + u"WorkingWithCleanupOptions.RemoveUnusedFields.docx");
        //ExEnd:RemoveUnusedFields
    }

    void RemoveContainingFields()
    {
        //ExStart:RemoveContainingFields
        //GistId:f39874821cb317d245a769c9ce346fea
        auto doc = MakeObject<Document>(MyDir + u"Table with fields.docx");

        doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveContainingFields);

        doc->get_MailMerge()->Execute(
            MakeArray<String>({u"FullName", u"Company", u"Address", u"Address2", u"City"}),
            MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(u"James Bond"), System::ObjectExt::Box<String>(u"MI5 Headquarters"),
                                                  System::ObjectExt::Box<String>(u"Milbank"), System::ObjectExt::Box<String>(u""),
                                                  System::ObjectExt::Box<String>(u"London")}));

        doc->Save(ArtifactsDir + u"WorkingWithCleanupOptions.RemoveContainingFields.docx");
        //ExEnd:RemoveContainingFields
    }

    void RemoveEmptyTableRows()
    {
        //ExStart:RemoveEmptyTableRows
        //GistId:f39874821cb317d245a769c9ce346fea
        auto doc = MakeObject<Document>(MyDir + u"Table with fields.docx");

        doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveEmptyTableRows);

        doc->get_MailMerge()->Execute(
            MakeArray<String>({u"FullName", u"Company", u"Address", u"Address2", u"City"}),
            MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<String>(u"James Bond"), System::ObjectExt::Box<String>(u"MI5 Headquarters"),
                                                  System::ObjectExt::Box<String>(u"Milbank"), System::ObjectExt::Box<String>(u""),
                                                  System::ObjectExt::Box<String>(u"London")}));

        doc->Save(ArtifactsDir + u"WorkingWithCleanupOptions.RemoveEmptyTableRows.docx");
        //ExEnd:RemoveEmptyTableRows
    }

    void RemoveUnmergedRegions()
    {
        //ExStart:RemoveUnmergedRegions
        //GistId:f39874821cb317d245a769c9ce346fea
        auto doc = MakeObject<Document>(MyDir + u"Mail merge destination - Northwind suppliers.docx");

        auto data = MakeObject<EmptyDataSource>();
        //ExStart:MailMergeCleanupOptions
        doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveUnusedRegions);
        // doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveContainingFields);
        // doc->get_MailMerge()->set_CleanupOptions(doc->get_MailMerge()->get_CleanupOptions() | MailMergeCleanupOptions::RemoveStaticFields);
        // doc->get_MailMerge()->set_CleanupOptions(doc->get_MailMerge()->get_CleanupOptions() | MailMergeCleanupOptions::RemoveEmptyParagraphs);
        // doc->get_MailMerge()->set_CleanupOptions(doc->get_MailMerge()->get_CleanupOptions() | MailMergeCleanupOptions::RemoveUnusedFields);
        //ExEnd:MailMergeCleanupOptions

        // Merge the data with the document by executing mail merge which will have no effect as there is no data.
        // However the regions found in the document will be removed automatically as they are unused.
        doc->get_MailMerge()->ExecuteWithRegions(data);

        doc->Save(ArtifactsDir + u"WorkingWithCleanupOptions.RemoveUnmergedRegions.docx");
        //ExEnd:RemoveUnmergedRegions
    }

    void RemoveRowsFromTable()
    {
        //ExStart:RemoveRowsFromTable
        auto doc = MakeObject<Document>(MyDir + u"Mail merge destination - Northwind suppliers.docx");

        auto data = MakeObject<EmptyDataSource>();
        doc->get_MailMerge()->set_CleanupOptions(MailMergeCleanupOptions::RemoveUnusedRegions |
                                                 MailMergeCleanupOptions::RemoveEmptyTableRows);

        doc->get_MailMerge()->set_MergeDuplicateRegions(true);
        doc->get_MailMerge()->ExecuteWithRegions(data);

        doc->Save(ArtifactsDir + u"WorkingWithCleanupOptions.RemoveRowsFromTable.docx");
        //ExEnd:RemoveRowsFromTable
    }

private:
    /// <summary>
    /// Stands in for the empty .NET DataSet the cleanup examples merge against.
    /// Aspose.Words for C++ has no System.Data, so mail merge with regions always
    /// takes an IMailMergeDataSource - one that yields no rows is enough here.
    /// </summary>
    class EmptyDataSource : public IMailMergeDataSource
    {
    public:
        String get_TableName() override
        {
            return String::Empty;
        }

        bool MoveNext() override
        {
            return false;
        }

        bool GetValue(String fieldName, SharedPtr<System::Object>& fieldValue) override
        {
            ASPOSE_UNUSED(fieldName);
            fieldValue.reset();

            return false;
        }

        SharedPtr<IMailMergeDataSource> GetChildDataSource(String tableName) override
        {
            ASPOSE_UNUSED(tableName);
            return nullptr;
        }
    };
};

}} // namespace DocsExamples::Mail_Merge_and_Reporting
