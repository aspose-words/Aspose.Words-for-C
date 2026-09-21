#pragma once

#include <cstdint>
#include <functional>
#include <iostream>
#include <Aspose.Words.Cpp/Body.h>
#include <Aspose.Words.Cpp/BreakType.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/Field.h>
#include <Aspose.Words.Cpp/Fields/FieldAddressBlock.h>
#include <Aspose.Words.Cpp/Fields/FieldAdvance.h>
#include <Aspose.Words.Cpp/Fields/FieldAsk.h>
#include <Aspose.Words.Cpp/Fields/FieldAuthor.h>
#include <Aspose.Words.Cpp/Fields/FieldCollection.h>
#include <Aspose.Words.Cpp/Fields/FieldEnd.h>
#include <Aspose.Words.Cpp/Fields/FieldHyperlink.h>
#include <Aspose.Words.Cpp/Fields/FieldIf.h>
#include <Aspose.Words.Cpp/Fields/FieldIfComparisonResult.h>
#include <Aspose.Words.Cpp/Fields/FieldIncludeText.h>
#include <Aspose.Words.Cpp/Fields/FieldMergeField.h>
#include <Aspose.Words.Cpp/Fields/FieldOptions.h>
#include <Aspose.Words.Cpp/Fields/FieldSeparator.h>
#include <Aspose.Words.Cpp/Fields/FieldStart.h>
#include <Aspose.Words.Cpp/Fields/FieldTA.h>
#include <Aspose.Words.Cpp/Fields/FieldToa.h>
#include <Aspose.Words.Cpp/Fields/FieldType.h>
#include <Aspose.Words.Cpp/Fields/FieldUnknown.h>
#include <Aspose.Words.Cpp/Fields/FieldUpdateCultureSource.h>
#include <Aspose.Words.Cpp/Fields/IFieldUpdateCultureProvider.h>
#include <Aspose.Words.Cpp/Font.h>
#include <Aspose.Words.Cpp/Fields/FieldArgumentBuilder.h>
#include <Aspose.Words.Cpp/Fields/FieldBuilder.h>
#include <Aspose.Words.Cpp/Fields/IFieldResultFormatter.h>
#include <Aspose.Words.Cpp/Fields/GeneralFormat.h>
#include <Aspose.Words.Cpp/CalendarType.h>
#include <Aspose.Words.Cpp/HeaderFooterType.h>
#include <Aspose.Words.Cpp/MailMerging/MailMerge.h>
#include <Aspose.Words.Cpp/MailMerging/MappedDataFieldCollection.h>
#include <Aspose.Words.Cpp/CompositeNode.h>
#include <Aspose.Words.Cpp/Node.h>
#include <Aspose.Words.Cpp/NodeCollection.h>
#include <Aspose.Words.Cpp/NodeType.h>
#include <Aspose.Words.Cpp/Paragraph.h>
#include <Aspose.Words.Cpp/Range.h>
#include <Aspose.Words.Cpp/Run.h>
#include <Aspose.Words.Cpp/Saving/SaveOutputParameters.h>
#include <Aspose.Words.Cpp/Section.h>
#include <system/action.h>
#include <system/array.h>
#include <system/collections/ienumerable.h>
#include <system/collections/list.h>
#include <system/date_time.h>
#include <system/enum.h>
#include <system/enumerator_adapter.h>
#include <system/exceptions.h>
#include <system/func.h>
#include <system/globalization/culture_info.h>
#include <system/globalization/date_time_format_info.h>
#include <system/linq/enumerable.h>
#include <system/object_ext.h>
#include <system/text/regularexpressions/group.h>
#include <system/text/regularexpressions/group_collection.h>
#include <system/text/regularexpressions/match.h>
#include <system/text/regularexpressions/regex.h>
#include <system/threading/thread.h>

#include "DocsExamplesBase.h"

using System::ArrayPtr;
using System::MakeArray;
using System::MakeObject;
using System::SharedPtr;
using System::String;

using namespace Aspose::Words;
using namespace Aspose::Words::Fields;

namespace DocsExamples { namespace Programming_with_Documents {

class WorkingWithFields : public DocsExamplesBase
{
public:
    void ChangeFieldUpdateCultureSource()
    {
        //ExStart:ChangeFieldUpdateCultureSource
        //GistId:9e90defe4a7bcafb004f73a2ef236986
        //ExStart:DocumentBuilderInsertField
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        // Insert content with German locale.
        builder->get_Font()->set_LocaleId(1031);
        builder->InsertField(u"MERGEFIELD Date1 \\@ \"dddd, d MMMM yyyy\"");
        builder->Write(u" - ");
        builder->InsertField(u"MERGEFIELD Date2 \\@ \"dddd, d MMMM yyyy\"");
        //ExEnd:DocumentBuilderInsertField

        // Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from
        // set the culture used during field update to the culture used by the field.
        doc->get_FieldOptions()->set_FieldUpdateCultureSource(FieldUpdateCultureSource::FieldCode);
        doc->get_MailMerge()->Execute(MakeArray<String>({u"Date2"}),
                                      MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<System::DateTime>(System::DateTime(2011, 1, 1))}));

        doc->Save(ArtifactsDir + u"WorkingWithFields.ChangeFieldUpdateCultureSource.docx");
        //ExEnd:ChangeFieldUpdateCultureSource
    }

    void SpecifyLocaleAtFieldLevel()
    {
        //ExStart:SpecifyLocaleAtFieldLevel
        //GistId:1cf07762df56f15067d6aef90b14b3db
        auto builder = MakeObject<DocumentBuilder>();

        SharedPtr<Field> field = builder->InsertField(FieldType::FieldDate, true);
        field->set_LocaleId(1049);

        builder->get_Document()->Save(ArtifactsDir + u"WorkingWithFields.SpecifylocaleAtFieldlevel.docx");
        //ExEnd:SpecifyLocaleAtFieldLevel
    }

    void ReplaceHyperlinks()
    {
        //ExStart:ReplaceHyperlinks
        //GistId:0213851d47551e83af42233f4d075cf6
        auto doc = MakeObject<Document>(MyDir + u"Hyperlinks.docx");

        for (const auto& field : System::IterateOver(doc->get_Range()->get_Fields()))
        {
            if (field->get_Type() == FieldType::FieldHyperlink)
            {
                auto hyperlink = System::ExplicitCast<FieldHyperlink>(field);

                // Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
                if (hyperlink->get_SubAddress() != nullptr)
                {
                    continue;
                }

                hyperlink->set_Address(u"http://www.aspose.com");
                hyperlink->set_Result(u"Aspose - The .NET & Java Component Publisher");
            }
        }

        doc->Save(ArtifactsDir + u"WorkingWithFields.ReplaceHyperlinks.docx");
        //ExEnd:ReplaceHyperlinks
    }

    void RenameMergeFields()
    {
        //ExStart:RenameMergeFields
        //GistId:bf0f8a6b40b69a5274ab3553315e147f
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u"MERGEFIELD MyMergeField1 \\* MERGEFORMAT");
        builder->InsertField(u"MERGEFIELD MyMergeField2 \\* MERGEFORMAT");

        for (const auto& f : System::IterateOver(doc->get_Range()->get_Fields()))
        {
            if (f->get_Type() == FieldType::FieldMergeField)
            {
                auto mergeField = System::ExplicitCast<FieldMergeField>(f);
                mergeField->set_FieldName(mergeField->get_FieldName() + u"_Renamed");
                mergeField->Update();
            }
        }

        doc->Save(ArtifactsDir + u"WorkingWithFields.RenameMergeFields.docx");
        //ExEnd:RenameMergeFields
    }

    //ExStart:MergeField
    /// <summary>
    /// Represents a facade object for a merge field in a Microsoft Word document.
    /// </summary>
    class MergeField : public System::Object
    {
    public:
        /// <summary>
        /// Gets the name of the merge field.
        /// </summary>
        String get_Name()
        {
            return (System::ExplicitCast<FieldStart>(mFieldStart))->GetField()->get_Result().Replace(u"«", u"").Replace(u"»", u"");
        }

        /// <summary>
        /// Sets the name of the merge field.
        /// </summary>
        void set_Name(String value)
        {
            // Merge field name is stored in the field result which is a Run
            // node between field separator and field end.
            auto fieldResult = System::ExplicitCast<Aspose::Words::Run>(mFieldSeparator->get_NextSibling());
            fieldResult->set_Text(String::Format(u"«{0}»", value));

            // But sometimes the field result can consist of more than one run, delete these runs.
            RemoveSameParent(fieldResult->get_NextSibling(), mFieldEnd);

            UpdateFieldCode(value);
        }

        MergeField(SharedPtr<FieldStart> fieldStart)
            : gRegex(MakeObject<System::Text::RegularExpressions::Regex>(u"\\s*(?<start>MERGEFIELD\\s|)(\\s|)(?<name>\\S+)\\s+"))
        {
            if (fieldStart == nullptr)
            {
                throw System::ArgumentNullException(u"fieldStart");
            }
            if (fieldStart->get_FieldType() != FieldType::FieldMergeField)
            {
                throw System::ArgumentException(u"Field start type must be FieldMergeField.");
            }

            mFieldStart = fieldStart;

            // Find the field separator node.
            mFieldSeparator = fieldStart->GetField()->get_Separator();
            if (mFieldSeparator == nullptr)
            {
                throw System::InvalidOperationException(u"Cannot find field separator.");
            }

            mFieldEnd = fieldStart->GetField()->get_End();
        }

    private:
        SharedPtr<Node> mFieldStart;
        SharedPtr<Node> mFieldSeparator;
        SharedPtr<Node> mFieldEnd;
        SharedPtr<System::Text::RegularExpressions::Regex> gRegex;

        void UpdateFieldCode(String fieldName)
        {
            // Field code is stored in a Run node between field start and field separator.
            auto fieldCode = System::ExplicitCast<Aspose::Words::Run>(mFieldStart->get_NextSibling());

            SharedPtr<System::Text::RegularExpressions::Match> match =
                gRegex->Match((System::ExplicitCast<FieldStart>(mFieldStart))->GetField()->GetFieldCode());

            String newFieldCode = String::Format(u" {0}{1} ", match->get_Groups()->idx_get(u"start")->get_Value(), fieldName);
            fieldCode->set_Text(newFieldCode);

            // But sometimes the field code can consist of more than one run, delete these runs.
            RemoveSameParent(fieldCode->get_NextSibling(), mFieldSeparator);
        }

        /// <summary>
        /// Removes nodes from start up to but not including the end node.
        /// Start and end are assumed to have the same parent.
        /// </summary>
        void RemoveSameParent(SharedPtr<Node> startNode, SharedPtr<Node> endNode)
        {
            if (endNode != nullptr && startNode->get_ParentNode() != endNode->get_ParentNode())
            {
                throw System::ArgumentException(u"Start and end nodes are expected to have the same parent.");
            }

            SharedPtr<Node> curChild = startNode;
            while (curChild != nullptr && curChild != endNode)
            {
                SharedPtr<Node> nextChild = curChild->get_NextSibling();
                curChild->Remove();
                curChild = nextChild;
            }
        }
    };
    //ExEnd:MergeField

    void RemoveField()
    {
        //ExStart:RemoveField
        //GistId:8c604665c1b97795df7a1e665f6b44ce
        auto doc = MakeObject<Document>(MyDir + u"Various fields.docx");

        SharedPtr<Field> field = doc->get_Range()->get_Fields()->idx_get(0);
        field->Remove();
        //ExEnd:RemoveField
    }

    void InsertTOAFieldWithoutDocumentBuilder()
    {
        //ExStart:InsertToaFieldWithoutDocumentBuilder
        //GistId:1cf07762df56f15067d6aef90b14b3db
        auto doc = MakeObject<Document>();
        auto para = MakeObject<Paragraph>(doc);

        // We want to insert TA and TOA fields like this:
        // { TA  \c 1 \l "Value 0" }
        // { TOA  \c 1 }

        auto fieldTA = System::ExplicitCast<FieldTA>(para->AppendField(FieldType::FieldTOAEntry, false));
        fieldTA->set_EntryCategory(u"1");
        fieldTA->set_LongCitation(u"Value 0");

        doc->get_FirstSection()->get_Body()->AppendChild(para);

        para = MakeObject<Paragraph>(doc);

        auto fieldToa = System::ExplicitCast<FieldToa>(para->AppendField(FieldType::FieldTOA, false));
        fieldToa->set_EntryCategory(u"1");
        doc->get_FirstSection()->get_Body()->AppendChild(para);

        fieldToa->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertTOAFieldWithoutDocumentBuilder.docx");
        //ExEnd:InsertToaFieldWithoutDocumentBuilder
    }

    void InsertNestedFields()
    {
        //ExStart:InsertNestedFields
        //GistId:1cf07762df56f15067d6aef90b14b3db
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        for (int i = 0; i < 5; i++)
        {
            builder->InsertBreak(BreakType::PageBreak);
        }

        builder->MoveToHeaderFooter(HeaderFooterType::FooterPrimary);

        // We want to insert a field like this:
        // { IF {PAGE} <> {NUMPAGES} "See Next Page" "Last Page" }
        SharedPtr<Field> field = builder->InsertField(u"IF ");
        builder->MoveTo(field->get_Separator());
        builder->InsertField(u"PAGE");
        builder->Write(u" <> ");
        builder->InsertField(u"NUMPAGES");
        builder->Write(u" \"See Next Page\" \"Last Page\" ");

        field->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertNestedFields.docx");
        //ExEnd:InsertNestedFields
    }

    void InsertMergeFieldUsingDOM()
    {
        //ExStart:InsertMergeFieldUsingDom
        //GistId:1cf07762df56f15067d6aef90b14b3db
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto para = System::ExplicitCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(0));

        builder->MoveTo(para);

        // We want to insert a merge field like this:
        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }

        auto field = System::ExplicitCast<FieldMergeField>(builder->InsertField(FieldType::FieldMergeField, false));

        // { " MERGEFIELD Test1" }
        field->set_FieldName(u"Test1");

        // { " MERGEFIELD Test1 \\b Test2" }
        field->set_TextBefore(u"Test2");

        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 }
        field->set_TextAfter(u"Test3");

        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m" }
        field->set_IsMapped(true);

        // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }
        field->set_IsVerticalFormatting(true);

        // Finally update this merge field
        field->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertMergeFieldUsingDOM.docx");
        //ExEnd:InsertMergeFieldUsingDom
    }

    void InsertMailMergeAddressBlockFieldUsingDOM()
    {
        //ExStart:InsertAddressBlockFieldUsingDom
        //GistId:1cf07762df56f15067d6aef90b14b3db
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto para = System::ExplicitCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(0));

        builder->MoveTo(para);

        // We want to insert a mail merge address block like this:
        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }

        auto field = System::ExplicitCast<FieldAddressBlock>(builder->InsertField(FieldType::FieldAddressBlock, false));

        // { ADDRESSBLOCK \\c 1" }
        field->set_IncludeCountryOrRegionName(u"1");

        // { ADDRESSBLOCK \\c 1 \\d" }
        field->set_FormatAddressOnCountryOrRegion(true);

        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 }
        field->set_ExcludedCountryOrRegionName(u"Test2");

        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 }
        field->set_NameAndAddressFormat(u"Test3");

        // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }
        field->set_LanguageId(u"Test 4");

        field->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertMailMergeAddressBlockFieldUsingDOM.docx");
        //ExEnd:InsertAddressBlockFieldUsingDom
    }

    void InsertFieldIncludeTextWithoutDocumentBuilder()
    {
        //ExStart:InsertFieldIncludeTextWithoutDocumentBuilder
        //GistId:1cf07762df56f15067d6aef90b14b3db
        auto doc = MakeObject<Document>();

        auto para = MakeObject<Paragraph>(doc);

        // We want to insert an INCLUDETEXT field like this:
        // { INCLUDETEXT  "file path" }

        auto fieldIncludeText = System::ExplicitCast<FieldIncludeText>(para->AppendField(FieldType::FieldIncludeText, false));
        fieldIncludeText->set_BookmarkName(u"bookmark");
        fieldIncludeText->set_SourceFullName(MyDir + u"IncludeText.docx");

        doc->get_FirstSection()->get_Body()->AppendChild(para);

        fieldIncludeText->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertIncludeFieldWithoutDocumentBuilder.docx");
        //ExEnd:InsertFieldIncludeTextWithoutDocumentBuilder
    }

    void InsertFieldNone()
    {
        //ExStart:InsertFieldNone
        //GistId:1cf07762df56f15067d6aef90b14b3db
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        auto field = System::ExplicitCast<FieldUnknown>(builder->InsertField(FieldType::FieldNone, false));

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertFieldNone.docx");
        //ExEnd:InsertFieldNone
    }

    void InsertField()
    {
        //ExStart:InsertField
        //GistId:1cf07762df56f15067d6aef90b14b3db
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u"MERGEFIELD MyFieldName \\* MERGEFORMAT");

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertField.docx");
        //ExEnd:InsertField
    }

    void InsertAuthorField()
    {
        //ExStart:InsertAuthorField
        //GistId:1cf07762df56f15067d6aef90b14b3db
        auto doc = MakeObject<Document>();

        auto para = System::ExplicitCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(0));

        // We want to insert an AUTHOR field like this:
        // { AUTHOR Test1 }

        auto field = System::ExplicitCast<FieldAuthor>(para->AppendField(FieldType::FieldAuthor, false));
        field->set_AuthorName(u"Test1");
        // { AUTHOR Test1 }

        field->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertAuthorField.docx");
        //ExEnd:InsertAuthorField
    }

    void InsertASKFieldWithOutDocumentBuilder()
    {
        //ExStart:InsertAskFieldWithoutDocumentBuilder
        //GistId:1cf07762df56f15067d6aef90b14b3db
        auto doc = MakeObject<Document>();

        auto para = System::ExplicitCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(0));

        // We want to insert an Ask field like this:
        // { ASK \"Test 1\" Test2 \\d Test3 \\o }

        auto field = System::ExplicitCast<FieldAsk>(para->AppendField(FieldType::FieldAsk, false));

        // { ASK \"Test 1\" " }
        field->set_BookmarkName(u"Test 1");

        // { ASK \"Test 1\" Test2 }
        field->set_PromptText(u"Test2");

        // { ASK \"Test 1\" Test2 \\d Test3 }
        field->set_DefaultResponse(u"Test3");

        // { ASK \"Test 1\" Test2 \\d Test3 \\o }
        field->set_PromptOnceOnMailMerge(true);

        field->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertASKFieldWithOutDocumentBuilder.docx");
        //ExEnd:InsertAskFieldWithoutDocumentBuilder
    }

    void InsertAdvanceFieldWithOutDocumentBuilder()
    {
        //ExStart:InsertAdvanceFieldWithoutDocumentBuilder
        //GistId:1cf07762df56f15067d6aef90b14b3db
        auto doc = MakeObject<Document>();

        auto para = System::ExplicitCast<Paragraph>(doc->GetChildNodes(NodeType::Paragraph, true)->idx_get(0));

        // We want to insert an Advance field like this:
        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }

        auto field = System::ExplicitCast<FieldAdvance>(para->AppendField(FieldType::FieldAdvance, false));

        // { ADVANCE \\d 10 " }
        field->set_DownOffset(u"10");

        // { ADVANCE \\d 10 \\l 10 }
        field->set_LeftOffset(u"10");

        // { ADVANCE \\d 10 \\l 10 \\r -3.3 }
        field->set_RightOffset(u"-3.3");

        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 }
        field->set_UpOffset(u"0");

        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 }
        field->set_HorizontalPosition(u"100");

        // { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }
        field->set_VerticalPosition(u"100");

        field->Update();

        doc->Save(ArtifactsDir + u"WorkingWithFields.InsertAdvanceFieldWithOutDocumentBuilder.docx");
        //ExEnd:InsertAdvanceFieldWithoutDocumentBuilder
    }

    void GetMailMergeFieldNames()
    {
        //ExStart:GetFieldNames
        //GistId:b4bab1bf22437a86d8062e91cf154494
        auto doc = MakeObject<Document>();

        ArrayPtr<String> fieldNames = doc->get_MailMerge()->GetFieldNames();
        //ExEnd:GetFieldNames
        std::cout << (String(u"\nDocument have ") + fieldNames->get_Length() + u" fields.") << std::endl;
    }

    void MappedDataFields()
    {
        //ExStart:MappedDataFields
        //GistId:b4bab1bf22437a86d8062e91cf154494
        auto doc = MakeObject<Document>();

        doc->get_MailMerge()->get_MappedDataFields()->Add(u"MyFieldName_InDocument", u"MyFieldName_InDataSource");
        //ExEnd:MappedDataFields
    }

    void DeleteFields()
    {
        //ExStart:DeleteFields
        //GistId:f39874821cb317d245a769c9ce346fea
        auto doc = MakeObject<Document>();

        doc->get_MailMerge()->DeleteFields();
        //ExEnd:DeleteFields
    }

    void FieldUpdateCulture()
    {
        //ExStart:FieldUpdateCulture
        //GistId:79b46682fbfd7f02f64783b163ed95fc
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(FieldType::FieldTime, true);

        doc->get_FieldOptions()->set_FieldUpdateCultureSource(FieldUpdateCultureSource::FieldCode);
        doc->get_FieldOptions()->set_FieldUpdateCultureProvider(MakeObject<WorkingWithFields::FieldUpdateCultureProvider>());

        doc->Save(ArtifactsDir + u"WorkingWithFields.FieldUpdateCulture.pdf");
        //ExEnd:FieldUpdateCulture
    }

    //ExStart:FieldUpdateCultureProvider
    //GistId:79b46682fbfd7f02f64783b163ed95fc
    class FieldUpdateCultureProvider : public IFieldUpdateCultureProvider
    {
    public:
        SharedPtr<System::Globalization::CultureInfo> GetCulture(String name, SharedPtr<Field> field) override
        {
            ASPOSE_UNUSED(field);
            if (name == u"ru-RU")
            {
                auto culture = MakeObject<System::Globalization::CultureInfo>(name, false);
                SharedPtr<System::Globalization::DateTimeFormatInfo> format = culture->get_DateTimeFormat();
                format->set_MonthNames(MakeArray<String>({u"месяц 1", u"месяц 2", u"месяц 3", u"месяц 4", u"месяц 5", u"месяц 6", u"месяц 7", u"месяц 8",
                                                          u"месяц 9", u"месяц 10", u"месяц 11", u"месяц 12", u""}));
                format->set_MonthGenitiveNames(format->get_MonthNames());
                format->set_AbbreviatedMonthNames(MakeArray<String>(
                    {u"мес 1", u"мес 2", u"мес 3", u"мес 4", u"мес 5", u"мес 6", u"мес 7", u"мес 8", u"мес 9", u"мес 10", u"мес 11", u"мес 12", u""}));
                format->set_AbbreviatedMonthGenitiveNames(format->get_AbbreviatedMonthNames());
                format->set_DayNames(MakeArray<String>(
                    {u"день недели 7", u"день недели 1", u"день недели 2", u"день недели 3", u"день недели 4", u"день недели 5", u"день недели 6"}));
                format->set_AbbreviatedDayNames(MakeArray<String>({u"день 7", u"день 1", u"день 2", u"день 3", u"день 4", u"день 5", u"день 6"}));
                format->set_ShortestDayNames(MakeArray<String>({u"д7", u"д1", u"д2", u"д3", u"д4", u"д5", u"д6"}));
                format->set_AMDesignator(u"До полудня");
                format->set_PMDesignator(u"После полудня");
                const String pattern = u"yyyy MM (MMMM) dd (dddd) hh:mm:ss tt";
                format->set_LongDatePattern(pattern);
                format->set_LongTimePattern(pattern);
                format->set_ShortDatePattern(pattern);
                format->set_ShortTimePattern(pattern);
                return culture;
            }
            else if (name == u"en-US")
            {
                return MakeObject<System::Globalization::CultureInfo>(name, false);
            }
            else
            {
                return nullptr;
            }
        }
    };
    //ExEnd:FieldUpdateCultureProvider

    void FieldDisplayResults()
    {
        //ExStart:FieldDisplayResults
        //GistId:bf0f8a6b40b69a5274ab3553315e147f
        //ExStart:UpdateDocFields
        //GistId:08db64c4d86842c4afd1ecb925ed07c4
        auto document = MakeObject<Document>(MyDir + u"Various fields.docx");

        document->UpdateFields();
        //ExEnd:UpdateDocFields

        for (const auto& field : System::IterateOver(document->get_Range()->get_Fields()))
        {
            std::cout << field->get_DisplayResult() << std::endl;
        }
        //ExEnd:FieldDisplayResults
    }

    void EvaluateIFCondition()
    {
        //ExStart:EvaluateIfCondition
        //GistId:79b46682fbfd7f02f64783b163ed95fc
        auto builder = MakeObject<DocumentBuilder>();

        auto field = System::ExplicitCast<FieldIf>(builder->InsertField(u"IF 1 = 1", nullptr));
        FieldIfComparisonResult actualResult = field->EvaluateCondition();

        std::cout << System::EnumGetName(actualResult) << std::endl;
        //ExEnd:EvaluateIfCondition
    }

    void ConvertFieldsInParagraph()
    {
        //ExStart:UnlinkFieldsInParagraph
        //GistId:f3592014d179ecb43905e37b2a68bc92
        auto doc = MakeObject<Document>(MyDir + u"Linked fields.docx");

        // Pass the appropriate parameters to convert all IF fields to text that are encountered only in the last
        // paragraph of the document.
        doc->get_FirstSection()
            ->get_Body()
            ->get_LastParagraph()
            ->get_Range()
            ->get_Fields()
            ->LINQ_Where([](SharedPtr<Field> f) { return f->get_Type() == FieldType::FieldIf; })
            ->LINQ_ToList()
            ->ForEach(std::function<void(SharedPtr<Field>)>([](SharedPtr<Field> f) { f->Unlink(); }));

        doc->Save(ArtifactsDir + u"WorkingWithFields.TestFile.docx");
        //ExEnd:UnlinkFieldsInParagraph
    }

    void ConvertFieldsInDocument()
    {
        //ExStart:UnlinkFieldsInDocument
        //GistId:f3592014d179ecb43905e37b2a68bc92
        auto doc = MakeObject<Document>(MyDir + u"Linked fields.docx");

        // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to text.
        doc->get_Range()
            ->get_Fields()
            ->LINQ_Where([](SharedPtr<Field> f) { return f->get_Type() == FieldType::FieldIf; })
            ->LINQ_ToList()
            ->ForEach(std::function<void(SharedPtr<Field> f)>([](SharedPtr<Field> f) { f->Unlink(); }));

        // Save the document with fields transformed to disk
        doc->Save(ArtifactsDir + u"WorkingWithFields.ConvertFieldsInDocument.docx");
        //ExEnd:UnlinkFieldsInDocument
    }

    void ConvertFieldsInBody()
    {
        //ExStart:UnlinkFieldsInBody
        //GistId:f3592014d179ecb43905e37b2a68bc92
        auto doc = MakeObject<Document>(MyDir + u"Linked fields.docx");

        // Pass the appropriate parameters to convert PAGE fields encountered to text only in the body of the first section.
        doc->get_FirstSection()
            ->get_Body()
            ->get_Range()
            ->get_Fields()
            ->LINQ_Where([](SharedPtr<Field> f) { return f->get_Type() == FieldType::FieldPage; })
            ->LINQ_ToList()
            ->ForEach(std::function<void(SharedPtr<Field> f)>([](SharedPtr<Field> f) { f->Unlink(); }));

        doc->Save(ArtifactsDir + u"WorkingWithFields.ConvertFieldsInBody.docx");
        //ExEnd:UnlinkFieldsInBody
    }

    void FieldCode()
    {
        //ExStart:FieldCode
        //GistId:7c2b7b650a88375b1d438746f78f0d64
        auto doc = MakeObject<Document>(MyDir + u"Hyperlinks.docx");

        for (const auto& field : System::IterateOver(doc->get_Range()->get_Fields()))
        {
            String fieldCode = field->GetFieldCode();
            String fieldResult = field->get_Result();
        }
        //ExEnd:FieldCode
    }

    void UnlinkFields()
    {
        //ExStart:UnlinkFields
        //GistId:f3592014d179ecb43905e37b2a68bc92
        auto doc = MakeObject<Document>(MyDir + u"Various fields.docx");
        doc->UnlinkFields();
        //ExEnd:UnlinkFields
    }

    //ExStart:ConvertFieldsToStaticText
    //GistId:f3592014d179ecb43905e37b2a68bc92
    /// <summary>
    /// Converts any fields of the specified type found in the descendants of the node into static text.
    /// </summary>
    /// <param name="compositeNode">The node in which all descendants of the specified FieldType will be converted to static text.</param>
    /// <param name="targetFieldType">The FieldType of the field to convert to static text.</param>
    void ConvertFieldsToStaticText(SharedPtr<CompositeNode> compositeNode, FieldType targetFieldType)
    {
        compositeNode->get_Range()
            ->get_Fields()
            ->LINQ_Where([&targetFieldType](SharedPtr<Field> f) { return f->get_Type() == targetFieldType; })
            ->LINQ_ToList()
            ->ForEach(std::function<void(SharedPtr<Field>)>([](SharedPtr<Field> f) { f->Unlink(); }));
    }
    //ExEnd:ConvertFieldsToStaticText

    void InsertFieldUsingFieldBuilder()
    {
        //ExStart:InsertFieldUsingFieldBuilder
        //GistId:1cf07762df56f15067d6aef90b14b3db
        auto doc = MakeObject<Document>();

        // Prepare IF field with two nested MERGEFIELD fields: { IF "left expression" = "right expression" "Firstname: { MERGEFIELD firstname }" "Lastname: { MERGEFIELD lastname }"}
        auto fieldBuilder =
            MakeObject<FieldBuilder>(FieldType::FieldIf)
                ->AddArgument(u"left expression")
                ->AddArgument(u"=")
                ->AddArgument(u"right expression")
                ->AddArgument(MakeObject<FieldArgumentBuilder>()
                                  ->AddText(u"Firstname: ")
                                  ->AddField(MakeObject<FieldBuilder>(FieldType::FieldMergeField)->AddArgument(u"firstname")))
                ->AddArgument(MakeObject<FieldArgumentBuilder>()
                                  ->AddText(u"Lastname: ")
                                  ->AddField(MakeObject<FieldBuilder>(FieldType::FieldMergeField)->AddArgument(u"lastname")));

        // Insert IF field in exact location
        SharedPtr<Field> field = fieldBuilder->BuildAndInsert(doc->get_FirstSection()->get_Body()->get_FirstParagraph());
        field->Update();

        doc->Save(ArtifactsDir + u"Field.InsertFieldUsingFieldBuilder.docx");
        //ExEnd:InsertFieldUsingFieldBuilder
    }

    void FieldResultFormatting()
    {
        //ExStart:FieldResultFormatting
        //GistId:79b46682fbfd7f02f64783b163ed95fc
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);
        auto formatter = MakeObject<FieldResultFormatter>(u"${0}", u"Date: {0}", u"Item # {0}:");
        doc->get_FieldOptions()->set_ResultFormatter(formatter);

        // Our field result formatter applies a custom format to newly created fields of three types of formats.
        // Field result formatters apply new formatting to fields as they are updated,
        // which happens as soon as we create them using this InsertField method overload.
        // 1 -  Numeric:
        builder->InsertField(u" = 2 + 3 \\# $###");

        ASSERT_EQ(u"$5", doc->get_Range()->get_Fields()->idx_get(0)->get_Result());
        ASSERT_EQ(1, formatter->CountFormatInvocations(FieldResultFormatter::FormatInvocationType::Numeric));

        // 2 -  Date/time:
        builder->InsertField(u"DATE \\@ \"d MMMM yyyy\"");

        ASSERT_TRUE(doc->get_Range()->get_Fields()->idx_get(1)->get_Result().StartsWith(u"Date: "));
        ASSERT_EQ(1, formatter->CountFormatInvocations(FieldResultFormatter::FormatInvocationType::DateTime));

        // 3 -  General:
        builder->InsertField(u"QUOTE \"2\" \\* Ordinal");

        ASSERT_EQ(u"Item # 2:", doc->get_Range()->get_Fields()->idx_get(2)->get_Result());
        ASSERT_EQ(1, formatter->CountFormatInvocations(FieldResultFormatter::FormatInvocationType::General));

        formatter->PrintFormatInvocations();
        //ExEnd:FieldResultFormatting
    }

    //ExStart:FieldResultFormatter
    //GistId:79b46682fbfd7f02f64783b163ed95fc
    /// <summary>
    /// When fields with formatting are updated, this formatter will override their formatting
    /// with a custom format, while tracking every invocation.
    /// </summary>
    class FieldResultFormatter : public IFieldResultFormatter
    {
    public:
        enum class FormatInvocationType
        {
            Numeric,
            DateTime,
            General,
            All
        };

        FieldResultFormatter(const String& numberFormat, const String& dateFormat, const String& generalFormat)
            : mNumberFormat(numberFormat), mDateFormat(dateFormat), mGeneralFormat(generalFormat)
        {
            mFormatInvocations = MakeObject<System::Collections::Generic::List<SharedPtr<FormatInvocation>>>();
        }

        String FormatNumeric(double value, String format) override
        {
            if (String::IsNullOrEmpty(mNumberFormat))
            {
                return nullptr;
            }

            String newValue = String::Format(mNumberFormat, value);
            mFormatInvocations->Add(MakeObject<FormatInvocation>(FormatInvocationType::Numeric, System::ObjectExt::Box<double>(value), format, newValue));
            return newValue;
        }

        String FormatDateTime(System::DateTime value, String format, Aspose::Words::CalendarType calendarType) override
        {
            if (String::IsNullOrEmpty(mDateFormat))
            {
                return nullptr;
            }

            String newValue = String::Format(mDateFormat, value);
            mFormatInvocations->Add(MakeObject<FormatInvocation>(
                FormatInvocationType::DateTime,
                System::ObjectExt::Box<String>(String::Format(u"{0} ({1})", value, System::EnumGetName(calendarType))),
                format,
                newValue));
            return newValue;
        }

        String Format(String value, GeneralFormat format) override
        {
            return FormatCore(System::ObjectExt::Box<String>(value), format);
        }

        String Format(double value, GeneralFormat format) override
        {
            return FormatCore(System::ObjectExt::Box<double>(value), format);
        }

        int CountFormatInvocations(FormatInvocationType formatInvocationType)
        {
            if (formatInvocationType == FormatInvocationType::All)
            {
                return mFormatInvocations->get_Count();
            }

            int count = 0;
            for (const auto& f : System::IterateOver(mFormatInvocations))
            {
                if (f->FormatInvocationType_ == formatInvocationType)
                {
                    count++;
                }
            }
            return count;
        }

        static String GetInvocationTypeName(FormatInvocationType type)
        {
            switch (type)
            {
            case FormatInvocationType::Numeric:
                return u"Numeric";
            case FormatInvocationType::DateTime:
                return u"DateTime";
            case FormatInvocationType::General:
                return u"General";
            default:
                return u"All";
            }
        }

        void PrintFormatInvocations()
        {
            for (const auto& f : System::IterateOver(mFormatInvocations))
            {
                std::cout << String::Format(u"Invocation type:\t{0}\n\tOriginal value:\t\t{1}\n\tOriginal format:\t{2}\n\tNew value:\t\t\t{3}\n",
                                            GetInvocationTypeName(f->FormatInvocationType_),
                                            f->Value,
                                            f->OriginalFormat,
                                            f->NewValue)
                          << std::endl;
            }
        }

    private:
        class FormatInvocation : public System::Object
        {
        public:
            FieldResultFormatter::FormatInvocationType FormatInvocationType_;
            SharedPtr<System::Object> Value;
            String OriginalFormat;
            String NewValue;

            FormatInvocation(FieldResultFormatter::FormatInvocationType formatInvocationType,
                             const SharedPtr<System::Object>& value,
                             const String& originalFormat,
                             const String& newValue)
                : FormatInvocationType_(formatInvocationType), Value(value), OriginalFormat(originalFormat), NewValue(newValue)
            {
            }
        };

        String FormatCore(const SharedPtr<System::Object>& value, GeneralFormat format)
        {
            if (String::IsNullOrEmpty(mGeneralFormat))
            {
                return nullptr;
            }

            String newValue = String::Format(mGeneralFormat, value);
            mFormatInvocations->Add(MakeObject<FormatInvocation>(FormatInvocationType::General, value, System::EnumGetName(format), newValue));
            return newValue;
        }

        String mNumberFormat;
        String mDateFormat;
        String mGeneralFormat;
        SharedPtr<System::Collections::Generic::List<SharedPtr<FormatInvocation>>> mFormatInvocations;
    };
    //ExEnd:FieldResultFormatter

    void ChangeLocale()
    {
        //ExStart:ChangeLocale
        //GistId:9e90defe4a7bcafb004f73a2ef236986
        auto doc = MakeObject<Document>();
        auto builder = MakeObject<DocumentBuilder>(doc);

        builder->InsertField(u"MERGEFIELD Date");

        // Store the current culture so it can be set back once mail merge is complete.
        SharedPtr<System::Globalization::CultureInfo> currentCulture = System::Threading::Thread::get_CurrentThread()->get_CurrentCulture();
        // Set to German language so dates and numbers are formatted using this culture during mail merge.
        System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(MakeObject<System::Globalization::CultureInfo>(u"de-DE"));

        doc->get_MailMerge()->Execute(MakeArray<String>({u"Date"}),
                                      MakeArray<SharedPtr<System::Object>>({System::ObjectExt::Box<System::DateTime>(System::DateTime::get_Now())}));

        System::Threading::Thread::get_CurrentThread()->set_CurrentCulture(currentCulture);

        doc->Save(ArtifactsDir + u"WorkingWithFields.ChangeLocale.docx");
        //ExEnd:ChangeLocale
    }
};

}} // namespace DocsExamples::Programming_with_Documents
